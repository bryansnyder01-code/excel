"""Billing with Stripe: a Pro plan with a higher AI limit. No SDK: form-encoded calls to the
Stripe API with httpx, and webhook signatures checked by hand.

Billing is on when stripe_secret_key and stripe_price_pro are set. Without them,
GET /plan says billing_enabled: false and checkout/portal answer 503.

Webhook events handled: checkout.session.completed (the user becomes Pro),
customer.subscription.updated (Pro while active/trialing/past_due, with the renewal date),
customer.subscription.deleted (back to free).
"""

from __future__ import annotations

import hashlib
import hmac
import json
import logging
import time
from collections.abc import AsyncIterator

import httpx
from fastapi import APIRouter, Depends, Header, HTTPException, Request
from fastapi.concurrency import run_in_threadpool

from app.api.accounts import current_user, get_store
from app.config import get_settings
from app.services.store import Store
from app.services.usage import UsageStore, get_usage_store, limit_for_user

log = logging.getLogger(__name__)

billing_router = APIRouter(prefix="/api/billing", tags=["billing"])

STRIPE_API = "https://api.stripe.com/v1"
SIGNATURE_TOLERANCE = 5 * 60  # seconds
ACTIVE = {"active", "trialing", "past_due"}

BILLING_OFF = "Billing isn't set up on this server."
STRIPE_FAILED = "Stripe couldn't be reached. Please try again in a minute."
NO_SUBSCRIPTION = "You don't have a subscription to manage yet."
BAD_SIGNATURE = "Invalid Stripe signature."


async def stripe_http() -> AsyncIterator[httpx.AsyncClient]:
    """HTTP client for the Stripe API. Tests override this with a mock transport."""
    async with httpx.AsyncClient(timeout=20) as client:
        yield client


def billing_enabled() -> bool:
    s = get_settings()
    return bool(s.stripe_secret_key and s.stripe_price_pro)


async def _stripe(http: httpx.AsyncClient, method: str, path: str, data: dict | None = None) -> dict:
    """One Stripe API call; raises 502 with a plain sentence when it fails."""
    try:
        r = await http.request(method, f"{STRIPE_API}{path}", data=data,
                               headers={"Authorization": f"Bearer {get_settings().stripe_secret_key}"})
        body = r.json()
    except (httpx.HTTPError, ValueError):
        log.exception("Stripe %s %s failed", method, path)
        raise HTTPException(502, STRIPE_FAILED) from None
    if r.status_code >= 400:
        log.error("Stripe %s %s: %s %s", method, path, r.status_code, body.get("error") if isinstance(body, dict) else body)
        raise HTTPException(502, STRIPE_FAILED)
    return body


@billing_router.get("/plan")
def plan(user: dict = Depends(current_user), store: Store = Depends(get_store),
         usage: UsageStore = Depends(get_usage_store)):
    info = store.billing(user["id"])
    renews = info["renews_at"]
    return {"plan": info["plan"], "limit": limit_for_user(user), "used": usage.used(f"user:{user['id']}"),
            "renews_at": int(renews * 1000) if renews else None, "billing_enabled": billing_enabled()}


async def _customer(store: Store, user: dict, http: httpx.AsyncClient) -> str:
    """The user's Stripe customer, created on first checkout and reused after."""
    info = await run_in_threadpool(store.billing, user["id"])
    if info["stripe_customer_id"]:
        return info["stripe_customer_id"]
    customer = await _stripe(http, "POST", "/customers",
                             {"email": user["email"], "metadata[user_id]": user["id"]})
    await run_in_threadpool(store.set_customer, user["id"], customer["id"])
    return customer["id"]


@billing_router.post("/checkout")
async def checkout(user: dict = Depends(current_user), store: Store = Depends(get_store),
                   http: httpx.AsyncClient = Depends(stripe_http)):
    """A Stripe Checkout page for the Pro subscription: {url}."""
    if not billing_enabled():
        raise HTTPException(503, BILLING_OFF)
    settings = get_settings()
    app_url = settings.app_url.rstrip("/")
    customer = await _customer(store, user, http)
    session = await _stripe(http, "POST", "/checkout/sessions", {
        "mode": "subscription",
        "customer": customer,
        "client_reference_id": user["id"],
        "line_items[0][price]": settings.stripe_price_pro,
        "line_items[0][quantity]": "1",
        "success_url": f"{app_url}/?billing=success",
        "cancel_url": f"{app_url}/?billing=cancel",
        "subscription_data[metadata][user_id]": user["id"],
    })
    return {"url": session["url"]}


@billing_router.post("/portal")
async def portal(user: dict = Depends(current_user), store: Store = Depends(get_store),
                 http: httpx.AsyncClient = Depends(stripe_http)):
    """The Stripe customer portal (change card, cancel): {url}."""
    if not billing_enabled():
        raise HTTPException(503, BILLING_OFF)
    customer = (await run_in_threadpool(store.billing, user["id"]))["stripe_customer_id"]
    if not customer:
        raise HTTPException(400, NO_SUBSCRIPTION)
    session = await _stripe(http, "POST", "/billing_portal/sessions",
                            {"customer": customer, "return_url": f"{get_settings().app_url.rstrip('/')}/"})
    return {"url": session["url"]}


# ---------------------------------------------------------------------------
# Webhook
# ---------------------------------------------------------------------------


def verify_signature(payload: bytes, header: str | None, secret: str, now: float | None = None) -> bool:
    """Stripe-Signature: "t=<unix>,v1=<hex>[,v1=...]"; v1 = HMAC-SHA256(secret, "<t>.<payload>")."""
    if not header or not secret:
        return False
    parts = [p.split("=", 1) for p in header.split(",") if "=" in p]
    stamps = [v for k, v in parts if k.strip() == "t"]
    sigs = [v.strip() for k, v in parts if k.strip() == "v1"]
    try:
        t = int(stamps[0])
    except (IndexError, ValueError):
        return False
    if abs((now if now is not None else time.time()) - t) > SIGNATURE_TOLERANCE:
        return False
    expected = hmac.new(secret.encode(), f"{t}.".encode() + payload, hashlib.sha256).hexdigest()
    return any(hmac.compare_digest(expected, s) for s in sigs)


def _period_end(sub: dict) -> float | None:
    """current_period_end, which newer API versions moved onto the subscription items."""
    end = sub.get("current_period_end")
    if end is None:
        items = (sub.get("items") or {}).get("data") or []
        end = items[0].get("current_period_end") if items else None
    return float(end) if end else None


def _user_for(store: Store, obj: dict) -> str | None:
    meta = obj.get("metadata") or {}
    user_id = obj.get("client_reference_id") or meta.get("user_id")
    if user_id:
        return user_id
    return store.user_by_customer(obj["customer"]) if obj.get("customer") else None


@billing_router.post("/webhook")
async def webhook(request: Request, stripe_signature: str | None = Header(default=None),
                  store: Store = Depends(get_store), http: httpx.AsyncClient = Depends(stripe_http)):
    payload = await request.body()
    if not verify_signature(payload, stripe_signature, get_settings().stripe_webhook_secret):
        raise HTTPException(400, BAD_SIGNATURE)
    try:
        event = json.loads(payload)
        kind, obj = event["type"], event["data"]["object"]
    except (ValueError, KeyError, TypeError):
        raise HTTPException(400, "That event couldn't be read.") from None

    if kind == "checkout.session.completed":
        user_id = await run_in_threadpool(_user_for, store, obj)
        if user_id:
            if obj.get("customer"):
                await run_in_threadpool(store.set_customer, user_id, obj["customer"])
            renews = None
            if obj.get("subscription") and billing_enabled():
                try:
                    renews = _period_end(await _stripe(http, "GET", f"/subscriptions/{obj['subscription']}"))
                except HTTPException:
                    pass  # the subscription.updated event that follows fills it in
            await run_in_threadpool(store.set_plan, user_id, "pro", renews)
    elif kind in ("customer.subscription.updated", "customer.subscription.created",
                  "customer.subscription.deleted"):
        user_id = await run_in_threadpool(_user_for, store, obj)
        if user_id:
            if kind != "customer.subscription.deleted" and obj.get("status") in ACTIVE:
                await run_in_threadpool(store.set_plan, user_id, "pro", _period_end(obj))
            else:
                await run_in_threadpool(store.set_plan, user_id, "free", None)
    return {"received": True}
