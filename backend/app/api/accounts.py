"""Accounts (password, reset links, Google sign-in), cloud sync of saved workbooks, edit
invites and collaborators, and public share links.

Errors are always {"detail": "<user-facing sentence>"}. Signed-in routes take
`Authorization: Bearer <token>`. Docs are opaque JSON to the backend (the frontend's
SavedDoc); only a shared doc's `workbook` is ever served publicly. A doc has one owner and
any number of editors and viewers (people who opened its editor or viewer invite link).
Editors read and write it, except locked sheets/ranges; viewers only read (and comment).
"""

from __future__ import annotations

import json
import logging
import re
import time
from collections.abc import AsyncIterator
from functools import lru_cache
from typing import Literal

import httpx
from fastapi import APIRouter, BackgroundTasks, Depends, Header, HTTPException, Request, Response
from fastapi.concurrency import run_in_threadpool
from pydantic import BaseModel, ValidationError

from app.config import get_settings
from app.services import protection
from app.services.mailer import send_mail, smtp_configured
from app.services.store import EmailTaken, Store

log = logging.getLogger(__name__)

auth_router = APIRouter(prefix="/api/auth", tags=["accounts"])
docs_router = APIRouter(prefix="/api/docs", tags=["accounts"])
shared_router = APIRouter(prefix="/api/shared", tags=["accounts"])
invites_router = APIRouter(prefix="/api/invites", tags=["accounts"])

MAX_DOC_BYTES = 5 * 1024 * 1024
MIN_PASSWORD = 8
_EMAIL = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

SIGN_IN_AGAIN = "Please sign in again."
DOC_NOT_FOUND = "Workbook not found."
SHARE_NOT_FOUND = "This share link doesn't exist or was turned off."
INVITE_NOT_FOUND = "This invite link doesn't exist or was turned off."
VIEW_ONLY = "You can view this workbook but not change it."
RESET_EXPIRED = "That reset link has expired. Ask for a new one."
SHORT_PASSWORD = "Use at least 8 characters for your password."
GOOGLE_OFF = "Google sign-in isn't set up on this server."
GOOGLE_FAILED = "Google sign-in failed. Please try again."

GOOGLE_TOKENINFO = "https://oauth2.googleapis.com/tokeninfo"
GOOGLE_ISSUERS = {"accounts.google.com", "https://accounts.google.com"}


@lru_cache
def _store_for(data_dir: str, session_days: float = 30) -> Store:
    return Store(data_dir, session_days)


def get_store() -> Store:
    """The database for the configured data_dir, opened on first use. Tests override this."""
    settings = get_settings()
    return _store_for(settings.data_dir, settings.session_days)


async def google_http() -> AsyncIterator[httpx.AsyncClient]:
    """HTTP client for checking Google ID tokens. Tests override this with a mock transport."""
    async with httpx.AsyncClient(timeout=10) as client:
        yield client


def _bearer(authorization: str | None) -> str | None:
    if authorization and authorization[:7].lower() == "bearer ":
        return authorization[7:].strip() or None
    return None


def current_user(authorization: str | None = Header(default=None), store: Store = Depends(get_store)) -> dict:
    user = store.user_for_token(_bearer(authorization))
    if not user:
        raise HTTPException(401, SIGN_IN_AGAIN, headers={"WWW-Authenticate": "Bearer"})
    return user


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------


class Credentials(BaseModel):
    email: str = ""
    password: str = ""


@auth_router.post("/signup")
def signup(body: Credentials, store: Store = Depends(get_store)):
    email = body.email.strip().lower()
    if not _EMAIL.match(email):
        raise HTTPException(422, "Enter a valid email address.")
    if len(body.password) < MIN_PASSWORD:
        raise HTTPException(422, SHORT_PASSWORD)
    try:
        user = store.create_user(email, body.password)
    except EmailTaken:
        raise HTTPException(409, "An account with that email already exists.") from None
    return {"token": store.create_session(user["id"]), "user": user}


@auth_router.post("/login")
def login(body: Credentials, store: Store = Depends(get_store)):
    email = body.email.strip().lower()
    if store.login_locked(email):
        raise HTTPException(429, "Too many attempts. Try again in a few minutes.")
    user = store.check_login(email, body.password)
    store.record_login(email, ok=user is not None)
    if not user:
        raise HTTPException(401, "Wrong email or password.")
    return {"token": store.create_session(user["id"]), "user": user}


@auth_router.post("/logout", status_code=204)
def logout(authorization: str | None = Header(default=None), user: dict = Depends(current_user),
           store: Store = Depends(get_store)):
    store.delete_session(_bearer(authorization))
    return Response(status_code=204)


@auth_router.get("/me")
def me(user: dict = Depends(current_user)):
    return user


@auth_router.get("/config")
def auth_config():
    """What the sign-in screen can offer on this server."""
    return {"google_client_id": get_settings().google_client_id or None, "password_reset": True}


class ForgotIn(BaseModel):
    email: str = ""


def _send_reset(email: str, url: str) -> None:
    settings = get_settings()
    if not smtp_configured(settings):
        log.warning("Password reset link for %s: %s", email, url)
        return
    send_mail(settings, email, "Reset your SheetCraft password",
              f"Someone asked to reset the password for your SheetCraft account.\n\n"
              f"Open this link within an hour to choose a new one:\n{url}\n\n"
              f"If it wasn't you, ignore this email; your password stays the same.\n")


@auth_router.post("/forgot", status_code=204)
def forgot(body: ForgotIn, background: BackgroundTasks, store: Store = Depends(get_store)):
    """Email a reset link. Always 204, so the response doesn't reveal whether the account exists."""
    email = body.email.strip().lower()
    if _EMAIL.match(email) and store.allow_reset_request(email):
        user = store.user_by_email(email)
        if user:
            url = f"{get_settings().app_url.rstrip('/')}/?reset={store.create_reset(user['id'])}"
            background.add_task(_send_reset, email, url)
    return Response(status_code=204)


class ResetIn(BaseModel):
    token: str = ""
    password: str = ""


@auth_router.post("/reset")
def reset(body: ResetIn, store: Store = Depends(get_store)):
    """Set a new password from an emailed link. Ends every other session and signs this one in."""
    if len(body.password) < MIN_PASSWORD:
        raise HTTPException(422, SHORT_PASSWORD)
    user = store.reset_password(body.token, body.password) if body.token else None
    if not user:
        raise HTTPException(400, RESET_EXPIRED)
    return {"token": store.create_session(user["id"]), "user": user}


class GoogleIn(BaseModel):
    credential: str = ""


async def _google_email(credential: str, client_id: str, http: httpx.AsyncClient) -> str | None:
    """The verified email in a Google ID token, or None."""
    try:
        r = await http.get(GOOGLE_TOKENINFO, params={"id_token": credential})
        claims = r.json() if r.status_code == 200 else None
    except (httpx.HTTPError, ValueError):
        return None
    if not isinstance(claims, dict):
        return None
    try:
        expires = float(claims.get("exp") or 0)
    except (TypeError, ValueError):
        return None
    email = str(claims.get("email") or "").strip().lower()
    if (claims.get("aud") != client_id or claims.get("iss") not in GOOGLE_ISSUERS
            or str(claims.get("email_verified")).lower() != "true" or expires <= time.time()
            or not _EMAIL.match(email)):
        return None
    return email


@auth_router.post("/google")
async def google_sign_in(body: GoogleIn, store: Store = Depends(get_store),
                         http: httpx.AsyncClient = Depends(google_http)):
    """Sign in (or sign up) with a Google ID token from Google Identity Services."""
    client_id = get_settings().google_client_id
    if not client_id:
        raise HTTPException(503, GOOGLE_OFF)
    email = await _google_email(body.credential, client_id, http) if body.credential else None
    if not email:
        raise HTTPException(401, GOOGLE_FAILED)
    user = await run_in_threadpool(store.user_by_email, email, True)
    return {"token": store.create_session(user["id"]), "user": user}


# ---------------------------------------------------------------------------
# Cloud sync
# ---------------------------------------------------------------------------


class DocIn(BaseModel):
    title: str
    updated: float  # ms since epoch, set by the client
    data: dict


def _own_doc(store: Store, user: dict, doc_id: str) -> None:
    if store.doc_owner(doc_id) != user["id"]:
        raise HTTPException(404, DOC_NOT_FOUND)


def _can_open(store: Store, user: dict, doc_id: str) -> str:
    role = store.doc_role(user["id"], doc_id)
    if not role:
        raise HTTPException(404, DOC_NOT_FOUND)
    return role


@docs_router.get("")
def list_docs(user: dict = Depends(current_user), store: Store = Depends(get_store)):
    return store.list_docs(user["id"])


@docs_router.get("/{doc_id}")
def get_doc(doc_id: str, user: dict = Depends(current_user), store: Store = Depends(get_store)):
    """The owner's or an editor's copy (they're the same doc)."""
    doc = store.get_doc(user["id"], doc_id)
    if not doc:
        raise HTTPException(404, DOC_NOT_FOUND)
    return doc


@docs_router.put("/{doc_id}")
async def put_doc(doc_id: str, request: Request, user: dict = Depends(current_user),
                  store: Store = Depends(get_store)):
    """Create or replace a doc (last write wins). The body is read by hand to enforce the size cap."""
    too_large = HTTPException(413, "This workbook is too large to sync (5 MB max).")
    if int(request.headers.get("content-length") or 0) > MAX_DOC_BYTES:
        raise too_large
    raw = b""
    async for chunk in request.stream():
        raw += chunk
        if len(raw) > MAX_DOC_BYTES:
            raise too_large
    try:
        body = DocIn.model_validate(json.loads(raw))
    except (ValueError, ValidationError):
        raise HTTPException(422, "That workbook data couldn't be read.") from None
    if not doc_id or len(doc_id) > 200:
        raise HTTPException(404, DOC_NOT_FOUND)
    role = await run_in_threadpool(store.doc_role, user["id"], doc_id)
    if role == "viewer":
        raise HTTPException(403, VIEW_ONLY)
    if role == "editor":
        problem = await run_in_threadpool(_locked_change, store, doc_id, body.data)
        if problem:
            raise HTTPException(403, problem)
    saved = await run_in_threadpool(store.put_doc, user["id"], doc_id, body.title, body.updated, body.data)
    if not saved:
        raise HTTPException(404, DOC_NOT_FOUND)
    return {"id": doc_id, "updated": body.updated}


def _locked_change(store: Store, doc_id: str, data: dict) -> str | None:
    """Why an editor's PUT would change locked cells (None when it doesn't). Compared with the open
    live room's copy when there is one on this instance, else with the stored workbook."""
    from app.api.live import local_workbook  # live imports this module

    current = local_workbook(doc_id)
    if current is None:
        doc = store.live_doc(doc_id)
        current = doc["data"].get("workbook") if doc and isinstance(doc["data"], dict) else None
    if not isinstance(current, dict) or not protection.has_locks(current):
        return None
    return protection.check_replace(current, data.get("workbook"))


@docs_router.delete("/{doc_id}", status_code=204)
def delete_doc(doc_id: str, user: dict = Depends(current_user), store: Store = Depends(get_store)):
    """The owner deletes it for everyone; an editor leaves it."""
    store.delete_doc(user["id"], doc_id)
    return Response(status_code=204)


@docs_router.post("/{doc_id}/share")
def share_doc(doc_id: str, user: dict = Depends(current_user), store: Store = Depends(get_store)):
    _own_doc(store, user, doc_id)
    return {"token": store.share_doc(doc_id)}


@docs_router.delete("/{doc_id}/share", status_code=204)
def unshare_doc(doc_id: str, user: dict = Depends(current_user), store: Store = Depends(get_store)):
    _own_doc(store, user, doc_id)
    store.unshare_doc(doc_id)
    return Response(status_code=204)


# ---------------------------------------------------------------------------
# Edit invites and collaborators
# ---------------------------------------------------------------------------


class InviteIn(BaseModel):
    role: Literal["editor", "viewer"] = "editor"


@docs_router.post("/{doc_id}/invite")
def invite_doc(doc_id: str, body: InviteIn | None = None, user: dict = Depends(current_user),
               store: Store = Depends(get_store)):
    """The invite link for a role (editor by default); one active link per role."""
    _own_doc(store, user, doc_id)
    role = body.role if body else "editor"
    return {"token": store.invite_doc(doc_id, role), "role": role}


@docs_router.get("/{doc_id}/invites")
def list_invites(doc_id: str, user: dict = Depends(current_user), store: Store = Depends(get_store)):
    """{editor: token | null, viewer: token | null}, for the owner."""
    _own_doc(store, user, doc_id)
    return store.invites(doc_id)


@docs_router.delete("/{doc_id}/invite", status_code=204)
def uninvite_doc(doc_id: str, role: Literal["editor", "viewer"] | None = None, user: dict = Depends(current_user),
                 store: Store = Depends(get_store)):
    """Turns an invite link off (?role=editor|viewer; both without it); people who already joined keep access."""
    _own_doc(store, user, doc_id)
    store.uninvite_doc(doc_id, role)
    return Response(status_code=204)


@docs_router.get("/{doc_id}/members")
def list_members(doc_id: str, user: dict = Depends(current_user), store: Store = Depends(get_store)):
    _can_open(store, user, doc_id)
    return store.members(doc_id)


@docs_router.delete("/{doc_id}/members/{member_id}", status_code=204)
def remove_member(doc_id: str, member_id: str, user: dict = Depends(current_user),
                  store: Store = Depends(get_store)):
    """The owner removes an editor, or an editor removes themself."""
    role = _can_open(store, user, doc_id)
    if role != "owner" and member_id != user["id"]:
        raise HTTPException(403, "Only the owner can remove other people.")
    if member_id == store.doc_owner(doc_id):
        raise HTTPException(400, "The owner can't be removed. Delete the workbook instead.")
    store.remove_member(doc_id, member_id)
    return Response(status_code=204)


class RoleIn(BaseModel):
    role: Literal["editor", "viewer"]


@docs_router.patch("/{doc_id}/members/{member_id}")
async def set_member_role(doc_id: str, member_id: str, body: RoleIn, user: dict = Depends(current_user),
                          store: Store = Depends(get_store)):
    """The owner makes a member an editor or a viewer. Their open live sessions switch at once."""
    _own_doc(store, user, doc_id)
    if member_id == user["id"]:
        raise HTTPException(400, "You own this workbook, so your role can't change.")
    if not await run_in_threadpool(store.set_member_role, doc_id, member_id, body.role):
        raise HTTPException(404, "That person isn't a member of this workbook.")
    from app.api.live import notify  # live imports this module

    await notify(doc_id, {"kind": "role", "user_id": member_id, "role": body.role})
    return {"id": member_id, "role": body.role}


@invites_router.post("/{token}/accept")
def accept_invite(token: str, user: dict = Depends(current_user), store: Store = Depends(get_store)):
    """Join the invite's doc with its role (an editor link upgrades a viewer; a viewer link never
    downgrades an editor). Returns {id, title, role}."""
    doc = store.accept_invite(token, user["id"])
    if not doc:
        raise HTTPException(404, INVITE_NOT_FOUND)
    return {"id": doc["id"], "title": doc["title"], "role": store.doc_role(user["id"], doc["id"])}


# ---------------------------------------------------------------------------
# Public share links (no sign-in)
# ---------------------------------------------------------------------------


@shared_router.get("/{token}")
def get_shared(token: str, store: Store = Depends(get_store)):
    doc = store.shared_doc(token)
    workbook = doc["data"].get("workbook") if doc and isinstance(doc["data"], dict) else None
    if not workbook:
        raise HTTPException(404, SHARE_NOT_FOUND)
    return {"title": doc["title"], "workbook": workbook}
