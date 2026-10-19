"""Sliding sessions, password reset, Google sign-in, edit invites/members and live co-editing."""

import logging
import time

import httpx
import pytest
from fastapi.testclient import TestClient
from starlette.websockets import WebSocketDisconnect

from app.api import live
from app.api.accounts import get_store, google_http
from app.config import get_settings
from app.main import app
from tests.test_api import tiny_workbook
from tests.test_more import auth, saved_doc, signup

client = TestClient(app)


def put_doc(headers: dict, doc_id: str = "doc1", title: str = "Budget"):
    r = client.put(f"/api/docs/{doc_id}", headers=headers,
                   json={"title": title, "updated": 1700000000000, "data": saved_doc(doc_id, title)})
    assert r.status_code == 200, r.text


# ---------------------------------------------------------------------------
# Sessions
# ---------------------------------------------------------------------------


def test_sessions_slide_and_expire():
    alice = signup("alice@example.com")
    store = get_store()
    assert store.user_for_token(alice["token"]) == alice["user"]
    assert store.user_for_token("") is None and store.user_for_token(None) is None
    assert store.user_for_token("nope") is None

    # Used 29 days ago: still valid, and the use pushes last_used forward.
    store._run("UPDATE sessions SET last_used = ?", (time.time() - 29 * 86400,))
    assert client.get("/api/auth/me", headers=auth(alice["token"])).status_code == 200
    assert store._run("SELECT last_used FROM sessions")[0]["last_used"] > time.time() - 60

    store._run("UPDATE sessions SET last_used = ?", (time.time() - 31 * 86400,))
    r = client.get("/api/auth/me", headers=auth(alice["token"]))
    assert r.status_code == 401 and r.json()["detail"] == "Please sign in again."
    assert store.user_for_token(alice["token"]) is None


# ---------------------------------------------------------------------------
# Password reset
# ---------------------------------------------------------------------------


def reset_links(caplog) -> list[str]:
    return [r.getMessage().split(": ", 1)[1] for r in caplog.records
            if r.getMessage().startswith("Password reset link for")]


def test_password_reset_flow(caplog):
    caplog.set_level(logging.WARNING)
    alice = signup("alice@example.com")
    other_session = client.post("/api/auth/login", json={"email": "alice@example.com",
                                                         "password": "correct horse"}).json()["token"]

    # Unknown and malformed emails look exactly the same and send nothing.
    assert client.post("/api/auth/forgot", json={"email": "nobody@example.com"}).status_code == 204
    assert client.post("/api/auth/forgot", json={"email": "junk"}).status_code == 204
    assert reset_links(caplog) == []

    assert client.post("/api/auth/forgot", json={"email": " Alice@Example.com "}).status_code == 204
    [link] = reset_links(caplog)
    assert link.startswith("http://localhost:3000/?reset=")
    token = link.split("reset=")[1]

    short = client.post("/api/auth/reset", json={"token": token, "password": "short"})
    assert short.status_code == 422 and short.json()["detail"] == "Use at least 8 characters for your password."
    bad = client.post("/api/auth/reset", json={"token": "wrong", "password": "new password"})
    assert bad.status_code == 400 and bad.json()["detail"] == "That reset link has expired. Ask for a new one."

    ok = client.post("/api/auth/reset", json={"token": token, "password": "new password"})
    assert ok.status_code == 200 and ok.json()["user"] == alice["user"]
    assert client.get("/api/auth/me", headers=auth(ok.json()["token"])).status_code == 200
    # Every other session ends; the old password stops working, the new one works.
    assert client.get("/api/auth/me", headers=auth(alice["token"])).status_code == 401
    assert client.get("/api/auth/me", headers=auth(other_session)).status_code == 401
    assert client.post("/api/auth/login", json={"email": "alice@example.com",
                                                "password": "correct horse"}).status_code == 401
    assert client.post("/api/auth/login", json={"email": "alice@example.com",
                                                "password": "new password"}).status_code == 200

    # Single use.
    again = client.post("/api/auth/reset", json={"token": token, "password": "another one"})
    assert again.status_code == 400


def test_password_reset_expiry_and_rate_limit(caplog):
    caplog.set_level(logging.WARNING)
    signup("bob@example.com")
    for _ in range(5):
        assert client.post("/api/auth/forgot", json={"email": "bob@example.com"}).status_code == 204
    links = reset_links(caplog)
    assert len(links) == 3  # the rest are dropped silently

    get_store()._run("UPDATE resets SET created = ?", (time.time() - 3601,))
    r = client.post("/api/auth/reset", json={"token": links[0].split("reset=")[1], "password": "new password"})
    assert r.status_code == 400


def test_reset_sends_email_when_smtp_configured(monkeypatch):
    sent = []
    settings = get_settings()
    monkeypatch.setattr(settings, "smtp_host", "smtp.example.com")
    monkeypatch.setattr(settings, "smtp_from", "noreply@example.com")
    monkeypatch.setattr(settings, "app_url", "https://sheets.example.com/")
    monkeypatch.setattr("app.api.accounts.send_mail", lambda s, to, subject, body: sent.append((to, body)))
    signup("carol@example.com")
    assert client.post("/api/auth/forgot", json={"email": "carol@example.com"}).status_code == 204
    [(to, body)] = sent
    assert to == "carol@example.com" and "https://sheets.example.com/?reset=" in body


# ---------------------------------------------------------------------------
# Google sign-in
# ---------------------------------------------------------------------------


def google_claims(**overrides) -> dict:
    claims = {"aud": "client-123", "iss": "https://accounts.google.com", "email": "Gina@Example.com",
              "email_verified": "true", "exp": str(int(time.time()) + 600)}
    claims.update(overrides)
    return claims


@pytest.fixture
def google(monkeypatch):
    """Answers tokeninfo with whatever claims the test puts in `replies` (keyed by credential)."""
    monkeypatch.setattr(get_settings(), "google_client_id", "client-123")
    replies: dict[str, dict | None] = {}

    def handler(request: httpx.Request) -> httpx.Response:
        assert request.url.host == "oauth2.googleapis.com" and request.url.path == "/tokeninfo"
        claims = replies.get(request.url.params["id_token"])
        return httpx.Response(200, json=claims) if claims else httpx.Response(400, json={"error": "invalid_token"})

    async def mock_http():
        async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as c:
            yield c

    app.dependency_overrides[google_http] = mock_http
    yield replies
    app.dependency_overrides.pop(google_http, None)


def test_auth_config(monkeypatch):
    assert client.get("/api/auth/config").json() == {"google_client_id": None, "password_reset": True}
    monkeypatch.setattr(get_settings(), "google_client_id", "client-123")
    assert client.get("/api/auth/config").json() == {"google_client_id": "client-123", "password_reset": True}


def test_google_sign_in_off_by_default():
    r = client.post("/api/auth/google", json={"credential": "x"})
    assert r.status_code == 503 and r.json()["detail"] == "Google sign-in isn't set up on this server."


def test_google_sign_in(google):
    google["good"] = google_claims()
    r = client.post("/api/auth/google", json={"credential": "good"})
    assert r.status_code == 200, r.text
    user = r.json()["user"]
    assert user["email"] == "gina@example.com"
    assert client.get("/api/auth/me", headers=auth(r.json()["token"])).json() == user
    # Same account next time; no password login until one is set.
    assert client.post("/api/auth/google", json={"credential": "good"}).json()["user"] == user
    assert client.post("/api/auth/login", json={"email": "gina@example.com", "password": ""}).status_code == 401

    # An existing password account signs in with Google too.
    alice = signup("alice@example.com")
    google["alice"] = google_claims(email="alice@example.com", iss="accounts.google.com")
    assert client.post("/api/auth/google", json={"credential": "alice"}).json()["user"] == alice["user"]

    google["aud"] = google_claims(aud="someone-else")
    google["iss"] = google_claims(iss="evil.example.com")
    google["unverified"] = google_claims(email_verified="false")
    google["expired"] = google_claims(exp=str(int(time.time()) - 5))
    for credential in ("aud", "iss", "unverified", "expired", "unknown", ""):
        r = client.post("/api/auth/google", json={"credential": credential})
        assert r.status_code == 401 and r.json()["detail"] == "Google sign-in failed. Please try again.", credential


# ---------------------------------------------------------------------------
# Invites and members
# ---------------------------------------------------------------------------


def test_invites_members_and_permissions():
    alice, bob, eve = signup("alice@example.com"), signup("bob@example.com"), signup("eve@example.com")
    a, b, e = auth(alice["token"]), auth(bob["token"]), auth(eve["token"])
    put_doc(a)

    assert client.post("/api/docs/doc1/invite", headers=b).status_code == 404  # not the owner
    token = client.post("/api/docs/doc1/invite", headers=a).json()["token"]
    assert client.post("/api/docs/doc1/invite", headers=a).json()["token"] == token

    assert client.post(f"/api/invites/{token}/accept").status_code == 401
    bad = client.post("/api/invites/nope/accept", headers=b)
    assert bad.status_code == 404 and bad.json()["detail"] == "This invite link doesn't exist or was turned off."
    assert client.post(f"/api/invites/{token}/accept", headers=b).json() == {"id": "doc1", "title": "Budget", "role": "editor"}
    assert client.post(f"/api/invites/{token}/accept", headers=b).status_code == 200  # idempotent
    assert client.post(f"/api/invites/{token}/accept", headers=a).json() == {"id": "doc1", "title": "Budget", "role": "owner"}

    members = [{"id": alice["user"]["id"], "email": "alice@example.com", "role": "owner"},
               {"id": bob["user"]["id"], "email": "bob@example.com", "role": "editor"}]
    assert client.get("/api/docs/doc1/members", headers=a).json() == members
    assert client.get("/api/docs/doc1/members", headers=b).json() == members
    assert client.get("/api/docs/doc1/members", headers=e).status_code == 404

    # Bob sees it in his list, reads it and writes it; it stays Alice's.
    [listed] = client.get("/api/docs", headers=b).json()
    assert listed["role"] == "editor" and listed["owner_email"] == "alice@example.com" and listed["id"] == "doc1"
    assert client.get("/api/docs/doc1", headers=b).json()["title"] == "Budget"
    r = client.put("/api/docs/doc1", headers=b, json={"title": "Bob's edit", "updated": 1700000001000,
                                                      "data": saved_doc("doc1", "Bob's edit")})
    assert r.status_code == 200
    assert client.get("/api/docs", headers=a).json()[0] | {"updated": 0} == {
        "id": "doc1", "title": "Bob's edit", "updated": 0, "sheets": 1, "shared": False,
        "role": "owner", "owner_email": "alice@example.com"}
    assert client.put("/api/docs/doc1", headers=e, json={"title": "x", "updated": 1, "data": {}}).status_code == 404
    assert client.post("/api/docs/doc1/share", headers=b).status_code == 404  # sharing stays the owner's

    # Turning the invite off keeps Bob but stops new people.
    assert client.delete("/api/docs/doc1/invite", headers=b).status_code == 404
    assert client.delete("/api/docs/doc1/invite", headers=a).status_code == 204
    assert client.post(f"/api/invites/{token}/accept", headers=e).status_code == 404
    assert client.get("/api/docs/doc1", headers=b).status_code == 200

    # Bob can't remove Alice or others; leaving via DELETE keeps the doc for Alice.
    assert client.delete(f"/api/docs/doc1/members/{alice['user']['id']}", headers=b).status_code == 403
    assert client.delete("/api/docs/doc1", headers=b).status_code == 204
    assert client.get("/api/docs/doc1", headers=b).status_code == 404
    assert client.get("/api/docs", headers=b).json() == []
    assert client.get("/api/docs/doc1", headers=a).status_code == 200

    # The owner removes an editor; a member removes themself.
    token = client.post("/api/docs/doc1/invite", headers=a).json()["token"]
    client.post(f"/api/invites/{token}/accept", headers=b)
    client.post(f"/api/invites/{token}/accept", headers=e)
    assert client.delete(f"/api/docs/doc1/members/{bob['user']['id']}", headers=a).status_code == 204
    assert client.delete(f"/api/docs/doc1/members/{eve['user']['id']}", headers=e).status_code == 204
    assert client.delete(f"/api/docs/doc1/members/{alice['user']['id']}", headers=a).status_code == 400
    assert [m["role"] for m in client.get("/api/docs/doc1/members", headers=a).json()] == ["owner"]

    # The owner deleting it removes it for everyone, invites included.
    client.post(f"/api/invites/{token}/accept", headers=b)
    assert client.delete("/api/docs/doc1", headers=a).status_code == 204
    assert client.get("/api/docs/doc1", headers=b).status_code == 404
    assert client.post(f"/api/invites/{token}/accept", headers=b).status_code == 404


# ---------------------------------------------------------------------------
# Live co-editing
# ---------------------------------------------------------------------------


def two_editors():
    alice, bob = signup("alice@example.com"), signup("bob@example.com")
    put_doc(auth(alice["token"]))
    token = client.post("/api/docs/doc1/invite", headers=auth(alice["token"])).json()["token"]
    client.post(f"/api/invites/{token}/accept", headers=auth(bob["token"]))
    return alice, bob


def test_live_rejects_bad_tokens_and_strangers():
    alice, _ = two_editors()
    eve = signup("eve@example.com")
    with TestClient(app) as c:
        for url, code in [("/api/live/doc1?token=nope", 4401), ("/api/live/doc1", 4401),
                          (f"/api/live/doc1?token={eve['token']}", 4404),
                          (f"/api/live/missing?token={alice['token']}", 4404)]:
            with c.websocket_connect(url) as ws, pytest.raises(WebSocketDisconnect) as closed:
                ws.receive_json()
            assert closed.value.code == code, url


def test_live_editing_session():
    alice, bob = two_editors()
    sheet = tiny_workbook("=1")["sheets"][0]["name"]
    with TestClient(app) as c:
        with c.websocket_connect(f"/api/live/doc1?token={alice['token']}") as wa:
            state = wa.receive_json()
            assert state["type"] == "state" and state["version"] == 0 and state["peers"] == []
            assert state["workbook"] == tiny_workbook("=B2*0.15")
            assert state["you"] == {**alice["user"], "color": live.PALETTE[0]}

            with c.websocket_connect(f"/api/live/doc1?token={bob['token']}") as wb:
                bstate = wb.receive_json()
                assert bstate["you"]["color"] == live.PALETTE[1]
                assert bstate["peers"] == [{**alice["user"], "color": live.PALETTE[0], "sel": None}]
                assert wa.receive_json() == {"type": "joined", "user": bstate["you"]}

                # Patch: Bob writes a cell past the end and renames a header.
                cells = [{"sheet": sheet, "r": 3, "c": 4, "cell": {"value": 42}},
                         {"sheet": sheet, "r": 1, "c": 0, "cell": None},
                         {"sheet": "Nope", "r": 1, "c": 0, "cell": {"value": "ignored"}}]
                headers = [{"sheet": sheet, "c": 1, "header": "Total"}]
                wb.send_json({"type": "patch", "cells": cells, "headers": headers})
                assert wb.receive_json() == {"type": "ack", "version": 1}
                assert wa.receive_json() == {"type": "patch", "version": 1, "from": bob["user"]["id"],
                                             "cells": cells, "headers": headers}

                # Presence goes to the others only.
                sel = {"sheet": sheet, "r1": 1, "c1": 0, "r2": 2, "c2": 1}
                wa.send_json({"type": "presence", "sel": sel})
                assert wb.receive_json() == {"type": "presence", "user": state["you"], "sel": sel}

                # Malformed messages are reported and change nothing.
                wa.send_text("not json")
                assert wa.receive_json()["type"] == "error"
                wa.send_json({"type": "patch", "cells": [{"sheet": sheet, "r": "x", "c": 0, "cell": None}]})
                assert wa.receive_json()["type"] == "error"
                wa.send_json({"type": "snapshot", "workbook": {"title": "no sheets key"}})
                assert wa.receive_json()["type"] == "error"
                wa.send_json({"type": "dance"})
                assert wa.receive_json()["type"] == "error"
                wa.send_text("x" * (live.MAX_MESSAGE_BYTES + 1))
                assert "too large" in wa.receive_json()["message"]

                # Snapshot: Alice replaces the workbook.
                snap = tiny_workbook("=B2*0.2")
                snap["title"] = "Renamed live"
                wa.send_json({"type": "snapshot", "workbook": snap})
                assert wa.receive_json() == {"type": "ack", "version": 2}
                assert wb.receive_json() == {"type": "snapshot", "version": 2, "from": alice["user"]["id"],
                                             "workbook": snap}

                wb.send_json({"type": "patch", "cells": [{"sheet": sheet, "r": 2, "c": 1, "cell": {"value": 7}}]})
                assert wb.receive_json()["version"] == 3
                assert wa.receive_json()["type"] == "patch"
            assert wa.receive_json() == {"type": "left", "user_id": bob["user"]["id"]}

            # A late joiner gets the current state, colour 2 is free again.
            with c.websocket_connect(f"/api/live/doc1?token={bob['token']}") as wb:
                again = wb.receive_json()
                assert again["version"] == 3 and again["you"]["color"] == live.PALETTE[1]
                assert again["peers"][0]["sel"] == sel
                rows = again["workbook"]["sheets"][0]["rows"]
                assert rows[1]["cells"][1] == {"value": 7}
                wa.receive_json()  # joined

    # Saved once the last person left.
    assert "doc1" not in live.rooms
    doc = client.get("/api/docs/doc1", headers=auth(alice["token"])).json()
    assert doc["title"] == "Renamed live" and doc["data"]["title"] == "Renamed live"
    assert doc["data"]["workbook"]["sheets"][0]["rows"][1]["cells"][1] == {"value": 7}
    assert doc["data"]["messages"][0]["text"] == "secret"  # the rest of the doc is untouched


def test_live_patch_grows_the_sheet_and_saves_while_editing(monkeypatch):
    monkeypatch.setattr(live, "SAVE_EVERY", 0.05)
    alice, _ = two_editors()
    sheet = tiny_workbook("=1")["sheets"][0]["name"]
    with TestClient(app) as c, c.websocket_connect(f"/api/live/doc1?token={alice['token']}") as ws:
        ws.receive_json()
        ws.send_json({"type": "patch", "cells": [{"sheet": sheet, "r": 4, "c": 5, "cell": {"value": "far"}}],
                      "headers": [{"sheet": sheet, "c": 6, "header": "New"}]})
        assert ws.receive_json()["type"] == "ack"
        deadline = time.time() + 5
        while True:
            wb = client.get("/api/docs/doc1", headers=auth(alice["token"])).json()["data"]["workbook"]
            if len(wb["sheets"][0]["rows"]) == 4 or time.time() > deadline:
                break
            time.sleep(0.05)
        s = wb["sheets"][0]
        assert len(s["rows"]) == 4 and s["rows"][3]["cells"][5] == {"value": "far"}
        assert s["rows"][3]["cells"][0] == {"value": None}
        assert len(s["columns"]) == 7 and s["columns"][6]["header"] == "New" and s["columns"][5]["header"] == ""
