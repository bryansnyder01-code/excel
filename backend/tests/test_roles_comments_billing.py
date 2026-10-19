"""Viewer/editor roles, locked cells, structural ops over the live socket, cell history and activity,
comments with @mentions, the Redis live broker (against a fake RESP server) and Stripe billing."""

import asyncio
import hashlib
import hmac
import json
import threading
import time
from urllib.parse import parse_qs

import httpx
import pytest
from fastapi.testclient import TestClient

from app.api import live
from app.api.billing import stripe_http, verify_signature
from app.config import get_settings
from app.main import app
from app.services import broker as broker_mod
from app.services.broker import RedisBroker, RedisConfig, encode_command, version_of
from app.services.usage import limit_for_user
from tests.test_api import tiny_workbook
from tests.test_more import auth, saved_doc, signup

client = TestClient(app)
VIEW_ONLY = "You can view this workbook but not change it."


def locked_workbook() -> dict:
    wb = tiny_workbook("=B2*0.15")
    wb["sheets"][0]["locked_ranges"] = ["C2", "A1:C1"]
    wb["sheets"].append({"name": "Rates", "locked": True, "columns": [{"header": "Rate"}],
                         "rows": [{"cells": [{"value": 0.15}]}]})
    return wb


def setup_doc(workbook: dict | None = None):
    """Alice owns doc1; Bob is an editor, Vera a viewer (each through their own invite link)."""
    alice, bob, vera = signup("alice@example.com"), signup("bob@example.com"), signup("vera@example.com")
    a = auth(alice["token"])
    data = saved_doc("doc1")
    if workbook is not None:
        data["workbook"] = workbook
    assert client.put("/api/docs/doc1", headers=a, json={"title": "Budget", "updated": 1, "data": data}).status_code == 200
    editor = client.post("/api/docs/doc1/invite", headers=a).json()
    viewer = client.post("/api/docs/doc1/invite", headers=a, json={"role": "viewer"}).json()
    assert editor["role"] == "editor" and viewer["role"] == "viewer" and editor["token"] != viewer["token"]
    assert client.post(f"/api/invites/{editor['token']}/accept", headers=auth(bob["token"])).status_code == 200
    assert client.post(f"/api/invites/{viewer['token']}/accept", headers=auth(vera["token"])).status_code == 200
    return alice, bob, vera, editor["token"], viewer["token"]


def put(headers, workbook, doc_id="doc1"):
    data = saved_doc(doc_id)
    data["workbook"] = workbook
    return client.put(f"/api/docs/{doc_id}", headers=headers, json={"title": "Budget", "updated": 2, "data": data})


def recv_until(ws, kind: str, limit: int = 20) -> dict:
    for _ in range(limit):
        msg = ws.receive_json()
        if msg["type"] == kind:
            return msg
    raise AssertionError(f"no {kind} message")


# ---------------------------------------------------------------------------
# Roles
# ---------------------------------------------------------------------------


def test_roles_invites_and_permissions():
    alice, bob, vera, editor_token, viewer_token = setup_doc()
    a, b, v = auth(alice["token"]), auth(bob["token"]), auth(vera["token"])

    # One active token per role; the owner lists them; nobody else can.
    assert client.post("/api/docs/doc1/invite", headers=a, json={"role": "viewer"}).json()["token"] == viewer_token
    assert client.post("/api/docs/doc1/invite", headers=a, json={}).json() == {"token": editor_token, "role": "editor"}
    assert client.get("/api/docs/doc1/invites", headers=a).json() == {"editor": editor_token, "viewer": viewer_token}
    assert client.get("/api/docs/doc1/invites", headers=b).status_code == 404
    assert client.post("/api/docs/doc1/invite", headers=a, json={"role": "owner"}).status_code == 422

    members = client.get("/api/docs/doc1/members", headers=v).json()
    assert [(m["email"], m["role"]) for m in members] == [
        ("alice@example.com", "owner"), ("bob@example.com", "editor"), ("vera@example.com", "viewer")]
    assert client.get("/api/docs", headers=v).json()[0]["role"] == "viewer"
    assert client.get("/api/docs", headers=b).json()[0]["role"] == "editor"

    # Viewers read but can't write.
    assert client.get("/api/docs/doc1", headers=v).status_code == 200
    r = put(v, tiny_workbook("=1"))
    assert r.status_code == 403 and r.json()["detail"] == VIEW_ONLY
    assert put(b, tiny_workbook("=2")).status_code == 200

    # A viewer link never downgrades an editor; an editor link upgrades a viewer.
    client.post(f"/api/invites/{viewer_token}/accept", headers=b)
    assert client.get("/api/docs", headers=b).json()[0]["role"] == "editor"
    client.post(f"/api/invites/{editor_token}/accept", headers=v)
    assert client.get("/api/docs", headers=v).json()[0]["role"] == "editor"

    # The owner changes roles.
    vid = vera["user"]["id"]
    assert client.patch(f"/api/docs/doc1/members/{vid}", headers=b, json={"role": "viewer"}).status_code == 404
    assert client.patch(f"/api/docs/doc1/members/{vid}", headers=a, json={"role": "boss"}).status_code == 422
    assert client.patch(f"/api/docs/doc1/members/{alice['user']['id']}", headers=a,
                        json={"role": "viewer"}).status_code == 400
    assert client.patch("/api/docs/doc1/members/nobody", headers=a, json={"role": "viewer"}).status_code == 404
    r = client.patch(f"/api/docs/doc1/members/{vid}", headers=a, json={"role": "viewer"})
    assert r.status_code == 200 and r.json() == {"id": vid, "role": "viewer"}
    assert put(v, tiny_workbook("=3")).status_code == 403

    # Turning off one role's link leaves the other; no role turns off both.
    assert client.delete("/api/docs/doc1/invite?role=viewer", headers=a).status_code == 204
    assert client.get("/api/docs/doc1/invites", headers=a).json() == {"editor": editor_token, "viewer": None}
    client.post("/api/docs/doc1/invite", headers=a, json={"role": "viewer"})
    assert client.delete("/api/docs/doc1/invite", headers=a).status_code == 204
    assert client.get("/api/docs/doc1/invites", headers=a).json() == {"editor": None, "viewer": None}


def test_live_viewer_and_role_changes():
    alice, bob, vera, *_ = setup_doc()
    sheet = "Sales"
    with TestClient(app) as c:
        with c.websocket_connect(f"/api/live/doc1?token={vera['token']}") as wv:
            state = wv.receive_json()
            assert state["role"] == "viewer"
            wb = c.websocket_connect(f"/api/live/doc1?token={bob['token']}").__enter__()
            assert wb.receive_json()["role"] == "editor"
            wv.receive_json()  # bob joined

            patch = {"type": "patch", "cells": [{"sheet": sheet, "r": 1, "c": 1, "cell": {"value": 5}}]}
            for msg in (patch, {"type": "snapshot", "workbook": tiny_workbook("=1")},
                        {"type": "ops", "ops": [{"op": "insert_rows", "sheet": sheet, "at": 1, "count": 1}]}):
                wv.send_json(msg)
                assert wv.receive_json() == {"type": "error", "message": VIEW_ONLY}
            # Viewers still share their selection.
            sel = {"sheet": sheet, "r1": 1, "c1": 1, "r2": 1, "c2": 1}
            wv.send_json({"type": "presence", "sel": sel})
            assert wb.receive_json()["sel"] == sel

            # The owner makes Bob a viewer: his open session switches at once.
            r = c.patch(f"/api/docs/doc1/members/{bob['user']['id']}", headers=auth(alice["token"]),
                        json={"role": "viewer"})
            assert r.status_code == 200
            assert wb.receive_json() == {"type": "role", "role": "viewer"}
            wb.send_json(patch)
            assert wb.receive_json() == {"type": "error", "message": VIEW_ONLY}
            wb.__exit__(None, None, None)
    assert client.get("/api/docs/doc1", headers=auth(alice["token"])).json()["data"]["workbook"] == tiny_workbook("=B2*0.15")


# ---------------------------------------------------------------------------
# Locked cells
# ---------------------------------------------------------------------------


def test_put_protection():
    alice, bob, *_ = setup_doc(locked_workbook())
    a, b = auth(alice["token"]), auth(bob["token"])

    changed = locked_workbook()
    changed["sheets"][0]["rows"][0]["cells"][1]["value"] = 250  # B2: not locked
    assert put(b, changed).status_code == 200

    for edit, message in [
        (lambda wb: wb["sheets"][0]["rows"][0]["cells"].__setitem__(2, {"value": 1}),
         "C2 on Sales is locked. Only the owner can change it."),
        (lambda wb: wb["sheets"][0]["columns"][0].__setitem__("header", "Thing"),
         "A1 on Sales is locked. Only the owner can change it."),
        (lambda wb: wb["sheets"][1]["rows"][0]["cells"][0].__setitem__("value", 0.2),
         'The sheet "Rates" is locked. Only the owner can change it.'),
        (lambda wb: wb["sheets"].pop(), 'The sheet "Rates" is locked. Only the owner can change it.'),
        (lambda wb: wb["sheets"][0].__setitem__("locked_ranges", []),
         "Only the owner can change which cells are locked."),
    ]:
        wb = json.loads(json.dumps(changed))
        edit(wb)
        r = put(b, wb)
        assert r.status_code == 403 and r.json()["detail"] == message
        assert put(a, wb).status_code == 200  # the owner can do all of it
        assert put(a, changed).status_code == 200


def test_live_protection_and_ops():
    alice, bob, *_ = setup_doc(locked_workbook())
    with TestClient(app) as c:
        with c.websocket_connect(f"/api/live/doc1?token={alice['token']}") as wa:
            wa.receive_json()
            wb = c.websocket_connect(f"/api/live/doc1?token={bob['token']}").__enter__()
            wb.receive_json()
            wa.receive_json()  # joined

            # Bob's patch touching C2 is refused whole: B2 in the same message isn't applied either.
            wb.send_json({"type": "patch", "cells": [{"sheet": "Sales", "r": 1, "c": 1, "cell": {"value": 1}},
                                                     {"sheet": "Sales", "r": 1, "c": 2, "cell": {"value": 2}}]})
            assert wb.receive_json() == {"type": "error", "message": "C2 on Sales is locked. Only the owner can change it."}
            wb.send_json({"type": "patch", "headers": [{"sheet": "Sales", "c": 0, "header": "X"}]})
            assert "A1 on Sales is locked" in wb.receive_json()["message"]
            wb.send_json({"type": "patch", "cells": [{"sheet": "Rates", "r": 1, "c": 0, "cell": {"value": 2}}]})
            assert wb.receive_json()["message"] == 'The sheet "Rates" is locked. Only the owner can change it.'
            snap = locked_workbook()
            snap["sheets"][0]["rows"][0]["cells"][2] = {"value": 0}
            wb.send_json({"type": "snapshot", "workbook": snap})
            assert wb.receive_json()["type"] == "error"
            for ops in ([{"op": "delete_rows", "sheet": "Sales", "at": 1, "count": 1}],
                        [{"op": "insert_cols", "sheet": "Sales", "at": 1, "count": 1}],  # inside A1:C1
                        [{"op": "insert_rows", "sheet": "Rates", "at": 1, "count": 1}],
                        [{"op": "insert_rows", "sheet": "Sales", "at": 1, "count": 1},
                         {"op": "delete_rows", "sheet": "Sales", "at": 2, "count": 1}]):  # the moved row
                wb.send_json({"type": "ops", "ops": ops})
                assert wb.receive_json()["type"] == "error", ops
            for bad in ([], [{"op": "explode", "sheet": "Sales", "at": 1, "count": 1}],
                        [{"op": "insert_rows", "sheet": "Nope", "at": 1, "count": 1}],
                        [{"op": "insert_rows", "sheet": "Sales", "at": 0, "count": 1}],
                        [{"op": "insert_cols", "sheet": "Sales", "at": 0, "count": 0}]):
                wb.send_json({"type": "ops", "ops": bad})
                assert wb.receive_json()["type"] == "error", bad

            # Moving the whole locked range is fine: Bob inserts two rows above row 2.
            ops = [{"op": "insert_rows", "sheet": "Sales", "at": 1, "count": 2}]
            wb.send_json({"type": "ops", "ops": ops})
            assert wb.receive_json() == {"type": "ack", "version": 1}
            assert wa.receive_json() == {"type": "ops", "version": 1, "from": bob["user"]["id"], "ops": ops}
            # Now C4 is the locked cell, and its formula moved with it.
            wb.send_json({"type": "patch", "cells": [{"sheet": "Sales", "r": 3, "c": 2, "cell": {"value": 1}}]})
            assert "C4 on Sales is locked" in wb.receive_json()["message"]
            wb.send_json({"type": "patch", "cells": [{"sheet": "Sales", "r": 1, "c": 2, "cell": {"value": "ok"}}]})
            assert wb.receive_json()["type"] == "ack"
            wa.receive_json()

            # The owner can change anything.
            wa.send_json({"type": "patch", "cells": [{"sheet": "Rates", "r": 1, "c": 0, "cell": {"value": 0.2}}]})
            assert wa.receive_json()["type"] == "ack"
            wa.send_json({"type": "ops", "ops": [{"op": "delete_cols", "sheet": "Sales", "at": 0, "count": 1}]})
            assert wa.receive_json()["type"] == "ack"
            assert wb.receive_json()["type"] == "patch"
            assert wb.receive_json()["ops"][0]["op"] == "delete_cols"
            wb.__exit__(None, None, None)

    saved = client.get("/api/docs/doc1", headers=auth(alice["token"])).json()["data"]["workbook"]
    sales, rates = saved["sheets"]
    assert [c["header"] for c in sales["columns"]] == ["Amount", "Tax"]
    assert sales["locked_ranges"] == ["B4", "A1:B1"]
    assert sales["rows"][2]["cells"] == [{"value": 100}, {"value": "=A4*0.15"}]  # grid row 3 = Excel row 4
    assert sales["rows"][0]["cells"] == [{"value": None}, {"value": "ok"}]
    assert rates["rows"][0]["cells"][0]["value"] == 0.2


# ---------------------------------------------------------------------------
# History and activity
# ---------------------------------------------------------------------------


def test_history_and_activity(monkeypatch):
    alice, bob, vera, *_ = setup_doc()
    with TestClient(app) as c:
        with c.websocket_connect(f"/api/live/doc1?token={bob['token']}") as ws:
            ws.receive_json()
            for value in (1, 2):
                ws.send_json({"type": "patch", "cells": [{"sheet": "Sales", "r": 1, "c": 1, "cell": {"value": value}}],
                              "headers": [{"sheet": "Sales", "c": 2, "header": "VAT"}]})
                assert ws.receive_json()["type"] == "ack"
            ws.send_json({"type": "patch", "cells": [{"sheet": "Sales", "r": 1, "c": 1, "cell": None}]})
            ws.receive_json()
            ws.send_json({"type": "ops", "ops": [{"op": "insert_rows", "sheet": "Sales", "at": 1, "count": 2},
                                                 {"op": "delete_cols", "sheet": "Sales", "at": 2, "count": 1}]})
            ws.receive_json()
            ws.send_json({"type": "snapshot", "workbook": tiny_workbook("=1")})
            ws.receive_json()

    v = auth(vera["token"])
    hist = client.get("/api/docs/doc1/history", params={"sheet": "Sales", "cell": "b2"}, headers=v).json()
    bob_user = {"id": bob["user"]["id"], "email": "bob@example.com"}
    assert [h["value"] for h in hist] == [None, {"value": 2}, {"value": 1}]
    assert all(h["user"] == bob_user and isinstance(h["at"], int) for h in hist)
    assert client.get("/api/docs/doc1/history", params={"sheet": "Sales", "cell": "C1"},
                      headers=v).json()[0]["value"] == {"value": "VAT"}
    assert client.get("/api/docs/doc1/history", params={"sheet": "Sales", "cell": "nope"}, headers=v).status_code == 422
    eve = signup("eve@example.com")
    assert client.get("/api/docs/doc1/history", params={"sheet": "Sales", "cell": "B2"},
                      headers=auth(eve["token"])).status_code == 404

    feed = client.get("/api/docs/doc1/activity", params={"limit": 4}, headers=v).json()
    assert [(f["kind"], f["sheet"], f["cell"]) for f in feed] == [
        ("replace", None, None), ("restructure", "Sales", None), ("edit", "Sales", "B2"), ("edit", "Sales", "B2")]
    assert feed[0]["summary"] == "Replaced the whole workbook"
    assert feed[1]["summary"] == "Inserted 2 rows at row 2 on Sales; Deleted 1 column at column C on Sales"
    assert feed[2]["summary"] == "Cleared B2 on Sales" and feed[3]["summary"] == "Set B2 on Sales to 2"
    assert feed[0]["user"] == bob_user

    # Pruned to the newest HISTORY_KEEP entries per doc.
    monkeypatch.setattr("app.services.store.HISTORY_KEEP", 3)
    from app.api.accounts import get_store

    get_store().add_history("doc1", bob["user"]["id"], [{"sheet": "S", "r": 1, "c": 0, "value": {"value": 9}}])
    assert len(client.get("/api/docs/doc1/activity", headers=v).json()) == 3


# ---------------------------------------------------------------------------
# Comments
# ---------------------------------------------------------------------------


def test_comments_and_mentions(monkeypatch):
    alice, bob, vera, *_ = setup_doc()
    a, b, v = auth(alice["token"]), auth(bob["token"]), auth(vera["token"])
    sent = []
    monkeypatch.setattr("app.api.comments.send_or_log", lambda s, to, subject, body: sent.append((to, subject, body)))
    monkeypatch.setattr(get_settings(), "app_url", "https://sheets.example.com/")

    with TestClient(app) as c, c.websocket_connect(f"/api/live/doc1?token={alice['token']}") as wa:
        wa.receive_json()
        # Viewers may comment; mentions of members (and the owner) email them, not strangers or yourself.
        body = "Is this right @Bob@example.com? cc @alice@example.com, @eve@example.com and @vera@example.com."
        r = c.post("/api/docs/doc1/comments", headers=v, json={"sheet": "Sales", "cell": "c2", "body": body})
        assert r.status_code == 200, r.text
        thread = r.json()
        assert thread["sheet"] == "Sales" and thread["cell"] == "C2" and thread["resolved"] is False
        [first] = thread["comments"]
        assert first["author"] == {"id": vera["user"]["id"], "email": "vera@example.com"} and first["body"] == body
        assert "edited" not in first and isinstance(first["created"], int)
        assert wa.receive_json() == {"type": "comments", "thread_id": thread["id"]}
        assert sorted(to for to, _, _ in sent) == ["alice@example.com", "bob@example.com"]
        to, subject, text = sent[0]
        assert subject == "vera@example.com mentioned you in Budget"
        assert "https://sheets.example.com/?doc=doc1&cell=Sales%21C2" in text and body in text

        # Reply, resolve, re-open by replying.
        r = c.post("/api/docs/doc1/comments", headers=b, json={"thread_id": thread["id"], "body": "Yes"})
        assert [x["body"] for x in r.json()["comments"]] == [body, "Yes"]
        wa.receive_json()
        assert c.patch(f"/api/docs/doc1/comments/{thread['id']}", headers=b,
                       json={"resolved": True}).json()["resolved"] is True
        wa.receive_json()
        reply = c.post("/api/docs/doc1/comments", headers=a, json={"thread_id": thread["id"], "body": "Reopen"}).json()
        assert reply["resolved"] is False
        wa.receive_json()

    # Validation.
    assert client.post("/api/docs/doc1/comments", headers=a, json={"sheet": "Sales", "cell": "C2", "body": " "}).status_code == 422
    assert client.post("/api/docs/doc1/comments", headers=a,
                       json={"sheet": "Sales", "cell": "C2", "body": "x" * 2001}).status_code == 422
    assert client.post("/api/docs/doc1/comments", headers=a, json={"sheet": "Sales", "cell": "?", "body": "x"}).status_code == 422
    assert client.post("/api/docs/doc1/comments", headers=a, json={"thread_id": "nope", "body": "x"}).status_code == 404
    eve = signup("eve@example.com")
    assert client.get("/api/docs/doc1/comments", headers=auth(eve["token"])).status_code == 404

    # Edit (author only) and delete (author or owner); an empty thread disappears.
    comments = reply["comments"]
    vera_c, bob_c, alice_c = comments
    url = f"/api/docs/doc1/comments/{thread['id']}"
    assert client.patch(f"{url}/{bob_c['id']}", headers=a, json={"body": "hijack"}).status_code == 403
    edited = client.patch(f"{url}/{bob_c['id']}", headers=b, json={"body": "Yes, checked"}).json()
    assert edited["comments"][1]["body"] == "Yes, checked" and edited["comments"][1]["edited"]
    assert client.delete(f"{url}/{alice_c['id']}", headers=b).status_code == 403
    assert client.delete(f"{url}/{bob_c['id']}", headers=b).status_code == 204
    assert client.delete(f"{url}/{vera_c['id']}", headers=a).status_code == 204
    assert client.delete(f"{url}/{vera_c['id']}", headers=a).status_code == 404
    [left] = client.get("/api/docs/doc1/comments", headers=v).json()
    assert [x["body"] for x in left["comments"]] == ["Reopen"]
    assert client.delete(f"{url}/{alice_c['id']}", headers=a).status_code == 204
    assert client.get("/api/docs/doc1/comments", headers=v).json() == []


# ---------------------------------------------------------------------------
# Redis broker, against a fake RESP server
# ---------------------------------------------------------------------------


def _id_key(entry_id: str) -> tuple[int, int]:
    ms, _, seq = entry_id.partition("-")
    return int(ms), int(seq or 0)


class FakeRedis:
    """Just enough Redis for RedisBroker: AUTH, SELECT, PING, XADD (MAXLEN ~), XREAD (COUNT, BLOCK), XREVRANGE."""

    def __init__(self, password: str | None = None):
        self.password = password
        self.streams: dict[str, list[tuple[str, list[str]]]] = {}
        self.commands: list[list[str]] = []
        self.last = (0, 0)
        self.loop = asyncio.new_event_loop()
        self.ready = threading.Event()
        self.thread = threading.Thread(target=self._run, daemon=True)
        self.thread.start()
        self.ready.wait(5)

    def _run(self):
        asyncio.set_event_loop(self.loop)
        self.cond = asyncio.Condition()
        self.server = self.loop.run_until_complete(asyncio.start_server(self._client, "127.0.0.1", 0))
        self.port = self.server.sockets[0].getsockname()[1]
        self.ready.set()
        self.loop.run_forever()

    def stop(self):
        async def _shutdown():
            self.server.close()
            tasks = [t for t in asyncio.all_tasks() if t is not asyncio.current_task()]
            for t in tasks:
                t.cancel()
            await asyncio.gather(*tasks, return_exceptions=True)

        asyncio.run_coroutine_threadsafe(_shutdown(), self.loop).result(5)
        self.loop.call_soon_threadsafe(self.loop.stop)
        self.thread.join(5)
        self.loop.close()

    def entries(self, key: str) -> list[tuple[str, dict]]:
        return [(i, json.loads(dict(zip(f[::2], f[1::2]))["e"])) for i, f in list(self.streams.get(key, []))]

    @staticmethod
    async def _read_command(reader) -> list[str]:
        line = await reader.readline()
        if not line:
            raise ConnectionError
        assert line[:1] == b"*"
        args = []
        for _ in range(int(line[1:-2])):
            n = int((await reader.readline())[1:-2])
            args.append((await reader.readexactly(n + 2))[:-2].decode())
        return args

    @staticmethod
    def _encode(value) -> bytes:
        if value is None:
            return b"*-1\r\n"
        if isinstance(value, list):
            return b"*%d\r\n" % len(value) + b"".join(FakeRedis._encode(v) for v in value)
        b = str(value).encode()
        return b"$%d\r\n%s\r\n" % (len(b), b)

    async def _client(self, reader, writer):
        authed = self.password is None
        try:
            while True:
                args = await self._read_command(reader)
                self.commands.append(args)
                cmd = args[0].upper()
                if cmd == "AUTH":
                    authed = args[-1] == self.password
                    writer.write(b"+OK\r\n" if authed else b"-WRONGPASS invalid password\r\n")
                elif not authed:
                    writer.write(b"-NOAUTH Authentication required.\r\n")
                elif cmd in ("SELECT", "PING"):
                    writer.write(b"+OK\r\n" if cmd == "SELECT" else b"+PONG\r\n")
                elif cmd == "XADD":
                    writer.write(self._encode(await self._xadd(args)))
                elif cmd == "XREVRANGE":
                    entries = self.streams.get(args[1], [])
                    writer.write(self._encode([[i, f] for i, f in entries[::-1][:int(args[5])]]))
                elif cmd == "XREAD":
                    writer.write(self._encode(await self._xread(args)))
                else:
                    writer.write(b"-ERR unknown command\r\n")
                await writer.drain()
        except (ConnectionError, asyncio.IncompleteReadError, asyncio.CancelledError):
            pass
        finally:
            writer.close()

    async def _xadd(self, args):
        key, rest = args[1], args[2:]
        maxlen = None
        if rest[0].upper() == "MAXLEN":
            rest = rest[1:]
            if rest[0] in ("~", "="):
                rest = rest[1:]
            maxlen, rest = int(rest[0]), rest[1:]
        assert rest[0] == "*"
        ms = int(time.time() * 1000)
        self.last = (ms, 0) if ms > self.last[0] else (self.last[0], self.last[1] + 1)
        entry_id = f"{self.last[0]}-{self.last[1]}"
        stream = self.streams.setdefault(key, [])
        stream.append((entry_id, rest[1:]))
        if maxlen is not None:
            del stream[:-maxlen]
        async with self.cond:
            self.cond.notify_all()
        return entry_id

    async def _xread(self, args):
        upper = [a.upper() for a in args]
        count = int(args[upper.index("COUNT") + 1]) if "COUNT" in upper else 1000
        block = int(args[upper.index("BLOCK") + 1]) if "BLOCK" in upper else None
        s = upper.index("STREAMS")
        key, after = args[s + 1], args[s + 2]
        after_key = self.last if after == "$" else _id_key(after)
        deadline = time.monotonic() + (block or 0) / 1000
        while True:
            found = [[i, f] for i, f in self.streams.get(key, []) if _id_key(i) > after_key][:count]
            if found:
                return [[key, found]]
            if block is None or time.monotonic() >= deadline:
                return None
            async with self.cond:
                try:
                    await asyncio.wait_for(self.cond.wait(), max(0.01, deadline - time.monotonic()))
                except asyncio.TimeoutError:
                    pass


@pytest.fixture
def fake_redis():
    server = FakeRedis(password="s3cret")
    yield server
    server.stop()


def test_resp_helpers():
    assert encode_command("XADD", "k", 5) == b"*3\r\n$4\r\nXADD\r\n$1\r\nk\r\n$1\r\n5\r\n"
    cfg = RedisConfig.from_url("rediss://user:p%40ss@cache.example.com:6380/3")
    assert (cfg.host, cfg.port, cfg.username, cfg.password, cfg.db, cfg.tls) == (
        "cache.example.com", 6380, "user", "p@ss", 3, True)
    assert RedisConfig.from_url("redis://localhost").db == 0
    with pytest.raises(ValueError):
        RedisConfig.from_url("http://localhost")
    assert version_of("1700000000000-2") == 1700000000000002 and version_of(None) == 0


def test_redis_broker_orders_events_across_instances(fake_redis):
    url = f"redis://:s3cret@127.0.0.1:{fake_redis.port}/2"

    async def scenario():
        one, two = RedisBroker(url, block_ms=200), RedisBroker(url, block_ms=200)
        got_one, got_two = [], []
        await one.subscribe("d", lambda i, e: got_one.append((i, e["n"])))
        await two.subscribe("d", lambda i, e: got_two.append((i, e["n"])))
        for n in range(10):
            await (one if n % 2 else two).publish("d", {"n": n})
        for _ in range(100):
            if len(got_one) == len(got_two) == 10:
                break
            await asyncio.sleep(0.02)
        assert await one.position("d") == got_one[-1][0]
        await one.close()
        await two.close()
        return got_one, got_two

    got_one, got_two = asyncio.run(scenario())
    assert got_one == got_two and [n for _, n in got_one] == list(range(10))
    assert ["AUTH", "s3cret"] in fake_redis.commands and ["SELECT", "2"] in fake_redis.commands
    xadd = next(c for c in fake_redis.commands if c[0] == "XADD")
    assert xadd[:6] == ["XADD", "sheetcraft:live:d", "MAXLEN", "~", "10000", "*"]

    async def wrong_password():
        b = RedisBroker(f"redis://:nope@127.0.0.1:{fake_redis.port}")
        with pytest.raises(broker_mod.RespError):
            await b.publish("d", {"n": 1})

    asyncio.run(wrong_password())


def _publish_as_other_instance(url: str, doc_id: str, event: dict) -> str:
    async def go():
        b = RedisBroker(url)
        await b.publish(doc_id, {**event, "origin": "other-instance"})
        pos = await b.position(doc_id)
        await b.close()
        return pos
    return asyncio.run(go())


def test_live_editing_through_redis(fake_redis, monkeypatch):
    url = f"redis://:s3cret@127.0.0.1:{fake_redis.port}/0"
    monkeypatch.setattr(get_settings(), "redis_url", url)
    alice, bob, *_ = setup_doc()
    key = "sheetcraft:live:doc1"
    other = {"id": bob["user"]["id"], "email": "bob@example.com", "color": live.PALETTE[1]}
    with TestClient(app) as c:
        with c.websocket_connect(f"/api/live/doc1?token={alice['token']}") as wa:
            state = wa.receive_json()
            assert state["version"] == 0 and state["role"] == "owner"

            wa.send_json({"type": "patch", "cells": [{"sheet": "Sales", "r": 1, "c": 1, "cell": {"value": 7}}]})
            ack = wa.receive_json()
            assert ack["type"] == "ack" and ack["version"] == version_of(fake_redis.entries(key)[-1][0])

            # Bob, on another instance, joins and edits: Alice's instance relays it and says who's here.
            _publish_as_other_instance(url, "doc1", {"kind": "joined", "conn": "bob-conn", "user": other})
            assert wa.receive_json() == {"type": "joined", "user": other}
            entry = _publish_as_other_instance(url, "doc1", {
                "kind": "ops", "conn": "bob-conn", "from": bob["user"]["id"],
                "ops": [{"op": "insert_rows", "sheet": "Sales", "at": 1, "count": 1}]})
            msg = wa.receive_json()
            assert msg == {"type": "ops", "version": version_of(entry), "from": bob["user"]["id"],
                           "ops": [{"op": "insert_rows", "sheet": "Sales", "at": 1, "count": 1}]}
            sel = {"sheet": "Sales", "r1": 1, "c1": 0, "r2": 1, "c2": 0}
            _publish_as_other_instance(url, "doc1", {"kind": "presence", "conn": "bob-conn", "user": other, "sel": sel})
            assert wa.receive_json() == {"type": "presence", "user": other, "sel": sel}
            for _ in range(50):
                here = [e for _, e in fake_redis.entries(key) if e["kind"] == "here"]
                if here:
                    break
                time.sleep(0.02)
            assert here[0]["peers"][0]["user"]["id"] == alice["user"]["id"]

            # A second tab of Alice's sees Bob as a peer (remote) with his selection.
            with c.websocket_connect(f"/api/live/doc1?token={alice['token']}") as wa2:
                st = wa2.receive_json()
                assert {"sel": sel, **other} in st["peers"]
                assert st["workbook"]["sheets"][0]["rows"][1]["cells"][1] == {"value": 7}
            _publish_as_other_instance(url, "doc1", {"kind": "left", "conn": "bob-conn", "user_id": bob["user"]["id"]})
            assert recv_until(wa, "left") == {"type": "left", "user_id": bob["user"]["id"]}

    saved = client.get("/api/docs/doc1", headers=auth(alice["token"])).json()["data"]["workbook"]
    assert saved["sheets"][0]["rows"][0] == {"kind": "data", "cells": []}
    assert saved["sheets"][0]["rows"][1]["cells"][1] == {"value": 7}

    # Edits made elsewhere while nobody here had the doc open are replayed when it opens again.
    entry = _publish_as_other_instance(url, "doc1", {
        "kind": "patch", "conn": "bob-conn", "from": bob["user"]["id"],
        "cells": [{"sheet": "Sales", "r": 1, "c": 0, "cell": {"value": "late"}}], "headers": []})
    with TestClient(app) as c, c.websocket_connect(f"/api/live/doc1?token={alice['token']}") as wa:
        assert wa.receive_json()["type"] == "state"
        msg = recv_until(wa, "patch")
        assert msg["version"] == version_of(entry) and msg["cells"][0]["cell"] == {"value": "late"}
    saved = client.get("/api/docs/doc1", headers=auth(alice["token"])).json()["data"]["workbook"]
    assert saved["sheets"][0]["rows"][0]["cells"][0] == {"value": "late"}
    for _ in range(100):  # the room closes once its last save is done
        if not live.rooms:
            break
        time.sleep(0.02)
    assert live.rooms == {}


# ---------------------------------------------------------------------------
# Billing
# ---------------------------------------------------------------------------


@pytest.fixture
def stripe(monkeypatch):
    s = get_settings()
    monkeypatch.setattr(s, "stripe_secret_key", "sk_test_123")
    monkeypatch.setattr(s, "stripe_webhook_secret", "whsec_abc")
    monkeypatch.setattr(s, "stripe_price_pro", "price_pro")
    monkeypatch.setattr(s, "app_url", "https://sheets.example.com")
    calls: list[tuple[str, str, dict]] = []

    def handler(request: httpx.Request) -> httpx.Response:
        assert request.url.host == "api.stripe.com"
        assert request.headers["authorization"] == "Bearer sk_test_123"
        form = {k: v[0] for k, v in parse_qs(request.content.decode()).items()}
        calls.append((request.method, request.url.path, form))
        if request.url.path == "/v1/customers":
            return httpx.Response(200, json={"id": "cus_1"})
        if request.url.path == "/v1/checkout/sessions":
            return httpx.Response(200, json={"id": "cs_1", "url": "https://checkout.stripe.com/c/cs_1"})
        if request.url.path == "/v1/billing_portal/sessions":
            return httpx.Response(200, json={"url": "https://billing.stripe.com/p/1"})
        if request.url.path == "/v1/subscriptions/sub_1":
            return httpx.Response(200, json={"id": "sub_1", "items": {"data": [{"current_period_end": 1900000000}]}})
        return httpx.Response(404, json={"error": {"message": "nope"}})

    async def mock_http():
        async with httpx.AsyncClient(transport=httpx.MockTransport(handler)) as c:
            yield c

    app.dependency_overrides[stripe_http] = mock_http
    yield calls
    app.dependency_overrides.pop(stripe_http, None)


def signed(event: dict, secret: str = "whsec_abc", t: int | None = None) -> tuple[bytes, dict]:
    payload = json.dumps(event).encode()
    t = int(time.time()) if t is None else t
    sig = hmac.new(secret.encode(), f"{t}.".encode() + payload, hashlib.sha256).hexdigest()
    return payload, {"Stripe-Signature": f"t={t},v1=deadbeef,v1={sig}", "Content-Type": "application/json"}


def test_billing_off_by_default():
    alice = signup("alice@example.com")
    a = auth(alice["token"])
    assert client.get("/api/billing/plan", headers=a).json() == {
        "plan": "free", "limit": 200, "used": 0, "renews_at": None, "billing_enabled": False}
    for path in ("/api/billing/checkout", "/api/billing/portal"):
        r = client.post(path, headers=a)
        assert r.status_code == 503 and r.json()["detail"] == "Billing isn't set up on this server."
    assert client.get("/api/billing/plan").status_code == 401


def test_billing_checkout_webhooks_and_limits(stripe):
    alice = signup("alice@example.com")
    a, uid = auth(alice["token"]), alice["user"]["id"]
    assert client.post("/api/billing/portal", headers=a).status_code == 400  # no customer yet

    r = client.post("/api/billing/checkout", headers=a)
    assert r.status_code == 200 and r.json() == {"url": "https://checkout.stripe.com/c/cs_1"}
    (m1, p1, customer), (m2, p2, session) = stripe
    assert (m1, p1, customer) == ("POST", "/v1/customers", {"email": "alice@example.com", "metadata[user_id]": uid})
    assert p2 == "/v1/checkout/sessions" and session == {
        "mode": "subscription", "customer": "cus_1", "client_reference_id": uid,
        "line_items[0][price]": "price_pro", "line_items[0][quantity]": "1",
        "success_url": "https://sheets.example.com/?billing=success",
        "cancel_url": "https://sheets.example.com/?billing=cancel", "subscription_data[metadata][user_id]": uid}
    client.post("/api/billing/checkout", headers=a)
    assert [p for _, p, _ in stripe].count("/v1/customers") == 1  # the customer is reused

    # Webhooks: bad signatures are refused.
    done = {"type": "checkout.session.completed",
            "data": {"object": {"client_reference_id": uid, "customer": "cus_1", "subscription": "sub_1"}}}
    payload, headers = signed(done, secret="wrong")
    assert client.post("/api/billing/webhook", content=payload, headers=headers).status_code == 400
    payload, headers = signed(done, t=int(time.time()) - 400)
    assert client.post("/api/billing/webhook", content=payload, headers=headers).status_code == 400
    assert client.post("/api/billing/webhook", content=payload).status_code == 400

    payload, headers = signed(done)
    assert client.post("/api/billing/webhook", content=payload, headers=headers).json() == {"received": True}
    plan = client.get("/api/billing/plan", headers=a).json()
    assert plan == {"plan": "pro", "limit": 2000, "used": 0, "renews_at": 1900000000000, "billing_enabled": True}
    assert limit_for_user(alice["user"]) == 2000 and limit_for_user(None) == 30

    portal = client.post("/api/billing/portal", headers=a).json()
    assert portal == {"url": "https://billing.stripe.com/p/1"}
    assert stripe[-1][2] == {"customer": "cus_1", "return_url": "https://sheets.example.com/"}

    # Renewal moves the date; a lapsed or deleted subscription goes back to free (found by customer id).
    sub = {"id": "sub_1", "customer": "cus_1", "status": "active", "current_period_end": 1950000000}
    payload, headers = signed({"type": "customer.subscription.updated", "data": {"object": sub}})
    client.post("/api/billing/webhook", content=payload, headers=headers)
    assert client.get("/api/billing/plan", headers=a).json()["renews_at"] == 1950000000000
    payload, headers = signed({"type": "customer.subscription.updated", "data": {"object": {**sub, "status": "unpaid"}}})
    client.post("/api/billing/webhook", content=payload, headers=headers)
    assert client.get("/api/billing/plan", headers=a).json()["plan"] == "free"
    payload, headers = signed({"type": "customer.subscription.updated", "data": {"object": sub}})
    client.post("/api/billing/webhook", content=payload, headers=headers)
    payload, headers = signed({"type": "customer.subscription.deleted", "data": {"object": sub}})
    client.post("/api/billing/webhook", content=payload, headers=headers)
    plan = client.get("/api/billing/plan", headers=a).json()
    assert plan["plan"] == "free" and plan["renews_at"] is None and plan["limit"] == 200

    # Unknown events are acknowledged and ignored.
    payload, headers = signed({"type": "invoice.paid", "data": {"object": {}}})
    assert client.post("/api/billing/webhook", content=payload, headers=headers).status_code == 200


def test_verify_signature_edge_cases():
    payload = b'{"a":1}'
    sig = hmac.new(b"k", b"100." + payload, hashlib.sha256).hexdigest()
    assert verify_signature(payload, f"t=100,v1={sig}", "k", now=100)
    assert verify_signature(payload, f"t=100,v1={sig}", "k", now=399)
    assert not verify_signature(payload, f"t=100,v1={sig}", "k", now=401)
    assert not verify_signature(payload, f"t=100,v1={sig}", "", now=100)
    assert not verify_signature(payload, "v1=abc", "k", now=100)
    assert not verify_signature(payload, "t=x,v1=abc", "k", now=100)
    assert not verify_signature(b'{"a":2}', f"t=100,v1={sig}", "k", now=100)


def test_docs_with_connections():
    from app.api.accounts import get_store

    alice = signup("alice@example.com")
    wb = tiny_workbook("=1")
    wb["connections"] = [{"id": "c1", "url": "https://example.com/x.csv", "refresh_minutes": 30}]
    put(auth(alice["token"]), wb, "live1")
    wb["connections"] = [{"id": "c2", "url": "https://example.com/y.csv", "refresh_minutes": None}]
    put(auth(alice["token"]), wb, "manual")
    put(auth(alice["token"]), tiny_workbook("=2"), "plain")
    assert get_store().docs_with_connections() == ["live1"]
