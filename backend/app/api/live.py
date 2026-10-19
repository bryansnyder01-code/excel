"""Live co-editing over a WebSocket: /api/live/{doc_id}?token=<session token>.

Each instance keeps one in-memory room per open doc: its copy of the workbook (loaded from the
stored doc on first join), a version and its connected peers. Clients send small cell patches,
structural ops (insert/delete rows and columns), whole-workbook snapshots and selection
presence. The receiving instance checks a message (shape, the sender's role, locked cells),
records it in the cell history, and publishes it as an event through the broker
(app/services/broker.py). Every instance with the room open applies the broker's events in
the broker's order, relays them to its own peers and acks the sender. With the default
in-memory broker that's all immediate and in one process; with Redis it's ordered by the
doc's stream, so every instance's copy converges.

Roles: owner and editor change the workbook; a viewer gets {"type":"error"} for patch,
snapshot and ops. Non-owners can't change locked sheets or locked ranges.

Saving: every instance with the room open saves its copy (at most every SAVE_EVERY seconds
while edits happen, and when its last peer leaves). With Redis each save carries the stream
id it includes and the store ignores a save that isn't newer than the stored one, so saves
are idempotent and never go backwards; a room opened later loads the saved copy and replays
the stream from that id.

Each peer has its own send queue, so every client sees changes in order and a slow
connection never holds up the others.

Close codes: 4401 not signed in, 4404 no such doc or no access.
"""

from __future__ import annotations

import asyncio
import json
import logging
import time
import uuid
from dataclasses import dataclass, field

from fastapi import APIRouter, Depends, WebSocket, WebSocketDisconnect
from fastapi.concurrency import run_in_threadpool
from pydantic import ValidationError

from app.api.accounts import get_store
from app.schemas import WorkbookSpec
from app.services import protection
from app.services.broker import INSTANCE_ID, Broker, get_broker, version_of
from app.services.formula_shift import OPS, apply_ops, cell_name
from app.services.store import Store

log = logging.getLogger(__name__)

live_router = APIRouter(tags=["live"])

MAX_MESSAGE_BYTES = 5 * 1024 * 1024
SAVE_EVERY = 2.0  # seconds
MAX_ROW = 100_000  # grid rows a patch may reach (keeps a typo from allocating millions of rows)
MAX_COL = 1_000
MAX_OPS = 100
PALETTE = ["#E8590C", "#1C7ED6", "#2F9E44", "#AE3EC9", "#F08C00", "#0CA678", "#E64980", "#5C7CFA"]

CLOSE_UNAUTHORIZED = 4401
CLOSE_NOT_FOUND = 4404

VIEW_ONLY = "You can view this workbook but not change it."


@dataclass(eq=False)
class Peer:
    ws: WebSocket
    user: dict
    color: str
    role: str = "editor"
    sel: dict | None = None
    conn: str = field(default_factory=lambda: uuid.uuid4().hex)
    queue: asyncio.Queue = field(default_factory=asyncio.Queue)

    def info(self) -> dict:
        return {"id": self.user["id"], "email": self.user["email"], "color": self.color}

    def send(self, message: dict | str) -> None:
        self.queue.put_nowait(message if isinstance(message, str) else json.dumps(message, ensure_ascii=False))


@dataclass(eq=False)
class Room:
    doc_id: str
    store: Store
    workbook: dict
    broker: Broker
    version: int = 0
    seq: str | None = None  # last applied stream entry id (RedisBroker)
    peers: list[Peer] = field(default_factory=list)
    # Peers connected to other instances: conn id -> {"user": info, "sel": sel}.
    remote: dict[str, dict] = field(default_factory=dict)
    dirty: bool = False
    last_saved: float = 0.0
    save_task: asyncio.Task | None = None
    loop: asyncio.AbstractEventLoop | None = field(default_factory=asyncio.get_running_loop)

    def color_for(self, user_id: str) -> str:
        """A user's existing colour (second tab), else the first palette colour nobody here has."""
        everyone = [p.info() for p in self.peers] + [x["user"] for x in self.remote.values()]
        for info in everyone:
            if info["id"] == user_id:
                return info["color"]
        used = {info["color"] for info in everyone}
        free = [c for c in PALETTE if c not in used]
        return free[0] if free else PALETTE[len(everyone) % len(PALETTE)]

    def broadcast(self, message: dict, exclude: str | None = None) -> None:
        """Send to this instance's peers, except the connection `exclude`."""
        text = json.dumps(message, ensure_ascii=False)
        for p in self.peers:
            if p.conn != exclude:
                p.send(text)

    def local(self, conn: str | None) -> Peer | None:
        return next((p for p in self.peers if p.conn == conn), None)

    def peer_list(self) -> list[dict]:
        return [{**p.info(), "sel": p.sel} for p in self.peers] + \
               [{**x["user"], "sel": x["sel"]} for x in self.remote.values()]

    async def publish(self, event: dict) -> None:
        await self.broker.publish(self.doc_id, {**event, "origin": INSTANCE_ID})

    # -- applying broker events (same order on every instance) ------------------

    def apply_event(self, entry_id: str | None, ev: dict) -> None:
        kind, conn, mine = ev.get("kind"), ev.get("conn"), ev.get("origin") == INSTANCE_ID
        if kind in ("patch", "snapshot", "ops"):
            if kind == "patch":
                apply_patch(self.workbook, ev["cells"], ev["headers"])
                out = {"cells": ev["cells"], "headers": ev["headers"]}
            elif kind == "snapshot":
                self.workbook = ev["workbook"]
                out = {"workbook": ev["workbook"]}
            else:
                apply_ops(self.workbook, ev["ops"])
                out = {"ops": ev["ops"]}
            self.changed(entry_id)
            self.broadcast({"type": kind, "version": self.version, "from": ev["from"], **out}, exclude=conn)
            sender = self.local(conn) if mine else None
            if sender:
                sender.send({"type": "ack", "version": self.version})
        elif kind == "presence":
            if not mine and conn in self.remote:
                self.remote[conn]["sel"] = ev.get("sel")
            self.broadcast({"type": "presence", "user": ev["user"], "sel": ev.get("sel")}, exclude=conn)
        elif kind == "joined":
            if not mine:
                self.remote[conn] = {"user": ev["user"], "sel": None}
                # Tell the newcomer's instance who is here already.
                if self.peers:
                    here = [{"conn": p.conn, "user": p.info(), "sel": p.sel} for p in self.peers]
                    asyncio.get_running_loop().create_task(self.publish({"kind": "here", "peers": here}))
            self.broadcast({"type": "joined", "user": ev["user"]}, exclude=conn)
        elif kind == "here":
            if mine:
                return
            for x in ev.get("peers") or []:
                if x["conn"] in self.remote or self.local(x["conn"]):
                    continue
                self.remote[x["conn"]] = {"user": x["user"], "sel": x.get("sel")}
                self.broadcast({"type": "joined", "user": x["user"]})
                if x.get("sel"):
                    self.broadcast({"type": "presence", "user": x["user"], "sel": x["sel"]})
        elif kind == "left":
            self.remote.pop(conn, None)
            user_id = ev["user_id"]
            if not any(p.user["id"] == user_id for p in self.peers) and \
                    not any(x["user"]["id"] == user_id for x in self.remote.values()):
                self.broadcast({"type": "left", "user_id": user_id})
        elif kind == "comments":
            self.broadcast({"type": "comments", "thread_id": ev.get("thread_id")})
        elif kind == "role":
            for p in self.peers:
                if p.user["id"] == ev["user_id"]:
                    p.role = ev["role"]
                    p.send({"type": "role", "role": ev["role"]})

    # -- persistence ---------------------------------------------------------

    def changed(self, entry_id: str | None = None) -> None:
        if entry_id:
            self.seq, self.version = entry_id, version_of(entry_id)
        else:
            self.version += 1
        self.dirty = True
        if self.save_task is None or self.save_task.done():
            delay = max(0.0, self.last_saved + SAVE_EVERY - time.monotonic())
            self.save_task = asyncio.create_task(self._save_after(delay))

    async def _save_after(self, delay: float) -> None:
        if delay:
            await asyncio.sleep(delay)
        await self.save()

    def save_now(self) -> None:
        """Save on the calling thread (the room's final save)."""
        if not self.dirty:
            return
        self.dirty = False
        try:
            self.store.save_workbook(self.doc_id, self.workbook, time.time() * 1000, self.seq)
        except Exception:  # noqa: BLE001
            log.exception("Couldn't save live doc %s", self.doc_id)

    async def save(self) -> None:
        if not self.dirty:
            return
        self.dirty = False
        self.last_saved = time.monotonic()
        workbook = json.loads(json.dumps(self.workbook))  # a copy: edits keep landing while the thread writes
        try:
            await run_in_threadpool(self.store.save_workbook, self.doc_id, workbook, time.time() * 1000, self.seq)
        except Exception:  # noqa: BLE001 - a failed save must not drop the connection
            log.exception("Couldn't save live doc %s", self.doc_id)


rooms: dict[str, Room] = {}


def _usable(room: Room | None, store: Store) -> Room | None:
    """An open room for this store and event loop (a room left over from a stopped loop, as when a
    server or a test client shut down mid-close, is stale and gets replaced)."""
    if room and room.store is store and room.loop is asyncio.get_running_loop():
        return room
    return None


async def _join_room(store: Store, doc_id: str) -> Room | None:
    room = _usable(rooms.get(doc_id), store)
    if room:
        return room
    broker = get_broker()
    position = await broker.position(doc_id)  # before loading, so nothing published meanwhile is missed
    doc = await run_in_threadpool(store.live_doc, doc_id)
    if doc is None:
        return None
    room = _usable(rooms.get(doc_id), store)  # someone else may have opened it while we were loading
    if room:
        return room
    workbook = doc["data"].get("workbook") if isinstance(doc["data"], dict) else None
    if not isinstance(workbook, dict):
        workbook = {"title": doc["title"], "sheets": []}
    room = rooms[doc_id] = Room(doc_id, store, workbook, broker)
    start = None
    if broker.ordered_ids:
        start = doc["seq"] or position or "0-0"
        room.seq = doc["seq"]
        room.version = version_of(doc["seq"] or position)
    await broker.subscribe(doc_id, room.apply_event, start)
    return room


async def notify(doc_id: str, event: dict) -> None:
    """Send an event (e.g. {"kind": "comments", "thread_id": ...}) to the doc's live room on every instance.
    Errors are logged: callers have already done their real work."""
    try:
        await get_broker().publish(doc_id, {**event, "origin": INSTANCE_ID})
    except Exception:  # noqa: BLE001
        log.exception("Couldn't notify live room %s", doc_id)


def local_workbook(doc_id: str) -> dict | None:
    """This instance's live copy of a doc's workbook, if the room is open here."""
    room = rooms.get(doc_id)
    return room.workbook if room else None


# ---------------------------------------------------------------------------
# Checking and applying changes
# ---------------------------------------------------------------------------


class BadMessage(ValueError):
    pass


def _int(v, lo: int, hi: int, what: str) -> int:
    if not isinstance(v, int) or isinstance(v, bool) or not lo <= v <= hi:
        raise BadMessage(f"{what} must be a whole number from {lo} to {hi}.")
    return v


def _clean_patch(msg: dict) -> tuple[list[dict], list[dict]]:
    cells, headers = msg.get("cells") or [], msg.get("headers") or []
    if not isinstance(cells, list) or not isinstance(headers, list):
        raise BadMessage("cells and headers must be lists.")
    out_cells, out_headers = [], []
    for c in cells:
        if not isinstance(c, dict) or not isinstance(c.get("sheet"), str):
            raise BadMessage("Each cell change needs a sheet name.")
        cell = c.get("cell")
        if cell is not None and not isinstance(cell, dict):
            raise BadMessage("A cell must be an object or null.")
        out_cells.append({"sheet": c["sheet"], "r": _int(c.get("r"), 0, MAX_ROW, "r"),
                          "c": _int(c.get("c"), 0, MAX_COL, "c"), "cell": cell})
    for h in headers:
        if not isinstance(h, dict) or not isinstance(h.get("sheet"), str) or not isinstance(h.get("header"), str):
            raise BadMessage("Each header change needs a sheet name and header text.")
        out_headers.append({"sheet": h["sheet"], "c": _int(h.get("c"), 0, MAX_COL, "c"), "header": h["header"]})
    return out_cells, out_headers


def _clean_ops(msg: dict, workbook: dict) -> list[dict]:
    ops = msg.get("ops")
    if not isinstance(ops, list) or not ops:
        raise BadMessage("ops must be a non-empty list.")
    if len(ops) > MAX_OPS:
        raise BadMessage(f"Send at most {MAX_OPS} ops at a time.")
    names = {s.get("name") for s in workbook.get("sheets") or [] if isinstance(s, dict)}
    out = []
    for o in ops:
        if not isinstance(o, dict) or o.get("op") not in OPS:
            raise BadMessage("Each op must be insert_rows, delete_rows, insert_cols or delete_cols.")
        if not isinstance(o.get("sheet"), str) or o["sheet"] not in names:
            raise BadMessage("Each op needs the name of a sheet in this workbook.")
        rows = o["op"].endswith("_rows")
        at = _int(o.get("at"), 1 if rows else 0, MAX_ROW if rows else MAX_COL, "at")
        count = _int(o.get("count"), 1, MAX_ROW if rows else MAX_COL, "count")
        out.append({"op": o["op"], "sheet": o["sheet"], "at": at, "count": count})
    return out


def _sheet(workbook: dict, name: str) -> dict | None:
    for s in workbook.get("sheets") or []:
        if isinstance(s, dict) and s.get("name") == name:
            return s
    return None


def _ensure_columns(sheet: dict, n: int) -> list:
    columns = sheet.setdefault("columns", [])
    while len(columns) < n:
        columns.append({"header": ""})
    return columns


def apply_patch(workbook: dict, cells: list[dict], headers: list[dict]) -> None:
    """Grid row r >= 1 is rows[r-1]; rows, cells and columns grow as needed. Unknown sheets are skipped."""
    for h in headers:
        sheet = _sheet(workbook, h["sheet"])
        if sheet is not None:
            col = _ensure_columns(sheet, h["c"] + 1)[h["c"]]
            if isinstance(col, dict):
                col["header"] = h["header"]
    for ch in cells:
        sheet = _sheet(workbook, ch["sheet"])
        if sheet is None or ch["r"] < 1:
            continue
        _ensure_columns(sheet, ch["c"] + 1)
        rows = sheet.setdefault("rows", [])
        while len(rows) < ch["r"]:
            rows.append({"cells": []})
        row = rows[ch["r"] - 1]
        if not isinstance(row, dict):
            row = rows[ch["r"] - 1] = {"cells": []}
        row_cells = row.setdefault("cells", [])
        while len(row_cells) <= ch["c"]:
            row_cells.append({"value": None})
        row_cells[ch["c"]] = ch["cell"] if ch["cell"] is not None else {"value": None}


def _clean_sel(sel) -> dict | None:
    if sel is None:
        return None
    if not isinstance(sel, dict) or not isinstance(sel.get("sheet"), str):
        raise BadMessage("sel needs a sheet name.")
    out = {"sheet": sel["sheet"]}
    for k in ("r1", "c1", "r2", "c2"):
        out[k] = _int(sel.get(k), 0, 1_048_576, k)
    return out


# -- history -------------------------------------------------------------------


def _short(value) -> str:
    text = value if isinstance(value, str) else json.dumps(value, ensure_ascii=False)
    text = text if len(text) <= 60 else text[:57] + "..."
    return f'"{text}"' if isinstance(value, str) and not value.startswith("=") else text


def _edit_summary(sheet: str, r: int, c: int, cell: dict | None) -> str:
    where = f"{cell_name(r, c)} on {sheet}"
    if not cell:
        return f"Cleared {where}"
    value = cell.get("value")
    if value is None or value == "":
        styled = any(v not in (None, False) for k, v in cell.items() if k != "value")
        return f"Formatted {where}" if styled else f"Cleared {where}"
    return f"Set {where} to {_short(value)}"


def _ops_summary(ops: list[dict]) -> str:
    parts = []
    for o in ops:
        n, at = o["count"], o["at"]
        if o["op"].endswith("_rows"):
            what, where = f"{n} row{'s' if n > 1 else ''}", f"row {at + 1}"  # grid row at = Excel row at + 1
        else:
            what, where = f"{n} column{'s' if n > 1 else ''}", f"column {cell_name(0, at)[:-1]}"
        verb = "Inserted" if o["op"].startswith("insert") else "Deleted"
        parts.append(f"{verb} {what} at {where} on {o['sheet']}")
    return "; ".join(parts)


def _record(room: Room, user_id: str, entries: list[dict]) -> None:
    try:
        room.store.add_history(room.doc_id, user_id, entries)
    except Exception:  # noqa: BLE001 - history must never block an edit
        log.exception("Couldn't record history for %s", room.doc_id)


async def handle(room: Room, peer: Peer, raw: str) -> None:
    """Check one client message and publish it. Raises BadMessage (nothing is changed) when it is
    malformed, the sender is a viewer, or it touches locked cells and the sender isn't the owner."""
    if len(raw) > MAX_MESSAGE_BYTES:
        raise BadMessage("That change is too large (5 MB max).")
    try:
        msg = json.loads(raw)
    except ValueError:
        raise BadMessage("That message isn't valid JSON.") from None
    if not isinstance(msg, dict):
        raise BadMessage("Messages must be JSON objects.")
    kind, me = msg.get("type"), peer.user["id"]
    if kind in ("patch", "snapshot", "ops") and peer.role == "viewer":
        raise BadMessage(VIEW_ONLY)
    guarded = peer.role != "owner"
    if kind == "patch":
        cells, headers = _clean_patch(msg)
        if guarded and (problem := protection.check_patch(room.workbook, cells, headers)):
            raise BadMessage(problem)
        _record(room, me, [
            *({"sheet": h["sheet"], "r": 0, "c": h["c"], "value": {"value": h["header"]},
               "summary": _edit_summary(h["sheet"], 0, h["c"], {"value": h["header"]})} for h in headers),
            *({"sheet": c["sheet"], "r": c["r"], "c": c["c"], "value": c["cell"],
               "summary": _edit_summary(c["sheet"], c["r"], c["c"], c["cell"])} for c in cells),
        ])
        await room.publish({"kind": "patch", "conn": peer.conn, "from": me, "cells": cells, "headers": headers})
    elif kind == "ops":
        ops = _clean_ops(msg, room.workbook)
        if guarded and (problem := protection.check_ops(room.workbook, ops)):
            raise BadMessage(problem)
        _record(room, me, [{"sheet": ops[0]["sheet"], "kind": "restructure", "summary": _ops_summary(ops)}])
        await room.publish({"kind": "ops", "conn": peer.conn, "from": me, "ops": ops})
    elif kind == "snapshot":
        workbook = msg.get("workbook")
        try:
            WorkbookSpec.model_validate(workbook)
        except ValidationError:
            raise BadMessage("That workbook couldn't be read.") from None
        if guarded and (problem := protection.check_replace(room.workbook, workbook)):
            raise BadMessage(problem)
        _record(room, me, [{"kind": "replace", "summary": "Replaced the whole workbook"}])
        await room.publish({"kind": "snapshot", "conn": peer.conn, "from": me, "workbook": workbook})
    elif kind == "presence":
        peer.sel = _clean_sel(msg.get("sel"))
        await room.publish({"kind": "presence", "conn": peer.conn, "user": peer.info(), "sel": peer.sel})
    else:
        raise BadMessage("Unknown message type.")


# ---------------------------------------------------------------------------
# The socket
# ---------------------------------------------------------------------------


async def _pump(peer: Peer) -> None:
    """Send queued messages to one client, in order."""
    try:
        while True:
            await peer.ws.send_text(await peer.queue.get())
    except Exception:  # noqa: BLE001 - the socket closed; the receive loop cleans up
        pass


@live_router.websocket("/api/live/{doc_id}")
async def live(websocket: WebSocket, doc_id: str, token: str = "", store: Store = Depends(get_store)):
    # Accept before closing so browsers see the close code (a refused handshake shows only 1006).
    await websocket.accept()
    user = await run_in_threadpool(store.user_for_token, token)
    if not user:
        await websocket.close(CLOSE_UNAUTHORIZED, "Please sign in again.")
        return
    role = await run_in_threadpool(store.doc_role, user["id"], doc_id)
    room = await _join_room(store, doc_id) if role else None
    if room is None:
        await websocket.close(CLOSE_NOT_FOUND, "Workbook not found.")
        return

    peer = Peer(websocket, user, room.color_for(user["id"]), role)
    peer.send({"type": "state", "version": room.version, "workbook": room.workbook, "you": peer.info(),
               "peers": room.peer_list(), "role": role})
    room.peers.append(peer)
    sender = asyncio.create_task(_pump(peer))
    try:
        await room.publish({"kind": "joined", "conn": peer.conn, "user": peer.info()})
        while True:
            message = await websocket.receive()
            if message["type"] == "websocket.disconnect":
                break
            try:
                if message.get("text") is None:
                    raise BadMessage("Send JSON text frames.")
                await handle(room, peer, message["text"])
            except BadMessage as e:
                peer.send({"type": "error", "message": str(e)})
    except (WebSocketDisconnect, RuntimeError):
        pass
    finally:
        room.peers.remove(peer)
        sender.cancel()
        last = not room.peers
        if last:
            # The last peer here: close the room before awaiting anything, so a rejoin (or a server
            # shutting down) never sees a half-closed room, and save synchronously (one small SQLite write).
            if room.save_task and not room.save_task.done():
                room.save_task.cancel()
            if rooms.get(doc_id) is room:
                del rooms[doc_id]
            room.save_now()
        try:
            await room.publish({"kind": "left", "conn": peer.conn, "user_id": user["id"]})
            if last:
                await room.broker.unsubscribe(doc_id, room.apply_event)
        except Exception:  # noqa: BLE001
            log.exception("Couldn't announce that a peer left %s", doc_id)
