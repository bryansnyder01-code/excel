"""Fan-out for live editing rooms, so several backend instances can serve the same workbook.

A broker carries room events (dicts) for a doc to every instance that has the doc open, in
one global order. Each instance applies the events, in that order, to its own copy of the
workbook and forwards them to its own WebSocket peers, so every copy converges.

* InMemoryBroker (default): one process. `publish` calls the room's handler right away, so
  behaviour is exactly the single-instance one.
* RedisBroker (settings.redis_url): one Redis Stream per doc, `sheetcraft:live:<doc>`.
  `publish` is XADD (MAXLEN ~ 10000); each subscribed instance runs XREAD BLOCK on its own
  connection and hands every entry, with its id, to the room. The entry id is the global
  order (and the room's version). RESP2 is spoken directly over asyncio streams: no client
  library needed.

Handlers are plain functions `handler(entry_id: str | None, event: dict) -> None`.
"""

from __future__ import annotations

import asyncio
import json
import logging
import ssl
import uuid
from collections.abc import Callable
from dataclasses import dataclass
from functools import lru_cache
from urllib.parse import unquote, urlparse

log = logging.getLogger(__name__)

Handler = Callable[[str | None, dict], None]

STREAM_PREFIX = "sheetcraft:live:"
STREAM_MAXLEN = 10_000
BLOCK_MS = 5_000
RECONNECT_DELAY = 1.0

# This process: events carry it as "origin" so an instance knows which peers are its own.
INSTANCE_ID = uuid.uuid4().hex


class Broker:
    ordered_ids = False  # True when handlers get stream entry ids (RedisBroker)

    async def publish(self, doc_id: str, event: dict) -> None:
        raise NotImplementedError

    async def subscribe(self, doc_id: str, handler: Handler, after: str | None = None) -> None:
        raise NotImplementedError

    async def unsubscribe(self, doc_id: str, handler: Handler | None = None) -> None:
        """Stop delivering the doc's events; with `handler`, only if it is still the subscribed one."""
        raise NotImplementedError

    async def position(self, doc_id: str) -> str | None:
        """The id of the newest event for the doc, or None (always None in memory)."""
        return None

    async def close(self) -> None:
        pass


class InMemoryBroker(Broker):
    def __init__(self) -> None:
        self._handlers: dict[str, Handler] = {}

    async def publish(self, doc_id: str, event: dict) -> None:
        handler = self._handlers.get(doc_id)
        if handler is not None:
            handler(None, event)

    async def subscribe(self, doc_id: str, handler: Handler, after: str | None = None) -> None:
        self._handlers[doc_id] = handler

    async def unsubscribe(self, doc_id: str, handler: Handler | None = None) -> None:
        if handler is None or self._handlers.get(doc_id) == handler:
            self._handlers.pop(doc_id, None)


# ---------------------------------------------------------------------------
# RESP2 over asyncio
# ---------------------------------------------------------------------------


class RespError(Exception):
    """An error reply from the server (-ERR ...)."""


@dataclass
class RedisConfig:
    host: str = "localhost"
    port: int = 6379
    username: str | None = None
    password: str | None = None
    db: int = 0
    tls: bool = False

    @classmethod
    def from_url(cls, url: str) -> RedisConfig:
        u = urlparse(url)
        if u.scheme not in ("redis", "rediss"):
            raise ValueError("redis_url must start with redis:// or rediss://")
        db = (u.path or "/").strip("/")
        return cls(host=u.hostname or "localhost", port=u.port or 6379,
                   username=unquote(u.username) if u.username else None,
                   password=unquote(u.password) if u.password else None,
                   db=int(db) if db else 0, tls=u.scheme == "rediss")


def encode_command(*args) -> bytes:
    out = [b"*%d\r\n" % len(args)]
    for a in args:
        b = a if isinstance(a, bytes) else str(a).encode()
        out.append(b"$%d\r\n%s\r\n" % (len(b), b))
    return b"".join(out)


async def read_reply(reader: asyncio.StreamReader):
    line = await reader.readline()
    if not line.endswith(b"\r\n"):
        raise ConnectionError("Redis closed the connection.")
    kind, rest = line[:1], line[1:-2]
    if kind == b"+":
        return rest.decode()
    if kind == b"-":
        raise RespError(rest.decode())
    if kind == b":":
        return int(rest)
    if kind == b"$":
        n = int(rest)
        if n < 0:
            return None
        data = await reader.readexactly(n + 2)
        return data[:-2].decode()
    if kind == b"*":
        n = int(rest)
        if n < 0:
            return None
        return [await read_reply(reader) for _ in range(n)]
    raise ConnectionError(f"Unexpected reply from Redis: {line[:40]!r}")


class RespConnection:
    def __init__(self, reader: asyncio.StreamReader, writer: asyncio.StreamWriter):
        self.reader, self.writer = reader, writer

    @classmethod
    async def open(cls, cfg: RedisConfig) -> RespConnection:
        reader, writer = await asyncio.open_connection(
            cfg.host, cfg.port, ssl=ssl.create_default_context() if cfg.tls else None)
        conn = cls(reader, writer)
        try:
            if cfg.password is not None:
                await conn.command(*(("AUTH", cfg.username, cfg.password) if cfg.username else ("AUTH", cfg.password)))
            if cfg.db:
                await conn.command("SELECT", cfg.db)
        except BaseException:
            conn.close()
            raise
        return conn

    async def command(self, *args):
        self.writer.write(encode_command(*args))
        await self.writer.drain()
        return await read_reply(self.reader)

    def close(self) -> None:
        try:
            self.writer.close()
        except Exception:  # noqa: BLE001
            pass


class RedisBroker(Broker):
    """See the module docstring. Connections belong to the event loop that opened them; if the loop
    changes (tests start a fresh one), they're reopened."""

    ordered_ids = True

    def __init__(self, url: str, maxlen: int = STREAM_MAXLEN, block_ms: int = BLOCK_MS):
        self.cfg = RedisConfig.from_url(url)
        self.maxlen, self.block_ms = maxlen, block_ms
        self._loop: asyncio.AbstractEventLoop | None = None
        self._conn: RespConnection | None = None
        self._lock: asyncio.Lock | None = None
        self._readers: dict[str, asyncio.Task] = {}
        self._handlers: dict[str, Handler] = {}

    @staticmethod
    def key(doc_id: str) -> str:
        return STREAM_PREFIX + doc_id

    def _bind_loop(self) -> None:
        loop = asyncio.get_running_loop()
        if self._loop is not loop:
            if self._conn:
                self._conn.close()
            self._loop, self._conn, self._lock, self._readers, self._handlers = loop, None, asyncio.Lock(), {}, {}

    async def _command(self, *args):
        """One command on the shared connection, reconnecting once if it dropped."""
        self._bind_loop()
        async with self._lock:
            for attempt in (1, 2):
                try:
                    if self._conn is None:
                        self._conn = await RespConnection.open(self.cfg)
                    return await self._conn.command(*args)
                except (ConnectionError, OSError, asyncio.IncompleteReadError):
                    if self._conn:
                        self._conn.close()
                    self._conn = None
                    if attempt == 2:
                        raise

    async def publish(self, doc_id: str, event: dict) -> None:
        await self._command("XADD", self.key(doc_id), "MAXLEN", "~", self.maxlen, "*",
                            "e", json.dumps(event, ensure_ascii=False))

    async def position(self, doc_id: str) -> str | None:
        reply = await self._command("XREVRANGE", self.key(doc_id), "+", "-", "COUNT", 1)
        return reply[0][0] if reply else None

    async def subscribe(self, doc_id: str, handler: Handler, after: str | None = None) -> None:
        self._bind_loop()
        await self.unsubscribe(doc_id)
        start = after or await self.position(doc_id) or "0-0"
        self._handlers[doc_id] = handler
        self._readers[doc_id] = asyncio.create_task(self._read(doc_id, handler, start))

    async def unsubscribe(self, doc_id: str, handler: Handler | None = None) -> None:
        if handler is not None and self._handlers.get(doc_id) != handler:
            return
        self._handlers.pop(doc_id, None)
        task = self._readers.pop(doc_id, None)
        if task:
            task.cancel()
            try:
                await task
            except (asyncio.CancelledError, Exception):  # noqa: BLE001
                pass

    async def _read(self, doc_id: str, handler: Handler, last: str) -> None:
        """XREAD BLOCK on a dedicated connection; every entry goes to the handler in stream order."""
        conn: RespConnection | None = None
        key = self.key(doc_id)
        try:
            while True:
                try:
                    if conn is None:
                        conn = await RespConnection.open(self.cfg)
                    reply = await conn.command("XREAD", "COUNT", 100, "BLOCK", self.block_ms, "STREAMS", key, last)
                except (ConnectionError, OSError, asyncio.IncompleteReadError, RespError):
                    log.warning("Live stream %s: Redis connection lost, retrying", key)
                    if conn:
                        conn.close()
                    conn = None
                    await asyncio.sleep(RECONNECT_DELAY)
                    continue
                for _stream, entries in reply or []:
                    for entry_id, fields in entries:
                        last = entry_id
                        data = dict(zip(fields[::2], fields[1::2]))
                        try:
                            handler(entry_id, json.loads(data.get("e") or "{}"))
                        except Exception:  # noqa: BLE001 - one bad entry must not stop the room
                            log.exception("Live stream %s: couldn't apply entry %s", key, entry_id)
        finally:
            if conn:
                conn.close()

    async def close(self) -> None:
        for doc_id in list(self._readers):
            await self.unsubscribe(doc_id)
        if self._conn:
            self._conn.close()
            self._conn = None


def version_of(entry_id: str | None) -> int:
    """A stream id as a growing integer for clients' `version` (ms * 1000 + sequence)."""
    if not entry_id:
        return 0
    ms, _, seq = entry_id.partition("-")
    return int(ms) * 1000 + min(int(seq or 0), 999)


_memory = InMemoryBroker()


@lru_cache
def _redis_broker(url: str) -> RedisBroker:
    return RedisBroker(url)


def get_broker() -> Broker:
    """RedisBroker when settings.redis_url is set, else the process-wide InMemoryBroker."""
    from app.config import get_settings

    url = getattr(get_settings(), "redis_url", "") or ""
    return _redis_broker(url) if url else _memory
