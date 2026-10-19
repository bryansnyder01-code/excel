"""Per-user / per-IP counts of AI requests over a rolling 24 hours (SQLite, same database file as accounts).

Each Claude-backed request inserts one row; rows older than the window are pruned as we go.
"""

from __future__ import annotations

import sqlite3
import threading
import time
from functools import lru_cache
from pathlib import Path

DB_NAME = "sheetcraft.db"
WINDOW = 24 * 60 * 60

SCHEMA = """
CREATE TABLE IF NOT EXISTS ai_usage (
    key TEXT,
    at REAL
);
CREATE INDEX IF NOT EXISTS ai_usage_key ON ai_usage (key, at);
"""


class UsageStore:
    def __init__(self, data_dir: str | Path):
        path = Path(data_dir)
        path.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._db = sqlite3.connect(path / DB_NAME, check_same_thread=False, timeout=10)
        with self._lock:
            self._db.executescript(SCHEMA)

    def used(self, key: str) -> int:
        with self._lock:
            row = self._db.execute(
                "SELECT COUNT(*) FROM ai_usage WHERE key = ? AND at > ?", (key, time.time() - WINDOW)
            ).fetchone()
        return row[0]

    def take(self, key: str, limit: int) -> tuple[bool, int]:
        """Count one request for `key` unless it already used `limit`. Returns (allowed, used before)."""
        now = time.time()
        with self._lock, self._db:
            self._db.execute("DELETE FROM ai_usage WHERE at <= ?", (now - WINDOW,))
            used = self._db.execute(
                "SELECT COUNT(*) FROM ai_usage WHERE key = ? AND at > ?", (key, now - WINDOW)
            ).fetchone()[0]
            if used >= limit:
                return False, used
            self._db.execute("INSERT INTO ai_usage VALUES (?, ?)", (key, now))
        return True, used


@lru_cache
def _usage_for(data_dir: str) -> UsageStore:
    return UsageStore(data_dir)


def get_usage_store() -> UsageStore:
    from app.config import get_settings

    return _usage_for(get_settings().data_dir)


def limit_for_user(user: dict | None) -> int:
    """AI requests per rolling 24 hours for this caller: anonymous (None), signed in, or signed in on Pro.
    `user` is {id, ...} as returned by Store.session_user / user_for_token."""
    from app.config import get_settings

    settings = get_settings()
    if not user:
        return settings.ai_limit_anonymous
    from app.api.accounts import get_store  # local: accounts isn't needed for plain usage counting

    try:
        plan = get_store().billing(user["id"])["plan"]
    except Exception:  # noqa: BLE001 - a billing lookup problem must never block AI requests
        plan = "free"
    return settings.ai_limit_pro if plan == "pro" else settings.ai_limit_signed_in
