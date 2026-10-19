"""SQLite storage for signed-in users' own templates (same database file as accounts)."""

from __future__ import annotations

import json
import sqlite3
import threading
import time
import uuid
from functools import lru_cache
from pathlib import Path

DB_NAME = "sheetcraft.db"

SCHEMA = """
CREATE TABLE IF NOT EXISTS user_templates (
    id TEXT PRIMARY KEY,
    user_id TEXT,
    title TEXT,
    description TEXT,
    data TEXT,
    created REAL
);
CREATE INDEX IF NOT EXISTS user_templates_user ON user_templates (user_id, created);
"""


class TemplateLimit(ValueError):
    pass


class UserTemplates:
    def __init__(self, data_dir: str | Path):
        path = Path(data_dir)
        path.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._db = sqlite3.connect(path / DB_NAME, check_same_thread=False, timeout=10)
        self._db.row_factory = sqlite3.Row
        with self._lock:
            self._db.executescript(SCHEMA)

    def list(self, user_id: str) -> list[dict]:
        with self._lock:
            rows = self._db.execute(
                "SELECT id, title, description, data, created FROM user_templates WHERE user_id = ? "
                "ORDER BY created DESC", (user_id,),
            ).fetchall()
        return [
            {"id": r["id"], "title": r["title"], "description": r["description"],
             "sheets": len(json.loads(r["data"]).get("sheets", [])), "created": int(r["created"] * 1000)}
            for r in rows
        ]

    def get(self, user_id: str, template_id: str) -> dict | None:
        with self._lock:
            row = self._db.execute(
                "SELECT data FROM user_templates WHERE id = ? AND user_id = ?", (template_id, user_id)
            ).fetchone()
        return json.loads(row["data"]) if row else None

    def add(self, user_id: str, title: str, description: str, workbook: dict, limit: int) -> str:
        template_id = uuid.uuid4().hex
        with self._lock, self._db:
            n = self._db.execute("SELECT COUNT(*) FROM user_templates WHERE user_id = ?", (user_id,)).fetchone()[0]
            if n >= limit:
                raise TemplateLimit(limit)
            self._db.execute(
                "INSERT INTO user_templates VALUES (?, ?, ?, ?, ?, ?)",
                (template_id, user_id, title, description, json.dumps(workbook, ensure_ascii=False), time.time()),
            )
        return template_id

    def delete(self, user_id: str, template_id: str) -> bool:
        with self._lock, self._db:
            return self._db.execute(
                "DELETE FROM user_templates WHERE id = ? AND user_id = ?", (template_id, user_id)
            ).rowcount > 0


@lru_cache
def _templates_for(data_dir: str) -> UserTemplates:
    return UserTemplates(data_dir)


def get_user_templates() -> UserTemplates:
    from app.config import get_settings

    return _templates_for(get_settings().data_dir)
