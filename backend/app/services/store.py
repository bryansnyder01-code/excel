"""SQLite storage for accounts, sessions, synced workbooks, share links and collaborators (stdlib only).

One connection per Store, shared across threads behind a lock: requests are short and
SQLite serialises writes anyway. The database file is created on first use.

Security notes
* Passwords are hashed with scrypt (n=2**14, r=8, p=1) and a random 16-byte salt.
* Session tokens and password reset tokens are random; only their SHA-256 is stored,
  so a leaked database can't be used to sign in.
* Sessions slide: each use pushes the expiry out (last_used is written at most hourly).
* Accounts created through Google sign-in have an empty pw_hash, which never matches a
  password, until the owner sets one with a reset link.
"""

from __future__ import annotations

import hashlib
import hmac
import json
import secrets
import sqlite3
import threading
import time
import uuid
from collections import defaultdict, deque
from pathlib import Path

DB_NAME = "sheetcraft.db"

SCRYPT = dict(n=2**14, r=8, p=1, dklen=32)

# Login throttle: more than this many failures for one email inside the window locks it out.
MAX_FAILED_LOGINS = 5
FAILED_LOGIN_WINDOW = 15 * 60

# Password reset: links last an hour; at most this many emails per address per hour.
RESET_TTL = 60 * 60
MAX_RESET_REQUESTS = 3
TOUCH_EVERY = 60 * 60  # seconds between last_used writes for one session
HISTORY_KEEP = 5000  # cell history entries kept per doc

SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id TEXT PRIMARY KEY,
    email TEXT UNIQUE COLLATE NOCASE,
    pw_hash TEXT,
    created REAL
);
CREATE TABLE IF NOT EXISTS sessions (
    token_hash TEXT PRIMARY KEY,
    user_id TEXT,
    created REAL,
    last_used REAL
);
CREATE TABLE IF NOT EXISTS docs (
    id TEXT,
    user_id TEXT,
    title TEXT,
    updated REAL,
    data TEXT,
    PRIMARY KEY (id)
);
CREATE TABLE IF NOT EXISTS shares (
    token TEXT PRIMARY KEY,
    doc_id TEXT,
    created REAL
);
CREATE TABLE IF NOT EXISTS resets (
    token_hash TEXT PRIMARY KEY,
    user_id TEXT,
    created REAL,
    used INTEGER DEFAULT 0
);
CREATE TABLE IF NOT EXISTS invites (
    token TEXT PRIMARY KEY,
    doc_id TEXT,
    created REAL
);
CREATE TABLE IF NOT EXISTS members (
    doc_id TEXT,
    user_id TEXT,
    role TEXT,
    created REAL,
    PRIMARY KEY (doc_id, user_id)
);
CREATE TABLE IF NOT EXISTS cell_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    doc_id TEXT,
    sheet TEXT,
    r INTEGER,
    c INTEGER,
    user_id TEXT,
    at INTEGER,
    value TEXT,
    kind TEXT DEFAULT 'edit',
    summary TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS comment_threads (
    id TEXT PRIMARY KEY,
    doc_id TEXT,
    sheet TEXT,
    cell TEXT,
    resolved INTEGER DEFAULT 0,
    created INTEGER
);
CREATE TABLE IF NOT EXISTS comments (
    id TEXT PRIMARY KEY,
    thread_id TEXT,
    doc_id TEXT,
    author_id TEXT,
    body TEXT,
    created INTEGER,
    edited INTEGER
);
CREATE TABLE IF NOT EXISTS live_state (
    doc_id TEXT PRIMARY KEY,
    seq TEXT
);
CREATE INDEX IF NOT EXISTS history_doc ON cell_history (doc_id, id);
CREATE INDEX IF NOT EXISTS history_cell ON cell_history (doc_id, sheet, r, c, id);
CREATE INDEX IF NOT EXISTS threads_doc ON comment_threads (doc_id, created);
CREATE INDEX IF NOT EXISTS comments_thread ON comments (thread_id, created);
CREATE INDEX IF NOT EXISTS docs_user ON docs (user_id, updated);
CREATE INDEX IF NOT EXISTS shares_doc ON shares (doc_id);
CREATE INDEX IF NOT EXISTS invites_doc ON invites (doc_id);
CREATE INDEX IF NOT EXISTS members_user ON members (user_id);
CREATE INDEX IF NOT EXISTS sessions_user ON sessions (user_id);
"""


class EmailTaken(ValueError):
    pass


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.scrypt(password.encode(), salt=salt, **SCRYPT)
    return f"scrypt${salt.hex()}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        _, salt_hex, digest_hex = stored.split("$")
        digest = hashlib.scrypt(password.encode(), salt=bytes.fromhex(salt_hex), **SCRYPT)
    except ValueError:
        return False
    return hmac.compare_digest(digest.hex(), digest_hex)


# Compared against when the email is unknown, so a miss takes as long as a wrong password.
_DUMMY_HASH = hash_password(secrets.token_urlsafe(16))


def _token_hash(token: str) -> str:
    return hashlib.sha256(token.encode()).hexdigest()


class Store:
    def __init__(self, data_dir: str | Path, session_days: float = 30):
        path = Path(data_dir)
        path.mkdir(parents=True, exist_ok=True)
        self.session_ttl = session_days * 24 * 60 * 60
        self._lock = threading.Lock()
        self._db = sqlite3.connect(path / DB_NAME, check_same_thread=False)
        self._db.row_factory = sqlite3.Row
        with self._lock, self._db:
            self._db.executescript(SCHEMA)
            # Databases from before sliding sessions: count their sessions as used when created.
            if "last_used" not in {r["name"] for r in self._db.execute("PRAGMA table_info(sessions)")}:
                self._db.execute("ALTER TABLE sessions ADD COLUMN last_used REAL")
            self._db.execute("UPDATE sessions SET last_used = created WHERE last_used IS NULL")
            # Later additions: invite roles, and billing fields on users.
            for table, column, decl in (("invites", "role", "TEXT DEFAULT 'editor'"),
                                        ("users", "plan", "TEXT DEFAULT 'free'"),
                                        ("users", "stripe_customer_id", "TEXT"),
                                        ("users", "plan_renews_at", "REAL")):
                if column not in {r["name"] for r in self._db.execute(f"PRAGMA table_info({table})")}:
                    self._db.execute(f"ALTER TABLE {table} ADD COLUMN {column} {decl}")
            self._db.execute("CREATE INDEX IF NOT EXISTS users_customer ON users (stripe_customer_id)")
        # email -> timestamps of recent failed logins / reset emails (in memory: resets on
        # restart, which is fine).
        self._failures: dict[str, deque[float]] = defaultdict(deque)
        self._reset_requests: dict[str, deque[float]] = defaultdict(deque)

    def _run(self, sql: str, args: tuple = ()) -> list[sqlite3.Row]:
        with self._lock, self._db:  # `with db` commits, or rolls back on error
            return self._db.execute(sql, args).fetchall()

    # -- users and sessions --------------------------------------------------

    def create_user(self, email: str, password: str) -> dict:
        user_id = uuid.uuid4().hex
        try:
            self._run("INSERT INTO users (id, email, pw_hash, created) VALUES (?, ?, ?, ?)", (user_id, email, hash_password(password), time.time()))
        except sqlite3.IntegrityError as e:
            raise EmailTaken(email) from e
        return {"id": user_id, "email": email}

    def user_by_email(self, email: str, create: bool = False) -> dict | None:
        """The account for an email; with create=True, make a password-less one if there is none."""
        rows = self._run("SELECT id, email FROM users WHERE email = ?", (email,))
        if rows:
            return dict(rows[0])
        if not create:
            return None
        user_id = uuid.uuid4().hex
        self._run("INSERT OR IGNORE INTO users (id, email, pw_hash, created) VALUES (?, ?, '', ?)", (user_id, email, time.time()))
        return self.user_by_email(email)

    def check_login(self, email: str, password: str) -> dict | None:
        rows = self._run("SELECT id, email, pw_hash FROM users WHERE email = ?", (email,))
        if not rows:
            verify_password(password, _DUMMY_HASH)
            return None
        if not rows[0]["pw_hash"]:  # signed up with Google and never set a password
            verify_password(password, _DUMMY_HASH)
            return None
        if not verify_password(password, rows[0]["pw_hash"]):
            return None
        return {"id": rows[0]["id"], "email": rows[0]["email"]}

    def login_locked(self, email: str) -> bool:
        with self._lock:
            recent = self._failures[email]
            while recent and recent[0] < time.time() - FAILED_LOGIN_WINDOW:
                recent.popleft()
            return len(recent) >= MAX_FAILED_LOGINS

    def record_login(self, email: str, ok: bool) -> None:
        with self._lock:
            if ok:
                self._failures.pop(email, None)
            else:
                self._failures[email].append(time.time())

    def create_session(self, user_id: str) -> str:
        token = secrets.token_urlsafe(32)
        now = time.time()
        self._run("INSERT INTO sessions (token_hash, user_id, created, last_used) VALUES (?, ?, ?, ?)",
                  (_token_hash(token), user_id, now, now))
        return token

    def session_user(self, token: str) -> dict | None:
        """The signed-in user, or None when the token is unknown or unused for session_days."""
        token_hash, now = _token_hash(token), time.time()
        with self._lock, self._db:
            row = self._db.execute(
                "SELECT u.id, u.email, s.last_used FROM sessions s JOIN users u ON u.id = s.user_id "
                "WHERE s.token_hash = ?", (token_hash,),
            ).fetchone()
            if not row:
                return None
            if now - (row["last_used"] or 0) > self.session_ttl:
                self._db.execute("DELETE FROM sessions WHERE token_hash = ?", (token_hash,))
                return None
            if now - (row["last_used"] or 0) > TOUCH_EVERY:
                self._db.execute("UPDATE sessions SET last_used = ? WHERE token_hash = ?", (now, token_hash))
        return {"id": row["id"], "email": row["email"]}

    def user_for_token(self, token: str | None) -> dict | None:
        """Resolve a bearer token to {id, email}; None for a missing, unknown or expired one. Never raises."""
        if not token or not isinstance(token, str):
            return None
        try:
            return self.session_user(token)
        except Exception:  # noqa: BLE001 - callers only want "signed in or not"
            return None

    def delete_session(self, token: str) -> None:
        self._run("DELETE FROM sessions WHERE token_hash = ?", (_token_hash(token),))

    # -- password reset ------------------------------------------------------

    def allow_reset_request(self, email: str) -> bool:
        """Counts a reset email for this address; False once it has had MAX_RESET_REQUESTS this hour."""
        with self._lock:
            recent = self._reset_requests[email]
            while recent and recent[0] < time.time() - RESET_TTL:
                recent.popleft()
            if len(recent) >= MAX_RESET_REQUESTS:
                return False
            recent.append(time.time())
            return True

    def create_reset(self, user_id: str) -> str:
        token = secrets.token_urlsafe(32)
        self._run("INSERT INTO resets VALUES (?, ?, ?, 0)", (_token_hash(token), user_id, time.time()))
        return token

    def reset_password(self, token: str, password: str) -> dict | None:
        """Use a reset token: set the password and end every session. None if the token is bad or spent."""
        pw_hash = hash_password(password)
        with self._lock, self._db:
            row = self._db.execute(
                "SELECT r.user_id, u.email FROM resets r JOIN users u ON u.id = r.user_id "
                "WHERE r.token_hash = ? AND r.used = 0 AND r.created > ?",
                (_token_hash(token), time.time() - RESET_TTL),
            ).fetchone()
            if not row:
                return None
            user_id = row["user_id"]
            self._db.execute("UPDATE resets SET used = 1 WHERE user_id = ?", (user_id,))
            self._db.execute("UPDATE users SET pw_hash = ? WHERE id = ?", (pw_hash, user_id))
            self._db.execute("DELETE FROM sessions WHERE user_id = ?", (user_id,))
        return {"id": user_id, "email": row["email"]}

    # -- docs ------------------------------------------------------------------

    def list_docs(self, user_id: str) -> list[dict]:
        """Docs the user owns or edits, newest first."""
        rows = self._run(
            "SELECT d.id, d.title, d.updated, d.data, u.email AS owner_email, "
            "CASE WHEN d.user_id = ? THEN 'owner' ELSE (SELECT m.role FROM members m WHERE m.doc_id = d.id AND m.user_id = ?) END AS role, "
            "EXISTS (SELECT 1 FROM shares WHERE doc_id = d.id) AS shared "
            "FROM docs d LEFT JOIN users u ON u.id = d.user_id "
            "WHERE d.user_id = ? OR d.id IN (SELECT doc_id FROM members WHERE user_id = ?) "
            "ORDER BY d.updated DESC",
            (user_id, user_id, user_id, user_id),
        )
        return [
            {"id": r["id"], "title": r["title"], "updated": r["updated"], "sheets": _sheet_count(r["data"]),
             "shared": bool(r["shared"]), "role": r["role"], "owner_email": r["owner_email"]}
            for r in rows
        ]

    def doc_role(self, user_id: str, doc_id: str) -> str | None:
        """'owner', 'editor', 'viewer', or None when the user can't open the doc (or it doesn't exist)."""
        rows = self._run(
            "SELECT CASE WHEN d.user_id = ? THEN 'owner' ELSE m.role END AS role FROM docs d "
            "LEFT JOIN members m ON m.doc_id = d.id AND m.user_id = ? WHERE d.id = ?",
            (user_id, user_id, doc_id),
        )
        return rows[0]["role"] if rows else None

    def get_doc(self, user_id: str, doc_id: str) -> dict | None:
        if not self.doc_role(user_id, doc_id):
            return None
        rows = self._run("SELECT id, title, updated, data FROM docs WHERE id = ?", (doc_id,))
        if not rows:
            return None
        r = rows[0]
        return {"id": r["id"], "title": r["title"], "updated": r["updated"], "data": json.loads(r["data"])}

    def doc_owner(self, doc_id: str) -> str | None:
        rows = self._run("SELECT user_id FROM docs WHERE id = ?", (doc_id,))
        return rows[0]["user_id"] if rows else None

    def put_doc(self, user_id: str, doc_id: str, title: str, updated: float, data: dict) -> bool:
        """Create or replace (last write wins). Editors write too; the doc stays the owner's.
        False when the id belongs to someone who hasn't shared it with this user, or who made them a viewer."""
        with self._lock, self._db:
            owner = self._db.execute("SELECT user_id FROM docs WHERE id = ?", (doc_id,)).fetchone()
            if owner and owner["user_id"] != user_id:
                member = self._db.execute(
                    "SELECT role FROM members WHERE doc_id = ? AND user_id = ?", (doc_id, user_id)).fetchone()
                if not member or member["role"] != "editor":
                    return False
                user_id = owner["user_id"]
            self._db.execute(
                "INSERT OR REPLACE INTO docs VALUES (?, ?, ?, ?, ?)",
                (doc_id, user_id, title, updated, json.dumps(data, ensure_ascii=False)),
            )
        return True

    def docs_with_connections(self) -> list[str]:
        """Ids of docs whose stored workbook has a data connection with refresh_minutes set
        (the scheduler refreshes those). A LIKE prefilter keeps the JSON parsing to likely docs."""
        rows = self._run("SELECT id, data FROM docs WHERE data LIKE '%refresh_minutes%'")
        out = []
        for r in rows:
            try:
                wb = json.loads(r["data"]).get("workbook") or {}
                conns = wb.get("connections") or []
            except (ValueError, AttributeError):
                continue
            if any(isinstance(c, dict) and c.get("refresh_minutes") for c in conns if isinstance(conns, list)):
                out.append(r["id"])
        return out

    def live_doc(self, doc_id: str) -> dict | None:
        """{title, data, seq} read together, for loading a live room: `seq` is the live stream position
        the stored workbook includes (None when it was never saved from a shared stream)."""
        rows = self._run("SELECT d.title, d.data, l.seq FROM docs d LEFT JOIN live_state l ON l.doc_id = d.id "
                         "WHERE d.id = ?", (doc_id,))
        if not rows:
            return None
        return {"title": rows[0]["title"], "data": json.loads(rows[0]["data"]), "seq": rows[0]["seq"]}

    def save_workbook(self, doc_id: str, workbook: dict, updated: float, seq: str | None = None) -> bool:
        """Write a live-edited workbook into an existing doc's data.workbook (title follows it).
        With `seq` (a live stream entry id), the write only happens when it is newer than the stored one,
        so several backend instances saving the same room never go backwards."""
        with self._lock, self._db:
            row = self._db.execute("SELECT title, data FROM docs WHERE id = ?", (doc_id,)).fetchone()
            if not row:
                return False
            if seq is not None:
                stored = self._db.execute("SELECT seq FROM live_state WHERE doc_id = ?", (doc_id,)).fetchone()
                if stored and stored["seq"] and stream_id_key(stored["seq"]) >= stream_id_key(seq):
                    return False
                self._db.execute("INSERT OR REPLACE INTO live_state VALUES (?, ?)", (doc_id, seq))
            try:
                data = json.loads(row["data"])
            except ValueError:
                data = {}
            if not isinstance(data, dict):
                data = {}
            title = workbook.get("title") if isinstance(workbook.get("title"), str) else row["title"]
            data["workbook"] = workbook
            if "title" in data:
                data["title"] = title
            if "updated" in data:
                data["updated"] = updated
            self._db.execute("UPDATE docs SET title = ?, updated = ?, data = ? WHERE id = ?",
                             (title, updated, json.dumps(data, ensure_ascii=False), doc_id))
        return True

    def delete_doc(self, user_id: str, doc_id: str) -> None:
        """The owner deletes the doc for everyone; an editor or viewer only leaves it."""
        with self._lock, self._db:
            n = self._db.execute("DELETE FROM docs WHERE id = ? AND user_id = ?", (doc_id, user_id)).rowcount
            if n:
                for table in ("shares", "invites", "members", "cell_history", "comment_threads", "comments",
                              "live_state"):
                    self._db.execute(f"DELETE FROM {table} WHERE doc_id = ?", (doc_id,))
            else:
                self._db.execute("DELETE FROM members WHERE doc_id = ? AND user_id = ?", (doc_id, user_id))

    # -- share links -------------------------------------------------------

    def share_doc(self, doc_id: str) -> str:
        with self._lock, self._db:
            row = self._db.execute("SELECT token FROM shares WHERE doc_id = ?", (doc_id,)).fetchone()
            if row:
                return row["token"]
            token = secrets.token_urlsafe(16)
            self._db.execute("INSERT INTO shares VALUES (?, ?, ?)", (token, doc_id, time.time()))
            return token

    def unshare_doc(self, doc_id: str) -> None:
        self._run("DELETE FROM shares WHERE doc_id = ?", (doc_id,))

    def shared_doc(self, token: str) -> dict | None:
        rows = self._run(
            "SELECT d.title, d.data FROM shares s JOIN docs d ON d.id = s.doc_id WHERE s.token = ?", (token,)
        )
        if not rows:
            return None
        return {"title": rows[0]["title"], "data": json.loads(rows[0]["data"])}


    # -- edit invites and members -------------------------------------------

    def invite_doc(self, doc_id: str, role: str = "editor") -> str:
        """The doc's invite token for `role` ('editor' or 'viewer'); one active token per role."""
        with self._lock, self._db:
            row = self._db.execute("SELECT token FROM invites WHERE doc_id = ? AND COALESCE(role, 'editor') = ?",
                                   (doc_id, role)).fetchone()
            if row:
                return row["token"]
            token = secrets.token_urlsafe(16)
            self._db.execute("INSERT INTO invites (token, doc_id, created, role) VALUES (?, ?, ?, ?)",
                             (token, doc_id, time.time(), role))
            return token

    def uninvite_doc(self, doc_id: str, role: str | None = None) -> None:
        if role:
            self._run("DELETE FROM invites WHERE doc_id = ? AND COALESCE(role, 'editor') = ?", (doc_id, role))
        else:
            self._run("DELETE FROM invites WHERE doc_id = ?", (doc_id,))

    def invites(self, doc_id: str) -> dict:
        """{editor: token | None, viewer: token | None}."""
        out = {"editor": None, "viewer": None}
        for r in self._run("SELECT token, COALESCE(role, 'editor') AS role FROM invites WHERE doc_id = ?", (doc_id,)):
            if r["role"] in out:
                out[r["role"]] = r["token"]
        return out

    def accept_invite(self, token: str, user_id: str) -> dict | None:
        """Give the user the invite's role on its doc. An editor link upgrades a viewer; a viewer link
        never downgrades an editor. Returns {id, title, role}, or None for a bad token."""
        with self._lock, self._db:
            row = self._db.execute(
                "SELECT d.id, d.title, d.user_id, COALESCE(i.role, 'editor') AS role FROM invites i "
                "JOIN docs d ON d.id = i.doc_id WHERE i.token = ?",
                (token,),
            ).fetchone()
            if not row:
                return None
            if row["user_id"] == user_id:
                return {"id": row["id"], "title": row["title"], "role": "owner"}
            current = self._db.execute("SELECT role FROM members WHERE doc_id = ? AND user_id = ?",
                                       (row["id"], user_id)).fetchone()
            role = row["role"]
            if current is None:
                self._db.execute("INSERT INTO members VALUES (?, ?, ?, ?)", (row["id"], user_id, role, time.time()))
            elif current["role"] == "editor" or current["role"] == role:
                role = current["role"]
            else:
                self._db.execute("UPDATE members SET role = ? WHERE doc_id = ? AND user_id = ?",
                                 (role, row["id"], user_id))
        return {"id": row["id"], "title": row["title"], "role": role}

    def members(self, doc_id: str) -> list[dict]:
        """The owner first, then editors and viewers in the order they joined."""
        rows = self._run(
            "SELECT u.id, u.email, 'owner' AS role, 0 AS ord, 0 AS joined FROM docs d JOIN users u ON u.id = d.user_id "
            "WHERE d.id = ? UNION ALL "
            "SELECT u.id, u.email, m.role, 1, m.created FROM members m JOIN users u ON u.id = m.user_id "
            "WHERE m.doc_id = ? ORDER BY ord, joined",
            (doc_id, doc_id),
        )
        return [{"id": r["id"], "email": r["email"], "role": r["role"]} for r in rows]

    def set_member_role(self, doc_id: str, user_id: str, role: str) -> bool:
        """Change a member's role ('editor' / 'viewer'). False when they aren't a member."""
        with self._lock, self._db:
            return self._db.execute("UPDATE members SET role = ? WHERE doc_id = ? AND user_id = ?",
                                    (role, doc_id, user_id)).rowcount > 0

    def remove_member(self, doc_id: str, user_id: str) -> None:
        self._run("DELETE FROM members WHERE doc_id = ? AND user_id = ?", (doc_id, user_id))

    # -- cell history and activity ---------------------------------------------

    def add_history(self, doc_id: str, user_id: str, entries: list[dict]) -> None:
        """Record entries {sheet, r, c, value, kind, summary} (r, c, sheet may be None) at the same moment,
        then keep only the newest HISTORY_KEEP for the doc."""
        if not entries:
            return
        at = int(time.time() * 1000)
        with self._lock, self._db:
            self._db.executemany(
                "INSERT INTO cell_history (doc_id, sheet, r, c, user_id, at, value, kind, summary) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)",
                [(doc_id, e.get("sheet"), e.get("r"), e.get("c"), user_id, at,
                  None if e.get("value") is None else json.dumps(e["value"], ensure_ascii=False),
                  e.get("kind", "edit"), e.get("summary", "")) for e in entries],
            )
            cut = self._db.execute("SELECT id FROM cell_history WHERE doc_id = ? ORDER BY id DESC LIMIT 1 OFFSET ?",
                                   (doc_id, HISTORY_KEEP)).fetchone()
            if cut:
                self._db.execute("DELETE FROM cell_history WHERE doc_id = ? AND id <= ?", (doc_id, cut["id"]))

    def cell_history(self, doc_id: str, sheet: str, r: int, c: int, limit: int = 50) -> list[dict]:
        rows = self._run(
            "SELECT h.at, h.value, h.user_id, u.email FROM cell_history h LEFT JOIN users u ON u.id = h.user_id "
            "WHERE h.doc_id = ? AND h.sheet = ? AND h.r = ? AND h.c = ? AND h.kind = 'edit' "
            "ORDER BY h.id DESC LIMIT ?", (doc_id, sheet, r, c, limit))
        return [{"user": {"id": x["user_id"], "email": x["email"]}, "at": x["at"],
                 "value": None if x["value"] is None else json.loads(x["value"])} for x in rows]

    def activity(self, doc_id: str, limit: int = 50) -> list[dict]:
        rows = self._run(
            "SELECT h.at, h.sheet, h.r, h.c, h.kind, h.summary, h.user_id, u.email FROM cell_history h "
            "LEFT JOIN users u ON u.id = h.user_id WHERE h.doc_id = ? ORDER BY h.id DESC LIMIT ?", (doc_id, limit))
        return [{"user": {"id": x["user_id"], "email": x["email"]}, "at": x["at"], "sheet": x["sheet"],
                 "cell": _a1(x["r"], x["c"]), "kind": x["kind"], "summary": x["summary"]} for x in rows]

    # -- comments --------------------------------------------------------------

    def _thread_rows(self, where: str, args: tuple) -> list[dict]:
        threads = self._run(f"SELECT id, sheet, cell, resolved FROM comment_threads WHERE {where} "
                            "ORDER BY created, id", args)
        if not threads:
            return []
        ids = [t["id"] for t in threads]
        marks = ",".join("?" * len(ids))
        comments = self._run(
            f"SELECT c.id, c.thread_id, c.body, c.created, c.edited, c.author_id, u.email FROM comments c "
            f"LEFT JOIN users u ON u.id = c.author_id WHERE c.thread_id IN ({marks}) ORDER BY c.created, c.rowid",
            tuple(ids))
        by_thread: dict[str, list] = {i: [] for i in ids}
        for c in comments:
            item = {"id": c["id"], "author": {"id": c["author_id"], "email": c["email"]}, "body": c["body"],
                    "created": c["created"]}
            if c["edited"]:
                item["edited"] = c["edited"]
            by_thread[c["thread_id"]].append(item)
        return [{"id": t["id"], "sheet": t["sheet"], "cell": t["cell"], "resolved": bool(t["resolved"]),
                 "comments": by_thread[t["id"]]} for t in threads]

    def threads(self, doc_id: str) -> list[dict]:
        return self._thread_rows("doc_id = ?", (doc_id,))

    def thread(self, doc_id: str, thread_id: str) -> dict | None:
        rows = self._thread_rows("doc_id = ? AND id = ?", (doc_id, thread_id))
        return rows[0] if rows else None

    def add_comment(self, doc_id: str, user_id: str, body: str, sheet: str | None = None, cell: str | None = None,
                    thread_id: str | None = None) -> str | None:
        """Add a comment to `thread_id` (re-opening it), or start a thread at sheet/cell.
        Returns the thread id, or None when thread_id isn't one of the doc's threads."""
        now = int(time.time() * 1000)
        with self._lock, self._db:
            if thread_id:
                if not self._db.execute("UPDATE comment_threads SET resolved = 0 WHERE id = ? AND doc_id = ?",
                                        (thread_id, doc_id)).rowcount:
                    return None
            else:
                thread_id = uuid.uuid4().hex
                self._db.execute("INSERT INTO comment_threads VALUES (?, ?, ?, ?, 0, ?)",
                                 (thread_id, doc_id, sheet, cell, now))
            self._db.execute("INSERT INTO comments VALUES (?, ?, ?, ?, ?, ?, NULL)",
                             (uuid.uuid4().hex, thread_id, doc_id, user_id, body, now))
        return thread_id

    def resolve_thread(self, doc_id: str, thread_id: str, resolved: bool) -> bool:
        with self._lock, self._db:
            return self._db.execute("UPDATE comment_threads SET resolved = ? WHERE id = ? AND doc_id = ?",
                                    (int(resolved), thread_id, doc_id)).rowcount > 0

    def comment_author(self, doc_id: str, thread_id: str, comment_id: str) -> str | None:
        rows = self._run("SELECT author_id FROM comments WHERE id = ? AND thread_id = ? AND doc_id = ?",
                         (comment_id, thread_id, doc_id))
        return rows[0]["author_id"] if rows else None

    def edit_comment(self, doc_id: str, thread_id: str, comment_id: str, body: str) -> bool:
        with self._lock, self._db:
            return self._db.execute(
                "UPDATE comments SET body = ?, edited = ? WHERE id = ? AND thread_id = ? AND doc_id = ?",
                (body, int(time.time() * 1000), comment_id, thread_id, doc_id)).rowcount > 0

    def delete_comment(self, doc_id: str, thread_id: str, comment_id: str) -> bool:
        """Delete one comment; a thread left empty goes too. Returns whether the thread still exists."""
        with self._lock, self._db:
            self._db.execute("DELETE FROM comments WHERE id = ? AND thread_id = ? AND doc_id = ?",
                             (comment_id, thread_id, doc_id))
            left = self._db.execute("SELECT 1 FROM comments WHERE thread_id = ? LIMIT 1", (thread_id,)).fetchone()
            if not left:
                self._db.execute("DELETE FROM comment_threads WHERE id = ? AND doc_id = ?", (thread_id, doc_id))
            return bool(left)

    # -- billing -----------------------------------------------------------------

    def billing(self, user_id: str) -> dict:
        """{plan, stripe_customer_id, renews_at (seconds or None)} for a user ('free' when unknown)."""
        rows = self._run("SELECT plan, stripe_customer_id, plan_renews_at FROM users WHERE id = ?", (user_id,))
        if not rows:
            return {"plan": "free", "stripe_customer_id": None, "renews_at": None}
        r = rows[0]
        return {"plan": r["plan"] or "free", "stripe_customer_id": r["stripe_customer_id"],
                "renews_at": r["plan_renews_at"]}

    def set_customer(self, user_id: str, customer_id: str) -> None:
        self._run("UPDATE users SET stripe_customer_id = ? WHERE id = ?", (customer_id, user_id))

    def user_by_customer(self, customer_id: str) -> str | None:
        rows = self._run("SELECT id FROM users WHERE stripe_customer_id = ?", (customer_id,))
        return rows[0]["id"] if rows else None

    def set_plan(self, user_id: str, plan: str, renews_at: float | None) -> bool:
        with self._lock, self._db:
            return self._db.execute("UPDATE users SET plan = ?, plan_renews_at = ? WHERE id = ?",
                                    (plan, renews_at, user_id)).rowcount > 0


def stream_id_key(entry_id: str) -> tuple[int, int]:
    """Redis stream ids ("1700000000000-3") as comparable tuples."""
    ms, _, seq = str(entry_id).partition("-")
    try:
        return int(ms), int(seq or 0)
    except ValueError:
        return 0, 0


def _a1(r: int | None, c: int | None) -> str | None:
    if r is None or c is None:
        return None
    letters, n = "", c + 1
    while n > 0:
        n, m = divmod(n - 1, 26)
        letters = chr(65 + m) + letters
    return f"{letters}{r + 1}"


def _sheet_count(data: str) -> int:
    try:
        return len(json.loads(data)["workbook"]["sheets"])
    except (ValueError, KeyError, TypeError):
        return 0
