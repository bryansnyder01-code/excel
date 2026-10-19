"""Background jobs, run once a minute from the app's lifespan (when settings.background_jobs is on):

* scheduled refresh of live data connections (workbook.connections with refresh_minutes), and
* scheduled email reports.

When a live room is open for a doc, refreshed data goes through the room so everyone connected gets a
snapshot; otherwise it is saved straight to the stored doc.

The job reads docs through `DocSource`, a small interface, so it can be wired to whatever the store
offers. `StoreDocSource` uses the store's public methods, plus `Store.docs_with_connections()` when the
store has it; until then it finds candidate docs with a read-only query of the docs table.
"""

from __future__ import annotations

import asyncio
import copy
import logging
import re
import sqlite3
import time
from collections.abc import Awaitable, Callable
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Protocol

from fastapi.concurrency import run_in_threadpool
from pydantic import ValidationError

from app.config import Settings, get_settings
from app.schemas import WorkbookSpec
from app.services.fetcher import FetchError, FetchResult, fetch_table
from app.services.formula_validator import sanitize_sheet_name, validate_workbook
from app.services.mailer import send_mail, smtp_configured
from app.services.reports import ReportStore, compose, key_numbers, recalculated_values

log = logging.getLogger(__name__)

TICK_SECONDS = 60
MIN_REFRESH_MINUTES = 5  # a connection never refreshes more often than this, whatever it asks for
MAX_DOCS_PER_TICK = 200

Fetch = Callable[..., Awaitable[FetchResult]]


# ---------------------------------------------------------------------------
# Docs
# ---------------------------------------------------------------------------


class DocSource(Protocol):
    def docs_with_connections(self) -> list[str]:
        """Ids of docs whose workbook has at least one connection with refresh_minutes."""

    def load(self, doc_id: str) -> tuple[str, dict] | None:
        """(owner id, workbook dict) or None when the doc is gone."""

    def save(self, doc_id: str, workbook: dict) -> bool:
        ...


class StoreDocSource:
    def __init__(self, store, data_dir: str | Path):
        self.store = store
        self.db_path = Path(data_dir) / "sheetcraft.db"

    def docs_with_connections(self) -> list[str]:
        if hasattr(self.store, "docs_with_connections"):
            return list(self.store.docs_with_connections())[:MAX_DOCS_PER_TICK]
        if not self.db_path.exists():
            return []
        # Read-only fallback: a cheap text match, confirmed by load() + is_due().
        db = sqlite3.connect(f"file:{self.db_path}?mode=ro", uri=True, timeout=10)
        try:
            rows = db.execute("SELECT id FROM docs WHERE data LIKE '%\"refresh_minutes\":%' LIMIT ?",
                              (MAX_DOCS_PER_TICK,)).fetchall()
        except sqlite3.Error:
            return []
        finally:
            db.close()
        return [r[0] for r in rows]

    def load(self, doc_id: str) -> tuple[str, dict] | None:
        owner = self.store.doc_owner(doc_id)
        doc = self.store.get_doc(owner, doc_id) if owner else None
        if not doc or not isinstance(doc.get("data"), dict) or not isinstance(doc["data"].get("workbook"), dict):
            return None
        return owner, doc["data"]["workbook"]

    def save(self, doc_id: str, workbook: dict) -> bool:
        return self.store.save_workbook(doc_id, workbook, time.time() * 1000)


def _room(doc_id: str):
    try:
        from app.api.live import rooms
    except ImportError:  # pragma: no cover
        return None
    return rooms.get(doc_id)


def current_workbook(source: DocSource, doc_id: str) -> tuple[str | None, dict] | None:
    """The freshest copy: the live room's when one is open, else the stored doc."""
    room = _room(doc_id)
    if room is not None:
        loaded = source.load(doc_id)
        return (loaded[0] if loaded else None), room.workbook
    return source.load(doc_id)


async def publish_workbook(source: DocSource, doc_id: str, workbook: dict) -> None:
    """Through the open live room when there is one (everyone connected gets a snapshot), else saved directly."""
    room = _room(doc_id)
    if room is None:
        await run_in_threadpool(source.save, doc_id, workbook)
    elif hasattr(room, "publish") and hasattr(room, "broker"):
        # Rooms backed by a broker apply snapshots in stream order on every instance (and save them).
        await room.publish({"kind": "snapshot", "conn": None, "from": "system", "workbook": workbook})
    else:
        room.workbook = workbook
        room.changed()  # saves it shortly
        room.broadcast({"type": "snapshot", "version": room.version, "from": "system", "workbook": workbook})


# ---------------------------------------------------------------------------
# Live data refresh
# ---------------------------------------------------------------------------


def is_due(conn: dict, now_ms: int) -> bool:
    minutes = conn.get("refresh_minutes")
    if not isinstance(minutes, (int, float)) or isinstance(minutes, bool) or minutes <= 0:
        return False
    last = conn.get("last_refreshed")
    if not isinstance(last, (int, float)):
        return True
    return now_ms - last >= max(minutes, MIN_REFRESH_MINUTES) * 60_000


def apply_fetch(workbook: dict, conn_id: str, result: FetchResult, now_ms: int) -> bool:
    """Replace the connection's target sheet's columns and rows (creating the sheet), set last_refreshed."""
    conn = next((c for c in workbook.get("connections") or [] if isinstance(c, dict) and c.get("id") == conn_id), None)
    if conn is None:
        return False
    sheets = workbook.setdefault("sheets", [])
    target = str(conn.get("sheet") or conn.get("name") or "Data")
    sheet = next((s for s in sheets if isinstance(s, dict) and str(s.get("name", "")).lower() == target.lower()), None)
    if sheet is None:
        name = sanitize_sheet_name(target, {str(s.get("name", "")).lower() for s in sheets if isinstance(s, dict)})
        conn["sheet"] = name
        sheet = {"name": name, "columns": [], "rows": []}
        sheets.append(sheet)
    old = {c.get("header"): c for c in sheet.get("columns") or [] if isinstance(c, dict)}
    columns = []
    for i, header in enumerate(result.columns):
        col = {"header": header, "format": result.formats[i].value if i < len(result.formats) else "text"}
        prev = old.get(header)
        if prev:  # keep the user's width, decimals and currency for a column that is still there
            for k in ("width", "decimals", "currency"):
                if prev.get(k) is not None:
                    col[k] = prev[k]
            if prev.get("format") and prev.get("format") != "text" and col["format"] != "text":
                col["format"] = prev["format"]
        columns.append(col)
    sheet["columns"] = columns
    sheet["rows"] = [{"cells": [{"value": v} for v in row]} for row in result.rows]
    conn["last_refreshed"] = now_ms
    return True


_attempts: dict[tuple[str, str], int] = {}  # (doc, connection) -> last attempt (ms), so failures back off


async def refresh_doc(source: DocSource, doc_id: str, fetch: Fetch = fetch_table, now_ms: int | None = None,
                      force: bool = False) -> int:
    """Fetch every due connection of one doc and publish the result. Returns how many were refreshed."""
    now_ms = now_ms or int(time.time() * 1000)
    found = await run_in_threadpool(current_workbook, source, doc_id)
    if not found:
        return 0
    _, workbook = found
    due = [c for c in workbook.get("connections") or [] if isinstance(c, dict) and (force or is_due(c, now_ms))]
    results: list[tuple[str, FetchResult]] = []
    for conn in due:
        key = (doc_id, str(conn.get("id")))
        interval = max(conn.get("refresh_minutes") or MIN_REFRESH_MINUTES, MIN_REFRESH_MINUTES) * 60_000
        if not force and now_ms - _attempts.get(key, 0) < interval:
            continue
        _attempts[key] = now_ms
        try:
            results.append((str(conn.get("id")), await fetch(conn.get("kind"), conn.get("url"), conn.get("path"))))
        except FetchError as e:
            log.info("Refresh of %s / %s failed: %s", doc_id, conn.get("id"), e)
        except Exception:  # noqa: BLE001 - one bad source must not stop the others
            log.exception("Refresh of %s / %s failed", doc_id, conn.get("id"))
    if not results:
        return 0
    # Apply to the freshest copy: people may have edited while the fetches ran.
    found = await run_in_threadpool(current_workbook, source, doc_id)
    if not found:
        return 0
    workbook = copy.deepcopy(found[1])
    n = sum(apply_fetch(workbook, cid, result, now_ms) for cid, result in results)
    if not n:
        return 0
    try:
        WorkbookSpec.model_validate(workbook)
    except ValidationError:
        log.warning("Refreshed workbook %s didn't validate; not saved", doc_id)
        return 0
    await publish_workbook(source, doc_id, workbook)
    return n


async def refresh_due_connections(source: DocSource, fetch: Fetch = fetch_table, now_ms: int | None = None) -> int:
    ids = await run_in_threadpool(source.docs_with_connections)
    total = 0
    for doc_id in ids:
        try:
            total += await refresh_doc(source, doc_id, fetch, now_ms)
        except Exception:  # noqa: BLE001
            log.exception("Refreshing doc %s failed", doc_id)
    return total


# ---------------------------------------------------------------------------
# Email reports
# ---------------------------------------------------------------------------


def _filename(title: str) -> str:
    return re.sub(r"[^\w\- ]+", "", title).strip().replace(" ", "_") or "workbook"


async def deliver_report(report: dict, source: DocSource, engine_factory: Callable[[], object] | None,
                         settings: Settings, usage=None) -> bool:
    """Build and send (or, without SMTP, log) one report. False when the doc is gone or unreadable."""
    from app.services.llm_engine import EngineError
    from app.services.pdf_export import PdfUnavailable, workbook_to_pdf
    from app.services.workbook_builder import workbook_to_bytes

    doc_id = report["doc_id"]
    found = await run_in_threadpool(current_workbook, source, doc_id)
    if not found or not found[0]:
        return False
    owner, raw = found
    try:
        spec = WorkbookSpec.model_validate(copy.deepcopy(raw))
    except ValidationError:
        log.warning("Report %s: workbook %s couldn't be read", report["id"], doc_id)
        return False
    validate_workbook(spec)
    values = await run_in_threadpool(recalculated_values, spec)
    numbers, uncalculated = key_numbers(spec, report.get("sheet"), values)

    summary = None
    if report.get("include_summary") and engine_factory is not None:
        allowed = True
        if settings.ai_limit_enabled and usage is not None:
            allowed, _ = usage.take(f"user:{owner}", settings.ai_limit_signed_in)
        if allowed:
            try:
                summary = await engine_factory().report_summary(spec.title, numbers[:20000])
            except EngineError as e:
                log.info("Report %s: no AI summary (%s)", report["id"], e)
            except Exception:  # noqa: BLE001 - e.g. no API key configured: send without a summary
                log.exception("Report %s: AI summary failed", report["id"])
        else:
            log.info("Report %s: owner is out of AI requests, sending without a summary", report["id"])

    attachments, attached = [], None
    name = _filename(spec.title)
    if report.get("attach") == "pdf":
        try:
            attachments.append((f"{name}.pdf", await run_in_threadpool(workbook_to_pdf, spec), "application/pdf"))
        except PdfUnavailable:
            log.info("Report %s: PDF unavailable, attaching .xlsx instead", report["id"])
    if report.get("attach") in ("xlsx", "pdf") and not attachments:
        attachments.append((f"{name}.xlsx", workbook_to_bytes(spec),
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
    if attachments:
        attached = attachments[0][0]

    subject, body = compose(spec, report, doc_id, settings.app_url, numbers, uncalculated, summary, attached)
    if not smtp_configured(settings):
        log.warning("Report email (SMTP not configured) to %s\nSubject: %s\nAttachments: %s\n\n%s",
                    ", ".join(report["recipients"]), subject, attached or "none", body)
        return True
    for to in report["recipients"]:
        await run_in_threadpool(send_mail, settings, to, subject, body, attachments)
    return True


async def send_due_reports(source: DocSource, reports: ReportStore, engine_factory, settings: Settings,
                           usage=None, now: float | None = None) -> int:
    now = time.time() if now is None else now
    sent = 0
    for report in await run_in_threadpool(reports.due, now):
        # Claim it first (reschedules it), so a slow or failing send, or a second backend instance, can't send
        # it twice.
        if not await run_in_threadpool(reports.claim, report["id"], report["next_send"], now):
            continue
        try:
            ok = await deliver_report(report, source, engine_factory, settings, usage)
        except Exception:  # noqa: BLE001
            log.exception("Report %s failed", report["id"])
            continue
        if not ok and source.load(report["doc_id"]) is None:
            await run_in_threadpool(reports.delete_for_doc, report["doc_id"])  # the doc was deleted
        sent += ok
    return sent


# ---------------------------------------------------------------------------
# The loop
# ---------------------------------------------------------------------------


async def tick(engine_factory) -> None:
    from app.api.accounts import get_store
    from app.services.reports import get_report_store
    from app.services.usage import get_usage_store

    settings = get_settings()
    source = StoreDocSource(get_store(), settings.data_dir)
    try:
        await refresh_due_connections(source)
    except Exception:  # noqa: BLE001
        log.exception("Live data refresh failed")
    try:
        await send_due_reports(source, get_report_store(), engine_factory, settings, get_usage_store())
    except Exception:  # noqa: BLE001
        log.exception("Sending reports failed")


async def run_forever(engine_factory, interval: float = TICK_SECONDS) -> None:
    while True:
        await asyncio.sleep(interval)
        await tick(engine_factory)


@asynccontextmanager
async def background_jobs(app):
    """Start the once-a-minute loop for the app's lifetime (only when settings.background_jobs is on)."""
    task = None
    if getattr(get_settings(), "background_jobs", True):
        from app.api.routes import get_engine

        def engine_factory():
            return app.dependency_overrides.get(get_engine, get_engine)()

        task = asyncio.create_task(run_forever(engine_factory))
    try:
        yield
    finally:
        if task:
            task.cancel()
            try:
                await task
            except (asyncio.CancelledError, Exception):  # noqa: BLE001
                pass



def install_background_jobs(app) -> None:
    """Wrap the app's existing lifespan (whatever else it starts) with the background jobs."""
    inner = app.router.lifespan_context

    @asynccontextmanager
    async def lifespan(a):
        async with inner(a) as state, background_jobs(a):
            yield state

    app.router.lifespan_context = lifespan
