"""Scheduled email reports for synced workbooks: storage, scheduling and the email itself.

Reports live in their own SQLite table in the accounts database file. Times are stored as
seconds since the epoch and served as milliseconds. Schedules are in UTC.

Key numbers: the server has no formula engine of its own, so when LibreOffice is installed the
workbook is recalculated through it (the same way the PDF export works) and the email shows
calculated values. Without LibreOffice, literal values are shown, formula cells are listed as
formulas, and the email says the totals are in the attached file.
"""

from __future__ import annotations

import calendar
import json
import logging
import sqlite3
import subprocess
import tempfile
import threading
import time
import uuid
from datetime import datetime, timedelta, timezone
from functools import lru_cache
from pathlib import Path

from openpyxl import load_workbook

from app.schemas import RowKind, WorkbookSpec

log = logging.getLogger(__name__)

DB_NAME = "sheetcraft.db"
MAX_REPORTS_PER_DOC = 20
TABLE_ROWS = 10  # rows shown when a sheet has no total rows
MAX_SHEETS_IN_EMAIL = 8
CELL_WIDTH = 18

SCHEMA = """
CREATE TABLE IF NOT EXISTS reports (
    id TEXT PRIMARY KEY,
    doc_id TEXT,
    created_by TEXT,
    frequency TEXT,
    weekday INTEGER,
    day INTEGER,
    hour INTEGER,
    recipients TEXT,
    sheet TEXT,
    attach TEXT,
    include_summary INTEGER,
    created REAL,
    last_sent REAL,
    next_send REAL
);
CREATE INDEX IF NOT EXISTS reports_doc ON reports (doc_id);
CREATE INDEX IF NOT EXISTS reports_due ON reports (next_send);
"""

FIELDS = ("frequency", "weekday", "day", "hour", "recipients", "sheet", "attach", "include_summary")


class ReportLimit(ValueError):
    pass


def next_send(frequency: str, weekday: int | None, day: int | None, hour: int, after: float) -> float:
    """The first scheduled moment strictly after `after` (epoch seconds, UTC)."""
    now = datetime.fromtimestamp(after, timezone.utc)
    if frequency == "daily":
        t = now.replace(hour=hour, minute=0, second=0, microsecond=0)
        if t <= now:
            t += timedelta(days=1)
    elif frequency == "weekly":
        wd = weekday if weekday is not None else 0
        t = now.replace(hour=hour, minute=0, second=0, microsecond=0) + timedelta(days=(wd - now.weekday()) % 7)
        if t <= now:
            t += timedelta(days=7)
    else:  # monthly; day is 1-28 so every month has it
        d = day or 1
        t = now.replace(day=d, hour=hour, minute=0, second=0, microsecond=0)
        if t <= now:
            year, month = (now.year + 1, 1) if now.month == 12 else (now.year, now.month + 1)
            t = t.replace(year=year, month=month, day=min(d, calendar.monthrange(year, month)[1]))
    return t.timestamp()


class ReportStore:
    def __init__(self, data_dir: str | Path):
        path = Path(data_dir)
        path.mkdir(parents=True, exist_ok=True)
        self._lock = threading.Lock()
        self._db = sqlite3.connect(path / DB_NAME, check_same_thread=False, timeout=10)
        self._db.row_factory = sqlite3.Row
        with self._lock:
            self._db.executescript(SCHEMA)

    @staticmethod
    def _out(r: sqlite3.Row) -> dict:
        return {
            "id": r["id"], "doc_id": r["doc_id"], "frequency": r["frequency"], "weekday": r["weekday"],
            "day": r["day"], "hour": r["hour"], "recipients": json.loads(r["recipients"]), "sheet": r["sheet"],
            "attach": r["attach"], "include_summary": bool(r["include_summary"]), "created_by": r["created_by"],
            "created": int(r["created"] * 1000),
            "last_sent": int(r["last_sent"] * 1000) if r["last_sent"] else None,
            "next_send": int(r["next_send"] * 1000) if r["next_send"] else None,
        }

    def list(self, doc_id: str) -> list[dict]:
        with self._lock:
            rows = self._db.execute("SELECT * FROM reports WHERE doc_id = ? ORDER BY created", (doc_id,)).fetchall()
        return [self._out(r) for r in rows]

    def get(self, doc_id: str, report_id: str) -> dict | None:
        with self._lock:
            row = self._db.execute("SELECT * FROM reports WHERE id = ? AND doc_id = ?", (report_id, doc_id)).fetchone()
        return self._out(row) if row else None

    def create(self, doc_id: str, user_id: str, fields: dict, now: float | None = None) -> dict:
        now = time.time() if now is None else now
        report_id = uuid.uuid4().hex
        due = next_send(fields["frequency"], fields["weekday"], fields["day"], fields["hour"], now)
        with self._lock, self._db:
            n = self._db.execute("SELECT COUNT(*) FROM reports WHERE doc_id = ?", (doc_id,)).fetchone()[0]
            if n >= MAX_REPORTS_PER_DOC:
                raise ReportLimit(MAX_REPORTS_PER_DOC)
            self._db.execute(
                "INSERT INTO reports VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                (report_id, doc_id, user_id, fields["frequency"], fields["weekday"], fields["day"], fields["hour"],
                 json.dumps(fields["recipients"]), fields["sheet"], fields["attach"], int(fields["include_summary"]),
                 now, None, due),
            )
        return self.get(doc_id, report_id)

    def delete(self, doc_id: str, report_id: str) -> bool:
        with self._lock, self._db:
            return self._db.execute("DELETE FROM reports WHERE id = ? AND doc_id = ?", (report_id, doc_id)).rowcount > 0

    def delete_for_doc(self, doc_id: str) -> None:
        with self._lock, self._db:
            self._db.execute("DELETE FROM reports WHERE doc_id = ?", (doc_id,))

    def due(self, now: float, limit: int = 50) -> list[dict]:
        with self._lock:
            rows = self._db.execute(
                "SELECT * FROM reports WHERE next_send <= ? ORDER BY next_send LIMIT ?", (now, limit)).fetchall()
        return [self._out(r) for r in rows]

    def claim(self, report_id: str, due_ms: int, now: float) -> bool:
        """Mark a due report sent and schedule the next one, only if nobody else already did."""
        with self._lock, self._db:
            row = self._db.execute("SELECT * FROM reports WHERE id = ?", (report_id,)).fetchone()
            if not row or row["next_send"] is None or int(row["next_send"] * 1000) != due_ms:
                return False
            due = next_send(row["frequency"], row["weekday"], row["day"], row["hour"], now)
            return self._db.execute(
                "UPDATE reports SET last_sent = ?, next_send = ? WHERE id = ? AND next_send = ?",
                (now, due, report_id, row["next_send"])).rowcount > 0

    def mark_sent(self, report_id: str, now: float, reschedule: bool) -> None:
        with self._lock, self._db:
            row = self._db.execute("SELECT * FROM reports WHERE id = ?", (report_id,)).fetchone()
            if not row:
                return
            due = (next_send(row["frequency"], row["weekday"], row["day"], row["hour"], now)
                   if reschedule else row["next_send"])
            self._db.execute("UPDATE reports SET last_sent = ?, next_send = ? WHERE id = ?", (now, due, report_id))


@lru_cache
def _reports_for(data_dir: str) -> ReportStore:
    return ReportStore(data_dir)


def get_report_store() -> ReportStore:
    from app.config import get_settings

    return _reports_for(get_settings().data_dir)


# ---------------------------------------------------------------------------
# Calculated values (LibreOffice when available)
# ---------------------------------------------------------------------------


def recalculated_values(spec: WorkbookSpec, timeout: int = 60) -> dict[str, list[list]] | None:
    """Per sheet, the calculated grid (row 0 = header), or None without LibreOffice or on any failure."""
    from app.services.pdf_export import _soffice
    from app.services.workbook_builder import build_workbook

    exe = _soffice()
    if not exe:
        return None
    try:
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "workbook.xlsx"
            build_workbook(spec).save(src)
            out_dir = Path(tmp) / "out"
            cmd = [exe, f"-env:UserInstallation=file://{Path(tmp) / 'profile'}", "--headless", "--calc",
                   "--convert-to", "xlsx", "--outdir", str(out_dir), str(src)]
            subprocess.run(cmd, check=True, capture_output=True, timeout=timeout)
            wb = load_workbook(out_dir / "workbook.xlsx", data_only=True, read_only=True)
            names = {s.name for s in spec.sheets}
            values = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)]
                      for ws in wb.worksheets if ws.title in names}
            wb.close()
            return values
    except Exception:  # noqa: BLE001 - fall back to the uncalculated workbook
        log.exception("Couldn't recalculate a workbook for a report")
        return None


# ---------------------------------------------------------------------------
# The email
# ---------------------------------------------------------------------------


def _fmt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v.is_integer():
            return f"{int(v):,}"
        return f"{v:,.2f}"
    if isinstance(v, int):
        return f"{v:,}"
    if hasattr(v, "isoformat"):
        return v.isoformat()[:10] if getattr(v, "hour", 0) == 0 else v.isoformat(sep=" ")
    return str(v).replace("\n", " ")


def _cell(sheet, values: list[list] | None, r: int, c: int) -> tuple[str, bool]:
    """(display text, is an uncalculated formula) for spec row index r (Excel row r + 2), column c."""
    if values is not None:
        row = values[r + 1] if r + 1 < len(values) else []
        return _fmt(row[c] if c < len(row) else None), False
    cells = sheet.rows[r].cells
    cell = cells[c] if c < len(cells) else None
    v = cell.value if cell else None
    if isinstance(v, str) and v.startswith("="):
        from app.services.workbook_builder import is_ai_formula

        if is_ai_formula(v):
            return (cell.ai_value or ""), False
        return v, True
    return _fmt(v), False


def _table(lines: list[list[str]]) -> str:
    out = []
    for row in lines:
        out.append("  ".join((t[: CELL_WIDTH - 1] + "…" if len(t) > CELL_WIDTH else t).ljust(CELL_WIDTH) for t in row).rstrip())
    return "\n".join(out)


def key_numbers(spec: WorkbookSpec, sheet_name: str | None, values: dict[str, list[list]] | None) -> tuple[str, bool]:
    """Plain-text key numbers and whether any of them are uncalculated formulas."""
    sheets = [s for s in spec.sheets if s.name == sheet_name] if sheet_name else spec.sheets
    sheets = sheets or spec.sheets
    parts, uncalculated = [], False
    for sheet in sheets[:MAX_SHEETS_IN_EMAIL]:
        grid = values.get(sheet.name) if values is not None else None
        calc = grid if values is not None else None
        ncols = len(sheet.columns)
        totals = [i for i, row in enumerate(sheet.rows) if row.kind in (RowKind.total, RowKind.subtotal)]
        lines: list[str] = [sheet.name, "-" * len(sheet.name)]
        if totals and not sheet_name:
            for i in totals[:10]:
                texts = [_cell(sheet, calc, i, c) for c in range(ncols)]
                uncalculated |= any(f for _, f in texts)
                label = texts[0][0] or f"Row {i + 2}"
                pairs = [f"{sheet.columns[c].header}: {t}" for c, (t, _) in enumerate(texts) if c > 0 and t]
                lines.append(f"{label}: " + ", ".join(pairs) if pairs else label)
        else:
            shown = sheet.rows[:TABLE_ROWS]
            table = [[c.header for c in sheet.columns]]
            for i in range(len(shown)):
                texts = [_cell(sheet, calc, i, c) for c in range(ncols)]
                uncalculated |= any(f for _, f in texts)
                table.append([t for t, _ in texts])
            if sheet_name:  # the chosen sheet: also its total rows past the first 10
                for i in [t for t in totals if t >= TABLE_ROWS]:
                    texts = [_cell(sheet, calc, i, c) for c in range(ncols)]
                    uncalculated |= any(f for _, f in texts)
                    table.append([t for t, _ in texts])
            lines.append(_table(table))
            if len(sheet.rows) > TABLE_ROWS:
                lines.append(f"(first {TABLE_ROWS} of {len(sheet.rows)} rows)")
        parts.append("\n".join(lines))
    if len(sheets) > MAX_SHEETS_IN_EMAIL:
        parts.append(f"...and {len(sheets) - MAX_SHEETS_IN_EMAIL} more sheets.")
    return "\n\n".join(parts), uncalculated


def compose(spec: WorkbookSpec, report: dict, doc_id: str, app_url: str, numbers: str, uncalculated: bool,
            summary: str | None, attached: str | None) -> tuple[str, str]:
    subject = f"{spec.title} — {report['frequency']} report"
    body = [f"{spec.title}", ""]
    if summary:
        body += [summary, ""]
    body += ["Key numbers", "", numbers, ""]
    if uncalculated:
        body.append("Some cells are shown as formulas because they couldn't be calculated on the server. "
                    + ("The attached file has the calculated totals." if attached
                       else "Open the workbook to see the calculated totals."))
        body.append("")
    if attached:
        body += [f"Attached: {attached}", ""]
    body += [f"Open the workbook: {app_url.rstrip('/')}/?doc={doc_id}", "",
             "You get this email because someone with access to this workbook scheduled a "
             f"{report['frequency']} report for this address. To stop it, ask them to delete the report."]
    return subject, "\n".join(body)
