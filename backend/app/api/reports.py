"""Scheduled email reports for a synced workbook. Owners and editors of the doc manage them.

GET    /api/docs/{doc_id}/reports              -> [report]
POST   /api/docs/{doc_id}/reports              -> report
DELETE /api/docs/{doc_id}/reports/{rid}        -> 204
POST   /api/docs/{doc_id}/reports/{rid}/send   -> 204 (send now; the schedule is unchanged)
"""

from __future__ import annotations

import time
from typing import Literal

from fastapi import APIRouter, Depends, HTTPException, Request, Response
from fastapi.concurrency import run_in_threadpool
from pydantic import BaseModel, Field

from app.api.accounts import _EMAIL, _can_open, current_user, get_store
from app.api.routes import get_engine
from app.config import get_settings
from app.services.reports import ReportLimit, ReportStore, get_report_store
from app.services.scheduler import StoreDocSource, current_workbook, deliver_report
from app.services.store import Store
from app.services.usage import get_usage_store

reports_router = APIRouter(prefix="/api/docs/{doc_id}/reports", tags=["reports"])

REPORT_NOT_FOUND = "That report doesn't exist."
MAX_RECIPIENTS = 10


class ReportIn(BaseModel):
    frequency: Literal["daily", "weekly", "monthly"]
    weekday: int | None = Field(default=None, ge=0, le=6, description="weekly: 0 = Monday ... 6 = Sunday.")
    day: int | None = Field(default=None, ge=1, le=28, description="monthly: day of the month.")
    hour: int = Field(ge=0, le=23, description="Hour of the day, UTC.")
    recipients: list[str]
    sheet: str | None = None
    attach: Literal["xlsx", "pdf", "none"] = "xlsx"
    include_summary: bool = False


class ReportOut(BaseModel):
    id: str
    doc_id: str
    frequency: str
    weekday: int | None
    day: int | None
    hour: int
    recipients: list[str]
    sheet: str | None
    attach: str
    include_summary: bool
    created: int
    last_sent: int | None
    next_send: int | None


def _out(r: dict) -> ReportOut:
    return ReportOut(**{k: v for k, v in r.items() if k in ReportOut.model_fields})


@reports_router.get("", response_model=list[ReportOut])
def list_reports(doc_id: str, user: dict = Depends(current_user), store: Store = Depends(get_store),
                 reports: ReportStore = Depends(get_report_store)):
    _can_open(store, user, doc_id)
    return [_out(r) for r in reports.list(doc_id)]


@reports_router.post("", response_model=ReportOut)
def create_report(doc_id: str, body: ReportIn, user: dict = Depends(current_user),
                  store: Store = Depends(get_store), reports: ReportStore = Depends(get_report_store)):
    _can_open(store, user, doc_id)
    recipients: list[str] = []
    for raw in body.recipients:
        email = raw.strip().lower()
        if not _EMAIL.match(email) or len(email) > 254:
            raise HTTPException(422, f'"{raw.strip()}" isn\'t a valid email address.')
        if email not in recipients:
            recipients.append(email)
    if not recipients:
        raise HTTPException(422, "Add at least one email address.")
    if len(recipients) > MAX_RECIPIENTS:
        raise HTTPException(422, f"A report can go to at most {MAX_RECIPIENTS} people.")
    sheet = (body.sheet or "").strip() or None
    if sheet:
        doc = store.get_doc(user["id"], doc_id)
        workbook = doc["data"].get("workbook") if doc and isinstance(doc["data"], dict) else None
        names = [s.get("name") for s in (workbook or {}).get("sheets") or [] if isinstance(s, dict)]
        if sheet not in names:
            raise HTTPException(422, f'This workbook has no sheet called "{sheet}".')
    fields = {
        "frequency": body.frequency,
        "weekday": (body.weekday if body.weekday is not None else 0) if body.frequency == "weekly" else None,
        "day": (body.day or 1) if body.frequency == "monthly" else None,
        "hour": body.hour, "recipients": recipients, "sheet": sheet, "attach": body.attach,
        "include_summary": body.include_summary,
    }
    try:
        return _out(reports.create(doc_id, user["id"], fields))
    except ReportLimit as e:
        raise HTTPException(422, f"A workbook can have at most {e} scheduled reports.") from None


@reports_router.delete("/{report_id}", status_code=204)
def delete_report(doc_id: str, report_id: str, user: dict = Depends(current_user),
                  store: Store = Depends(get_store), reports: ReportStore = Depends(get_report_store)):
    _can_open(store, user, doc_id)
    if not reports.delete(doc_id, report_id):
        raise HTTPException(404, REPORT_NOT_FOUND)
    return Response(status_code=204)


@reports_router.post("/{report_id}/send", status_code=204)
async def send_report_now(doc_id: str, report_id: str, request: Request, user: dict = Depends(current_user),
                          store: Store = Depends(get_store), reports: ReportStore = Depends(get_report_store)):
    """Send it now (counts towards the owner's AI requests when it has a summary). The schedule is unchanged."""
    await run_in_threadpool(_can_open, store, user, doc_id)
    report = reports.get(doc_id, report_id)
    if not report:
        raise HTTPException(404, REPORT_NOT_FOUND)
    settings = get_settings()
    source = StoreDocSource(store, settings.data_dir)
    if not await run_in_threadpool(current_workbook, source, doc_id):
        raise HTTPException(409, "This workbook has nothing to report yet. Open it once so it syncs, then try again.")
    # The engine is only created when the report has an AI summary (tests override get_engine).
    engine_factory = request.app.dependency_overrides.get(get_engine, get_engine)
    ok = await deliver_report(report, source, engine_factory, settings, get_usage_store())
    if not ok:
        raise HTTPException(409, "This workbook couldn't be read for the report.")
    await run_in_threadpool(reports.mark_sent, report_id, time.time(), False)
    return Response(status_code=204)
