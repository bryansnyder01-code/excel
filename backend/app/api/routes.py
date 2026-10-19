"""HTTP routes for the workbook engine."""

from __future__ import annotations

import asyncio
import io
import json
import logging
import re
from collections.abc import Awaitable, Callable
from dataclasses import dataclass
from functools import lru_cache
from pathlib import Path

from fastapi import APIRouter, Depends, File, Form, Header, HTTPException, Query, Request, UploadFile
from fastapi.responses import Response, StreamingResponse
from pydantic import BaseModel

from app.api.accounts import _bearer, get_store
from app.config import get_settings
from app.schemas import (
    AIEvalRequest,
    AIEvalResponse,
    AIFillRequest,
    AIFillResponse,
    AskRequest,
    AskResult,
    CheckRequest,
    CheckResponse,
    EditRequest,
    ExplainRequest,
    ExplainResult,
    GenerateRequest,
    Issue,
    PivotRequest,
    WorkbookResponse,
    WorkbookSpec,
)
from app.services.checker import check_workbook
from app.services.formula_validator import validate_workbook
from app.services.importer import ImportError_, import_file, tables_to_workbook
from app.services.llm_engine import EngineError, WorkbookEngine
from app.services.pdf_export import PdfUnavailable, pdf_available, workbook_to_pdf
from app.services.pivot import PivotError, build_pivot
from app.services.usage import UsageStore, get_usage_store
from app.services.workbook_builder import workbook_to_bytes

log = logging.getLogger(__name__)

router = APIRouter(prefix="/api/workbooks", tags=["workbooks"])
templates_router = APIRouter(prefix="/api/templates", tags=["templates"])

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "templates"
MAX_UPLOAD_BYTES = 10 * 1024 * 1024
PDF_MIME = "application/pdf"

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@lru_cache
def get_engine() -> WorkbookEngine:
    return WorkbookEngine()


def _check_prompt(prompt: str) -> str:
    prompt = prompt.strip()
    limit = get_settings().max_prompt_chars
    if len(prompt) > limit:
        raise HTTPException(422, f"Please keep requests under {limit} characters.")
    return prompt


# ---------------------------------------------------------------------------
# AI usage limits: every Claude call counts one request against the caller's rolling 24 hours.
# ---------------------------------------------------------------------------


@dataclass
class AICaller:
    key: str
    signed_in: bool
    usage: UsageStore
    user: dict | None = None

    @property
    def limit(self) -> int:
        from app.services.usage import limit_for_user

        return limit_for_user(self.user)

    def spend(self) -> None:
        """Count one AI request, or raise 429 when today's allowance is used up."""
        if not get_settings().ai_limit_enabled:
            return
        allowed, _ = self.usage.take(self.key, self.limit)
        if not allowed:
            more = "Try again tomorrow." if self.signed_in else "Sign in for more, or try again tomorrow."
            raise HTTPException(429, f"You've used today's AI requests ({self.limit}). {more}")


def _client_ip(request: Request) -> str:
    if get_settings().trust_proxy:
        forwarded = request.headers.get("x-forwarded-for", "").split(",")[0].strip()
        if forwarded:
            return forwarded
    return request.client.host if request.client else "unknown"


def ai_caller(request: Request, authorization: str | None = Header(default=None),
              usage: UsageStore = Depends(get_usage_store)) -> AICaller:
    """Signed-in user when a valid Bearer token is sent, else the client IP."""
    token = _bearer(authorization)
    user = get_store().session_user(token) if token else None
    if user:
        return AICaller(key=f"user:{user['id']}", signed_in=True, usage=usage, user=user)
    return AICaller(key=f"ip:{_client_ip(request)}", signed_in=False, usage=usage)


@router.get("/usage")
def usage(caller: AICaller = Depends(ai_caller)):
    return {"used": caller.usage.used(caller.key), "limit": caller.limit, "signed_in": caller.signed_in}


@router.post("/generate", response_model=WorkbookResponse)
async def generate(req: GenerateRequest, engine: WorkbookEngine = Depends(get_engine),
                   caller: AICaller = Depends(ai_caller)):
    prompt = _check_prompt(req.prompt)
    caller.spend()
    try:
        return await engine.generate(prompt)
    except EngineError as e:
        raise HTTPException(502, str(e)) from e


@router.post("/edit", response_model=WorkbookResponse)
async def edit(req: EditRequest, engine: WorkbookEngine = Depends(get_engine), caller: AICaller = Depends(ai_caller)):
    prompt = _check_prompt(req.prompt)
    caller.spend()
    try:
        return await engine.edit(prompt, req.workbook)
    except EngineError as e:
        raise HTTPException(502, str(e)) from e


# ---------------------------------------------------------------------------
# Streaming variants: Server-Sent Events, one `data: <json>` line per event.
#   {"type": "progress", "text": "..."}  (any number)
#   {"type": "result", "workbook": {...}, "summary": "...", "issues": [...]}  or
#   {"type": "error", "message": "..."}  (exactly one, last)
# ---------------------------------------------------------------------------


def _sse(event: dict) -> str:
    return f"data: {json.dumps(event, ensure_ascii=False)}\n\n"


def _event_stream(run: Callable[[Callable[[str], Awaitable[None]]], Awaitable[WorkbookResponse]]) -> StreamingResponse:
    """Run the engine in a task and stream its progress lines, then the result or a user-safe error."""

    async def events():
        queue: asyncio.Queue[dict | None] = asyncio.Queue()

        async def progress(text: str) -> None:
            await queue.put({"type": "progress", "text": text})

        async def worker() -> None:
            try:
                result = await run(progress)
                await queue.put({"type": "result", **result.model_dump(mode="json")})
            except EngineError as e:
                await queue.put({"type": "error", "message": str(e)})
            except Exception:
                log.exception("Streaming generation failed")
                await queue.put({"type": "error", "message": "Something went wrong. Please try again."})
            finally:
                await queue.put(None)

        task = asyncio.create_task(worker())
        try:
            while (event := await queue.get()) is not None:
                yield _sse(event)
        finally:
            task.cancel()  # client went away: stop paying for the model call

    return StreamingResponse(
        events(), media_type="text/event-stream", headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"}
    )


def _error_stream(message: str) -> StreamingResponse:
    async def events():
        yield _sse({"type": "error", "message": message})

    return StreamingResponse(events(), media_type="text/event-stream", headers={"Cache-Control": "no-cache"})


@router.post("/generate/stream")
async def generate_stream(req: GenerateRequest, engine: WorkbookEngine = Depends(get_engine),
                          caller: AICaller = Depends(ai_caller)):
    try:
        prompt = _check_prompt(req.prompt)
    except HTTPException as e:
        return _error_stream(e.detail)
    caller.spend()  # over the limit: a plain HTTP 429 before any streaming starts
    return _event_stream(lambda progress: engine.generate(prompt, progress=progress))


@router.post("/edit/stream")
async def edit_stream(req: EditRequest, engine: WorkbookEngine = Depends(get_engine),
                      caller: AICaller = Depends(ai_caller)):
    try:
        prompt = _check_prompt(req.prompt)
    except HTTPException as e:
        return _error_stream(e.detail)
    caller.spend()
    return _event_stream(lambda progress: engine.edit(prompt, req.workbook, progress=progress))


class ValidateResponse(BaseModel):
    workbook: WorkbookSpec
    issues: list[Issue]


@router.post("/validate", response_model=ValidateResponse)
async def validate(workbook: WorkbookSpec):
    """Check a workbook after manual grid edits; returns it with safe auto-corrections applied."""
    issues = validate_workbook(workbook)
    return ValidateResponse(workbook=workbook, issues=issues)


@router.post("/export")
async def export(workbook: WorkbookSpec, format: str = Query("xlsx", pattern="^(xlsx|pdf)$")):
    """Compile the workbook to .xlsx (default) or .pdf and return it as a download."""
    validate_workbook(workbook)  # sanitises sheet names, guards divisions
    filename = re.sub(r"[^\w\- ]+", "", workbook.title).strip().replace(" ", "_") or "workbook"
    if format == "pdf":
        try:
            data, mime = workbook_to_pdf(workbook), PDF_MIME
        except PdfUnavailable as e:
            raise HTTPException(503, str(e)) from e
    else:
        data, mime = workbook_to_bytes(workbook), XLSX_MIME
    return Response(
        content=data,
        media_type=mime,
        headers={"Content-Disposition": f'attachment; filename="{filename}.{format}"'},
    )


@router.post("/import", response_model=WorkbookResponse)
async def import_workbook(file: UploadFile = File(...)):
    """Read an uploaded .xlsx or .csv into a workbook the grid and AI can work with."""
    data = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "Files can be at most 10 MB.")
    try:
        workbook, issues = import_file(file.filename or "upload.xlsx", data)
    except ImportError_ as e:
        raise HTTPException(422, str(e)) from e
    issues += validate_workbook(workbook, autocorrect=False)
    n = len(workbook.sheets)
    rows = sum(len(s.rows) for s in workbook.sheets)
    summary = f"Opened {file.filename}: {n} {'sheet' if n == 1 else 'sheets'}, {rows} rows. Ask me to clean it up, analyse it or add formulas."
    return WorkbookResponse(workbook=workbook, summary=summary, issues=issues)


@router.post("/pivot", response_model=WorkbookResponse)
async def pivot(req: PivotRequest):
    """Add a summary sheet (totals by category) built from live SUMIFS-style formulas."""
    try:
        workbook, summary = build_pivot(req.workbook, req)
    except PivotError as e:
        raise HTTPException(422, str(e)) from e
    issues = [i for i in validate_workbook(workbook) if i.sheet == workbook.sheets[-1].name]
    return WorkbookResponse(workbook=workbook, summary=summary, issues=issues)


@router.post("/explain", response_model=ExplainResult)
async def explain(req: ExplainRequest, engine: WorkbookEngine = Depends(get_engine),
                  caller: AICaller = Depends(ai_caller)):
    """Explain one cell in plain English and, when it's wrong, suggest a corrected value."""
    caller.spend()
    try:
        return await engine.explain(req)
    except EngineError as e:
        raise HTTPException(502, str(e)) from e


@router.post("/ask", response_model=AskResult)
async def ask(req: AskRequest, engine: WorkbookEngine = Depends(get_engine), caller: AICaller = Depends(ai_caller)):
    """Answer a question about the data, pointing at the cells that back up the answer."""
    req.question = _check_prompt(req.question)
    if not req.question:
        raise HTTPException(422, "Please type a question.")
    caller.spend()
    try:
        return await engine.ask(req)
    except EngineError as e:
        raise HTTPException(502, str(e)) from e


@router.post("/check", response_model=CheckResponse)
async def check(req: CheckRequest, engine: WorkbookEngine = Depends(get_engine),
                caller: AICaller = Depends(ai_caller)):
    """Find errors and likely mistakes, each with a suggested fix where one is clear."""
    if req.use_ai:
        caller.spend()
    return await check_workbook(req, engine if req.use_ai else None)


# ---------------------------------------------------------------------------
# AI fill (one value per selected row), =AI() cells, photo / PDF to table
# ---------------------------------------------------------------------------
MAX_FILL_ROWS = 200
MAX_EVAL_ITEMS = 100
MAX_CELL_CHARS = 1000  # each input cell is cut to this many characters before it reaches the model
MAX_AI_INPUT_CHARS = 200_000  # whole request


def _cell_text(v) -> str:
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v)[:MAX_CELL_CHARS]


_NUMBER_RE = re.compile(r"^[-+]?\d+(\.\d+)?([eE][-+]?\d+)?$")


def _to_number(text: str) -> float | None:
    s = re.sub(r"[\s,$€£¥₹]", "", text or "")
    pct = s.endswith("%")
    s = s.rstrip("%")
    if not _NUMBER_RE.match(s):
        return None
    n = float(s)
    return n / 100 if pct else n


@router.post("/ai-fill", response_model=AIFillResponse)
async def ai_fill(req: AIFillRequest, engine: WorkbookEngine = Depends(get_engine),
                  caller: AICaller = Depends(ai_caller)):
    """One AI-written value per row, in order (classify, extract, translate, score...)."""
    instruction = _check_prompt(req.instruction)
    if not instruction:
        raise HTTPException(422, "Please say what to fill in.")
    if not req.rows:
        raise HTTPException(422, "Select at least one row to fill.")
    if len(req.rows) > MAX_FILL_ROWS:
        raise HTTPException(422, f"AI fill works on up to {MAX_FILL_ROWS} rows at a time. Select fewer rows.")
    headers = [_cell_text(h) for h in req.headers]
    rows = [[_cell_text(v) for v in row] for row in req.rows]
    if sum(len(c) for row in rows for c in row) > MAX_AI_INPUT_CHARS:
        raise HTTPException(422, "Those rows hold too much text for one AI fill. Select fewer rows or columns.")
    caller.spend()
    try:
        answers = await engine.fill(instruction, headers, rows, req.output.value)
    except EngineError as e:
        raise HTTPException(502, str(e)) from e
    if req.output.value == "number":
        return AIFillResponse(values=[_to_number(a) for a in answers])
    return AIFillResponse(values=[a.strip() or None for a in answers])


@router.post("/ai-eval", response_model=AIEvalResponse)
async def ai_eval(req: AIEvalRequest, engine: WorkbookEngine = Depends(get_engine),
                  caller: AICaller = Depends(ai_caller)):
    """Results for the =AI(prompt, inputs...) cells in a workbook, keyed by the caller's keys."""
    if not req.items:
        return AIEvalResponse(results={})
    if len(req.items) > MAX_EVAL_ITEMS:
        raise HTTPException(422, f"Up to {MAX_EVAL_ITEMS} AI cells can be calculated at a time.")
    limit = get_settings().max_prompt_chars
    items: dict[str, dict] = {}
    for it in req.items:
        if len(it.prompt) > limit:
            raise HTTPException(422, f"Keep each AI prompt under {limit} characters.")
        items[it.key] = {"key": it.key, "prompt": it.prompt.strip(), "inputs": [_cell_text(v) for v in it.inputs[:50]]}
    if sum(len(i["prompt"]) + sum(len(v) for v in i["inputs"]) for i in items.values()) > MAX_AI_INPUT_CHARS:
        raise HTTPException(422, "Those AI cells hold too much text for one request. Calculate fewer at a time.")
    answerable = [i for i in items.values() if i["prompt"]]
    results = {k: "" for k in items}
    if answerable:
        caller.spend()
        try:
            found = await engine.evaluate(answerable)
        except EngineError as e:
            raise HTTPException(502, str(e)) from e
        results.update({k: v.strip() for k, v in found.items()})
    return AIEvalResponse(results=results)


MAX_IMAGE_BYTES = 5 * 1024 * 1024  # the model's per-image limit; larger photos are shrunk when Pillow is installed


def _sniff_media_type(data: bytes) -> str | None:
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if data.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "image/gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    if data[:1024].lstrip().startswith(b"%PDF-"):
        return "application/pdf"
    return None


def _shrink_image(data: bytes) -> bytes | None:
    """Re-encode a large photo as a JPEG under the model's size limit. None when Pillow isn't available."""
    try:
        from PIL import Image
    except ImportError:
        return None
    try:
        img = Image.open(io.BytesIO(data))
        img = img.convert("RGB")
        for side in (4000, 3000, 2400, 1800):
            img.thumbnail((side, side))
            buf = io.BytesIO()
            img.save(buf, "JPEG", quality=85)
            if buf.tell() <= MAX_IMAGE_BYTES:
                return buf.getvalue()
    except Exception:  # noqa: BLE001 - unreadable image: let the caller report it
        return None
    return None


@router.post("/extract", response_model=WorkbookResponse)
async def extract(file: UploadFile = File(...), hint: str | None = Form(default=None),
                  engine: WorkbookEngine = Depends(get_engine), caller: AICaller = Depends(ai_caller)):
    """Turn a photo, scan or PDF of a table into a workbook (one sheet per table found, up to 5)."""
    data = await file.read(MAX_UPLOAD_BYTES + 1)
    if len(data) > MAX_UPLOAD_BYTES:
        raise HTTPException(413, "Files can be at most 10 MB.")
    if not data:
        raise HTTPException(422, "That file is empty.")
    media_type = _sniff_media_type(data)  # trust the bytes, not the file name or browser's content type
    if not media_type:
        raise HTTPException(415, "Please upload a photo (PNG, JPEG, WebP or GIF) or a PDF.")
    if media_type != "application/pdf" and len(data) > MAX_IMAGE_BYTES:
        smaller = _shrink_image(data)
        if smaller is None:
            raise HTTPException(413, "Photos can be at most 5 MB. Try a smaller or cropped picture.")
        data, media_type = smaller, "image/jpeg"
    hint = (hint or "").strip()
    if len(hint) > get_settings().max_prompt_chars:
        raise HTTPException(422, f"Please keep the hint under {get_settings().max_prompt_chars} characters.")
    caller.spend()
    try:
        result = await engine.extract(data, media_type, hint or None)
        workbook = tables_to_workbook(result)
    except EngineError as e:
        raise HTTPException(502, str(e)) from e
    except ImportError_ as e:
        raise HTTPException(422, str(e)) from e
    issues = validate_workbook(workbook)
    rows = sum(len(s.rows) for s in workbook.sheets)
    n = len(workbook.sheets)
    summary = (result.summary or "").strip() or f"Read {n} {'table' if n == 1 else 'tables'} from {file.filename}."
    summary += f" ({n} {'sheet' if n == 1 else 'sheets'}, {rows} rows.)"
    return WorkbookResponse(workbook=workbook, summary=summary, issues=issues)


@router.get("/capabilities")
async def capabilities():
    return {"pdf": pdf_available()}


# ---------------------------------------------------------------------------
# Templates
# ---------------------------------------------------------------------------


@lru_cache
def _templates() -> dict[str, dict]:
    out = {}
    for path in sorted(TEMPLATES_DIR.glob("*.json")):
        t = json.loads(path.read_text())
        out[t["id"]] = t
    return out


@templates_router.get("")
async def list_templates():
    return [
        {k: t[k] for k in ("id", "title", "description", "category")} | {"sheets": len(t["workbook"]["sheets"])}
        for t in _templates().values()
    ]


@templates_router.get("/{template_id}", response_model=WorkbookSpec)
async def get_template(template_id: str):
    t = _templates().get(template_id)
    if not t:
        raise HTTPException(404, "No such template.")
    return t["workbook"]
