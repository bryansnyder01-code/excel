"""Signed-in users' own templates: save a workbook as a template, list, open and delete them.

Routes take `Authorization: Bearer <token>`; errors are {"detail": "<user-facing sentence>"}.
"""

from __future__ import annotations

import json

from fastapi import APIRouter, Depends, HTTPException, Request, Response
from fastapi.concurrency import run_in_threadpool
from pydantic import BaseModel, ValidationError

from app.api.accounts import current_user
from app.schemas import WorkbookSpec
from app.services.user_templates import TemplateLimit, UserTemplates, get_user_templates

my_templates_router = APIRouter(prefix="/api/templates/mine", tags=["templates"])

MAX_TEMPLATE_BYTES = 5 * 1024 * 1024
MAX_TEMPLATES = 50
TEMPLATE_NOT_FOUND = "Template not found."


class TemplateIn(BaseModel):
    title: str = ""
    description: str = ""
    workbook: WorkbookSpec


@my_templates_router.get("")
def list_my_templates(user: dict = Depends(current_user), store: UserTemplates = Depends(get_user_templates)):
    return store.list(user["id"])


@my_templates_router.post("")
async def save_my_template(request: Request, user: dict = Depends(current_user),
                           store: UserTemplates = Depends(get_user_templates)):
    """Save a workbook as a template. The body is read by hand to enforce the size cap."""
    too_large = HTTPException(413, "This template is too large to save (5 MB max).")
    if int(request.headers.get("content-length") or 0) > MAX_TEMPLATE_BYTES:
        raise too_large
    raw = b""
    async for chunk in request.stream():
        raw += chunk
        if len(raw) > MAX_TEMPLATE_BYTES:
            raise too_large
    try:
        body = TemplateIn.model_validate(json.loads(raw))
    except (ValueError, ValidationError):
        raise HTTPException(422, "That template couldn't be read.") from None
    title = (body.title.strip() or body.workbook.title.strip() or "My template")[:200]
    try:
        template_id = await run_in_threadpool(
            store.add, user["id"], title, body.description.strip()[:1000], body.workbook.model_dump(mode="json"),
            MAX_TEMPLATES,
        )
    except TemplateLimit:
        raise HTTPException(422, f"You can keep up to {MAX_TEMPLATES} templates.") from None
    return {"id": template_id}


@my_templates_router.get("/{template_id}", response_model=WorkbookSpec)
def get_my_template(template_id: str, user: dict = Depends(current_user),
                    store: UserTemplates = Depends(get_user_templates)):
    workbook = store.get(user["id"], template_id)
    if workbook is None:
        raise HTTPException(404, TEMPLATE_NOT_FOUND)
    return workbook


@my_templates_router.delete("/{template_id}", status_code=204)
def delete_my_template(template_id: str, user: dict = Depends(current_user),
                       store: UserTemplates = Depends(get_user_templates)):
    if not store.delete(user["id"], template_id):
        raise HTTPException(404, TEMPLATE_NOT_FOUND)
    return Response(status_code=204)
