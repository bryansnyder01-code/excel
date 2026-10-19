"""Cell comments (threads anchored to a sheet and cell), cell history and the activity feed.

Anyone who can open a doc (owner, editor or viewer) reads all of these and may comment.
After any comment change the doc's live room (on every instance) gets
{"type": "comments", "thread_id": ...} so open clients refetch.
@mentions: an `@person@example.com` token naming the owner or a member emails them a link to
the cell (in the background; logged instead when SMTP isn't set up).
"""

from __future__ import annotations

import logging
import re
from urllib.parse import quote

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, Query, Response
from fastapi.concurrency import run_in_threadpool
from pydantic import BaseModel

from app.api.accounts import DOC_NOT_FOUND, current_user, get_store
from app.config import get_settings
from app.services.formula_shift import cell_name, parse_cell
from app.services.mailer import send_or_log
from app.services.store import Store

log = logging.getLogger(__name__)

collab_router = APIRouter(prefix="/api/docs", tags=["collaboration"])

MAX_BODY = 2000
THREAD_NOT_FOUND = "That comment wasn't found. It may have been deleted."
BAD_CELL = "Give the cell as a reference like B5."
_MENTION = re.compile(r"(?<![\w.+-])@([^\s@]+@[^\s@]+\.[^\s@]+)")


def _role(store: Store, user: dict, doc_id: str) -> str:
    role = store.doc_role(user["id"], doc_id)
    if not role:
        raise HTTPException(404, DOC_NOT_FOUND)
    return role


def _cell(text: str) -> tuple[int, int]:
    rc = parse_cell(text or "")
    if not rc:
        raise HTTPException(422, BAD_CELL)
    return rc


async def _changed(doc_id: str, thread_id: str) -> None:
    from app.api.live import notify  # live imports accounts; keep the import graph simple

    await notify(doc_id, {"kind": "comments", "thread_id": thread_id})


# ---------------------------------------------------------------------------
# History and activity
# ---------------------------------------------------------------------------


@collab_router.get("/{doc_id}/history")
def cell_history(doc_id: str, sheet: str, cell: str, user: dict = Depends(current_user),
                 store: Store = Depends(get_store)):
    """Who changed one cell (Excel A1, e.g. B5; row 1 is the header row), newest first, at most 50."""
    _role(store, user, doc_id)
    r, c = _cell(cell)
    return store.cell_history(doc_id, sheet, r, c, 50)


@collab_router.get("/{doc_id}/activity")
def activity(doc_id: str, limit: int = Query(50, ge=1, le=200), user: dict = Depends(current_user),
             store: Store = Depends(get_store)):
    """Recent changes to the doc, newest first: cell edits, structural changes and whole-workbook replaces."""
    _role(store, user, doc_id)
    return store.activity(doc_id, limit)


# ---------------------------------------------------------------------------
# Comments
# ---------------------------------------------------------------------------


class CommentIn(BaseModel):
    sheet: str = ""
    cell: str = ""
    body: str = ""
    thread_id: str | None = None


class ResolveIn(BaseModel):
    resolved: bool


class EditIn(BaseModel):
    body: str = ""


def _clean_body(body: str) -> str:
    body = body.strip()
    if not body:
        raise HTTPException(422, "Write a comment first.")
    if len(body) > MAX_BODY:
        raise HTTPException(422, f"Keep comments under {MAX_BODY} characters.")
    return body


def mentions(body: str) -> list[str]:
    """Lower-cased emails mentioned as @name@example.com (trailing punctuation dropped), in order, once each."""
    out: list[str] = []
    for m in _MENTION.finditer(body):
        email = m.group(1).rstrip(".,;:!?)]}'\"").lower()
        if email not in out:
            out.append(email)
    return out


def _mail_mentions(store: Store, doc_id: str, author: dict, thread: dict, body: str,
                   background: BackgroundTasks) -> None:
    wanted = mentions(body)
    if not wanted:
        return
    people = {m["email"].lower(): m for m in store.members(doc_id)}
    doc = store.live_doc(doc_id)
    title = doc["title"] if doc else "a workbook"
    settings = get_settings()
    where = f"{thread['sheet']}!{thread['cell']}"
    link = f"{settings.app_url.rstrip('/')}/?doc={quote(doc_id, safe='')}&cell={quote(where, safe='')}"
    for email in wanted:
        person = people.get(email)
        if not person or person["id"] == author["id"]:
            continue
        background.add_task(
            send_or_log, settings, person["email"], f"{author['email']} mentioned you in {title}",
            f"{author['email']} mentioned you in a comment on {where} in \"{title}\":\n\n{body}\n\n"
            f"Open it: {link}\n")


@collab_router.get("/{doc_id}/comments")
def list_comments(doc_id: str, user: dict = Depends(current_user), store: Store = Depends(get_store)):
    _role(store, user, doc_id)
    return store.threads(doc_id)


@collab_router.post("/{doc_id}/comments")
async def add_comment(doc_id: str, body: CommentIn, background: BackgroundTasks,
                      user: dict = Depends(current_user), store: Store = Depends(get_store)):
    """Reply to `thread_id` (re-opening it), or start a thread on sheet + cell. Returns the thread."""
    await run_in_threadpool(_role, store, user, doc_id)
    text = _clean_body(body.body)
    if body.thread_id:
        thread_id = await run_in_threadpool(store.add_comment, doc_id, user["id"], text, None, None, body.thread_id)
        if not thread_id:
            raise HTTPException(404, THREAD_NOT_FOUND)
    else:
        sheet = body.sheet.strip()
        if not sheet or len(sheet) > 100:
            raise HTTPException(422, "Say which sheet the comment is on.")
        r, c = _cell(body.cell)
        thread_id = await run_in_threadpool(store.add_comment, doc_id, user["id"], text, sheet, cell_name(r, c))
    thread = await run_in_threadpool(store.thread, doc_id, thread_id)
    await run_in_threadpool(_mail_mentions, store, doc_id, user, thread, text, background)
    await _changed(doc_id, thread_id)
    return thread


@collab_router.patch("/{doc_id}/comments/{thread_id}")
async def resolve_thread(doc_id: str, thread_id: str, body: ResolveIn, user: dict = Depends(current_user),
                         store: Store = Depends(get_store)):
    await run_in_threadpool(_role, store, user, doc_id)
    if not await run_in_threadpool(store.resolve_thread, doc_id, thread_id, body.resolved):
        raise HTTPException(404, THREAD_NOT_FOUND)
    await _changed(doc_id, thread_id)
    return await run_in_threadpool(store.thread, doc_id, thread_id)


@collab_router.patch("/{doc_id}/comments/{thread_id}/{comment_id}")
async def edit_comment(doc_id: str, thread_id: str, comment_id: str, body: EditIn,
                       user: dict = Depends(current_user), store: Store = Depends(get_store)):
    """The author edits their comment's text (mentions aren't emailed again)."""
    await run_in_threadpool(_role, store, user, doc_id)
    text = _clean_body(body.body)
    author = await run_in_threadpool(store.comment_author, doc_id, thread_id, comment_id)
    if author is None:
        raise HTTPException(404, THREAD_NOT_FOUND)
    if author != user["id"]:
        raise HTTPException(403, "Only the person who wrote a comment can edit it.")
    await run_in_threadpool(store.edit_comment, doc_id, thread_id, comment_id, text)
    await _changed(doc_id, thread_id)
    return await run_in_threadpool(store.thread, doc_id, thread_id)


@collab_router.delete("/{doc_id}/comments/{thread_id}/{comment_id}", status_code=204)
async def delete_comment(doc_id: str, thread_id: str, comment_id: str, user: dict = Depends(current_user),
                         store: Store = Depends(get_store)):
    """The author or the doc's owner deletes a comment; a thread with no comments left goes too."""
    role = await run_in_threadpool(_role, store, user, doc_id)
    author = await run_in_threadpool(store.comment_author, doc_id, thread_id, comment_id)
    if author is None:
        raise HTTPException(404, THREAD_NOT_FOUND)
    if author != user["id"] and role != "owner":
        raise HTTPException(403, "Only the comment's author or the workbook's owner can delete it.")
    await run_in_threadpool(store.delete_comment, doc_id, thread_id, comment_id)
    await _changed(doc_id, thread_id)
    return Response(status_code=204)
