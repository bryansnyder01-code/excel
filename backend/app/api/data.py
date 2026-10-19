"""Live data: POST /api/data/fetch pulls a CSV, a public Google Sheet or a JSON API into columns and rows.

No sign-in needed; each client IP gets FETCH_LIMIT fetches per rolling hour (in memory).
"""

from __future__ import annotations

import threading
import time
from collections import defaultdict, deque
from typing import Literal

from fastapi import APIRouter, Depends, HTTPException, Request
from pydantic import BaseModel

from app.services.fetcher import FetchError, fetch_table, make_client

data_router = APIRouter(prefix="/api/data", tags=["data"])

FETCH_LIMIT = 60
FETCH_WINDOW = 60 * 60


class RateLimiter:
    """At most `limit` hits per key in a rolling window. In memory: resets on restart, which is fine here."""

    def __init__(self, limit: int, window: float):
        self.limit, self.window = limit, window
        self._hits: dict[str, deque[float]] = defaultdict(deque)
        self._lock = threading.Lock()

    def allow(self, key: str) -> bool:
        now = time.monotonic()
        with self._lock:
            hits = self._hits[key]
            while hits and hits[0] <= now - self.window:
                hits.popleft()
            if len(hits) >= self.limit:
                return False
            hits.append(now)
            if len(self._hits) > 10_000:  # forget idle clients
                for k in [k for k, v in self._hits.items() if not v]:
                    del self._hits[k]
            return True


fetch_limiter = RateLimiter(FETCH_LIMIT, FETCH_WINDOW)


async def fetch_http():
    """HTTP client for live data. Tests override this with an httpx MockTransport client."""
    async with make_client() as client:
        yield client


class FetchRequest(BaseModel):
    kind: Literal["csv", "gsheet", "json"]
    url: str
    path: str | None = None


class FetchResponse(BaseModel):
    columns: list[str]
    rows: list[list[str | float | None]]
    fetched_at: int
    truncated: bool = False


@data_router.post("/fetch", response_model=FetchResponse)
async def fetch(req: FetchRequest, request: Request, http=Depends(fetch_http)):
    from app.api.routes import _client_ip

    if not fetch_limiter.allow(_client_ip(request)):
        raise HTTPException(429, f"That's {FETCH_LIMIT} data refreshes this hour. Try again later.")
    try:
        result = await fetch_table(req.kind, req.url, req.path, client=http)
    except FetchError as e:
        raise HTTPException(422, str(e)) from e
    return FetchResponse(columns=result.columns, rows=result.rows, fetched_at=result.fetched_at,
                         truncated=result.truncated)
