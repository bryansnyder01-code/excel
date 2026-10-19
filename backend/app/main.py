"""SheetCraft AI backend entry point.  Run with:  uvicorn app.main:app --reload"""

import logging

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from app.api.accounts import auth_router, docs_router, invites_router, shared_router
from app.api.live import live_router
from app.api.comments import collab_router
from app.api.billing import billing_router
from app.api.my_templates import my_templates_router
from app.api.routes import router, templates_router
from app.api.data import data_router
from app.api.reports import reports_router
from app.services.scheduler import install_background_jobs
from app.config import get_settings

logging.basicConfig(level=logging.INFO)

settings = get_settings()
app = FastAPI(title="SheetCraft AI", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["Content-Disposition"],
)
app.include_router(router)
app.include_router(my_templates_router)  # before templates_router: /api/templates/mine wins over /{template_id}
app.include_router(templates_router)
app.include_router(auth_router)
app.include_router(docs_router)
app.include_router(shared_router)
app.include_router(invites_router)
app.include_router(live_router)
app.include_router(collab_router)
app.include_router(billing_router)
app.include_router(data_router)
app.include_router(reports_router)
install_background_jobs(app)  # live data refresh + scheduled email reports, once a minute


@app.get("/health")
async def health():
    return {"status": "ok", "model": settings.model}


if settings.static_dir:
    # Last, so every /api route above wins over a file of the same name.
    app.mount("/", StaticFiles(directory=settings.static_dir, html=True), name="web")
