"""
app/main.py

DocForge FastAPI application.

Registered routes:
  POST   /api/documents/upload
  GET    /api/documents/
  GET    /api/documents/{id}
  DELETE /api/documents/{id}
  POST   /api/generate/{doc_id}
  GET    /api/generate/{job_id}/status
  GET    /api/generate/document/{doc_id}
  GET    /api/export/{job_id}
  GET    /api/export/{job_id}/pdf
  GET    /api/export/{job_id}/preview
  GET    /health
  GET    /
"""

from __future__ import annotations

import logging
import time
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.config import settings
from app.api.routes import documents, generate, export

# ── Logging ────────────────────────────────────────────────────────────────────

logging.basicConfig(
    level=logging.DEBUG if settings.DEBUG else logging.INFO,
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("docforge")


# ── Lifespan ───────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    log.info("=" * 50)
    log.info("DocForge starting up")
    log.info("Upload dir : %s", settings.UPLOAD_DIR)
    log.info("Output dir : %s", settings.OUTPUT_DIR)
    log.info("AI model   : %s", settings.OPENAI_MODEL)
    log.info("Debug      : %s", settings.DEBUG)
    log.info("=" * 50)
    yield
    log.info("DocForge shutting down")


# ── App ────────────────────────────────────────────────────────────────────────

app = FastAPI(
    title="DocForge",
    description=(
        "AI-powered documentation automation platform.\n\n"
        "Upload a DOCX template → generate submission-ready documents "
        "with formatting fully preserved."
    ),
    version="1.0.0",
    lifespan=lifespan,
    docs_url="/docs",
    redoc_url="/redoc",
)

# ── CORS ───────────────────────────────────────────────────────────────────────

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],        # Restrict in production to your frontend origin
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ── Request timing middleware ──────────────────────────────────────────────────

@app.middleware("http")
async def request_timing(request: Request, call_next):
    start = time.perf_counter()
    response = await call_next(request)
    elapsed_ms = (time.perf_counter() - start) * 1000
    response.headers["X-Response-Time-Ms"] = f"{elapsed_ms:.1f}"
    log.debug(
        "%s %s → %d  (%.0f ms)",
        request.method, request.url.path, response.status_code, elapsed_ms,
    )
    return response

# ── Global exception handler ───────────────────────────────────────────────────

@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    log.exception("Unhandled exception: %s %s", request.method, request.url.path)
    return JSONResponse(
        status_code=500,
        content={"detail": "Internal server error. Check server logs."},
    )

# ── Routers ────────────────────────────────────────────────────────────────────

app.include_router(documents.router)
app.include_router(generate.router)
app.include_router(export.router)

# ── System endpoints ───────────────────────────────────────────────────────────

@app.get("/health", tags=["System"], summary="Health check")
async def health():
    return {
        "status": "ok",
        "app": settings.APP_NAME,
        "version": "1.0.0",
        "upload_dir": str(settings.UPLOAD_DIR),
        "output_dir": str(settings.OUTPUT_DIR),
    }


@app.get("/", tags=["System"], summary="Root")
async def root():
    return {
        "message": "DocForge API is running",
        "docs": "/docs",
        "health": "/health",
    }
