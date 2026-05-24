"""
app/api/routes/export.py

Document export and preview endpoints.

GET /api/export/{job_id}          — download rebuilt DOCX
GET /api/export/{job_id}/pdf      — download as PDF (LibreOffice conversion)
GET /api/export/{job_id}/preview  — JSON preview: sections, word counts, fidelity score
"""

from __future__ import annotations

import logging
import uuid
from pathlib import Path
from typing import Annotated

from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.responses import FileResponse
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.database import get_db
from app.models.document import Document, GenerationJob, JobStatus
from app.services.export.exporter import convert_to_pdf, pdf_available
from app.services.fidelity.scorer import score_documents

log    = logging.getLogger(__name__)
router = APIRouter(prefix="/api/export", tags=["Export"])


# ── Guard helper ───────────────────────────────────────────────────────────────

def _require_completed(job: GenerationJob | None, job_id: uuid.UUID) -> GenerationJob:
    """Raise appropriate HTTP error if job is not completed."""
    if job is None:
        raise HTTPException(404, "Job not found")
    if job.status in (JobStatus.PENDING, JobStatus.RUNNING):
        raise HTTPException(
            status.HTTP_409_CONFLICT,
            f"Job is still {job.status.value}. "
            f"Poll /api/generate/{job_id}/status.",
        )
    if job.status == JobStatus.FAILED:
        raise HTTPException(
            status.HTTP_422_UNPROCESSABLE_ENTITY,
            f"Job failed: {job.error}",
        )
    if not job.output_path:
        raise HTTPException(500, "Job completed but output_path is missing")

    output = Path(job.output_path)
    if not output.exists():
        raise HTTPException(404, "Output file no longer exists on disk")

    return job


# ── DOCX download ──────────────────────────────────────────────────────────────

@router.get(
    "/{job_id}",
    summary="Download the rebuilt DOCX",
    response_class=FileResponse,
)
async def download_docx(
    job_id: uuid.UUID,
    db:     Annotated[AsyncSession, Depends(get_db)],
) -> FileResponse:
    """Download the AI-generated, formatting-preserved DOCX file."""
    job = await db.get(GenerationJob, job_id)
    job = _require_completed(job, job_id)

    filename = f"docforge_{job_id}.docx"
    log.info("Serving DOCX: %s", filename)
    return FileResponse(
        path      = str(job.output_path),
        media_type= "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        filename  = filename,
        headers   = {"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ── PDF download ───────────────────────────────────────────────────────────────

@router.get(
    "/{job_id}/pdf",
    summary="Download the rebuilt document as PDF",
    response_class=FileResponse,
)
async def download_pdf(
    job_id: uuid.UUID,
    db:     Annotated[AsyncSession, Depends(get_db)],
) -> FileResponse:
    """
    Convert the output DOCX to PDF via LibreOffice and download.
    PDF is cached on disk after first conversion.
    Returns 503 if LibreOffice is not installed.
    """
    if not pdf_available():
        raise HTTPException(
            status.HTTP_503_SERVICE_UNAVAILABLE,
            "PDF export unavailable. Install LibreOffice on the server: "
            "sudo apt-get install -y libreoffice",
        )

    job = await db.get(GenerationJob, job_id)
    job = _require_completed(job, job_id)

    docx_path = Path(job.output_path)
    pdf_path  = docx_path.with_suffix(".pdf")

    if not pdf_path.exists():
        log.info("Converting job %s to PDF", job_id)
        pdf_path = await convert_to_pdf(docx_path)
        if pdf_path is None:
            raise HTTPException(500, "PDF conversion failed. Check server logs.")

    filename = f"docforge_{job_id}.pdf"
    log.info("Serving PDF: %s", filename)
    return FileResponse(
        path      = str(pdf_path),
        media_type= "application/pdf",
        filename  = filename,
        headers   = {"Content-Disposition": f'attachment; filename="{filename}"'},
    )


# ── Preview ────────────────────────────────────────────────────────────────────

@router.get(
    "/{job_id}/preview",
    summary="Preview output: section word counts + fidelity score",
)
async def preview_output(
    job_id: uuid.UUID,
    db:     Annotated[AsyncSession, Depends(get_db)],
) -> dict:
    """
    Returns a JSON preview of the generated document including:
      - Per-section word counts and content source (AI vs preserved)
      - Total word count
      - Fidelity score comparing original vs rebuilt
    """
    job = await db.get(GenerationJob, job_id)
    job = _require_completed(job, job_id)

    doc: Document | None = await db.get(Document, job.document_id)

    ai_results:   dict = job.ai_results  or {}
    instructions: dict = job.instructions or {}
    instructed    = {s["heading"] for s in instructions.get("sections", [])}

    preview_sections = []
    total_words      = 0

    for heading, content in ai_results.items():
        wc = len(content.split())
        total_words += wc
        preview_sections.append({
            "heading":    heading,
            "word_count": wc,
            "source":     "ai_generated" if heading in instructed else "auto_generated",
            "preview":    content[:200] + "..." if len(content) > 200 else content,
        })

    # Fidelity score (original vs rebuilt)
    fidelity = None
    if doc and doc.file_path:
        try:
            report   = score_documents(doc.file_path, job.output_path)
            fidelity = report.to_dict()
        except Exception as e:
            log.warning("Fidelity scoring failed: %s", e)

    return {
        "job_id":                  str(job_id),
        "status":                  job.status.value,
        "total_sections_generated": len(ai_results),
        "total_word_count":        total_words,
        "sections":                preview_sections,
        "fidelity":                fidelity,
    }
