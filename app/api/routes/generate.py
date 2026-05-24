"""
app/api/routes/generate.py

AI generation job endpoints.

POST /api/generate/{doc_id}             — create a generation job
GET  /api/generate/{job_id}/status      — poll job status
GET  /api/generate/document/{doc_id}    — list all jobs for a document
"""

from __future__ import annotations

import logging
import uuid
from typing import Annotated

from fastapi import APIRouter, BackgroundTasks, Depends, HTTPException, status
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.database import get_db
from app.models.document import Document, DocumentStatus, GenerationJob, JobStatus
from app.schemas.document import GenerateRequest, GenerateResponse, JobStatusResponse
from app.services.ai.rewrite_pipeline import run_rewrite_job

log    = logging.getLogger(__name__)
router = APIRouter(prefix="/api/generate", tags=["Generation"])


# ── Create job ─────────────────────────────────────────────────────────────────

@router.post(
    "/{doc_id}",
    response_model=GenerateResponse,
    status_code=status.HTTP_202_ACCEPTED,
    summary="Trigger AI generation for a document",
    description=(
        "Creates a generation job and returns immediately with a job_id.\n\n"
        "Poll GET /api/generate/{job_id}/status until status is 'completed', "
        "then download from GET /api/export/{job_id}."
    ),
)
async def create_generation_job(
    doc_id:           uuid.UUID,
    request:          GenerateRequest,
    background_tasks: BackgroundTasks,
    db:               Annotated[AsyncSession, Depends(get_db)],
) -> GenerateResponse:
    # 1. Load and validate document
    doc: Document | None = await db.get(Document, doc_id)
    if doc is None:
        raise HTTPException(404, "Document not found")

    if doc.status == DocumentStatus.UPLOADED:
        raise HTTPException(
            status.HTTP_409_CONFLICT,
            "Document is still being parsed. Wait for status='parsed'.",
        )
    if doc.status == DocumentStatus.PARSE_FAILED:
        raise HTTPException(
            status.HTTP_422_UNPROCESSABLE_ENTITY,
            f"Document parsing failed: {doc.parse_error}",
        )

    # 2. Validate requested headings exist in document structure
    if not request.replace_all:
        available: set[str] = {
            s["heading_text"]
            for s in (doc.structure or {}).get("sections", [])
        }
        bad = [
            instr.heading
            for instr in request.sections
            if instr.heading not in available
        ]
        if bad:
            raise HTTPException(
                status.HTTP_422_UNPROCESSABLE_ENTITY,
                {
                    "message":            "One or more headings not found in document",
                    "invalid_headings":   bad,
                    "available_headings": sorted(available),
                },
            )

    # 3. Serialise instructions
    section_instructions = [
        {
            "heading":       instr.heading,
            "instruction":   instr.instruction,
            "extra_context": instr.extra_context,
        }
        for instr in request.sections
    ]

    # 4. Create job record
    job = GenerationJob(
        document_id  = doc_id,
        instructions = {
            "global_context": request.global_context,
            "sections":       section_instructions,
            "replace_all":    request.replace_all,
        },
        status = JobStatus.PENDING,
    )
    db.add(job)
    await db.commit()
    await db.refresh(job)

    # 5. Schedule background pipeline
    background_tasks.add_task(
        run_rewrite_job,
        job_id               = job.id,
        doc_id               = doc_id,
        section_instructions = section_instructions,
        global_context       = request.global_context,
        replace_all          = request.replace_all,
        db                   = db,
    )

    log.info("Generation job %s created for document %s", job.id, doc_id)
    return GenerateResponse(job_id=job.id, status=job.status.value)


# ── Status ─────────────────────────────────────────────────────────────────────

@router.get(
    "/{job_id}/status",
    response_model=JobStatusResponse,
    summary="Poll generation job status",
    description="Status flow: pending → running → completed | failed",
)
async def get_job_status(
    job_id: uuid.UUID,
    db:     Annotated[AsyncSession, Depends(get_db)],
) -> JobStatusResponse:
    job: GenerationJob | None = await db.get(GenerationJob, job_id)
    if job is None:
        raise HTTPException(404, "Job not found")
    return JobStatusResponse.model_validate(job)


# ── List jobs for document ─────────────────────────────────────────────────────

@router.get(
    "/document/{doc_id}",
    response_model=list[JobStatusResponse],
    summary="List all generation jobs for a document",
)
async def list_jobs_for_document(
    doc_id: uuid.UUID,
    db:     Annotated[AsyncSession, Depends(get_db)],
) -> list[JobStatusResponse]:
    doc = await db.get(Document, doc_id)
    if doc is None:
        raise HTTPException(404, "Document not found")

    result = await db.execute(
        select(GenerationJob)
        .where(GenerationJob.document_id == doc_id)
        .order_by(GenerationJob.created_at.desc())
    )
    return [JobStatusResponse.model_validate(j) for j in result.scalars().all()]
