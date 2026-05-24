"""
app/api/routes/documents.py

Document management endpoints.

POST   /api/documents/upload    — upload DOCX/PDF, trigger background parse
GET    /api/documents/          — paginated list
GET    /api/documents/{doc_id}  — detail with structure + style profile
DELETE /api/documents/{doc_id}  — delete document + all jobs
"""

from __future__ import annotations

import logging
import uuid
from pathlib import Path
from typing import Annotated

import aiofiles
from fastapi import (
    APIRouter, BackgroundTasks, Depends,
    HTTPException, Query, UploadFile, status,
)
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.database import get_db
from app.models.document import Document, DocumentStatus
from app.schemas.document import DocumentDetail, DocumentUploadResponse
from app.services.docx.pipeline import process_document
from app.utils.file_utils import validate_upload, safe_upload_path

log    = logging.getLogger(__name__)
router = APIRouter(prefix="/api/documents", tags=["Documents"])

MAX_UPLOAD_BYTES = 50 * 1024 * 1024  # 50 MB hard limit


# ── Upload ─────────────────────────────────────────────────────────────────────

@router.post(
    "/upload",
    response_model=DocumentUploadResponse,
    status_code=status.HTTP_201_CREATED,
    summary="Upload a DOCX or PDF template",
    description=(
        "Accepts .docx or .pdf files up to 50 MB. "
        "Returns immediately with the document ID. "
        "Background parsing begins automatically — poll GET /api/documents/{id} "
        "until status is 'parsed' before triggering generation."
    ),
)
async def upload_document(
    file:             UploadFile,
    background_tasks: BackgroundTasks,
    db:               Annotated[AsyncSession, Depends(get_db)],
) -> DocumentUploadResponse:
    # 1. Validate type
    ext = validate_upload(file)

    # 2. Determine save path
    save_path = safe_upload_path(file.filename or "upload", ext)

    # 3. Stream to disk with size guard
    bytes_written = 0
    try:
        async with aiofiles.open(save_path, "wb") as out:
            while chunk := await file.read(65536):
                bytes_written += len(chunk)
                if bytes_written > MAX_UPLOAD_BYTES:
                    save_path.unlink(missing_ok=True)
                    raise HTTPException(
                        status_code=status.HTTP_413_REQUEST_ENTITY_TOO_LARGE,
                        detail="File exceeds 50 MB limit",
                    )
                await out.write(chunk)
    except HTTPException:
        raise
    except Exception as exc:
        save_path.unlink(missing_ok=True)
        log.exception("Failed to save upload: %s", exc)
        raise HTTPException(500, "Failed to save uploaded file")

    log.info("Saved upload: %s (%d bytes)", save_path.name, bytes_written)

    # 4. Create DB record
    doc = Document(
        filename  = file.filename or save_path.name,
        file_path = str(save_path),
        file_type = ext.lstrip("."),
        status    = DocumentStatus.UPLOADED,
    )
    db.add(doc)
    await db.commit()
    await db.refresh(doc)

    # 5. Schedule background parse
    background_tasks.add_task(process_document, doc.id, save_path, db)

    log.info("Document %s created, parse scheduled", doc.id)

    # Return a cleanly serialisable response ensuring `id` is a proper UUID string
    return {
        "id": str(doc.id),
        "filename": doc.filename,
        "status": str(doc.status),
        "created_at": doc.created_at,
    }


# ── List ───────────────────────────────────────────────────────────────────────

@router.get(
    "/",
    response_model=list[DocumentUploadResponse],
    summary="List all uploaded documents",
)
async def list_documents(
    db:    Annotated[AsyncSession, Depends(get_db)],
    skip:  int = Query(0,  ge=0),
    limit: int = Query(20, ge=1, le=100),
) -> list[DocumentUploadResponse]:
    result = await db.execute(
        select(Document)
        .order_by(Document.created_at.desc())
        .offset(skip)
        .limit(limit)
    )
    docs = result.scalars().all()
    # Ensure IDs are serialised as strings to avoid malformed UUIDs in responses
    return [
        {
            "id": str(d.id),
            "filename": d.filename,
            "status": str(d.status),
            "created_at": d.created_at,
        }
        for d in docs
    ]


# ── Detail ─────────────────────────────────────────────────────────────────────

@router.get(
    "/{doc_id}",
    response_model=DocumentDetail,
    summary="Get document detail including parsed structure and style profile",
)
async def get_document(
    doc_id: uuid.UUID,
    db:     Annotated[AsyncSession, Depends(get_db)],
) -> DocumentDetail:
    doc = await db.get(Document, doc_id)
    if doc is None:
        raise HTTPException(404, "Document not found")
    return DocumentDetail.model_validate(doc)


# ── Delete ─────────────────────────────────────────────────────────────────────

@router.delete(
    "/{doc_id}",
    status_code=status.HTTP_204_NO_CONTENT,
    summary="Delete a document and all its generation jobs",
)
async def delete_document(
    doc_id: uuid.UUID,
    db:     Annotated[AsyncSession, Depends(get_db)],
):
    doc = await db.get(Document, doc_id)
    if doc is None:
        raise HTTPException(404, "Document not found")

    # Remove file from disk
    try:
        Path(doc.file_path).unlink(missing_ok=True)
    except Exception:
        pass

    await db.delete(doc)
    await db.commit()
    log.info("Deleted document %s", doc_id)
