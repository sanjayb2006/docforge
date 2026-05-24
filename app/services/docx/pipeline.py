"""
app/services/docx/pipeline.py

Background document processing pipeline.

Called by the upload endpoint via FastAPI BackgroundTasks.
Runs after a file is saved to disk and a Document row exists in the DB.

Steps:
  1. parse_docx()              → full DocumentStructure (sections, tables, images, breaks)
  2. extract_style_profile()   → StyleProfile (fonts, margins, heading styles, TOC)
  3. Persist both to DB        → Document.status = PARSED

All exceptions are caught and stored in Document.parse_error so the
upload endpoint always returns 201 immediately regardless of parse speed.
"""

from __future__ import annotations

import logging
import uuid
from pathlib import Path

from sqlalchemy.ext.asyncio import AsyncSession

from app.models.document import Document, DocumentStatus
from app.services.docx.parser import parse_docx
from app.services.docx.style_extractor import extract_style_profile

log = logging.getLogger(__name__)


async def process_document(
    doc_id:    uuid.UUID,
    file_path: Path,
    db:        AsyncSession,
) -> None:
    """
    Parse a DOCX and persist the results to the Document row.

    Args:
        doc_id:    UUID of the Document row to update
        file_path: Absolute path to the saved upload on disk
        db:        Async SQLAlchemy session (shared from the request)
    """
    doc: Document | None = await db.get(Document, doc_id)
    if doc is None:
        log.error("process_document: Document %s not found", doc_id)
        return

    try:
        log.info("Processing document %s → %s", doc_id, file_path.name)

        # ── Step 1: Structure ─────────────────────────────────────────────────
        structure = parse_docx(file_path)
        log.info(
            "Parsed %d sections, %d tables, %d images from '%s'",
            len(structure.get("sections", [])),
            structure.get("table_count", 0),
            structure.get("image_count", 0),
            file_path.name,
        )

        # ── Step 2: Style profile ─────────────────────────────────────────────
        style_profile = extract_style_profile(file_path)
        log.info(
            "Style profile: %d heading styles, %d fonts, page_numbers=%s",
            len(style_profile.get("heading_styles", {})),
            len(style_profile.get("fonts_in_use", [])),
            style_profile.get("header_footer", {}).get("has_page_numbers"),
        )

        # ── Step 3: Persist ───────────────────────────────────────────────────
        doc.structure     = structure
        doc.style_profile = style_profile
        doc.status        = DocumentStatus.PARSED
        doc.parse_error   = None
        await db.commit()

        log.info("Document %s → PARSED", doc_id)

    except Exception as exc:
        log.exception("process_document failed for %s: %s", doc_id, exc)
        doc.status      = DocumentStatus.PARSE_FAILED
        doc.parse_error = str(exc)
        await db.commit()
