"""
app/services/ai/rewrite_pipeline.py

End-to-end AI rewrite job runner (runs as a BackgroundTask).

State machine:
  PENDING → RUNNING → COMPLETED
                    → FAILED

Steps:
  1. Load Document (structure + style_profile) from DB
  2. Call generate_sections() → {heading: ai_text}
  3. Persist ai_results to DB (useful for debugging partial runs)
  4. Call rebuild_docx() → output DOCX file
  5. validate_output() → sanity check
  6. Persist output_path + COMPLETED status
"""

from __future__ import annotations

import logging
import uuid
from datetime import datetime
from pathlib import Path

from sqlalchemy.ext.asyncio import AsyncSession

from app.config import settings
from app.models.document import Document, GenerationJob, JobStatus
from app.services.ai.generator import generate_sections
from app.services.docx.rebuilder import rebuild_docx, validate_output
from app.utils.file_utils import safe_output_path

log = logging.getLogger(__name__)


async def run_rewrite_job(
    job_id:               uuid.UUID,
    doc_id:               uuid.UUID,
    section_instructions: list[dict],
    global_context:       str,
    replace_all:          bool,
    db:                   AsyncSession,
) -> None:
    """
    Background task: run the full AI rewrite pipeline for one GenerationJob.

    Args:
        job_id:               UUID of the GenerationJob row
        doc_id:               UUID of the source Document
        section_instructions: [{"heading": ..., "instruction": ..., "extra_context": ...}]
        global_context:       Applied to every section prompt
        replace_all:          Generate content for ALL sections
        db:                   Async SQLAlchemy session
    """
    job: GenerationJob | None = await db.get(GenerationJob, job_id)
    doc: Document      | None = await db.get(Document,       doc_id)

    if job is None or doc is None:
        log.error("run_rewrite_job: job=%s or doc=%s not found in DB", job_id, doc_id)
        return

    # ── Mark running ──────────────────────────────────────────────────────────
    job.status = JobStatus.RUNNING
    await db.commit()

    try:
        structure:     dict = doc.structure     or {}
        style_profile: dict = doc.style_profile or {}
        original_path: Path = Path(doc.file_path)

        if not original_path.exists():
            raise FileNotFoundError(f"Original file missing: {original_path}")

        # ── Step 1: AI generation ─────────────────────────────────────────────
        log.info("Job %s: generating %d section instructions", job_id, len(section_instructions))
        ai_results: dict[str, str] = await generate_sections(
            structure=structure,
            style_profile=style_profile,
            section_instructions=section_instructions,
            global_context=global_context,
            replace_all=replace_all,
        )
        log.info("Job %s: AI generated %d sections", job_id, len(ai_results))

        # Persist intermediate results so they survive a crash
        job.ai_results = ai_results
        await db.commit()

        # ── Step 2: Rebuild DOCX ──────────────────────────────────────────────
        output_path = safe_output_path(str(job_id), suffix=".docx")
        log.info("Job %s: rebuilding DOCX → %s", job_id, output_path)

        rebuild_docx(
            original_path=original_path,
            output_path=output_path,
            structure=structure,
            ai_results=ai_results,
            style_profile=style_profile,
        )

        # ── Step 3: Validate ──────────────────────────────────────────────────
        if not validate_output(output_path):
            raise RuntimeError("Rebuilt DOCX failed validation (0 paragraphs or unreadable)")

        # ── Step 4: Persist success ───────────────────────────────────────────
        job.output_path  = str(output_path)
        job.status       = JobStatus.COMPLETED
        job.completed_at = datetime.utcnow()
        await db.commit()

        log.info("Job %s: COMPLETED → %s (%d bytes)",
                 job_id, output_path.name, output_path.stat().st_size)

    except Exception as exc:
        log.exception("Job %s FAILED: %s", job_id, exc)
        job.status       = JobStatus.FAILED
        job.error        = str(exc)
        job.completed_at = datetime.utcnow()
        await db.commit()
