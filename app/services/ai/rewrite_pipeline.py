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
  4. Snapshot original DOCX                         ← Step 0 instrumentation
  5. Call rebuild_docx() → output DOCX file
  6. Snapshot output DOCX + diff + log audit        ← Step 0 instrumentation
  7. validate_output() → sanity check
  8. Persist output_path + COMPLETED status

NO FUNCTIONAL CHANGES to the rebuild path.
The audit is observational only — it cannot cause a job to fail.
"""

from __future__ import annotations

import logging
import traceback
import uuid
from datetime import datetime
from pathlib import Path

from sqlalchemy.ext.asyncio import AsyncSession

from app.config import settings
from app.models.document import Document, GenerationJob, JobStatus
from app.services.ai.generator import generate_sections
from app.services.ai.sanitizer import sanitize_ai_results
from app.services.docx.rebuilder import rebuild_docx, validate_output as validate_rebuild
from app.services.docx.transformer import transform_docx, validate_output as validate_transform
from app.services.docx.rebuild_audit import (
    snapshot_document,
    diff_snapshots,
    log_audit,
    audit_to_dict,
)
from app.utils.file_utils import safe_output_path

log = logging.getLogger(__name__)


async def run_rewrite_job(
    job_id:               uuid.UUID,
    doc_id:               uuid.UUID,
    section_instructions: list[dict],
    global_context:       str,
    replace_all:          bool,
    db:                   AsyncSession,
    use_transformer:      bool = True,       # ← Step 1: opt-in, default ON
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
        use_transformer:      If True, use in-place transformer (high fidelity).
                              If False, use legacy rebuilder (semantic reconstruction).
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

        # ── Step 1b: Sanitize AI content ──────────────────────────────────────
        # Remove all pollution/debug/metadata content before transformer injection
        log.info("Job %s: sanitizing generated content", job_id)
        ai_results = sanitize_ai_results(ai_results, structure=structure)
        log.info("Job %s: sanitization complete — %d sections ready for injection", job_id, len(ai_results))

        # Persist intermediate results so they survive a crash
        job.ai_results = ai_results
        await db.commit()

        # ── Step 2: Snapshot ORIGINAL before rebuild ──────────────────────────
        # Zero-risk: read-only snapshot of original file
        original_snapshot = None
        try:
            original_snapshot = snapshot_document(original_path, label="original")
        except Exception as snap_exc:
            # Audit failures must NEVER block the rebuild
            log.warning("Job %s: pre-rebuild snapshot failed (non-fatal): %s", job_id, snap_exc)

        # ── Step 3: Build DOCX (transformer or rebuilder) ─────────────────────
        output_path = safe_output_path(str(job_id), suffix=".docx")
        engine_name = "transformer" if use_transformer else "rebuilder"
        log.info("Job %s: building DOCX via %s → %s", job_id, engine_name, output_path)

        if use_transformer:
            transform_docx(
                original_path=original_path,
                output_path=output_path,
                structure=structure,
                ai_results=ai_results,
                style_profile=style_profile,
            )
            validate_fn = validate_transform
        else:
            rebuild_docx(
                original_path=original_path,
                output_path=output_path,
                structure=structure,
                ai_results=ai_results,
                style_profile=style_profile,
            )
            validate_fn = validate_rebuild

        # ── Step 4: Snapshot OUTPUT + diff + log audit ────────────────────────
        # This is the core of Step 0: compare before/after and emit findings.
        rebuild_audit_dict: dict | None = None
        try:
            if original_snapshot is not None and output_path.exists():
                output_snapshot = snapshot_document(output_path, label="output")

                # Which sections had AI content (text legitimately differs)
                ai_section_headings = set(ai_results.keys())

                audit = diff_snapshots(
                    original=original_snapshot,
                    output=output_snapshot,
                    ai_sections=ai_section_headings,
                )

                # Emit all findings to server logs — grep for [AUDIT]
                log_audit(audit)

                # Serialise for optional storage
                rebuild_audit_dict = audit_to_dict(audit)

                log.info(
                    "Job %s: rebuild audit complete — severity=%s  "
                    "para_delta=%+d  style_mismatches=%d  spacing_anomalies=%d  "
                    "tables_lost=%d  images_lost=%d  headings_missing=%d",
                    job_id,
                    audit.severity,
                    audit.para_count_delta,
                    audit.style_mismatches,
                    audit.spacing_anomalies,
                    audit.tables_lost,
                    audit.images_lost,
                    len(audit.headings_missing),
                )
        except Exception as audit_exc:
            # Audit failures must NEVER block completion
            log.warning("Job %s: post-rebuild audit failed (non-fatal): %s", job_id, audit_exc)
            log.debug(traceback.format_exc())

        # ── Step 5: Validate ──────────────────────────────────────────────────
        if not validate_fn(output_path):
            raise RuntimeError(
                f"Output DOCX failed validation via {engine_name} "
                "(0 paragraphs or unreadable)"
            )

        # ── Step 6: Persist success ───────────────────────────────────────────
        job.output_path  = str(output_path)
        job.status       = JobStatus.COMPLETED
        job.completed_at = datetime.utcnow()

        # Attach audit data to job if available — surfaces in /api/export/{id}/preview
        if rebuild_audit_dict is not None:
            existing = job.ai_results or {}
            rebuild_audit_dict["engine"] = engine_name   # ← which path was used
            existing["__rebuild_audit__"] = rebuild_audit_dict
            job.ai_results = existing

        await db.commit()

        log.info(
            "Job %s: COMPLETED → %s (%d bytes)",
            job_id, output_path.name, output_path.stat().st_size,
        )

    except Exception as exc:
        log.exception("Job %s FAILED: %s", job_id, exc)
        job.status       = JobStatus.FAILED
        job.error        = f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}"
        job.completed_at = datetime.utcnow()
        await db.commit()
