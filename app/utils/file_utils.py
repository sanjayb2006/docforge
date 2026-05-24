"""
app/utils/file_utils.py

Centralised file handling helpers:
  - upload type/MIME validation
  - collision-safe filename generation
  - upload size enforcement
  - output path helpers
  - SHA-256 checksum
"""

from __future__ import annotations

import hashlib
import uuid
from pathlib import Path

from fastapi import HTTPException, UploadFile, status

from app.config import settings

ALLOWED_EXTENSIONS: set[str]  = {".docx", ".pdf"}
ALLOWED_MIME_TYPES: set[str]  = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "application/pdf",
}


def validate_upload(file: UploadFile) -> str:
    """
    Validate file extension and MIME type.
    Returns the normalised extension ('.docx' | '.pdf').
    Raises HTTP 422 on failure.
    """
    filename  = file.filename or ""
    ext       = Path(filename).suffix.lower()

    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=f"Unsupported file type '{ext}'. Allowed: {sorted(ALLOWED_EXTENSIONS)}",
        )

    content_type = file.content_type or ""
    # Be lenient: browsers sometimes send application/octet-stream
    if (
        content_type
        and content_type not in ALLOWED_MIME_TYPES
        and content_type != "application/octet-stream"
    ):
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=f"Unexpected content-type '{content_type}'.",
        )

    return ext


def safe_upload_path(original_filename: str, ext: str) -> Path:
    """
    Generate a collision-safe path inside UPLOAD_DIR.
    Format: <uuid_hex>_<sanitised_stem><ext>
    """
    stem      = Path(original_filename).stem
    safe_stem = "".join(c if c.isalnum() or c in "-_" else "_" for c in stem)[:60]
    name      = f"{uuid.uuid4().hex}_{safe_stem}{ext}"
    return settings.UPLOAD_DIR / name


def safe_output_path(job_id: str, suffix: str = ".docx") -> Path:
    """Deterministic output path for a generation job."""
    return settings.OUTPUT_DIR / f"{job_id}{suffix}"


def file_size_bytes(path: Path) -> int:
    return path.stat().st_size


def checksum_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()
