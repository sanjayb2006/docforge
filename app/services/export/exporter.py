"""
app/services/export/exporter.py

PDF export via LibreOffice headless.

LibreOffice is the gold-standard for DOCX→PDF on Linux — it honours Word
formatting (fonts, margins, spacing) without requiring a Windows environment.

Behaviour:
  - Converts DOCX to PDF in the same output directory
  - Caches the PDF on disk (subsequent requests serve the cached file)
  - Returns None gracefully if LibreOffice is not installed
  - Times out after settings.GENERATION_TIMEOUT_SEC seconds

Install LibreOffice on Ubuntu:
    sudo apt-get install -y libreoffice
"""

from __future__ import annotations

import asyncio
import logging
import shutil
from pathlib import Path

from app.config import settings

log = logging.getLogger(__name__)

# Candidate binary paths (checked in order)
_SOFFICE_CANDIDATES = [
    "soffice",
    "libreoffice",
    "/usr/bin/soffice",
    "/usr/bin/libreoffice",
    "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]

_soffice_path: str | None = None  # cached after first lookup


def _find_soffice() -> str | None:
    global _soffice_path
    if _soffice_path is not None:
        return _soffice_path
    for candidate in _SOFFICE_CANDIDATES:
        if shutil.which(candidate):
            _soffice_path = candidate
            log.info("LibreOffice found at: %s", _soffice_path)
            return _soffice_path
    log.warning("LibreOffice not found — PDF export disabled")
    return None


def pdf_available() -> bool:
    """True if LibreOffice is installed and accessible."""
    return _find_soffice() is not None


async def convert_to_pdf(docx_path: Path) -> Path | None:
    """
    Convert a DOCX file to PDF using LibreOffice headless.

    Args:
        docx_path: Path to the input DOCX

    Returns:
        Path to the generated PDF, or None if unavailable / failed.
    """
    soffice = _find_soffice()
    if soffice is None:
        return None

    output_dir   = docx_path.parent
    expected_pdf = output_dir / docx_path.with_suffix(".pdf").name

    cmd = [
        soffice,
        "--headless",
        "--norestore",
        "--convert-to", "pdf",
        "--outdir", str(output_dir),
        str(docx_path),
    ]
    log.info("PDF conversion: %s", " ".join(cmd))

    try:
        proc = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE,
            stderr=asyncio.subprocess.PIPE,
        )
        stdout, stderr = await asyncio.wait_for(
            proc.communicate(),
            timeout=settings.GENERATION_TIMEOUT_SEC,
        )

        if proc.returncode != 0:
            log.error(
                "LibreOffice exited %d: %s",
                proc.returncode,
                stderr.decode(errors="replace")[:500],
            )
            return None

        if expected_pdf.exists():
            log.info("PDF created: %s (%d bytes)", expected_pdf.name, expected_pdf.stat().st_size)
            return expected_pdf

        # Fallback: fuzzy name match (LibreOffice may lowercase the stem)
        for f in output_dir.glob("*.pdf"):
            if f.stem.lower() == docx_path.stem.lower():
                log.info("PDF found (fuzzy): %s", f)
                return f

        log.error("PDF not found after LibreOffice conversion")
        return None

    except asyncio.TimeoutError:
        log.error("LibreOffice timed out after %d s", settings.GENERATION_TIMEOUT_SEC)
        return None
    except Exception as exc:
        log.exception("PDF conversion error: %s", exc)
        return None
