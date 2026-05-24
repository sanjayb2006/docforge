"""
app/schemas/document.py

Pydantic v2 request and response schemas.
Kept thin — business logic lives in services, not schemas.
"""

from __future__ import annotations

import uuid
from datetime import datetime
from typing import Any

from pydantic import BaseModel, Field


# ── Document ───────────────────────────────────────────────────────────────────

class DocumentUploadResponse(BaseModel):
    id:         uuid.UUID
    filename:   str
    status:     str
    created_at: datetime

    model_config = {"from_attributes": True}


class DocumentDetail(BaseModel):
    id:            uuid.UUID
    filename:      str
    file_type:     str
    status:        str
    structure:     dict[str, Any] | None = None
    style_profile: dict[str, Any] | None = None
    parse_error:   str | None            = None
    created_at:    datetime

    model_config = {"from_attributes": True}


# ── Generation ─────────────────────────────────────────────────────────────────

class SectionInstruction(BaseModel):
    heading:       str           = Field(..., description="Exact heading text from parsed document")
    instruction:   str           = Field(..., description="What to write for this section")
    extra_context: str | None    = Field(None, description="Extra data, lab readings, measurements, etc.")


class GenerateRequest(BaseModel):
    global_context: str = Field(
        default="",
        description="Context applied to every section (subject name, student info, purpose)",
    )
    sections: list[SectionInstruction] = Field(
        ..., min_length=1,
        description="Per-section generation instructions",
    )
    replace_all: bool = Field(
        default=False,
        description="If True, regenerate ALL sections even without explicit instructions",
    )


class GenerateResponse(BaseModel):
    job_id: uuid.UUID
    status: str


class JobStatusResponse(BaseModel):
    id:           uuid.UUID
    status:       str
    error:        str | None      = None
    output_path:  str | None      = None
    created_at:   datetime
    completed_at: datetime | None = None

    model_config = {"from_attributes": True}
