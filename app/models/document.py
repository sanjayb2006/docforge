"""
app/models/document.py

SQLAlchemy ORM models.

Document        — uploaded template file + parsed structure/style data
GenerationJob   — one AI rewrite run against a Document
"""

from __future__ import annotations

import enum
import uuid
from datetime import datetime

from sqlalchemy import DateTime, Enum, ForeignKey, JSON, String, Text
from sqlalchemy.dialects.postgresql import UUID as PG_UUID
from sqlalchemy.types import TypeDecorator, CHAR
from sqlalchemy.orm import Mapped, mapped_column, relationship

from app.core.database import Base


# ── Enums ──────────────────────────────────────────────────────────────────────

class DocumentStatus(str, enum.Enum):
    UPLOADED     = "uploaded"       # file saved, parse not yet started
    PARSED       = "parsed"         # structure + style extracted successfully
    PARSE_FAILED = "parse_failed"   # extraction failed (see parse_error)


class JobStatus(str, enum.Enum):
    PENDING   = "pending"    # job queued
    RUNNING   = "running"    # AI generation in progress
    COMPLETED = "completed"  # output DOCX written
    FAILED    = "failed"     # error (see error field)


# ── Document ───────────────────────────────────────────────────────────────────

class Document(Base):
    """
    One uploaded template document.

    structure     — DocumentStructure dict (sections, paragraphs, tables, images)
    style_profile — StyleProfile dict (fonts, margins, spacing, headings)
    """
    __tablename__ = "documents"

    # Cross-dialect GUID type that stores UUIDs as 36-char strings on SQLite
    class GUID(TypeDecorator):
        impl = CHAR
        cache_ok = True

        def load_dialect_impl(self, dialect):
            if dialect.name == "postgresql":
                return dialect.type_descriptor(PG_UUID(as_uuid=True))
            else:
                return dialect.type_descriptor(CHAR(36))

        def process_bind_param(self, value, dialect):
            if value is None:
                return None
            if isinstance(value, uuid.UUID):
                # store as canonical string
                return str(value)
            # allow string input but normalise to UUID canonical form
            return str(uuid.UUID(value))

        def process_result_value(self, value, dialect):
            if value is None:
                return None
            return uuid.UUID(value)

    id: Mapped[uuid.UUID] = mapped_column(
        GUID(), primary_key=True, default=uuid.uuid4
    )
    filename: Mapped[str]  = mapped_column(String(255), nullable=False)
    file_path: Mapped[str] = mapped_column(String(1024), nullable=False)
    file_type: Mapped[str] = mapped_column(String(10), nullable=False)   # docx | pdf

    # Populated by the background parse pipeline
    structure:     Mapped[dict | None] = mapped_column(JSON, nullable=True)
    style_profile: Mapped[dict | None] = mapped_column(JSON, nullable=True)

    status:      Mapped[DocumentStatus] = mapped_column(
        Enum(DocumentStatus), default=DocumentStatus.UPLOADED, nullable=False
    )
    parse_error: Mapped[str | None] = mapped_column(Text, nullable=True)

    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, nullable=False)
    updated_at: Mapped[datetime] = mapped_column(
        DateTime, default=datetime.utcnow, onupdate=datetime.utcnow, nullable=False
    )

    jobs: Mapped[list[GenerationJob]] = relationship(
        "GenerationJob", back_populates="document", cascade="all, delete-orphan"
    )


# ── GenerationJob ──────────────────────────────────────────────────────────────

class GenerationJob(Base):
    """
    One AI generation run.

    instructions — user-provided GenerateRequest (section instructions, context)
    ai_results   — {heading: generated_text} dict from the AI pipeline
    output_path  — absolute path to the rebuilt DOCX on disk
    """
    __tablename__ = "generation_jobs"

    id: Mapped[uuid.UUID] = mapped_column(
        Document.GUID(), primary_key=True, default=uuid.uuid4
    )
    document_id: Mapped[uuid.UUID] = mapped_column(
        Document.GUID(), ForeignKey("documents.id"), nullable=False
    )

    instructions: Mapped[dict]      = mapped_column(JSON, nullable=False)
    ai_results:   Mapped[dict | None] = mapped_column(JSON, nullable=True)
    output_path:  Mapped[str | None]  = mapped_column(String(1024), nullable=True)

    status: Mapped[JobStatus] = mapped_column(
        Enum(JobStatus), default=JobStatus.PENDING, nullable=False
    )
    error: Mapped[str | None] = mapped_column(Text, nullable=True)

    created_at:   Mapped[datetime]      = mapped_column(DateTime, default=datetime.utcnow, nullable=False)
    completed_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)

    document: Mapped[Document] = relationship("Document", back_populates="jobs")
