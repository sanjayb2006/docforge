"""
app/core/database.py

Async SQLAlchemy engine, session factory, and declarative base.
"""

from sqlalchemy.ext.asyncio import (
    AsyncSession,
    async_sessionmaker,
    create_async_engine,
)
from sqlalchemy.orm import DeclarativeBase

from app.config import settings

# ── Engine ─────────────────────────────────────────────────────────────────────

engine = create_async_engine(
    "sqlite+aiosqlite:///./docforge.db",
    echo=settings.DEBUG,
    pool_pre_ping=True,
)

# ── Session factory ────────────────────────────────────────────────────────────

AsyncSessionLocal = async_sessionmaker(
    bind=engine,
    class_=AsyncSession,
    expire_on_commit=False,
)

# ── Declarative base ───────────────────────────────────────────────────────────

class Base(DeclarativeBase):
    """All ORM models inherit from this."""
    pass

# ── Dependency ─────────────────────────────────────────────────────────────────

async def get_db() -> AsyncSession:
    """
    FastAPI dependency that yields an async database session.
    """
    async with AsyncSessionLocal() as session:
        try:
            yield session
            await session.commit()
        except Exception:
            await session.rollback()
            raise
        finally:
            await session.close()