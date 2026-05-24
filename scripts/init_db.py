"""
scripts/init_db.py

Create all database tables directly (development / local use).
For production deployments use Alembic migrations.

Usage:
    python scripts/init_db.py
"""

import asyncio
import sys
from pathlib import Path

# Allow running from project root
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from app.core.database import engine, Base
from app.models.document import Document, GenerationJob  # registers models with Base


async def init() -> None:
    print(f"Creating tables in database...")
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)
    await engine.dispose()
    print("Done. Tables created:")
    for table in Base.metadata.sorted_tables:
        print(f"  ✓ {table.name}")


if __name__ == "__main__":
    asyncio.run(init())
