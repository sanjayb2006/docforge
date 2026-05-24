import asyncio
from pathlib import Path

import pytest
from httpx import AsyncClient

from app.main import app
from app.core.database import AsyncSessionLocal
from app.models.document import Document, DocumentStatus, GenerationJob, JobStatus
from app.config import settings


@pytest.mark.asyncio
async def test_end_to_end_generation(tmp_path, monkeypatch):
    # 1) create a small DOCX to upload
    docx_path = tmp_path / "sample.docx"
    from docx import Document as Docx

    d = Docx()
    d.add_paragraph("Test")
    d.save(docx_path)

    # 2) monkeypatch the heavy background pipeline to a lightweight stub
    async def fake_run_rewrite_job(*, job_id, doc_id, section_instructions, global_context, replace_all, db):
        # small delay to simulate work
        await asyncio.sleep(0.1)
        async with AsyncSessionLocal() as session:
            job = await session.get(GenerationJob, job_id)
            # create output file path
            out_path = settings.OUTPUT_DIR / f"{job_id}.docx"
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_bytes(b"DUMMY DOCX CONTENT")
            job.status = JobStatus.COMPLETED
            job.output_path = str(out_path)
            await session.commit()


    monkeypatch.setattr("app.services.ai.rewrite_pipeline.run_rewrite_job", fake_run_rewrite_job)

    async with AsyncClient(app=app, base_url="http://test") as client:
        # Upload file
        files = {"file": ("sample.docx", docx_path.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")}
        r = await client.post("/api/documents/upload", files=files)
        assert r.status_code == 201
        data = r.json()
        doc_id = data["id"]

        # mark document as parsed and give it a simple structure
        async with AsyncSessionLocal() as session:
            doc = await session.get(Document, doc_id)
            doc.status = DocumentStatus.PARSED
            doc.structure = {"sections": [{"heading_text": "Introduction"}]}
            await session.commit()

        # Trigger generation (request sections that exist)
        payload = {
            "global_context": "",
            "sections": [{"heading": "Introduction", "instruction": "Write intro."}],
            "replace_all": False,
        }
        r = await client.post(f"/api/generate/{doc_id}", json=payload)
        assert r.status_code == 202
        job_id = r.json()["job_id"]

        # Poll status
        for _ in range(50):
            r = await client.get(f"/api/generate/{job_id}/status")
            assert r.status_code == 200
            status = r.json().get("status")
            if status == "completed":
                break
            await asyncio.sleep(0.1)
        else:
            pytest.fail("Job did not complete in time")

        # Download generated file
        r = await client.get(f"/api/export/{job_id}")
        assert r.status_code == 200
        # write to tmp and validate
        out_file = tmp_path / f"out_{job_id}.docx"
        out_file.write_bytes(r.content)
        assert out_file.exists()
        assert out_file.stat().st_size > 0

        print("E2E flow succeeded: uploaded -> generated -> downloaded")
