<<<<<<< HEAD
# DocForge — AI Documentation Automation Platform

AI-powered backend that takes a DOCX template and generates submission-ready
documents with formatting fully preserved.

---

## Project Structure

```
docforge/
├── .env.example                        # copy to .env and fill in
├── alembic.ini                         # Alembic migration config
├── requirements.txt
├── README.md
│
├── alembic/
│   ├── env.py                          # async migration runner
│   └── versions/                       # migration files (auto-generated)
│
├── app/
│   ├── main.py                         # FastAPI app + routers + middleware
│   ├── config.py                       # settings (pydantic-settings + .env)
│   │
│   ├── api/routes/
│   │   ├── documents.py                # upload, list, detail, delete
│   │   ├── generate.py                 # create job, poll status, list jobs
│   │   └── export.py                   # download DOCX, PDF, preview + fidelity
│   │
│   ├── core/
│   │   └── database.py                 # async engine, session factory, Base
│   │
│   ├── models/
│   │   └── document.py                 # Document + GenerationJob ORM models
│   │
│   ├── schemas/
│   │   └── document.py                 # Pydantic request/response schemas
│   │
│   ├── services/
│   │   ├── docx/
│   │   │   ├── parser.py               # full structure extraction (v2)
│   │   │   ├── style_extractor.py      # inheritance-resolved style profiles (v2)
│   │   │   ├── rebuilder.py            # XML-level DOCX reconstruction (v2)
│   │   │   └── pipeline.py             # background parse orchestrator
│   │   │
│   │   ├── ai/
│   │   │   ├── prompt_builder.py       # section + bulk prompt construction
│   │   │   ├── generator.py            # OpenAI API calls
│   │   │   └── rewrite_pipeline.py     # end-to-end job runner (BackgroundTask)
│   │   │
│   │   ├── export/
│   │   │   └── exporter.py             # LibreOffice DOCX→PDF conversion
│   │   │
│   │   └── fidelity/
│   │       └── scorer.py               # 5-dimension fidelity scoring system
│   │
│   └── utils/
│       └── file_utils.py               # upload validation, path helpers
│
├── scripts/
│   └── init_db.py                      # create tables (dev use)
│
└── tests/
    ├── test_docx_pipeline.py           # parser + extractor + rebuilder unit tests
    └── test_stress.py                  # 36 stress tests across 10 real-world scenarios
```

---

## Quick Start

```bash
# 1. Clone and install
git clone <repo>
cd docforge
pip install -r requirements.txt

# 2. Configure
cp .env.example .env
# Edit .env: set DATABASE_URL, OPENAI_API_KEY

# 3. Database
python scripts/init_db.py

# 4. Run
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

Swagger UI → http://localhost:8000/docs

---

## API Workflow

```
1. POST /api/documents/upload
   → Upload a DOCX template
   → Returns: { id, status: "uploaded" }

2. GET /api/documents/{id}
   → Poll until status = "parsed"
   → Returns: structure (sections, tables, images) + style_profile

3. POST /api/generate/{doc_id}
   Body: {
     "global_context": "VTU lab report, BCSL657D DevOps",
     "sections": [
       { "heading": "1. Aim", "instruction": "Write aim for Jenkins CI/CD" },
       { "heading": "4. Result", "instruction": "Document successful pipeline run" }
     ]
   }
   → Returns: { job_id, status: "pending" }

4. GET /api/generate/{job_id}/status
   → Poll until status = "completed"

5. GET /api/export/{job_id}          → download DOCX
   GET /api/export/{job_id}/pdf      → download PDF (needs LibreOffice)
   GET /api/export/{job_id}/preview  → JSON with word counts + fidelity score
```

---

## Fidelity Score

Every rebuilt document is scored across 5 dimensions:

| Dimension           | Weight | Measures |
|---------------------|--------|----------|
| Structure Fidelity  | 30%    | Headings present, order correct, levels match |
| Formatting Fidelity | 25%    | Margins, body font, body size |
| Content Completeness| 20%    | Word count ratio, paragraph count |
| Element Preservation| 15%    | Tables and images survived |
| Style Consistency   | 10%    | Uniform fonts/sizes across body paragraphs |

Grade A+ (≥95) = submission-ready, no manual correction expected.

---

## What Is Preserved Through Rebuild

- ✅ All heading styles (font, size, bold, color, spacing)
- ✅ Body paragraph spacing (before/after/line)
- ✅ Paragraph indentation (left, right, first-line, hanging)
- ✅ Page margins and page size (from sectPr — untouched)
- ✅ Tables (verbatim XML clone — borders, shading, merged cells)
- ✅ Images (verbatim XML clone — relationship IDs preserved)
- ✅ Headers and footers (XML clone from original sections)
- ✅ TOC paragraphs (copied verbatim)
- ✅ Per-run formatting (bold, italic, font, size, color per run)
- ✅ Page breaks (re-inserted where original had them)
- ✅ Custom fonts (inventoried and flagged)
- ✅ Unicode and multi-script content

---

## PDF Export

Requires LibreOffice on the server:

```bash
sudo apt-get install -y libreoffice
```

PDFs are cached on disk after first conversion. Subsequent requests serve
the cached file.

---

## Running Tests

```bash
pytest tests/ -v
# 51 tests: 15 unit + 36 stress
```
=======
# docforge
>>>>>>> d91a98a86ae4ddb6ee2a9a6dac29fa5c9342d871
