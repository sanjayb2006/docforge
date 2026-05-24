"""In-place SQLite migration to canonicalise UUIDs.

Converts 32-char hex UUIDs (no hyphens) to canonical hyphenated form
for these columns: documents.id, generation_jobs.id, generation_jobs.document_id.

Idempotent and safe: skips rows already in hyphenated form. Backs up
the DB file before making changes.
"""
import shutil
import sqlite3
import uuid
import sys
from pathlib import Path
from datetime import datetime
import string

DB = Path('docforge.db')
if not DB.exists():
    print('Database file not found:', DB)
    sys.exit(1)

# Backup
timestamp = datetime.utcnow().strftime('%Y%m%d%H%M%S')
bak = DB.with_name(f"{DB.name}.{timestamp}.bak")
shutil.copy2(DB, bak)
print('Backup created at', bak)

hexd = set(string.hexdigits)

def is_hex32(s: str) -> bool:
    return isinstance(s, str) and len(s) == 32 and all(c in hexd for c in s)

conn = sqlite3.connect(DB)
cur = conn.cursor()
try:
    # Safety: disable FK enforcement while we rewrite PKs and FKs in a single transaction
    cur.execute('PRAGMA foreign_keys=OFF')
    conn.isolation_level = None
    cur.execute('BEGIN')

    docs = cur.execute('SELECT id FROM documents').fetchall()
    doc_map = {}
    docs_updated = 0
    for (did,) in docs:
        if did is None:
            continue
        if '-' in did and len(did) == 36:
            continue
        if is_hex32(did):
            new = str(uuid.UUID(hex=did))
            # Check for conflict
            conflict = cur.execute('SELECT 1 FROM documents WHERE id = ?', (new,)).fetchone()
            if conflict:
                print('Conflict: target document id already exists, skipping', new)
                continue
            cur.execute('UPDATE documents SET id = ? WHERE id = ?', (new, did))
            doc_map[did] = new
            docs_updated += 1

    jobs = cur.execute('SELECT id, document_id FROM generation_jobs').fetchall()
    jobs_updated = 0
    jobs_docid_updated = 0
    for jid, jdoc in jobs:
        # update job id if needed
        if jid and is_hex32(jid):
            newj = str(uuid.UUID(hex=jid))
            conflict = cur.execute('SELECT 1 FROM generation_jobs WHERE id = ?', (newj,)).fetchone()
            if conflict:
                print('Conflict: target job id exists, skipping', newj)
            else:
                cur.execute('UPDATE generation_jobs SET id = ? WHERE id = ?', (newj, jid))
                jobs_updated += 1
                jid = newj

        # update job.document_id if it matches a rewritten document id
        if jdoc and jdoc in doc_map:
            cur.execute('UPDATE generation_jobs SET document_id = ? WHERE id = ?', (doc_map[jdoc], jid))
            jobs_docid_updated += 1

    cur.execute('COMMIT')
    cur.execute('PRAGMA foreign_keys=ON')

    # Verify results
    docs_len36 = cur.execute("SELECT COUNT(*) FROM documents WHERE LENGTH(id)=36 AND id LIKE '%-%'").fetchone()[0]
    docs_total = cur.execute('SELECT COUNT(*) FROM documents').fetchone()[0]
    jobs_len36 = cur.execute("SELECT COUNT(*) FROM generation_jobs WHERE LENGTH(id)=36 AND id LIKE '%-%'").fetchone()[0]
    jobs_total = cur.execute('SELECT COUNT(*) FROM generation_jobs').fetchone()[0]

    print('Migration summary:')
    print('  documents updated:', docs_updated)
    print('  generation_jobs ids updated:', jobs_updated)
    print('  generation_jobs document_id updated:', jobs_docid_updated)
    print(f'  documents with hyphenated ids: {docs_len36}/{docs_total}')
    print(f'  jobs with hyphenated ids: {jobs_len36}/{jobs_total}')

finally:
    conn.close()

print('Migration complete. Backup preserved at', bak)
