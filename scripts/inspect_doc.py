import sqlite3
import json
from pathlib import Path

DB = Path('docforge.db')

def inspect(doc_id: str):
    if not DB.exists():
        print('DB not found:', DB)
        return
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    cur.execute('SELECT id, status, structure FROM documents WHERE id = ?', (doc_id,))
    row = cur.fetchone()
    print('row:', row)
    if not row:
        conn.close()
        return
    _, status, structure = row
    print('status:', status)
    try:
        print('structure (parsed):', json.loads(structure))
    except Exception:
        print('structure (raw):', structure)
    conn.close()

if __name__ == '__main__':
    import sys
    if len(sys.argv) < 2:
        print('Usage: inspect_doc.py <doc_id>')
    else:
        inspect(sys.argv[1])
