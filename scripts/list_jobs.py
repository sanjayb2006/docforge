import sqlite3
from pathlib import Path

DB = Path('docforge.db')

def list_jobs(limit=20):
    if not DB.exists():
        print('DB not found:', DB)
        return
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    cur.execute('SELECT id, document_id, status, output_path FROM generation_jobs LIMIT ?', (limit,))
    rows = cur.fetchall()
    for r in rows:
        print(r)
    conn.close()

if __name__ == '__main__':
    list_jobs()
