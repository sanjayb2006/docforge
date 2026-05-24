import sqlite3
from pathlib import Path

DB = Path('docforge.db')

def list_docs(limit=20):
    if not DB.exists():
        print('DB not found:', DB)
        return
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    cur.execute('SELECT id, filename, status, created_at FROM documents LIMIT ?', (limit,))
    rows = cur.fetchall()
    for r in rows:
        print(r)
    conn.close()

if __name__ == '__main__':
    list_docs()
