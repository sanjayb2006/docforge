import httpx
from pathlib import Path
from docx import Document as Docx

BASE_URL = 'http://127.0.0.1:8002'

p = Path('scripts/api_debug.docx')
p.parent.mkdir(parents=True, exist_ok=True)
d = Docx()
d.add_heading('Introduction', level=1)
d.add_paragraph('This is the intro paragraph.')
d.save(p)

with httpx.Client(timeout=30) as client:
    files = {
        'file': (
            'api_debug.docx',
            p.read_bytes(),
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )
    }
    r = client.post(f'{BASE_URL}/api/documents/upload', files=files)
    print('upload', r.status_code, r.text)
    if r.status_code != 201:
        raise SystemExit(1)
    doc_id = r.json()['id']

    # wait for parse to complete
    for i in range(30):
        r_status = client.get(f'{BASE_URL}/api/documents/{doc_id}')
        if r_status.status_code != 200:
            print('status fetch failed', r_status.status_code, r_status.text)
            raise SystemExit(1)
        data = r_status.json()
        print('doc status', i, data.get('status'))
        if data.get('status') == 'parsed':
            break
        if data.get('status') == 'parse_failed':
            print('parse failed', data)
            raise SystemExit(1)
    else:
        print('parse did not complete')
        raise SystemExit(1)

    payload = {
        'global_context': '',
        'sections': [
            {'heading': 'Introduction', 'instruction': 'Write intro.'}
        ],
        'replace_all': False,
    }
    r2 = client.post(f'{BASE_URL}/api/generate/{doc_id}', json=payload)
    print('generate', r2.status_code, r2.text)
