import httpx
from pathlib import Path
from docx import Document as Docx
import time

BASE_URL = 'http://127.0.0.1:8002'

p = Path('scripts/full_e2e.docx')
p.parent.mkdir(parents=True, exist_ok=True)
d = Docx()
d.add_heading('Introduction', level=1)
d.add_paragraph('This is the intro paragraph.')
d.save(p)

with httpx.Client(timeout=60) as client:
    files = {
        'file': (
            'full_e2e.docx',
            p.read_bytes(),
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        )
    }
    r = client.post(f'{BASE_URL}/api/documents/upload', files=files)
    print('upload', r.status_code, r.text)
    r.raise_for_status()
    doc_id = r.json()['id']

    # wait for parse to complete
    for i in range(60):
        r_status = client.get(f'{BASE_URL}/api/documents/{doc_id}')
        print('status', i, r_status.status_code, r_status.text)
        r_status.raise_for_status()
        data = r_status.json()
        if data.get('status') == 'parsed':
            break
        if data.get('status') == 'parse_failed':
            raise RuntimeError('parse failed: ' + str(data))
        time.sleep(0.5)
    else:
        raise RuntimeError('document did not parse in time')

    payload = {
        'global_context': '',
        'sections': [
            {'heading': 'Introduction', 'instruction': 'Write intro.'}
        ],
        'replace_all': False,
    }
    r2 = client.post(f'{BASE_URL}/api/generate/{doc_id}', json=payload)
    print('generate', r2.status_code, r2.text)
    r2.raise_for_status()
    job_id = r2.json()['job_id']

    for i in range(60):
        r3 = client.get(f'{BASE_URL}/api/generate/{job_id}/status')
        print('job status', i, r3.status_code, r3.text)
        r3.raise_for_status()
        status = r3.json().get('status')
        if status == 'completed':
            break
        if status == 'failed':
            raise RuntimeError('job failed: ' + str(r3.json()))
        time.sleep(0.5)
    else:
        raise RuntimeError('job did not complete in time')

    r4 = client.get(f'{BASE_URL}/api/export/{job_id}')
    print('export', r4.status_code)
    r4.raise_for_status()
    out = Path('scripts/downloaded_export.docx')
    out.write_bytes(r4.content)
    print('saved', out, out.stat().st_size)
