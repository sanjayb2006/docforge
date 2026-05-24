from docx import Document as Docx
from pathlib import Path
import httpx

# create temp docx
p = Path('test_upload.docx')
d = Docx()
d.add_paragraph('Hello')
d.save(p)

with httpx.Client(timeout=30) as c:
    files = {'file': ('test_upload.docx', p.read_bytes(), 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')}
    r = c.post('http://127.0.0.1:8001/api/documents/upload', files=files)
    print('status', r.status_code)
    print('body', r.text)
