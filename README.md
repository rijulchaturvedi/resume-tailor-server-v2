
# Resume Tailor Server — CORS like original

- CORS restricted to `chrome-extension://*` on `/tailor` (Flask-CORS)
- Accepts **JSON** (original flow) and **multipart** (v2 flow)
- Uses server-bundled `base_resume.docx` if no file is uploaded

Render:
- Build: `pip install -r requirements.txt`
- Start: `gunicorn app:app`
- Env: `OPENAI_API_KEY` (optional), `OPENAI_MODEL` (optional)
