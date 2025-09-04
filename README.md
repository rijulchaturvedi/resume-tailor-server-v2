# Resume Tailor Server (v2)
- GET /health -> {"status":"ok"}
- POST /tailor: multipart with
  - base_resume: .docx file (from extension)
  - payload: JSON { job_description, resume_config{summary_sentences,experience[],skills_categories[]}, options.style_rules[] }
Returns tailored .docx.

Render:
- Build: pip install -r requirements.txt
- Start: gunicorn app:app
- Env: OPENAI_API_KEY, (optional) OPENAI_MODEL=gpt-4o-mini
