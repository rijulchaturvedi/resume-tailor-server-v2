from flask import Flask, request, send_file, jsonify, make_response
from flask_cors import CORS
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
import io, os, json, re, logging
from typing import List, Dict, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeout

# ----------------------------
# App & CORS (mirrors original)
# ----------------------------
app = Flask(__name__)
# Allow chrome extension origins to hit /tailor (OPTIONS + POST)
CORS(app, resources={r"/tailor": {"origins": "chrome-extension://*"}})
app.logger.setLevel(logging.INFO)

# ----------------------------
# Model / OpenAI settings
# ----------------------------
MODEL = os.getenv("OPENAI_MODEL", "gpt-5")  # your chosen model name
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
USE_RESPONSES_API = os.getenv("USE_RESPONSES_API", "0").lower() in ("1", "true", "yes")
EXECUTOR = ThreadPoolExecutor(max_workers=int(os.getenv("OAI_THREADS", "2")))
OAI_ENABLED = os.getenv("USE_OPENAI", "1").lower() not in ("0", "false", "no")
OAI_BUDGET = float(os.getenv("OAI_BUDGET", "12"))  # seconds hard wall

try:
    from openai import OpenAI
    _openai_available = True
except Exception:
    _openai_available = False

# ----------------------------
# Text utils
# ----------------------------
def sanitize(text: str) -> str:
    if not text:
        return ""
    text = text.replace("—", "-").replace("–", "-")
    text = re.sub(r"[\u201c\u201d]", '"', text)
    text = re.sub(r"[\u2018\u2019]", "'", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

# ----------------------------
# OpenAI helpers (fail-safe)
# ----------------------------
def _get_client():
    if not OPENAI_API_KEY or not _openai_available or not OAI_ENABLED:
        return None
    try:
        # Fail fast; our outer thread guard enforces a hard wall too
        return OpenAI(api_key=OPENAI_API_KEY, timeout=12.0, max_retries=0)
    except Exception as e:
        app.logger.exception("OpenAI client init failed: %s", e)
        return None

def _gpt_call_chat(client, system, prompt) -> str:
    # Do not set temperature; some models accept only default
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": prompt},
        ],
    )
    return resp.choices[0].message.content.strip()

def _gpt_call_resp(client, system, prompt) -> str:
    # Optional: Responses API path
    resp = client.responses.create(
        model=MODEL,
        input=[{"role": "system", "content": system},
               {"role": "user", "content": prompt}],
    )
    if hasattr(resp, "output_text"):
        return str(resp.output_text).strip()
    if hasattr(resp, "output") and resp.output:
        parts = []
        for item in resp.output:
            content = getattr(item, "content", [])
            if content and hasattr(content[0], "text"):
                parts.append(content[0].text)
        if parts:
            return sanitize(" ".join(parts))
    return sanitize(str(resp))

def gpt(prompt: str, system: str = "You are a helpful writing assistant.") -> str:
    client = _get_client()
    if client is None:
        return "Placeholder output (model disabled or key missing)."
    try:
        if USE_RESPONSES_API:
            return EXECUTOR.submit(_gpt_call_resp, client, system, prompt).result(timeout=OAI_BUDGET)
        return EXECUTOR.submit(_gpt_call_chat, client, system, prompt).result(timeout=OAI_BUDGET)
    except FuturesTimeout:
        app.logger.warning("OpenAI call timed out; using placeholder.")
        return "Placeholder output due to timeout."
    except Exception as e:
        app.logger.warning("OpenAI error; using placeholder: %s", e)
        return "Placeholder output due to model error."

# ----------------------------
# DOCX helpers
# ----------------------------
def ensure_docx(doc_or_bytes):
    """Accept a python-docx Document, a file-like object, or raw bytes and return a Document."""
    try:
        if hasattr(doc_or_bytes, "paragraphs"):
            return doc_or_bytes  # already a Document
        if hasattr(doc_or_bytes, "read"):
            data = doc_or_bytes.read()
        else:
            data = doc_or_bytes
        bio = io.BytesIO(data)
        return Document(bio)
    except Exception as e:
        raise RuntimeError(f"Failed to open DOCX: {e}")

KNOWN_HEADINGS = {
    "PROFESSIONAL SUMMARY","SUMMARY","EXPERIENCE","WORK EXPERIENCE",
    "SKILLS","CORE SKILLS","TECHNICAL SKILLS","EDUCATION","PROJECTS",
    "CERTIFICATIONS","PUBLICATIONS","ACHIEVEMENTS"
}

def is_heading(text: str) -> bool:
    t = sanitize(text)
    return (not t) or t.isupper() or (t.upper() in KNOWN_HEADINGS)

def find_section_bounds(doc: Document, titles: List[str]) -> Optional[Tuple[int, int]]:
    titles_up = {t.upper() for t in titles}
    start = None
    for i, p in enumerate(doc.paragraphs):
        if sanitize(p.text).upper() in titles_up:
            start = i
            break
    if start is None:
        return None
    end = len(doc.paragraphs) - 1
    for j in range(start + 1, len(doc.paragraphs)):
        if is_heading(doc.paragraphs[j].text):
            end = j - 1
            break
    return (start, end)

def delete_range(doc: Document, start: int, end: int):
    for i in range(end, start - 1, -1):
        p = doc.paragraphs[i]._element
        p.getparent().remove(p)

def insert_paragraph_after(paragraph: Paragraph, text: str = "", style: Optional[str] = None) -> Paragraph:
    """
    Inserts a paragraph immediately after `paragraph`.
    If `style` is provided but missing in the document, fall back to plain
    paragraph and prefix a bullet if we intended a bullet style.
    """
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)

    run = None
    if text:
        run = new_para.add_run(text)

    if style:
        try:
            new_para.style = style
        except KeyError:
            # Style doesn't exist in this template; if it was a bullet style,
            # add a visible bullet prefix as a graceful fallback.
            if run and style.lower().startswith("list"):
                if not run.text.strip().startswith("•"):
                    run.text = f"• {run.text}"
            # else: silently keep default style

    return new_para

# ----------------------------
# Section writers
# ----------------------------
def write_summary(doc: Document, summary: str):
    sec = find_section_bounds(doc, ["PROFESSIONAL SUMMARY", "SUMMARY"])
    if sec:
        s, e = sec
        if e >= s + 1:
            delete_range(doc, s + 1, e)
        insert_paragraph_after(doc.paragraphs[s], sanitize(summary))
    else:
        # Append a new section at the end
        heading = doc.add_paragraph()
        heading.add_run("PROFESSIONAL SUMMARY").bold = True
        doc.add_paragraph(sanitize(summary))

def inject_skills(doc: Document, skills_map: Dict[str, list]):
    sec = find_section_bounds(doc, ["SKILLS", "CORE SKILLS", "TECHNICAL SKILLS"])
    if sec:
        s, e = sec
        if e >= s + 1:
            delete_range(doc, s + 1, e)
        anchor = doc.paragraphs[s]
    else:
        anchor = doc.add_paragraph()
        anchor.add_run("SKILLS").bold = True

    last = anchor
    for cat, arr in skills_map.items():
        last = insert_paragraph_after(last, f"{cat}:")
        if arr:
            last = insert_paragraph_after(last, ", ".join(arr))

def inject_bullets(doc: Document, bullets_by_company: Dict[str, list]):
    if not bullets_by_company:
        return
    exp_sec = find_section_bounds(doc, ["EXPERIENCE", "WORK EXPERIENCE"])
    start_idx = exp_sec[0] if exp_sec else 0
    end_idx = exp_sec[1] if exp_sec else len(doc.paragraphs) - 1

    i = start_idx
    while i <= end_idx and i < len(doc.paragraphs):
        t = sanitize(doc.paragraphs[i].text)
        match_comp = next((c for c in bullets_by_company.keys() if c and c.lower() in t.lower()), None)
        if match_comp is None:
            i += 1
            continue

        # Find old bullet block just after the company line
        j = i + 1
        while j <= end_idx and j < len(doc.paragraphs):
            txt = sanitize(doc.paragraphs[j].text)
            if is_heading(txt) or not txt:
                break
            j += 1

        # Delete old bullets
        if j - 1 >= i + 1:
            delete_range(doc, i + 1, j - 1)
            end_idx -= (j - 1) - (i + 1) + 1

        # Insert new bullets right after the company line
        anchor = doc.paragraphs[i]
        last = anchor
        for b in bullets_by_company.get(match_comp, []):
            last = insert_paragraph_after(last, sanitize(b), style="List Bullet")  # safe fallback inside

        # Advance past newly added bullets
        i = i + 1 + len(bullets_by_company.get(match_comp, []))

# ----------------------------
# Batch bullets (1 model call)
# ----------------------------
def gpt_bullets_batch(experience: list[dict], jd: str, style_rules: list[str]) -> dict[str, list[str]]:
    """
    Single OpenAI call that returns bullets for ALL companies as pure JSON:
    { "<company>": ["bullet1","bullet2", ...], ... }
    Falls back to placeholders on error/timeout.
    """
    entries = []
    for e in experience:
        k = int(e.get("bullets", 0) or 0)
        if k <= 0:
            continue
        comp = sanitize(e.get("company", ""))
        role = sanitize(e.get("role", ""))
        entries.append(f'- company: "{comp}"; role: "{role}"; bullets: {k}')
    if not entries:
        return {sanitize(e.get("company","")): [] for e in experience}

    rules_txt = (
        "\n".join(f"- {r}" for r in style_rules) if style_rules else
        "- Plain, specific tone\n- 20–28 words each\n- Metricize when truthful\n- Do not include company names inside bullets"
    )

    prompt = f"""Return ONLY JSON (no code fences) mapping company name to an array of bullet strings.
Entries:
{chr(10).join(entries)}

Rules:
{rules_txt}

Job description:
---
{jd[:4000]}
---

JSON schema example:
{{
  "Company A": ["bullet 1", "bullet 2"],
  "Company B": ["bullet 1", "bullet 2", "bullet 3"]
}}"""

    text = gpt(prompt, system="You produce strict JSON only. No commentary.")
    # Try parse; if it fails, build placeholders
    try:
        data = json.loads(text)
    except Exception:
        data = {}
        for e in experience:
            comp = sanitize(e.get("company",""))
            k = int(e.get("bullets", 0) or 0)
            role = sanitize(e.get("role",""))
            if k > 0:
                data[comp] = [
                    f"Delivered outcomes aligned with role '{role}', applying relevant skills; results tailored to the job description. (placeholder)"
                    for _ in range(k)
                ]
            else:
                data[comp] = []

    # Normalize counts + sanitize
    out = {}
    for e in experience:
        comp = sanitize(e.get("company",""))
        k = int(e.get("bullets", 0) or 0)
        bullets = [sanitize(b) for b in (data.get(comp) or [])][:k]
        out[comp] = bullets
    return out

# ----------------------------
# Routes
# ----------------------------
@app.route("/")
def index():
    return jsonify({"ok": True, "service": "resume-tailor-server", "endpoints": ["/health", "/tailor"]})

@app.route("/health")
def health():
    return jsonify({"ok": True, "model": MODEL})

@app.route("/tailor", methods=["POST", "OPTIONS"])
def tailor():
    # CORS preflight handled by Flask-CORS
    origin = request.headers.get("Origin", "*")

    # Accept either JSON body (original flow) or multipart (v2 flow)
    if request.content_type and "multipart/form-data" in request.content_type.lower():
        base_resume_file = request.files.get("base_resume")
        payload_part = request.form.get("payload")
        if not payload_part:
            return make_response(("missing payload", 400))
        try:
            data = json.loads(payload_part)
        except Exception:
            return make_response(("invalid payload json", 400))
    else:
        try:
            data = request.get_json(force=True, silent=False)
        except Exception:
            return make_response(("invalid json", 400))
        base_resume_file = None

    job_desc = sanitize((data or {}).get("job_description", ""))
    cfg = (data or {}).get("resume_config", {}) or {}
    summary_sentences = int(cfg.get("summary_sentences", 2))
    experience = cfg.get("experience", []) or []
    skills_categories = cfg.get("skills_categories", []) or []
    style_rules = (data.get("options", {}) or {}).get("style_rules", []) or []

    # Load base resume
    if base_resume_file:
        base_doc = ensure_docx(base_resume_file)
    else:
        base_path = os.path.join(os.path.dirname(__file__), "base_resume.docx")
        if not os.path.exists(base_path):
            return make_response(("server missing base_resume.docx", 500))
        with open(base_path, "rb") as f:
            base_doc = ensure_docx(f)

    # --- Generate content (OpenAI or placeholders)
    summary_prompt = (
        f"Write {summary_sentences} sentence professional summary aligned to this job description. "
        f"Plain, specific tone; no em/en dashes.\n---\n{job_desc[:4000]}"
    )
    summary = sanitize(gpt(summary_prompt))

    # Batch bullets in one model call to avoid timeouts
    bullets_by_company: Dict[str, list] = gpt_bullets_batch(experience, job_desc, style_rules)

    # Skills grouping (simple; backend enforces categories)
    skills_map = {c: [] for c in skills_categories}
    if skills_categories:
        skills_map[skills_categories[0]] = ["SQL", "Python", "Power BI", "Stakeholder management"]

    # --- Edit DOCX
    doc = base_doc
    write_summary(doc, summary)
    inject_skills(doc, skills_map)
    inject_bullets(doc, bullets_by_company)

    # --- Return DOCX
    out = io.BytesIO()
    doc.save(out)
    out.seek(0)
    filename = "Rijul_Chaturvedi_Tailored.docx"

    resp = make_response(
        send_file(
            out,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name=filename,
        )
    )
    # Echo origin explicitly (Flask-CORS also handles this, but mirrors original behavior)
    resp.headers["Access-Control-Allow-Origin"] = origin
    resp.headers["Vary"] = "Origin"
    return resp

# ---------------
# Local dev entry
# ---------------
if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
