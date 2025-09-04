from flask import Flask, request, send_file, jsonify, make_response
from flask_cors import CORS
from docx import Document
import io, os, json, re, logging
from typing import List, Dict, Tuple, Optional

app = Flask(__name__)
# Match original: allow chrome-extension origins to hit /tailor (OPTIONS + POST)
CORS(app, resources={r"/tailor": {"origins": "chrome-extension://*"}})
app.logger.setLevel(logging.INFO)

MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

try:
    from openai import OpenAI
    _openai_available = True
except Exception:
    _openai_available = False

def sanitize(text: str) -> str:
    if not text: return ""
    text = text.replace("—","-").replace("–","-")
    text = re.sub(r"[\u201c\u201d]", '"', text)
    text = re.sub(r"[\u2018\u2019]", "'", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def _get_client():
    if not OPENAI_API_KEY or not _openai_available:
        return None
    try:
        return OpenAI(api_key=OPENAI_API_KEY)
    except Exception as e:
        app.logger.exception("OpenAI client init failed: %s", e)
        return None

def gpt(prompt: str, system: str = "You are a helpful writing assistant.") -> str:
    client = _get_client()
    if client is None:
        return "Placeholder output (no OPENAI_API_KEY set)."
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[{"role":"system","content":system},{"role":"user","content":prompt}],
        temperature=0.2
    )
    return resp.choices[0].message.content.strip()

def ensure_docx(doc_or_bytes):
    """Accept a python-docx Document, a file-like object, or raw bytes and return a Document."""
    try:
        # If it behaves like a python-docx Document (has paragraphs), use it
        if hasattr(doc_or_bytes, "paragraphs"):
            return doc_or_bytes
        # If it's a Werkzeug/FileStorage or any file-like, read bytes
        if hasattr(doc_or_bytes, "read"):
            data = doc_or_bytes.read()
        else:
            data = doc_or_bytes
        bio = io.BytesIO(data)
        return Document(bio)
    except Exception as e:
        raise RuntimeError(f"Failed to open DOCX: {e}")

def write_summary(doc: Document, summary: str):
    # simple heuristic: replace first non-empty paragraph under SUMMARY/PROFESSIONAL SUMMARY if found, else prepend
    titles = ["PROFESSIONAL SUMMARY","SUMMARY"]
    idxs = {i: sanitize(p.text) for i,p in enumerate(doc.paragraphs)}
    start = next((i for i,t in idxs.items() if t.upper() in titles), None)
    if start is not None:
        # wipe all until next heading-like line
        j = start + 1
        while j < len(doc.paragraphs) and not sanitize(doc.paragraphs[j].text).isupper():
            # delete this paragraph
            p = doc.paragraphs[j]._element
            p.getparent().remove(p)
        # insert one paragraph
        doc.add_paragraph(summary)
    else:
        # prepend as new section
        p = doc.add_paragraph("PROFESSIONAL SUMMARY")
        p.runs[0].bold = True
        doc.add_paragraph(summary)

def inject_skills(doc: Document, skills_map: Dict[str, list]):
    # very basic: append/update SKILLS section
    titles = ["SKILLS","CORE SKILLS","TECHNICAL SKILLS"]
    idxs = {i: sanitize(p.text) for i,p in enumerate(doc.paragraphs)}
    start = next((i for i,t in idxs.items() if t.upper() in [x.upper() for x in titles]), None)
    if start is not None:
        # wipe content after heading until next heading-like
        j = start + 1
        while j < len(doc.paragraphs) and not sanitize(doc.paragraphs[j].text).isupper():
            p = doc.paragraphs[j]._element
            p.getparent().remove(p)
        # write categories
        for cat, arr in skills_map.items():
            doc.add_paragraph(f"{cat}:")
            if arr:
                doc.add_paragraph(", ".join(arr))
    else:
        p = doc.add_paragraph("SKILLS"); p.runs[0].bold = True
        for cat, arr in skills_map.items():
            doc.add_paragraph(f"{cat}:")
            if arr:
                doc.add_paragraph(", ".join(arr))

def inject_bullets(doc: Document, bullets_by_company: Dict[str, list]):
    # naive strategy: find company name lines, replace following bullet-like lines
    for i, para in enumerate(doc.paragraphs):
        t = sanitize(para.text)
        for comp, bullets in bullets_by_company.items():
            if comp and comp.lower() in t.lower():
                # remove subsequent bullet-ish paragraphs until blank or heading
                j = i + 1
                while j < len(doc.paragraphs):
                    txt = sanitize(doc.paragraphs[j].text)
                    if not txt or txt.isupper():
                        break
                    # delete this paragraph
                    p = doc.paragraphs[j]._element
                    p.getparent().remove(p)
                # insert new bullets
                for b in bullets:
                    doc.add_paragraph(b, style="List Bullet")

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

    # Accept either JSON body (original flow) or multipart (new flow)
    job_desc = ""
    experience = []
    skills_categories = []
    summary_sentences = 2
    style_rules = []

    base_resume_file = None
    if request.content_type and "multipart/form-data" in request.content_type.lower():
        base_resume_file = request.files.get("base_resume")
        payload_part = request.form.get("payload")
        if payload_part:
            data = json.loads(payload_part)
        else:
            return make_response(("missing payload", 400))
    else:
        # JSON
        try:
            data = request.get_json(force=True, silent=False)
        except Exception:
            return make_response(("invalid json", 400))

    job_desc = sanitize((data or {}).get("job_description",""))
    cfg = (data or {}).get("resume_config",{}) or {}
    summary_sentences = int(cfg.get("summary_sentences", 2))
    experience = cfg.get("experience",[]) or []
    skills_categories = cfg.get("skills_categories",[]) or []
    style_rules = (data.get("options",{}) or {}).get("style_rules",[]) or []

    # get base resume
    if base_resume_file:
        base_doc = ensure_docx(base_resume_file)
    else:
        # use server-bundled docx (new base resume placed on server)
        base_path = os.path.join(os.path.dirname(__file__), "base_resume.docx")
        if not os.path.exists(base_path):
            return make_response(("server missing base_resume.docx", 500))
        with open(base_path, "rb") as f:
            base_doc = ensure_docx(f)

    # Generate text (placeholder if no key)
    summary_prompt = f"Write {summary_sentences} sentence professional summary aligned to this JD. Plain tone, no em/en dashes.\n---\n{job_desc[:4000]}"
    summary = sanitize(gpt(summary_prompt))

    bullets_by_company: Dict[str, list] = {}
    for e in experience:
        comp = e.get("company","")
        role = e.get("role","")
        k = int(e.get("bullets", 0))
        if k <= 0:
            bullets_by_company[comp] = []
            continue
        bp = f"Write exactly {k} resume bullets for role '{role}'. 20-28 words each. Metricized, impact-focused. No company names inside bullets.\nAligned to JD:\n---\n{job_desc[:4000]}"
        raw = gpt(bp)
        # simple split
        parts = [sanitize(x) for x in re.split(r"[\n\r]+", raw) if sanitize(x)]
        # strip leading "1. " etc
        clean = []
        for t in parts:
            m = re.match(r"^\s*(\d+)[\.\)]\s+(.*)$", t)
            clean.append((m.group(2) if m else t))
        bullets_by_company[comp] = clean[:k]

    # skills grouping (very simple placeholder)
    skills_map = {c: [] for c in skills_categories}
    if skills_categories:
        skills_map[skills_categories[0]] = ["SQL","Python","Power BI","Stakeholder management"]

    # Edit docx
    doc = base_doc
    write_summary(doc, summary)
    inject_skills(doc, skills_map)
    inject_bullets(doc, bullets_by_company)

    out = io.BytesIO()
    doc.save(out); out.seek(0)
    filename = "Rijul_Chaturvedi_Tailored.docx"

    resp = make_response(send_file(out, as_attachment=True, download_name=filename))
    # Add explicit CORS for extension origin (mirrors original)
    resp.headers["Access-Control-Allow-Origin"] = origin
    resp.headers["Vary"] = "Origin"
    resp.headers["Content-Type"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    return resp
