from flask import Flask, request, send_file, jsonify, make_response
from flask_cors import CORS
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
import io, os, json, re, logging
from typing import List, Dict, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeout

# ----------------------------
# App & CORS (Chrome extension → /tailor)
# ----------------------------
app = Flask(__name__)
CORS(app, resources={r"/tailor": {"origins": "chrome-extension://*"}})
app.logger.setLevel(logging.INFO)

# ----------------------------
# Model / OpenAI settings
# ----------------------------
# Default to a fast/reliable model; override with OPENAI_MODEL=gpt-5 if you like.
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
USE_RESPONSES_API = os.getenv("USE_RESPONSES_API", "0").lower() in ("1", "true", "yes")
EXECUTOR = ThreadPoolExecutor(max_workers=int(os.getenv("OAI_THREADS", "2")))
OAI_ENABLED = os.getenv("USE_OPENAI", "1").lower() not in ("0", "false", "no")
OAI_BUDGET = float(os.getenv("OAI_BUDGET", "20"))                 # hard wall (s) per call
OAI_CLIENT_TIMEOUT = float(os.getenv("OAI_CLIENT_TIMEOUT", "20")) # SDK client timeout (s)

try:
    from openai import OpenAI
    _openai_available = True
except Exception:
    _openai_available = False

# ----------------------------
# Text & heading utils
# ----------------------------
def sanitize(text: str) -> str:
    if not text:
        return ""
    text = text.replace("—", "-").replace("–", "-")
    text = re.sub(r"[\u201c\u201d]", '"', text)
    text = re.sub(r"[\u2018\u2019]", "'", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def _norm_heading(text: str) -> str:
    t = sanitize(text).upper()
    t = t.replace("&", "AND")
    t = re.sub(r"[^A-Z0-9 ]+", " ", t)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t

KNOWN_HEADINGS = {
    "PROFESSIONAL SUMMARY","SUMMARY","EXPERIENCE","WORK EXPERIENCE",
    "SKILLS","CORE SKILLS","TECHNICAL SKILLS","SKILLS AND TOOLS","EDUCATION",
    "PROJECTS","CERTIFICATIONS","PUBLICATIONS","ACHIEVEMENTS"
}

def is_heading(text: str) -> bool:
    t = _norm_heading(text)
    return (not t) or t in KNOWN_HEADINGS or t.isupper()

def find_section_bounds(doc: Document, titles: List[str]) -> Optional[Tuple[int, int]]:
    titles_up = {_norm_heading(t) for t in titles}
    start = None
    for i, p in enumerate(doc.paragraphs):
        if _norm_heading(p.text) in titles_up:
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
    Insert a paragraph immediately after `paragraph`.
    If `style` doesn't exist, fall back to normal text and prepend "• " for bullet styles.
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
            if run and style.lower().startswith("list"):
                if not run.text.strip().startswith("•"):
                    run.text = f"• {run.text}"
    return new_para

# ----------------------------
# DOCX load helper
# ----------------------------
def ensure_docx(doc_or_bytes):
    """Accept a python-docx Document, a file-like, or raw bytes; return a Document."""
    try:
        if hasattr(doc_or_bytes, "paragraphs"):
            return doc_or_bytes
        if hasattr(doc_or_bytes, "read"):
            data = doc_or_bytes.read()
        else:
            data = doc_or_bytes
        bio = io.BytesIO(data)
        return Document(bio)
    except Exception as e:
        raise RuntimeError(f"Failed to open DOCX: {e}")

# ----------------------------
# Bullet helpers (quantification & detection)
# ----------------------------
QUANT_REGEX = re.compile(
    r"(\d|\$|%|\bhrs?\b|\bhour(s)?\b|\bday(s)?\b|\bweek(s)?\b|\bmonth(s)?\b|\byear(s)?\b|\b\d{1,3}k\b)",
    flags=re.IGNORECASE
)

def ensure_quantified(bullets: List[str], placeholder: str = "quantified impact: KPI TBD") -> List[str]:
    out = []
    for b in bullets:
        t = sanitize(b)
        if QUANT_REGEX.search(t):
            out.append(t)
        else:
            out.append(f"{t} ({placeholder})")
    return out

def is_bullet_para(p) -> bool:
    """Detect bullet/list paragraphs by style or leading glyphs/markers."""
    try:
        name = (p.style.name or "").lower()
    except Exception:
        name = ""
    if "bullet" in name or "list" in name:
        return True
    t = sanitize(p.text)
    if not t:
        return False
    if t.startswith(("•", "▪", "– ", "- ", "— ")):
        return True
    if re.match(r"^\s*\d+[\.\)]\s+", t):
        return True
    return False

# ----------------------------
# OpenAI helpers (fail-safe)
# ----------------------------
def _get_client():
    if not OPENAI_API_KEY or not _openai_available or not OAI_ENABLED:
        return None
    try:
        return OpenAI(api_key=OPENAI_API_KEY, timeout=OAI_CLIENT_TIMEOUT, max_retries=0)
    except Exception as e:
        app.logger.exception("OpenAI client init failed: %s", e)
        return None

def _gpt_call_chat(client, system, prompt) -> str:
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[
            {"role": "system", "content": system},
            {"role": "user", "content": prompt},
        ],
    )
    return resp.choices[0].message.content.strip()

def _gpt_call_resp(client, system, prompt) -> str:
    if not hasattr(client, "responses"):
        raise AttributeError("Responses API not available in this SDK")
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
        if USE_RESPONSES_API and hasattr(client, "responses"):
            return EXECUTOR.submit(_gpt_call_resp, client, system, prompt).result(timeout=OAI_BUDGET)
        return EXECUTOR.submit(_gpt_call_chat, client, system, prompt).result(timeout=OAI_BUDGET)
    except FuturesTimeout:
        app.logger.warning("OpenAI call timed out; using placeholder.")
        return "Placeholder output due to timeout."
    except Exception as e:
        app.logger.warning("OpenAI error; using placeholder: %s", e)
        return "Placeholder output due to model error."

# ----------------------------
# DOCX section writers
# ----------------------------
def write_summary(doc: Document, summary: str):
    sec = find_section_bounds(doc, ["PROFESSIONAL SUMMARY", "SUMMARY"])
    if sec:
        s, e = sec
        if e >= s + 1:
            delete_range(doc, s + 1, e)
        insert_paragraph_after(doc.paragraphs[s], sanitize(summary))
    else:
        heading = doc.add_paragraph()
        heading.add_run("PROFESSIONAL SUMMARY").bold = True
        doc.add_paragraph(sanitize(summary))

def inject_skills(doc: Document, new_skills: Dict[str, list]):
    """
    Keep existing skills content; append new items per category (dedup).
    If section doesn't exist, create SKILLS & TOOLS. Categories are lines ending with ":".
    """
    if not new_skills:
        return

    sec = find_section_bounds(doc, ["SKILLS", "CORE SKILLS", "TECHNICAL SKILLS", "SKILLS & TOOLS", "SKILLS AND TOOLS"])
    if sec is None:
        anchor = doc.add_paragraph()
        anchor.add_run("SKILLS & TOOLS").bold = True
        sec = find_section_bounds(doc, ["SKILLS & TOOLS", "SKILLS AND TOOLS"])
    s, e = sec

    # Parse existing categories and item lines between s+1..e
    existing_map: Dict[str, Tuple[int, set]] = {}
    current_cat = None
    for idx in range(s + 1, e + 1):
        line = sanitize(doc.paragraphs[idx].text)
        if not line:
            continue
        if line.endswith(":"):
            current_cat = line[:-1].strip()
            existing_map.setdefault(current_cat, (None, set()))
        else:
            if current_cat is not None:
                items = [x.strip() for x in line.split(",") if x.strip()]
                pi, aset = existing_map[current_cat]
                aset.update(items)
                existing_map[current_cat] = (idx, aset)

    # Append or create categories
    for cat, additions in new_skills.items():
        additions = [x for x in additions if x]
        if not additions:
            continue

        if cat in existing_map:
            items_idx, aset = existing_map[cat]
            to_add = [x for x in additions if x not in aset]
            if not to_add:
                continue
            if items_idx is not None:
                para = doc.paragraphs[items_idx]
                current = sanitize(para.text)
                joiner = (", " if current and not current.endswith(",") else "")
                para.text = f"{current}{joiner}{', '.join(to_add)}"
                aset.update(to_add)
                existing_map[cat] = (items_idx, aset)
            else:
                # Find category header line
                cat_idx = None
                for idx in range(s + 1, e + 1):
                    if sanitize(doc.paragraphs[idx].text).strip().lower() == f"{cat.lower()}:":
                        cat_idx = idx
                        break
                anchor = doc.paragraphs[cat_idx] if cat_idx is not None else doc.paragraphs[e]
                newp = insert_paragraph_after(anchor, ", ".join(to_add))
                existing_map[cat] = (doc.paragraphs.index(newp), set(to_add))
        else:
            # New category → append header + items at end of the section
            sec = find_section_bounds(doc, ["SKILLS", "CORE SKILLS", "TECHNICAL SKILLS", "SKILLS & TOOLS", "SKILLS AND TOOLS"])
            s, e = sec
            anchor = doc.paragraphs[e]
            header = insert_paragraph_after(anchor, f"{cat}:")
            items_p = insert_paragraph_after(header, ", ".join(additions))
            existing_map[cat] = (doc.paragraphs.index(items_p), set(additions))

def inject_bullets(doc: Document, bullets_by_company: Dict[str, list]):
    """
    Replace only the existing bullet paragraphs under each company with the new ones.
    """
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

        # Identify contiguous bullet paragraphs directly under this line
        j = i + 1
        while j <= end_idx and j < len(doc.paragraphs) and is_bullet_para(doc.paragraphs[j]):
            j += 1

        # Delete ONLY the bullet block (if any)
        if j - 1 >= i + 1:
            delete_range(doc, i + 1, j - 1)
            removed = (j - 1) - (i + 1) + 1
            end_idx -= removed

        # Insert new bullets right after the company line
        anchor = doc.paragraphs[i]
        last = anchor
        for b in bullets_by_company.get(match_comp, []):
            last = insert_paragraph_after(last, sanitize(b), style="List Bullet")  # safe fallback inside

        # Move past the newly inserted bullets
        i = i + 1 + len(bullets_by_company.get(match_comp, []))

# ----------------------------
# Skills heuristic (JD-driven)
# ----------------------------
SKILL_BANK = {
    "Business Analysis & Delivery": [
        "Requirements elicitation", "User stories", "Acceptance criteria", "UAT", "BPMN", "UML",
        "Process re-engineering", "Traceability (RTM)", "Fit–gap", "Prioritization (RICE, MoSCoW)"
    ],
    "Project & Program Management": [
        "Agile Scrum", "Kanban", "SDLC", "RACI", "RAID", "Sprint planning", "Roadmapping", "Stakeholder management"
    ],
    "Data, Analytics & BI": [
        "SQL", "Python", "Pandas", "NumPy", "Power BI", "Tableau", "A/B testing", "Forecasting", "Anomaly detection"
    ],
    "Cloud, Data & MLOps": [
        "AWS Lambda", "S3", "Redshift", "Airflow", "dbt", "Spark", "Databricks", "ETL/ELT orchestration"
    ],
    "AI, ML & GenAI": [
        "LLMs", "RAG", "Prompt design", "Hugging Face", "FAISS", "Model monitoring", "Inference optimization"
    ],
    "Enterprise Platforms & Integration": [
        "Salesforce", "NetSuite", "SAP", "ServiceNow", "REST APIs", "Postman", "Git", "Azure DevOps", "Jira"
    ],
    "Collaboration, Design & Stakeholder": [
        "Miro", "Figma", "Lucidchart", "Wireframing", "Prototyping", "Executive communication", "Workshops"
    ],
}

def pick_skills(jd: str, bank: Dict[str, List[str]], top_k: int = 8) -> Dict[str, List[str]]:
    jd_l = jd.lower()
    out = {}
    for cat, items in bank.items():
        hits = []
        for s in items:
            pat = re.escape(s.lower()).replace(r"\ ", r"\s+")
            if re.search(rf"(?<![A-Za-z0-9]){pat}(?![A-Za-z0-9])", jd_l):
                hits.append(s)
        seen = set()
        merged = [x for x in hits + items if (x not in seen and not seen.add(x))]
        out[cat] = merged[:top_k]
    return out

# ----------------------------
# Batch bullets (1 model call, quantified)
# ----------------------------
def gpt_bullets_batch(experience: List[Dict], jd: str, style_rules: List[str]) -> Dict[str, List[str]]:
    """
    Single model call; returns {company: [bullets]}.
    Enforces quantification where truthful; never fabricates numbers—adds a KPI marker if absent.
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
        "- 20–28 words each\n"
        "- Start with a strong verb; past tense\n"
        "- Include at least one quantifier when truthful (%, $, #, time, count); prefer real figures from the JD\n"
        "- If no truthful figure exists, mark with 'KPI TBD' instead of inventing numbers\n"
        "- No company names inside bullets\n"
        "- Avoid buzzwords; focus on actions and outcomes"
    )

    prompt = f"""Return ONLY JSON (no code fences) mapping company name to an array of bullet strings.
Entries:
{chr(10).join(entries)}

Rules:
{rules_txt}

Job description (trimmed):
---
{jd[:1500]}
---

JSON schema example:
{{
  "Company A": ["bullet 1", "bullet 2"],
  "Company B": ["bullet 1", "bullet 2", "bullet 3"]
}}"""

    text = gpt(prompt, system="You produce strict JSON only. No commentary.")
    # Try parse; if fail, build placeholders
    try:
        data = json.loads(text)
    except Exception:
        data = {}
        for e in experience:
            comp = sanitize(e.get("company",""))
            k = int(e.get("bullets", 0) or 0)
            role = sanitize(e.get("role",""))
            data[comp] = [f"Delivered outcomes in '{role}' aligned to the JD (KPI TBD)" for _ in range(max(k,0))]

    # Normalize counts + sanitize + ensure quant signal
    out = {}
    for e in experience:
        comp = sanitize(e.get("company",""))
        k = int(e.get("bullets", 0) or 0)
        bullets = [sanitize(b) for b in (data.get(comp) or [])][:k]
        bullets = ensure_quantified(bullets)
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
    enabled = bool(OPENAI_API_KEY) and OAI_ENABLED
    return jsonify({"ok": True, "model": MODEL, "openai_enabled": enabled})

@app.route("/tailor", methods=["POST", "OPTIONS"])
def tailor():
    origin = request.headers.get("Origin", "*")

    # JSON or multipart payloads
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

    # --- Generate content
    summary_prompt = (
        f"Write {summary_sentences} sentence professional summary aligned to the job description below. "
        f"Use concise, specific language; prefer quantified outcomes when truthful; avoid buzzwords and dashes.\n---\n{job_desc[:1500]}"
    )
    summary = sanitize(gpt(summary_prompt))

    bullets_by_company: Dict[str, List[str]] = gpt_bullets_batch(experience, job_desc, style_rules)

    desired_bank = {c: SKILL_BANK.get(c, []) for c in skills_categories}
    skills_map = pick_skills(job_desc, desired_bank) if skills_categories else {}

    # --- Edit DOCX
    doc = base_doc
    write_summary(doc, summary)
    inject_skills(doc, skills_map)          # keep originals; append deduped
    inject_bullets(doc, bullets_by_company) # replace bullet block only

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
    resp.headers["Access-Control-Allow-Origin"] = origin
    resp.headers["Vary"] = "Origin"
    return resp

# ---------------
# Local dev entry
# ---------------
if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
