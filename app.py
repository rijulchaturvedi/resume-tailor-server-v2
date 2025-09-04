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
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")  # override with gpt-5 if desired
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
USE_RESPONSES_API = os.getenv("USE_RESPONSES_API", "0").lower() in ("1", "true", "yes")
EXECUTOR = ThreadPoolExecutor(max_workers=int(os.getenv("OAI_THREADS", "2")))
OAI_ENABLED = os.getenv("USE_OPENAI", "1").lower() not in ("0", "false", "no")
OAI_BUDGET = float(os.getenv("OAI_BUDGET", "20"))                 # hard wall (s) per model call
OAI_CLIENT_TIMEOUT = float(os.getenv("OAI_CLIENT_TIMEOUT", "20")) # SDK client timeout (s)

# Behavior toggles
SHOW_KPI_PLACEHOLDER = os.getenv("SHOW_KPI_PLACEHOLDER", "1").lower() in ("1", "true", "yes")
BULLETS_STRICT_REPLACE = os.getenv("BULLETS_STRICT_REPLACE", "0").lower() in ("1", "true", "yes")

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
            if run and style.lower().startswith("list"):  # e.g., "List Bullet"
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
    if not SHOW_KPI_PLACEHOLDER:
        return [sanitize(b) for b in bullets]
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
# Metric harvesting (from resume blocks + JD)
# ----------------------------
NUM_PHRASE_REGEX = re.compile(
    r"(\$?\d+(?:\.\d+)?\s?(?:k|m|bn|b|million|billion)?|\d+\s?(?:%|percent)|\d+\+\b|"
    r"\b\d+\s?(?:days?|weeks?|months?|years?)|\$\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d{1,3}%|\d{1,3}\s?%)",
    re.IGNORECASE
)

def extract_numeric_phrases(text: str, max_phrases: int = 12) -> List[str]:
    found = []
    for m in NUM_PHRASE_REGEX.finditer(text or ""):
        val = sanitize(m.group(0))
        if val and val not in found:
            found.append(val)
        if len(found) >= max_phrases:
            break
    return found

def harvest_metrics_from_role_block(doc: Document, start_idx: int, end_idx: int) -> List[str]:
    buf = []
    for i in range(start_idx + 1, min(end_idx + 1, len(doc.paragraphs))):
        t = sanitize(doc.paragraphs[i].text)
        if not t or is_heading(t):
            break
        buf.append(t)
    return extract_numeric_phrases("  ".join(buf))

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
    Keep existing skills; merge duplicate category headers; append new items (dedup) into the first occurrence.
    """
    if not new_skills:
        return

    sec = find_section_bounds(doc, ["SKILLS", "CORE SKILLS", "TECHNICAL SKILLS", "SKILLS & TOOLS", "SKILLS AND TOOLS"])
    if sec is None:
        anchor = doc.add_paragraph()
        anchor.add_run("SKILLS & TOOLS").bold = True
        sec = find_section_bounds(doc, ["SKILLS & TOOLS", "SKILLS AND TOOLS"])
    s, e = sec

    # Pass 1: map first occurrence per category; collect duplicates to delete
    def norm_cat_name(name: str) -> str:
        return _norm_heading(name).replace("  ", " ").strip()

    first_idx: Dict[str, Dict[str, object]] = {}  # norm_cat -> {header:int, items_idx:Optional[int], items:set}
    to_delete: List[Tuple[int, int]] = []

    current_cat = None
    current_header_idx = None
    for idx in range(s + 1, e + 1):
        line = sanitize(doc.paragraphs[idx].text)
        if not line:
            continue
        if line.endswith(":"):
            key = norm_cat_name(line[:-1].strip())
            if key not in first_idx:
                first_idx[key] = {"header": idx, "items_idx": None, "items": set()}
                current_cat = key
                current_header_idx = idx
            else:
                # duplicate header: mark it (and its following items line if present) for deletion
                dup_start = idx
                dup_end = idx
                if idx + 1 <= e:
                    nxt = sanitize(doc.paragraphs[idx + 1].text)
                    if nxt and not nxt.endswith(":") and not is_heading(nxt):
                        dup_end = idx + 1
                        # merge its items into the first occurrence set
                        items = [x.strip() for x in nxt.split(",") if x.strip()]
                        first_idx[key]["items"].update(items)
                to_delete.append((dup_start, dup_end))
                current_cat = None
                current_header_idx = None
        else:
            if current_cat is not None:
                items = [x.strip() for x in line.split(",") if x.strip()]
                first_idx[current_cat]["items"].update(items)
                if first_idx[current_cat]["items_idx"] is None:
                    first_idx[current_cat]["items_idx"] = idx

    # Delete duplicates bottom-up
    for a, b in sorted(to_delete, key=lambda x: x[0], reverse=True):
        delete_range(doc, a, b)
        e -= (b - a + 1)

    # Pass 2: write back merged items for existing categories
    for key, meta in first_idx.items():
        h = int(meta["header"])
        it = meta["items_idx"]
        items_sorted = list(meta["items"])
        if items_sorted:
            if it is not None and it < len(doc.paragraphs):
                para = doc.paragraphs[it]
                para.text = ", ".join(items_sorted)
            else:
                header_para = doc.paragraphs[h]
                insert_paragraph_after(header_para, ", ".join(items_sorted))

    # Pass 3: append new skills per incoming categories
    for cat, additions in new_skills.items():
        additions = [x for x in additions if x]
        if not additions:
            continue
        key = norm_cat_name(cat)
        if key in first_idx:
            # Append dedup to first occurrence items
            h = int(first_idx[key]["header"])
            # Find existing items line right after header (if any) again
            target_items_idx = None
            if first_idx[key]["items_idx"] is not None:
                target_items_idx = int(first_idx[key]["items_idx"])
            else:
                # search next line
                if h + 1 < len(doc.paragraphs):
                    maybe = sanitize(doc.paragraphs[h + 1].text)
                    if maybe and not maybe.endswith(":") and not is_heading(maybe):
                        target_items_idx = h + 1

            if target_items_idx is not None and target_items_idx < len(doc.paragraphs):
                para = doc.paragraphs[target_items_idx]
                existing = [x.strip() for x in sanitize(para.text).split(",") if x.strip()]
                add = [x for x in additions if x not in existing]
                if add:
                    joiner = (", " if para.text and not para.text.strip().endswith(",") else "")
                    para.text = f"{sanitize(para.text)}{joiner}{', '.join(add)}"
            else:
                header_para = doc.paragraphs[h]
                insert_paragraph_after(header_para, ", ".join(additions))
        else:
            # New category → append header + items at end of section
            sec2 = find_section_bounds(doc, ["SKILLS", "CORE SKILLS", "TECHNICAL SKILLS", "SKILLS & TOOLS", "SKILLS AND TOOLS"])
            s2, e2 = sec2 if sec2 else (0, len(doc.paragraphs) - 1)
            anchor = doc.paragraphs[e2]
            header = insert_paragraph_after(anchor, f"{cat}:")
            insert_paragraph_after(header, ", ".join(additions))

def inject_bullets(doc: Document, bullets_by_company: Dict[str, list]):
    """
    Replace bullets under each company. If BULLETS_STRICT_REPLACE=1, remove ALL non-heading, non-empty
    lines under the role until a blank/heading; otherwise, remove only bullet-styled paragraphs.
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

        j = i + 1
        if BULLETS_STRICT_REPLACE:
            while j <= end_idx and j < len(doc.paragraphs):
                txt = sanitize(doc.paragraphs[j].text)
                if not txt or is_heading(txt):
                    break
                j += 1
        else:
            while j <= end_idx and j < len(doc.paragraphs) and is_bullet_para(doc.paragraphs[j]):
                j += 1

        # Delete the identified block
        if j - 1 >= i + 1:
            delete_range(doc, i + 1, j - 1)
            removed = (j - 1) - (i + 1) + 1
            end_idx -= removed

        # Insert new bullets right after the company line
        anchor = doc.paragraphs[i]
        last = anchor
        for b in bullets_by_company.get(match_comp, []):
            last = insert_paragraph_after(last, sanitize(b), style="List Bullet")

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
# Batch bullets (1 model call, uses harvested metrics)
# ----------------------------
def gpt_bullets_batch(
    experience: List[Dict],
    jd: str,
    style_rules: List[str],
    metrics_by_company: Dict[str, List[str]]
) -> Dict[str, List[str]]:
    """
    Returns {company: [bullets]} using only provided figures (resume/JD) when possible.
    If no figure fits, optional KPI marker is appended (SHOW_KPI_PLACEHOLDER).
    """
    entries = []
    for e in experience:
        k = int(e.get("bullets", 0) or 0)
        if k <= 0:
            continue
        comp = sanitize(e.get("company", ""))
        role = sanitize(e.get("role", ""))
        mx = metrics_by_company.get(comp, [])
        # We inline metrics as a hint list; model should weave them naturally (and only if truthful)
        metrics_str = ", ".join(mx) if mx else ""
        entries.append(f'- company: "{comp}"; role: "{role}"; bullets: {k}; metrics: [{metrics_str}]')
    if not entries:
        return {sanitize(e.get("company","")): [] for e in experience}

    rules_txt = (
        "\n".join(f"- {r}" for r in style_rules) if style_rules else
        "- 20–28 words each\n"
        "- Start with a strong verb; past tense\n"
        "- Prefer the provided metrics; use them naturally (%, $, counts, time)\n"
        "- Do NOT invent numbers; if none apply, leave unnumbered\n"
        "- No company names inside bullets\n"
        "- Avoid buzzwords; focus on actions and outcomes"
    )

    prompt = f"""Return ONLY JSON (no code fences) mapping company name to an array of bullet strings.

Entries (each entry may include 'metrics' that you can use verbatim):
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
    # Parse or fallback
    try:
        data = json.loads(text)
    except Exception:
        data = {}
        for e in experience:
            comp = sanitize(e.get("company",""))
            k = int(e.get("bullets", 0) or 0)
            role = sanitize(e.get("role",""))
            data[comp] = [f"Drove outcomes in '{role}' aligned to JD" for _ in range(max(k,0))]

    # Normalize + sanitize + optional KPI marker
    out = {}
    for e in experience:
        comp = sanitize(e.get("company",""))
        k = int(e.get("bullets", 0) or 0)
        bullets = [sanitize(b) for b in (data.get(comp) or [])][:k]
        bullets = ensure_quantified(bullets)  # add KPI marker only if enabled and bullet lacks quant
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

    # Build metrics_by_company from resume role blocks + JD
    metrics_by_company: Dict[str, List[str]] = {}
    exp_sec = find_section_bounds(base_doc, ["EXPERIENCE", "WORK EXPERIENCE"])
    start_idx = exp_sec[0] if exp_sec else 0
    end_idx = exp_sec[1] if exp_sec else len(base_doc.paragraphs) - 1

    for e in experience:
        comp = sanitize(e.get("company",""))
        comp_idx = None
        for i in range(start_idx, end_idx + 1):
            if comp and comp.lower() in sanitize(base_doc.paragraphs[i].text).lower():
                comp_idx = i
                break
        role_metrics = harvest_metrics_from_role_block(base_doc, comp_idx if comp_idx is not None else start_idx, end_idx) if comp_idx is not None else []
        jd_metrics = extract_numeric_phrases(job_desc)
        seen = set()
        merged = [x for x in role_metrics + jd_metrics if (x not in seen and not seen.add(x))][:20]
        metrics_by_company[comp] = merged

    # Batch bullets (quant-friendly, metrics-aware)
    bullets_by_company: Dict[str, List[str]] = gpt_bullets_batch(experience, job_desc, style_rules, metrics_by_company)

    # Skills grouping (JD-driven) — keep existing + append deduped, merge duplicate headers
    desired_bank = {c: SKILL_BANK.get(c, []) for c in skills_categories}
    skills_map = pick_skills(job_desc, desired_bank) if skills_categories else {}

    # --- Edit DOCX
    doc = base_doc
    write_summary(doc, summary)
    inject_skills(doc, skills_map)          # keep originals; merge & append deduped
    inject_bullets(doc, bullets_by_company) # replace under each role (strict mode optional)

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
