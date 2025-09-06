from flask import Flask, request, send_file, jsonify, make_response
from flask_cors import CORS
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
import io, os, json, re, logging
from typing import List, Dict, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeout

# ----------------------------
# App & CORS
# ----------------------------
app = Flask(__name__)
CORS(app, resources={r"/tailor": {"origins": "chrome-extension://*"}})
app.logger.setLevel(logging.INFO)

# ----------------------------
# OpenAI config
# ----------------------------
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")  # or "gpt-5" if you prefer
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
EXECUTOR = ThreadPoolExecutor(max_workers=int(os.getenv("OAI_THREADS", "2")))
OAI_ENABLED = os.getenv("USE_OPENAI", "1").strip().lower() not in ("0", "false", "no")
OAI_BUDGET = float(os.getenv("OAI_BUDGET", "20"))
OAI_CLIENT_TIMEOUT = float(os.getenv("OAI_CLIENT_TIMEOUT", "20"))
USE_RESPONSES_API = os.getenv("USE_RESPONSES_API", "0").strip().lower() in ("1","true","yes")

def _env_bool(name: str, default: bool=False) -> bool:
    v = os.getenv(name)
    if v is None: return default
    return v.strip().lower() in ("1","true","yes","on","y")

# Behavior toggles (overridable via JSON options)
SHOW_KPI_PLACEHOLDER = _env_bool("SHOW_KPI_PLACEHOLDER", True)
BULLETS_STRICT_REPLACE = _env_bool("BULLETS_STRICT_REPLACE", True)

try:
    from openai import OpenAI
    _openai_available = True
except Exception:
    _openai_available = False

# ----------------------------
# Text helpers
# ----------------------------
KNOWN_HEADINGS = {
    "PROFESSIONAL SUMMARY","SUMMARY",
    "EXPERIENCE","WORK EXPERIENCE","PROFESSIONAL EXPERIENCE",
    "SKILLS","CORE SKILLS","TECHNICAL SKILLS","SKILLS & TOOLS","SKILLS AND TOOLS",
    "EDUCATION","PROJECTS","CERTIFICATIONS","PUBLICATIONS","ACHIEVEMENTS"
}

def sanitize(text: str) -> str:
    if not text: return ""
    text = text.replace("\u00A0", " ").replace("\t", " ")
    text = text.replace("—","-").replace("–","-")
    text = re.sub(r"[\u201c\u201d]", '"', text)
    text = re.sub(r"[\u2018\u2019]", "'", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def _canon(s: str) -> str:
    s = sanitize(s).lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _norm_heading(text: str) -> str:
    t = sanitize(text).upper()
    t = t.replace("&","AND")
    t = re.sub(r"[^A-Z0-9 ]+", " ", t)
    t = re.sub(r"\s{2,}", " ", t).strip()
    return t

def is_major_heading(text: str) -> bool:
    t = _norm_heading(text)
    return (not t) or t in KNOWN_HEADINGS or t.isupper()

# ----------------------------
# DOCX utils
# ----------------------------
def ensure_docx(doc_or_bytes):
    try:
        if hasattr(doc_or_bytes, "paragraphs"):
            return doc_or_bytes
        if hasattr(doc_or_bytes, "read"):
            data = doc_or_bytes.read()
        else:
            data = doc_or_bytes
        return Document(io.BytesIO(data))
    except Exception as e:
        raise RuntimeError(f"Failed to open DOCX: {e}")

def delete_range(doc: Document, start: int, end: int):
    """Delete paragraphs from start to end (inclusive)"""
    for i in range(end, start-1, -1):
        if i < len(doc.paragraphs):
            p = doc.paragraphs[i]._element
            p.getparent().remove(p)

def insert_paragraph_after(paragraph: Paragraph, text: str = "", style: Optional[str] = None) -> Paragraph:
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
            # fallback bullet glyph if "List Bullet" style isn't present
            if run and style.lower().startswith("list") and not run.text.strip().startswith("•"):
                run.text = f"• {run.text}"
    return new_para

def find_section_bounds(doc: Document, titles: List[str]) -> Optional[Tuple[int,int]]:
    titles_up = {_norm_heading(t) for t in titles}
    start = None
    for i,p in enumerate(doc.paragraphs):
        if _norm_heading(p.text) in titles_up:
            start = i
            break
    if start is None:
        return None
    end = len(doc.paragraphs)-1
    for j in range(start+1, len(doc.paragraphs)):
        if is_major_heading(doc.paragraphs[j].text):
            end = j-1
            break
    return (start, end)

def find_all_section_bounds(doc: Document, titles: List[str]) -> List[Tuple[int,int]]:
    spans = []
    i = 0
    titles_up = {_norm_heading(t) for t in titles}
    while i < len(doc.paragraphs):
        if _norm_heading(doc.paragraphs[i].text) in titles_up:
            s = i
            e = len(doc.paragraphs)-1
            for j in range(i+1, len(doc.paragraphs)):
                if is_major_heading(doc.paragraphs[j].text):
                    e = j-1
                    break
            spans.append((s,e))
            i = e + 1
        else:
            i += 1
    return spans

def exactish_match(line: str, target: str) -> bool:
    L = _canon(line); T = _canon(target)
    if not L or not T: return False
    return L == T or (T in L)

def find_anchors_by_exact_headers(doc: Document, exact_headers: Dict[str,str]) -> List[Tuple[str,int]]:
    """ exact_headers: { company -> exact visible header line in the DOCX } """
    pairs = [(sanitize(k), v) for k, v in exact_headers.items()]
    found = []
    used = set()
    for comp, header in pairs:
        idx = None
        # strict equality on sanitized text
        for i, p in enumerate(doc.paragraphs):
            if sanitize(p.text) == sanitize(header):
                idx = i; break
        # fallback: canon contains
        if idx is None:
            for i, p in enumerate(doc.paragraphs):
                if exactish_match(p.text, header):
                    idx = i; break
        if idx is not None and idx not in used:
            found.append((comp, idx)); used.add(idx)
        else:
            app.logger.warning("Exact header NOT found for %s -> '%s'", comp, header)
    found.sort(key=lambda x: x[1])
    return found

def next_index_after(idx: int, arr: List[int]) -> Optional[int]:
    for a in arr:
        if a > idx: return a
    return None

def find_role_section_bounds(doc: Document, header_idx: int) -> Tuple[int, int]:
    """
    Find the bounds of content under a role header.
    Returns (content_start, content_end) where content_start is the first paragraph
    after the header and content_end is the last paragraph before the next major section.
    """
    content_start = header_idx + 1
    content_end = len(doc.paragraphs) - 1
    
    # Find the next major heading to determine where this role section ends
    for i in range(header_idx + 1, len(doc.paragraphs)):
        para_text = sanitize(doc.paragraphs[i].text)
        if is_major_heading(para_text) or _is_role_header(para_text):
            content_end = i - 1
            break
    
    return content_start, content_end

def _is_role_header(text: str) -> bool:
    """Check if a paragraph looks like a role header (contains role/company info)"""
    text = sanitize(text)
    # Look for patterns like "Title, Company" or "**Title,** Company"
    # This is a heuristic - you might need to adjust based on your resume format
    if not text:
        return False
    
    # Common indicators of role headers
    role_indicators = [
        "analyst", "manager", "engineer", "developer", "consultant", 
        "director", "specialist", "coordinator", "lead", "senior"
    ]
    
    text_lower = text.lower()
    has_role_keyword = any(keyword in text_lower for keyword in role_indicators)
    
    # Check if it has bold formatting or other formatting that suggests it's a header
    # (This is approximate since we're just looking at text)
    has_formatting_clues = any(marker in text for marker in ["**", "•", "-"])
    
    return has_role_keyword and (has_formatting_clues or len(text.split()) <= 10)

# ----------------------------
# KPI placeholder / quantification
# ----------------------------
QUANT_REGEX = re.compile(
    r"(\d|\$|%|\bhrs?\b|\bhour(s)?\b|\bday(s)?\b|\bweek(s)?\b|\bmonth(s)?\b|\byear(s)?\b|\b\d{1,3}k\b)",
    re.IGNORECASE
)
def ensure_quantified(bullets: List[str], placeholder: str = "quantified impact: KPI TBD") -> List[str]:
    if not SHOW_KPI_PLACEHOLDER:
        return [sanitize(b) for b in bullets]
    out = []
    for b in bullets:
        t = sanitize(b)
        out.append(t if QUANT_REGEX.search(t) else f"{t} ({placeholder})")
    return out

# ----------------------------
# OpenAI
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
        messages=[{"role":"system","content":system},{"role":"user","content":prompt}],
    )
    return resp.choices[0].message.content.strip()

def _gpt_call_resp(client, system, prompt) -> str:
    if not hasattr(client, "responses"):
        raise AttributeError("Responses API not available in this SDK")
    resp = client.responses.create(
        model=MODEL,
        input=[{"role":"system","content":system},{"role":"user","content":prompt}],
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
# Summary & Skills (keep originals; append new)
# ----------------------------
def write_summary(doc: Document, summary: str):
    sec = find_section_bounds(doc, ["PROFESSIONAL SUMMARY","SUMMARY"])
    if sec:
        s,e = sec
        if e >= s+1: delete_range(doc, s+1, e)
        insert_paragraph_after(doc.paragraphs[s], sanitize(summary))
    else:
        h = doc.add_paragraph(); h.add_run("PROFESSIONAL SUMMARY").bold = True
        doc.add_paragraph(sanitize(summary))

def parse_skills_section(doc: Document, s: int, e: int):
    order = []
    mapping: Dict[str, List[str]] = {}
    current = None
    i = s + 1
    while i <= e and i < len(doc.paragraphs):
        line = sanitize(doc.paragraphs[i].text)
        if not line:
            i += 1; continue
        if line.endswith(":"):
            current = _norm_heading(line[:-1].strip())
            if current not in mapping:
                mapping[current] = []
                order.append(current)
            j = i + 1
            if j <= e:
                items_line = sanitize(doc.paragraphs[j].text)
                if items_line and not items_line.endswith(":") and not is_major_heading(items_line):
                    mapping[current].extend([x.strip() for x in items_line.split(",") if x.strip()])
                    i = j
            i += 1
            continue
        if ":" in line:
            head, items = line.split(":", 1)
            key = _norm_heading(head.strip())
            if key not in mapping:
                mapping[key] = []
                order.append(key)
            mapping[key].extend([x.strip() for x in items.split(",") if x.strip()])
            i += 1
            continue
        if current:
            mapping[current].extend([x.strip() for x in line.split(",") if x.strip()])
        i += 1
    for k,v in mapping.items():
        vv=[]; seen=set()
        for x in v:
            if x and x.lower() not in seen:
                seen.add(x.lower()); vv.append(x)
        mapping[k] = vv
    return order, mapping

def rewrite_skills_section(doc: Document, s: int, e: int, order: List[str], mapping: Dict[str, List[str]]):
    if e >= s+1:
        delete_range(doc, s+1, e)
    anchor = doc.paragraphs[s]
    last = anchor
    for cat in order:
        header = insert_paragraph_after(last, f"{cat.title().replace(' And ',' & ')}:")
        items = mapping.get(cat, [])
        if items:
            last = insert_paragraph_after(header, ", ".join(items))
        else:
            last = header

def inject_skills(doc: Document, new_skills: Dict[str, list]):
    if not new_skills: return
    spans = find_all_section_bounds(doc, ["SKILLS","CORE SKILLS","TECHNICAL SKILLS","SKILLS & TOOLS","SKILLS AND TOOLS"])
    if not spans:
        h = doc.add_paragraph(); h.add_run("SKILLS & TOOLS").bold = True
        spans = find_all_section_bounds(doc, ["SKILLS & TOOLS","SKILLS AND TOOLS"])
    first_s = spans[0][0]
    last_e = spans[-1][1]
    order, mapping = parse_skills_section(doc, first_s, last_e)
    for cat, additions in new_skills.items():
        key = _norm_heading(cat)
        if key not in mapping:
            mapping[key] = []
            order.append(key)
        existing = set(x.lower() for x in mapping[key])
        for item in (additions or []):
            t = sanitize(item)
            if t and t.lower() not in existing:
                mapping[key].append(t)
                existing.add(t.lower())
    rewrite_skills_section(doc, first_s, last_e, order, mapping)

# ----------------------------
# GPT bullets (batch generation)
# ----------------------------
def gpt_bullets_batch(experience: List[Dict], jd: str, style_rules: List[str], metrics_by_company: Dict[str, List[str]]) -> Dict[str, List[str]]:
    entries = []
    for e in experience:
        k = int(e.get("bullets", 0) or 0)
        if k <= 0: continue
        comp = sanitize(e.get("company","")); role = sanitize(e.get("role",""))
        mx = metrics_by_company.get(comp, [])
        entries.append(f'- company: "{comp}"; role: "{role}"; bullets: {k}; metrics: [{", ".join(mx)}]')
    if not entries:
        return {sanitize(e.get("company","")): [] for e in experience}

    rules_txt = "\n".join(f"- {r}" for r in (style_rules or [
        "20–28 words each",
        "Start with a strong verb; past tense",
        "Use provided metrics naturally (%, $, counts, time); do NOT invent numbers",
        "No company names inside bullets",
        "Avoid buzzwords; focus on actions and outcomes",
    ]))

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
    try:
        data = json.loads(text)
    except Exception:
        data = {}
        for e in experience:
            comp = sanitize(e.get("company","")); k = int(e.get("bullets", 0) or 0)
            role = sanitize(e.get("role",""))
            data[comp] = [f"Drove outcomes in '{role}' aligned to JD" for _ in range(max(k,0))]

    out = {}
    for e in experience:
        comp = sanitize(e.get("company","")); k = int(e.get("bullets", 0) or 0)
        bullets = [sanitize(b) for b in (data.get(comp) or [])][:k]
        out[comp] = ensure_quantified(bullets)
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

@app.route("/tailor", methods=["POST","OPTIONS"])
def tailor():
    origin = request.headers.get("Origin", "*")

    # parse JSON or multipart
    if request.content_type and "multipart/form-data" in request.content_type.lower():
        base_resume_file = request.files.get("base_resume")
        payload_part = request.form.get("payload")
        if not payload_part: return make_response(("missing payload", 400))
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

    # toggles (can be overridden by options)
    global SHOW_KPI_PLACEHOLDER, BULLETS_STRICT_REPLACE
    SHOW_KPI_PLACEHOLDER = _env_bool("SHOW_KPI_PLACEHOLDER", True)
    BULLETS_STRICT_REPLACE = _env_bool("BULLETS_STRICT_REPLACE", True)
    opts = (data.get("options") or {})
    if "show_kpi_placeholder" in opts: SHOW_KPI_PLACEHOLDER = bool(opts["show_kpi_placeholder"])
    if "strict_replace" in opts: BULLETS_STRICT_REPLACE = bool(opts["strict_replace"])
    app.logger.info("Toggles -> strict_replace=%s, show_kpi_placeholder=%s",
                    BULLETS_STRICT_REPLACE, SHOW_KPI_PLACEHOLDER)

    # exact headers (company -> exact visible header line)
    exact_headers: Dict[str,str] = {}
    if "exact_headers" in opts and isinstance(opts["exact_headers"], dict):
        exact_headers = { sanitize(k): v for k, v in opts["exact_headers"].items() if v }

    # inputs
    job_desc = sanitize(data.get("job_description",""))
    cfg = (data.get("resume_config") or {}) or {}
    summary_sentences = int(cfg.get("summary_sentences", 2))
    experience = cfg.get("experience",[]) or []
    skills_categories = cfg.get("skills_categories",[]) or []
    style_rules = (opts.get("style_rules") or [])

    # base resume
    if base_resume_file:
        base_doc = ensure_docx(base_resume_file)
    else:
        base_path = os.path.join(os.path.dirname(__file__), "base_resume.docx")
        if not os.path.exists(base_path):
            return make_response(("server missing base_resume.docx", 500))
        with open(base_path, "rb") as f:
            base_doc = ensure_docx(f)

    # ---------- pass 1: locate anchors to harvest metrics (pre-edit)
    anchors_pre = find_anchors_by_exact_headers(base_doc, exact_headers) if exact_headers else []
    
    # extract numeric phrases from role block + JD to bias GPT toward quantification
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
            if len(found) >= max_phrases: break
        return found

    def harvest_metrics(doc: Document, start_idx: int, end_idx: int) -> List[str]:
        buf = []
        for i in range(start_idx + 1, min(end_idx + 1, len(doc.paragraphs))):
            t = sanitize(doc.paragraphs[i].text)
            if not t or is_major_heading(t): break
            buf.append(t)
        return extract_numeric_phrases("  ".join(buf))

    metrics_by_company: Dict[str, List[str]] = {}
    for comp, start_i in anchors_pre:
        content_start, content_end = find_role_section_bounds(base_doc, start_i)
        role_metrics = harvest_metrics(base_doc, start_i, content_end)
        jd_metrics = extract_numeric_phrases(job_desc)
        merged, seen = [], set()
        for val in role_metrics + jd_metrics:
            if val not in seen:
                seen.add(val); merged.append(val)
        metrics_by_company[comp] = merged[:20]

    # summary text
    summary_prompt = (
        f"Write {summary_sentences} sentence professional summary aligned to the job description below. "
        f"Use concise, specific language; prefer quantified outcomes; avoid buzzwords and dashes.\n---\n{job_desc[:1500]}"
    )
    summary = sanitize(gpt(summary_prompt))

    # bullets (single call for all roles)
    bullets_by_company = gpt_bullets_batch(experience, job_desc, style_rules, metrics_by_company)

    # ---------- mutate doc (these edits shift indices)
    write_summary(base_doc, summary)
    
    # skills: keep originals; append deduped based on JD
    SKILL_BANK = {
        "Business Analysis & Delivery": [
            "Requirements elicitation","User stories","Acceptance criteria","UAT","BPMN","UML",
            "Process re-engineering","Traceability (RTM)","Fit–gap","Prioritization (RICE, MoSCoW)"
        ],
        "Project & Program Management": [
            "Agile Scrum","Kanban","SDLC","RACI","RAID","Sprint planning","Roadmapping","Stakeholder management"
        ],
        "Data, Analytics & BI": [
            "SQL","Python","Pandas","NumPy","Power BI","Tableau","A/B testing","Forecasting","Anomaly detection"
        ],
        "Cloud, Data & MLOps": [
            "AWS Lambda","S3","Redshift","Airflow","dbt","Spark","Databricks","ETL/ELT orchestration"
        ],
        "AI, ML & GenAI": [
            "LLMs","RAG","Prompt design","Hugging Face","FAISS","Model monitoring","Inference optimization"
        ],
        "Enterprise Platforms & Integration": [
            "Salesforce","NetSuite","SAP","ServiceNow","REST APIs","Postman","Git","Azure DevOps","Jira"
        ],
        "Collaboration, Design & Stakeholder": [
            "Miro","Figma","Lucidchart","Wireframing","Prototyping","Executive communication","Workshops"
        ],
    }
    def pick_skills(jd: str, bank: Dict[str, List[str]], top_k: int = 8) -> Dict[str, List[str]]:
        jd_l = jd.lower(); out={}
        for cat, items in bank.items():
            hits=[]
            for s in items:
                pat = re.escape(s.lower()).replace(r"\ ", r"\s+")
                if re.search(rf"(?<![A-Za-z0-9]){pat}(?![A-Za-z0-9])", jd_l):
                    hits.append(s)
            seen=set(); merged=[x for x in hits+items if (x not in seen and not seen.add(x))]
            out[cat] = merged[:top_k]
        return out

    skills_categories = skills_categories or list(SKILL_BANK.keys())
    skills_map = pick_skills(job_desc, {c: SKILL_BANK.get(c, []) for c in skills_categories}) if skills_categories else {}
    inject_skills(base_doc, skills_map)

    # ---------- pass 2: re-anchor AFTER edits (fresh indices for insertion)
    anchors = find_anchors_by_exact_headers(base_doc, exact_headers) if exact_headers else []
    app.logger.info("Final anchors: %s", anchors)

    # ---- Replace content under each role section (process in reverse order to maintain indices)
    if anchors:
        for idx in range(len(anchors)-1, -1, -1):
            comp, header_idx = anchors[idx]
            
            # Find the content bounds for this role section
            content_start, content_end = find_role_section_bounds(base_doc, header_idx)
            
            app.logger.info(f"Processing {comp}: header_idx={header_idx}, content_start={content_start}, content_end={content_end}")
            
            if BULLETS_STRICT_REPLACE:
                # Delete all existing content under this role header
                if content_end >= content_start:
                    app.logger.info(f"Deleting range {content_start} to {content_end}")
                    delete_range(base_doc, content_start, content_end)
                
                # Insert new bullets right after the header
                insert_base = base_doc.paragraphs[header_idx]
                last = insert_base
                for bullet in bullets_by_company.get(comp, []):
                    last = insert_paragraph_after(last, sanitize(bullet), style="List Bullet")
                    app.logger.info(f"Inserted bullet: {sanitize(bullet)[:50]}...")
            else:
                # Append at end of the role block (after existing content)
                if content_end >= 0 and content_end < len(base_doc.paragraphs):
                    insert_base = base_doc.paragraphs[content_end]
                else:
                    insert_base = base_doc.paragraphs[header_idx]
                
                last = insert_base
                for bullet in bullets_by_company.get(comp, []):
                    last = insert_paragraph_after(last, sanitize(bullet), style="List Bullet")
    else:
        app.logger.warning("No anchors detected; bullets not injected.")

    # send file
    out = io.BytesIO(); base_doc.save(out); out.seek(0)
    resp = make_response(send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="Rijul_Chaturvedi_Tailored.docx",
    ))
    resp.headers["Access-Control-Allow-Origin"] = origin
    resp.headers["Vary"] = "Origin"
    return resp

if __name__ == "__main__":
    port = int(os.getenv("PORT","8000"))
    app.run(host="0.0.0.0", port=port)
