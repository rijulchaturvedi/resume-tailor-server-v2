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
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")  # set to "gpt-5" if you wish
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
USE_RESPONSES_API = os.getenv("USE_RESPONSES_API", "0").strip().lower() in ("1", "true", "yes")
EXECUTOR = ThreadPoolExecutor(max_workers=int(os.getenv("OAI_THREADS", "2")))
OAI_ENABLED = os.getenv("USE_OPENAI", "1").strip().lower() not in ("0", "false", "no")
OAI_BUDGET = float(os.getenv("OAI_BUDGET", "20"))
OAI_CLIENT_TIMEOUT = float(os.getenv("OAI_CLIENT_TIMEOUT", "20"))

def _env_bool(name: str, default: bool=False) -> bool:
    v = os.getenv(name)
    if v is None: return default
    return v.strip().lower() in ("1","true","yes","y","on")

# Behavior toggles (overridable via JSON options)
SHOW_KPI_PLACEHOLDER = _env_bool("SHOW_KPI_PLACEHOLDER", True)
BULLETS_STRICT_REPLACE = _env_bool("BULLETS_STRICT_REPLACE", True)  # strict ON by default

try:
    from openai import OpenAI
    _openai_available = True
except Exception:
    _openai_available = False

# ----------------------------
# Text / headings utils
# ----------------------------
def sanitize(text: str) -> str:
    if not text: return ""
    text = text.replace("\u00A0", " ").replace("\t", " ")
    text = text.replace("—","-").replace("–","-")
    text = re.sub(r"[\u201c\u201d]", '"', text)
    text = re.sub(r"[\u2018\u2019]", "'", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def _canon(s: str) -> str:
    # Normalization used for matching (lower, drop punctuation to words)
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

KNOWN_HEADINGS = {
    "PROFESSIONAL SUMMARY","SUMMARY",
    "EXPERIENCE","WORK EXPERIENCE","PROFESSIONAL EXPERIENCE",
    "SKILLS","CORE SKILLS","TECHNICAL SKILLS","SKILLS & TOOLS","SKILLS AND TOOLS",
    "EDUCATION","PROJECTS","CERTIFICATIONS","PUBLICATIONS","ACHIEVEMENTS"
}

def is_heading(text: str) -> bool:
    t = _norm_heading(text)
    return (not t) or t in KNOWN_HEADINGS or t.isupper()

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
        if is_heading(doc.paragraphs[j].text):
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
                if is_heading(doc.paragraphs[j].text):
                    e = j-1
                    break
            spans.append((s,e))
            i = e + 1
        else:
            i += 1
    return spans

def delete_range(doc: Document, start: int, end: int):
    for i in range(end, start-1, -1):
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
            # fallback bullet glyph if List style isn't in the template
            if run and style.lower().startswith("list") and not run.text.strip().startswith("•"):
                run.text = f"• {run.text}"
    return new_para

# ----------------------------
# DOCX load helper
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

# ----------------------------
# Quantification helpers
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

def is_bullet_para(p) -> bool:
    try:
        name = (p.style.name or "").lower()
    except Exception:
        name = ""
    if "bullet" in name or "list" in name:
        return True
    t = sanitize(p.text)
    if not t: return False
    if t.startswith(("•","▪","– ","- ","— ")): return True
    if re.match(r"^\s*\d+[\.\)]\s+", t): return True
    return False

# ----------------------------
# Role/header detection
# ----------------------------
MONTHS = r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|January|February|March|April|May|June|July|August|September|October|November|December"
ROLE_LINE_RE = re.compile(rf"\b({MONTHS})\b\s+\d{{4}}\s*-\s*(Present|\b({MONTHS})\b\s+\d{{4}})", re.IGNORECASE)

def looks_like_role_header(line: str) -> bool:
    t = sanitize(line)
    if not t: return False
    if ROLE_LINE_RE.search(t): return True
    return "," in t  # role, company, location pattern is common

def find_block_end(doc: Document, start_i: int, next_known_start: Optional[int], end_limit: int) -> int:
    limit = (next_known_start - 1) if next_known_start is not None else end_limit
    for j in range(start_i + 1, min(limit, len(doc.paragraphs) - 1) + 1):
        t = sanitize(doc.paragraphs[j].text)
        if not t:
            continue
        if ROLE_LINE_RE.search(t) or is_heading(t):
            return j - 1
    return limit

# ----------------------------
# Strong exact-header anchoring
# ----------------------------
def _tokens(s: str) -> set:
    return set(_canon(s).split())

def _contains_all(hay: set, need: set) -> bool:
    return bool(need) and need.issubset(hay)

def find_role_anchor_strict(doc: Document, header_text: str, role: str, company: str) -> Optional[int]:
    """
    Find the *true* role header paragraph:
      - Must look like a role header (date range or role/company comma pattern)
      - Must contain tokens from both role and company
      - Prefer exact/substring match over token-only
    Searches the whole doc for robustness.
    """
    if not header_text:
        return None
    target = _canon(header_text)
    role_t = _tokens(role)
    comp_t = _tokens(company)

    candidate_by_score: List[Tuple[int,int]] = []  # (score, idx) higher score wins

    for i in range(len(doc.paragraphs)):
        line = sanitize(doc.paragraphs[i].text)
        if not line: continue
        if not looks_like_role_header(line):  # Only consider plausible headers
            continue
        L = _canon(line)
        bag = _tokens(line)
        if not (_contains_all(bag, role_t) and _contains_all(bag, comp_t)):
            continue

        score = 0
        if L == target: score += 3
        elif target in L: score += 2
        # Token overlap bonus
        score += min(len(bag & (role_t | comp_t)), 6)

        # Prefer lines that also have a date pattern
        if ROLE_LINE_RE.search(line): score += 2

        candidate_by_score.append((score, i))

    if not candidate_by_score:
        return None
    candidate_by_score.sort(reverse=True)
    return candidate_by_score[0][1]

# ----------------------------
# Metrics (resume + JD)
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
        if len(found) >= max_phrases: break
    return found

def harvest_metrics_from_block(doc: Document, start_idx: int, boundary_end: int) -> List[str]:
    buf = []
    for i in range(start_idx + 1, min(boundary_end + 1, len(doc.paragraphs))):
        t = sanitize(doc.paragraphs[i].text)
        if not t or is_heading(t): break
        buf.append(t)
    return extract_numeric_phrases("  ".join(buf))

# ----------------------------
# OpenAI helpers
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
# Summary & Skills writers (keep originals; append deduped)
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
                if items_line and not items_line.endswith(":") and not is_heading(items_line):
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
# GPT bullets (batch)
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
# Helpers to clean the top-of-experience "intro bullets"
# ----------------------------
def purge_intro_bullets(doc: Document, exp_start: int, first_anchor_idx: int):
    """
    Delete anything between the EXPERIENCE heading and the first job header.
    This removes stray bullets that precede the first role.
    """
    s = exp_start + 1
    e = max(first_anchor_idx - 1, s - 1)
    if s <= e:
        delete_range(doc, s, e)

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

    # JSON or multipart
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

    # Toggles
    global SHOW_KPI_PLACEHOLDER, BULLETS_STRICT_REPLACE
    SHOW_KPI_PLACEHOLDER = _env_bool("SHOW_KPI_PLACEHOLDER", True)
    BULLETS_STRICT_REPLACE = _env_bool("BULLETS_STRICT_REPLACE", True)
    opts = (data.get("options") or {})
    if "show_kpi_placeholder" in opts: SHOW_KPI_PLACEHOLDER = bool(opts["show_kpi_placeholder"])
    if "strict_replace" in opts: BULLETS_STRICT_REPLACE = bool(opts["strict_replace"])
    app.logger.info("Toggles -> strict_replace=%s, show_kpi_placeholder=%s",
                    BULLETS_STRICT_REPLACE, SHOW_KPI_PLACEHOLDER)

    # Exact headers map (keys = company, values = visible header line in DOCX)
    exact_headers: Dict[str, str] = {}
    if "exact_headers" in opts and isinstance(opts["exact_headers"], dict):
        exact_headers = { sanitize(k): v for k,v in opts["exact_headers"].items() if v }

    # Inputs
    job_desc = sanitize((data or {}).get("job_description",""))
    cfg = (data or {}).get("resume_config",{}) or {}
    summary_sentences = int(cfg.get("summary_sentences", 2))
    experience = cfg.get("experience",[]) or []
    skills_categories = cfg.get("skills_categories",[]) or []
    style_rules = (opts.get("style_rules") or [])

    # Load base resume
    if base_resume_file:
        base_doc = ensure_docx(base_resume_file)
    else:
        base_path = os.path.join(os.path.dirname(__file__), "base_resume.docx")
        if not os.path.exists(base_path):
            return make_response(("server missing base_resume.docx", 500))
        with open(base_path, "rb") as f:
            base_doc = ensure_docx(f)

    # Summary
    summary_prompt = (
        f"Write {summary_sentences} sentence professional summary aligned to the job description below. "
        f"Use concise, specific language; prefer quantified outcomes; avoid buzzwords and dashes.\n---\n{job_desc[:1500]}"
    )
    summary = sanitize(gpt(summary_prompt))

    # Experience section bounds
    exp_sec = (find_section_bounds(base_doc, ["PROFESSIONAL EXPERIENCE"])
               or find_section_bounds(base_doc, ["WORK EXPERIENCE"])
               or find_section_bounds(base_doc, ["EXPERIENCE"]))
    if not exp_sec:
        exp_start, exp_end = 0, len(base_doc.paragraphs)-1
    else:
        exp_start, exp_end = exp_sec

    # ---- Anchor detection using strict header matcher
    anchors: List[Tuple[str,int]] = []
    used = set()
    for e in experience:
        k = int(e.get("bullets", 0) or 0)
        if k <= 0: 
            continue
        comp = sanitize(e.get("company",""))
        role = sanitize(e.get("role",""))
        hdr = exact_headers.get(comp) or f"{role}, {comp}"
        idx = find_role_anchor_strict(base_doc, hdr, role, comp)
        if idx is None:
            app.logger.warning("Strict header not found for %s / %s", comp, role)
            continue
        if idx not in used:
            anchors.append((comp, idx)); used.add(idx)

    anchors.sort(key=lambda x: x[1])
    app.logger.info("Final anchors: %s", anchors)

    # If we have at least one anchor, purge any intro bullets above it within Experience
    if anchors and exp_start < anchors[0][1]:
        purge_intro_bullets(base_doc, exp_start, anchors[0][1])

    # Metrics (resume block + JD), keyed by company
    full_end = len(base_doc.paragraphs)-1
    metrics_by_company: Dict[str, List[str]] = {}
    for i, (comp, start_i) in enumerate(anchors):
        next_start = anchors[i+1][1] if i+1 < len(anchors) else None
        boundary_end = find_block_end(base_doc, start_i, next_start, full_end)
        role_metrics = harvest_metrics_from_block(base_doc, start_i, boundary_end)
        jd_metrics = extract_numeric_phrases(job_desc)
        seen=set()
        merged=[x for x in role_metrics + jd_metrics if (x not in seen and not seen.add(x))][:20]
        metrics_by_company[comp] = merged
    for e in experience:
        comp = sanitize(e.get("company",""))
        if comp not in metrics_by_company:
            metrics_by_company[comp] = extract_numeric_phrases(job_desc)

    # Generate bullets
    bullets_by_company = gpt_bullets_batch(experience, job_desc, style_rules, metrics_by_company)

    # Skills (keep originals; append deduped)
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

    # Apply edits
    doc = base_doc
    write_summary(doc, summary)
    inject_skills(doc, skills_map)

    # Strict bullet replacement directly under each header
    if anchors:
        for idx in range(len(anchors)-1, -1, -1):  # bottom-up
            comp, start_i = anchors[idx]
            next_start = anchors[idx+1][1] if idx+1 < len(anchors) else None
            boundary_end = find_block_end(doc, start_i, next_start, full_end)

            # delete everything after the header down to the next header
            if BULLETS_STRICT_REPLACE and boundary_end >= start_i + 1:
                delete_range(doc, start_i + 1, boundary_end)
            elif not BULLETS_STRICT_REPLACE:
                j = start_i + 1
                while j <= boundary_end and j < len(doc.paragraphs) and is_bullet_para(doc.paragraphs[j]):
                    j += 1
                if j-1 >= start_i+1:
                    delete_range(doc, start_i+1, j-1)

            # insert bullets immediately *below* the header
            anchor_para = doc.paragraphs[start_i]
            last = anchor_para
            for b in bullets_by_company.get(comp, []):
                last = insert_paragraph_after(last, sanitize(b), style="List Bullet")
    else:
        app.logger.warning("No anchors detected; bullets not injected.")

    # Return file
    out = io.BytesIO(); doc.save(out); out.seek(0)
    resp = make_response(send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="Rijul_Chaturvedi_Tailored.docx",
    ))
    resp.headers["Access-Control-Allow-Origin"] = origin
    resp.headers["Vary"] = "Origin"
    return resp

# ---------------
# Local dev
# ---------------
if __name__ == "__main__":
    port = int(os.getenv("PORT","8000"))
    app.run(host="0.0.0.0", port=port)
