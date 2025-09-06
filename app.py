from flask import Flask, request, send_file, jsonify, make_response
from flask_cors import CORS
from docx import Document
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.shared import Pt
from docx.oxml.ns import qn
import io, os, json, re, logging
from typing import List, Dict, Tuple, Optional
from concurrent.futures import ThreadPoolExecutor, TimeoutError as FuturesTimeout
from datetime import datetime

# ----------------------------
# App & CORS
# ----------------------------
app = Flask(__name__)
CORS(app, resources={r"/tailor": {"origins": "chrome-extension://*"}})
app.logger.setLevel(logging.INFO)

# ----------------------------
# OpenAI config
# ----------------------------
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
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

# Behavior toggles
SHOW_KPI_PLACEHOLDER = _env_bool("SHOW_KPI_PLACEHOLDER", False)
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
    "EDUCATION","PROJECTS","CERTIFICATIONS","PUBLICATIONS","ACHIEVEMENTS","ACADEMIC PROJECTS"
}

def humanize_text(text: str) -> str:
    """Remove AI-style dashes and make text more natural"""
    if not text: return ""
    # Replace em/en dashes with simple hyphens or commas
    text = text.replace("—", ", ").replace("–", ", ")
    # Replace smart quotes with regular quotes
    text = re.sub(r"[\u201c\u201d]", '"', text)
    text = re.sub(r"[\u2018\u2019]", "'", text)
    # Remove excessive punctuation patterns common in AI text
    text = re.sub(r'\s*,\s*,+', ',', text)
    text = re.sub(r'\s*;\s*;+', ';', text)
    return text.strip()

def sanitize(text: str) -> str:
    if not text: return ""
    text = text.replace("\u00A0", " ").replace("\t", " ")
    text = humanize_text(text)
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
    return (not t) or t in KNOWN_HEADINGS or (t.isupper() and len(t.split()) <= 4)

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

def set_paragraph_font(paragraph: Paragraph, font_name: str = "Times New Roman", font_size: int = 9):
    """Set font for all runs in a paragraph"""
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(font_size)
        # Set the font for both ASCII and East Asian text
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)

def insert_paragraph_after(paragraph: Paragraph, text: str = "", style: Optional[str] = None) -> Paragraph:
    """Insert a new paragraph after the given paragraph with proper formatting"""
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = Paragraph(new_p, paragraph._parent)
    run = None
    if text:
        run = new_para.add_run(text)
        # Set font to Times New Roman, size 9
        run.font.name = "Times New Roman"
        run.font.size = Pt(9)
        # Ensure font is set for both ASCII and East Asian text
        rPr = run._element.get_or_add_rPr()
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:ascii'), "Times New Roman")
        rFonts.set(qn('w:hAnsi'), "Times New Roman")
    if style:
        try:
            new_para.style = style
        except KeyError:
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

def find_anchors_by_exact_headers(doc: Document, exact_headers: Dict[str,str]) -> List[Tuple[str,int]]:
    pairs = [(sanitize(k), v) for k, v in exact_headers.items()]
    found = []
    used = set()
    for comp, header in pairs:
        idx = None
        for i, p in enumerate(doc.paragraphs):
            if sanitize(p.text) == sanitize(header):
                idx = i; break
        if idx is None:
            for i, p in enumerate(doc.paragraphs):
                if _canon(header) in _canon(p.text):
                    idx = i; break
        if idx is not None and idx not in used:
            found.append((comp, idx)); used.add(idx)
        else:
            app.logger.warning("Header NOT found for %s -> '%s'", comp, header)
    found.sort(key=lambda x: x[1])
    return found

def is_bullet_point(text: str) -> bool:
    """Check if a paragraph is a bullet point"""
    text = text.strip()
    # Check for common bullet markers including the special bullet character
    bullet_markers = ['•', '●', '○', '■', '□', '▪', '▫', '-', '*', '>', '→', '·']
    for marker in bullet_markers:
        if text.startswith(marker):
            return True
    # Also check if it starts with a number followed by period or parenthesis (numbered list)
    if re.match(r'^\d+[.)]\s', text):
        return True
    return False

def find_role_section_bounds(doc: Document, header_idx: int) -> Tuple[int, int]:
    """Find bounds of ALL content under a role header, including all bullets"""
    content_start = header_idx + 1
    content_end = len(doc.paragraphs) - 1
    
    # Continue until we hit a major heading or another role header
    for i in range(header_idx + 1, len(doc.paragraphs)):
        para_text = sanitize(doc.paragraphs[i].text)
        
        # Stop at major headings
        if is_major_heading(para_text):
            content_end = i - 1
            break
        
        # Stop at another role header
        if _is_role_header(para_text, doc, i):
            content_end = i - 1
            break
    
    return content_start, content_end

def _is_role_header(text: str, doc: Document = None, idx: int = None) -> bool:
    """Enhanced check if paragraph is a role header"""
    text = sanitize(text).lower()
    if not text: return False
    
    # Check for bold formatting (strong indicator of headers)
    if doc and idx is not None and idx < len(doc.paragraphs):
        para = doc.paragraphs[idx]
        if para.runs and any(run.bold for run in para.runs):
            # If it has bold text and contains role/company indicators
            role_indicators = [
                "analyst", "manager", "engineer", "developer", "consultant", 
                "director", "specialist", "coordinator", "lead", "senior", "intern",
                "advisor", "architect", "administrator", "officer"
            ]
            if any(keyword in text for keyword in role_indicators):
                return True
    
    # Patterns that indicate role headers
    role_indicators = [
        "analyst", "manager", "engineer", "developer", "consultant", 
        "director", "specialist", "coordinator", "lead", "senior", "intern"
    ]
    
    # Date patterns
    has_date = bool(re.search(r'\b(19|20)\d{2}\b|present|current', text))
    
    # Company/location patterns
    has_location = bool(re.search(r',\s*[A-Z]{2}\b|,\s*(ny|ca|tx|wa|il|ma)\b', text, re.IGNORECASE))
    
    has_role = any(keyword in text for keyword in role_indicators)
    
    # Strong indicators: role + date or role + location
    if has_role and (has_date or has_location):
        return True
    
    # Check for pattern like "Title, Company, Location Date"
    if "," in text and (has_date or has_location):
        return True
    
    return False

# ----------------------------
# Professional Summary (Replace)
# ----------------------------
def replace_summary(doc: Document, summary: str):
    """Replace existing summary completely"""
    sec = find_section_bounds(doc, ["PROFESSIONAL SUMMARY","SUMMARY"])
    if sec:
        s, e = sec
        # Delete all content under summary heading
        if e >= s+1:
            delete_range(doc, s+1, e)
        # Insert new summary with proper formatting
        new_para = insert_paragraph_after(doc.paragraphs[s], sanitize(summary))
        set_paragraph_font(new_para)
    else:
        # Add new summary section if not exists
        h = doc.add_paragraph()
        run = h.add_run("PROFESSIONAL SUMMARY")
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(9)
        p = doc.add_paragraph(sanitize(summary))
        set_paragraph_font(p)

# ----------------------------
# Skills (Preserve & Append)
# ----------------------------
def parse_skills_section(doc: Document, s: int, e: int):
    """Parse existing skills maintaining structure"""
    order = []
    mapping: Dict[str, List[str]] = {}
    current = None
    i = s + 1
    
    while i <= e and i < len(doc.paragraphs):
        line = sanitize(doc.paragraphs[i].text)
        if not line:
            i += 1
            continue
            
        # Category with colon
        if line.endswith(":"):
            current = _norm_heading(line[:-1].strip())
            if current not in mapping:
                mapping[current] = []
                order.append(current)
            # Check next line for items
            j = i + 1
            if j <= e:
                items_line = sanitize(doc.paragraphs[j].text)
                if items_line and not items_line.endswith(":") and not is_major_heading(items_line):
                    mapping[current].extend([x.strip() for x in items_line.split(",") if x.strip()])
                    i = j
            i += 1
            continue
            
        # Inline category: items
        if ":" in line:
            head, items = line.split(":", 1)
            key = _norm_heading(head.strip())
            if key not in mapping:
                mapping[key] = []
                order.append(key)
            mapping[key].extend([x.strip() for x in items.split(",") if x.strip()])
            i += 1
            continue
            
        # Continuation of previous category
        if current:
            mapping[current].extend([x.strip() for x in line.split(",") if x.strip()])
        i += 1
    
    # Deduplicate within each category
    for k, v in mapping.items():
        seen = set()
        deduped = []
        for item in v:
            if item and item.lower() not in seen:
                seen.add(item.lower())
                deduped.append(item)
        mapping[k] = deduped
    
    return order, mapping

def rewrite_skills_section(doc: Document, s: int, e: int, order: List[str], mapping: Dict[str, List[str]]):
    """Rewrite skills section with merged content and proper formatting - inline style"""
    if e >= s+1:
        delete_range(doc, s+1, e)
    
    anchor = doc.paragraphs[s]
    last = anchor
    
    for cat in order:
        # Format category name
        formatted_cat = cat.title().replace(' And ', ' & ')
        items = mapping.get(cat, [])
        
        if items:
            # Create single paragraph with category: items format (inline)
            combined_text = f"{formatted_cat}: {', '.join(items)}"
            skills_para = insert_paragraph_after(last, combined_text)
            set_paragraph_font(skills_para)
            last = skills_para
        else:
            # Just the category if no items
            cat_para = insert_paragraph_after(last, f"{formatted_cat}:")
            set_paragraph_font(cat_para)
            last = cat_para

def merge_skills(doc: Document, new_skills: Dict[str, list]):
    """Preserve existing skills and append new ones"""
    if not new_skills:
        return
    
    spans = find_all_section_bounds(doc, ["SKILLS","CORE SKILLS","TECHNICAL SKILLS","SKILLS & TOOLS","SKILLS AND TOOLS"])
    
    if not spans:
        # Create skills section if not exists
        h = doc.add_paragraph()
        run = h.add_run("SKILLS & TOOLS")
        run.bold = True
        run.font.name = "Times New Roman"
        run.font.size = Pt(9)
        spans = find_all_section_bounds(doc, ["SKILLS & TOOLS","SKILLS AND TOOLS"])
    
    if not spans:
        return
    
    first_s = spans[0][0]
    last_e = spans[-1][1]
    
    # Parse existing skills
    order, mapping = parse_skills_section(doc, first_s, last_e)
    
    # Merge new skills (append only, no replacement)
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
    
    # Rewrite section with merged content
    rewrite_skills_section(doc, first_s, last_e, order, mapping)

# ----------------------------
# OpenAI Integration
# ----------------------------
def _get_client():
    if not OPENAI_API_KEY or not _openai_available or not OAI_ENABLED:
        return None
    try:
        return OpenAI(api_key=OPENAI_API_KEY, timeout=OAI_CLIENT_TIMEOUT, max_retries=0)
    except Exception as e:
        app.logger.exception("OpenAI client init failed: %s", e)
        return None

def _gpt_call(client, system, prompt) -> str:
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[{"role":"system","content":system},{"role":"user","content":prompt}],
        temperature=0.3  # Lower temperature for consistency
    )
    return resp.choices[0].message.content.strip()

def gpt(prompt: str, system: str = "You are a professional resume writer. Write in a natural, human style without AI markers.") -> str:
    client = _get_client()
    if client is None:
        return "Placeholder output (model disabled or key missing)."
    try:
        result = EXECUTOR.submit(_gpt_call, client, system, prompt).result(timeout=OAI_BUDGET)
        return humanize_text(result)
    except Exception as e:
        app.logger.warning("OpenAI error: %s", e)
        return "Placeholder output due to model error."

# ----------------------------
# Enhanced GPT Bullets Generation
# ----------------------------
def gpt_bullets_batch(experience: List[Dict], jd: str, style_rules: List[str], metrics_by_company: Dict[str, List[str]]) -> Dict[str, List[str]]:
    """Generate humanized, quantified bullets"""
    entries = []
    for e in experience:
        k = int(e.get("bullets", 0) or 0)
        if k <= 0: continue
        comp = sanitize(e.get("company",""))
        role = sanitize(e.get("role",""))
        mx = metrics_by_company.get(comp, [])
        entries.append(f'- company: "{comp}"; role: "{role}"; bullets: {k}; metrics: [{", ".join(mx)}]')
    
    if not entries:
        return {sanitize(e.get("company","")): [] for e in experience}
    
    # Enhanced style rules for human-like output
    rules_txt = "\n".join(f"- {r}" for r in (style_rules or [
        "20 to 28 words each bullet",
        "Start with strong action verb in past tense",
        "Include specific numbers and percentages naturally (use provided metrics)",
        "Never use em/en dashes, use commas or 'by' instead",
        "Write like a human, not AI (avoid overly formal language)",
        "Focus on measurable business impact and outcomes",
        "Use simple connecting words like 'and', 'by', 'through'",
        "Each bullet must contain at least one quantified metric"
    ]))
    
    prompt = f"""Generate resume bullets that sound natural and human-written. Return ONLY valid JSON.

CRITICAL: Write naturally without AI markers. Use numbers extensively. Replace dashes with commas or 'by'.

Entries:
{chr(10).join(entries)}

Style Rules:
{rules_txt}

Job Description Focus:
{jd[:1500]}

Return JSON like:
{{
  "Company Name": ["Achieved X by doing Y, resulting in Z% improvement", "Led team of N to deliver..."]
}}"""

    text = gpt(prompt, system="You are a professional resume writer. Generate human-sounding, quantified bullets. Return only valid JSON.")
    
    try:
        data = json.loads(text)
    except Exception:
        data = {}
        for e in experience:
            comp = sanitize(e.get("company",""))
            k = int(e.get("bullets", 0) or 0)
            data[comp] = [f"Delivered measurable outcomes aligned with role responsibilities" for _ in range(max(k,0))]
    
    # Process and humanize bullets
    out = {}
    for e in experience:
        comp = sanitize(e.get("company",""))
        k = int(e.get("bullets", 0) or 0)
        bullets = []
        
        for b in (data.get(comp) or [])[:k]:
            # Additional humanization
            b = humanize_text(b)
            # Ensure quantification
            if not re.search(r'\d+', b):
                b = f"{b} achieving measurable results"
            bullets.append(sanitize(b))
        
        out[comp] = bullets[:k]
    
    return out

# ----------------------------
# Metrics Extraction
# ----------------------------
def extract_numeric_phrases(text: str, max_phrases: int = 15) -> List[str]:
    """Extract quantified metrics from text"""
    NUM_REGEX = re.compile(
        r"(\$?\d+(?:\.\d+)?[KMB]?\b|\d+\+?\s*%|\d+\s*(?:hours?|days?|weeks?|months?|years?)|"
        r"\$\d{1,3}(?:,\d{3})*(?:\.\d+)?|\d{1,3}%|\d+x\b|\d+\s*million|\d+\s*billion)",
        re.IGNORECASE
    )
    
    found = []
    seen = set()
    for m in NUM_REGEX.finditer(text or ""):
        val = sanitize(m.group(0))
        if val and val.lower() not in seen:
            found.append(val)
            seen.add(val.lower())
            if len(found) >= max_phrases:
                break
    return found

# ----------------------------
# Main Route
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
    
    # Parse request
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
    
    # Options
    global SHOW_KPI_PLACEHOLDER, BULLETS_STRICT_REPLACE
    BULLETS_STRICT_REPLACE = True  # Always replace bullets
    opts = data.get("options", {})
    
    exact_headers: Dict[str,str] = {}
    if "exact_headers" in opts and isinstance(opts["exact_headers"], dict):
        exact_headers = {sanitize(k): v for k, v in opts["exact_headers"].items() if v}
    
    # Inputs
    job_desc = sanitize(data.get("job_description",""))
    cfg = data.get("resume_config", {})
    summary_sentences = int(cfg.get("summary_sentences", 2))
    experience = cfg.get("experience", [])
    skills_categories = cfg.get("skills_categories", [])
    style_rules = opts.get("style_rules", [])
    
    # Load base resume
    if base_resume_file:
        base_doc = ensure_docx(base_resume_file)
    else:
        base_path = os.path.join(os.path.dirname(__file__), "base_resume.docx")
        if not os.path.exists(base_path):
            return make_response(("server missing base_resume.docx", 500))
        with open(base_path, "rb") as f:
            base_doc = ensure_docx(f)
    
    # Extract metrics from existing content
    anchors_pre = find_anchors_by_exact_headers(base_doc, exact_headers) if exact_headers else []
    
    metrics_by_company: Dict[str, List[str]] = {}
    for comp, start_i in anchors_pre:
        content_start, content_end = find_role_section_bounds(base_doc, start_i)
        
        # Extract metrics from role section
        buf = []
        for i in range(content_start, min(content_end + 1, len(base_doc.paragraphs))):
            t = sanitize(base_doc.paragraphs[i].text)
            if t:
                buf.append(t)
        
        role_metrics = extract_numeric_phrases(" ".join(buf))
        jd_metrics = extract_numeric_phrases(job_desc)
        
        # Merge metrics
        merged = []
        seen = set()
        for val in role_metrics + jd_metrics:
            if val not in seen:
                seen.add(val)
                merged.append(val)
        
        metrics_by_company[comp] = merged[:20]
    
    # Generate professional summary
    summary_prompt = (
        f"Write exactly {summary_sentences} sentences for a professional summary. "
        f"Use natural language with specific achievements and numbers. "
        f"Avoid AI-style writing markers like dashes. Use 'and', 'through', 'by' as connectors.\n\n"
        f"Job Description:\n{job_desc[:1500]}"
    )
    summary = sanitize(gpt(summary_prompt))
    
    # Generate bullets
    bullets_by_company = gpt_bullets_batch(experience, job_desc, style_rules, metrics_by_company)
    
    # MODIFICATION 1 & 3: Replace summary completely
    replace_summary(base_doc, summary)
    
    # MODIFICATION 2: Preserve original skills, append new ones
    SKILL_BANK = {
        "Business Analysis & Delivery": [
            "Requirements elicitation","User stories","Acceptance criteria","UAT","BPMN","UML",
            "Process re-engineering","Traceability (RTM)","Business case development"
        ],
        "Project & Program Management": [
            "Agile Scrum","Kanban","SDLC","RACI","Sprint planning","Roadmapping","Stakeholder management"
        ],
        "Data Analytics & BI": [
            "SQL","Python","Pandas","NumPy","Power BI","Tableau","A/B testing","Forecasting"
        ],
        "Cloud Data & MLOps": [
            "AWS Lambda","S3","Redshift","Airflow","dbt","Spark","Databricks","ETL/ELT"
        ],
        "AI ML & GenAI": [
            "LLMs","RAG","Prompt design","Hugging Face","FAISS","Model monitoring"
        ],
        "Enterprise Platforms": [
            "Salesforce","NetSuite","SAP","ServiceNow","REST APIs","Jira","Confluence"
        ],
        "Collaboration Design": [
            "Miro","Figma","Lucidchart","Wireframing","Executive communication","Workshops"
        ]
    }
    
    def pick_relevant_skills(jd: str, bank: Dict[str, List[str]], top_k: int = 6) -> Dict[str, List[str]]:
        jd_lower = jd.lower()
        out = {}
        for cat, items in bank.items():
            relevant = []
            for skill in items:
                pattern = re.escape(skill.lower()).replace(r"\ ", r"\s*")
                if re.search(rf"\b{pattern}\b", jd_lower):
                    relevant.append(skill)
            
            # Take relevant skills first, then fill with others
            seen = set()
            merged = []
            for s in relevant + items:
                if s not in seen:
                    seen.add(s)
                    merged.append(s)
                    if len(merged) >= top_k:
                        break
            
            if merged:
                out[cat] = merged
        
        return out
    
    skills_categories = skills_categories or list(SKILL_BANK.keys())
    skills_map = pick_relevant_skills(job_desc, {c: SKILL_BANK.get(c, []) for c in skills_categories})
    merge_skills(base_doc, skills_map)
    
    # MODIFICATION 1: Replace bullets completely (delete old, insert new)
    anchors = find_anchors_by_exact_headers(base_doc, exact_headers) if exact_headers else []
    
    if anchors:
        # Process in reverse order to maintain indices
        for idx in range(len(anchors)-1, -1, -1):
            comp, header_idx = anchors[idx]
            
            # Find content bounds
            content_start, content_end = find_role_section_bounds(base_doc, header_idx)
            
            app.logger.info(f"Processing {comp}: header at {header_idx}, content {content_start}-{content_end}")
            
            # Simply delete ALL content between header and next section
            # This is cleaner than trying to identify individual bullets
            if content_end >= content_start and content_start < len(base_doc.paragraphs):
                # Delete everything in the content range
                delete_range(base_doc, content_start, content_end)
                app.logger.info(f"Deleted content from {content_start} to {content_end} for {comp}")
            
            # Insert new bullets right after the header
            insert_base = base_doc.paragraphs[header_idx]
            last = insert_base
            
            new_bullets = bullets_by_company.get(comp, [])
            for bullet in new_bullets:
                # Ensure bullet starts with bullet point
                bullet_text = sanitize(bullet)
                if not bullet_text.startswith("•"):
                    bullet_text = f"• {bullet_text}"
                new_para = insert_paragraph_after(last, bullet_text)
                set_paragraph_font(new_para)  # Set Times New Roman, 9pt
                last = new_para
            
            app.logger.info(f"Inserted {len(new_bullets)} new bullets for {comp}")
    
    # MODIFICATION 4: Generate filename WITHOUT "Tailored"
    today = datetime.now().strftime("%Y-%m-%d")
    filename = f"Rijul_Chaturvedi_{today}.docx"
    
    # Save and send
    out = io.BytesIO()
    base_doc.save(out)
    out.seek(0)
    
    resp = make_response(send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=filename,
    ))
    resp.headers["Access-Control-Allow-Origin"] = origin
    resp.headers["Vary"] = "Origin"
    return resp

if __name__ == "__main__":
    port = int(os.getenv("PORT", "8000"))
    app.run(host="0.0.0.0", port=port)
