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
MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")  # or "gpt-5"
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

# Per-request toggles (also read from JSON)
# DEFAULT STRICT REPLACE = ON so it actually replaces even if the extension forgets to send options.
SHOW_KPI_PLACEHOLDER = _env_bool("SHOW_KPI_PLACEHOLDER", True)
BULLETS_STRICT_REPLACE = _env_bool("BULLETS_STRICT_REPLACE", True)

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
    # normalize NBSP & tabs
    text = text.replace("\u00A0", " ").replace("\t", " ")
    # normalize dashes/quotes
    text = text.replace("—","-").replace("–","-")
    text = re.sub(r"[\u201c\u201d]", '"', text)
    text = re.sub(r"[\u2018\u2019]", "'", text)
    # collapse whitespace
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

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
            # Fallback bullet glyph if style missing
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
# Company/role detection
# ----------------------------
STOPWORDS = {"inc","llc","llp","ltd","plc","corp","co","company","university","the"}

def norm_tokens(s: str) -> List[str]:
    s = sanitize(s).lower()
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"[^a-z0-9\s&\-]", " ", s)
    toks = [t for t in s.split() if t and t not in STOPWORDS]
    return toks

def company_aliases(company: str) -> List[str]:
    c = sanitize(company)
    parts = [p.strip() for p in c.split(",") if p.strip()]
    core = parts[0] if parts else c
    aliases = set()
    aliases.add(" ".join(norm_tokens(c)))
    aliases.add(" ".join(norm_tokens(core)))
    core_words = norm_tokens(core)
    if core_words:
        aliases.add(" ".join(core_words[:2]))
    return [a for a in aliases if a]

MONTHS = r"Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|January|February|March|April|May|June|July|August|September|October|November|December"
ROLE_LINE_RE = re.compile(rf"\b({MONTHS})\b\s+\d{{4}}\s*-\s*(Present|\b({MONTHS})\b\s+\d{{4}})", re.IGNORECASE)

def scan_role_headers(doc: Document, exp_start: int, exp_end: int) -> List[Tuple[int, str, str]]:
    headers = []
    for i in range(exp_start+1, exp_end+1):
        t = sanitize(doc.paragraphs[i].text)
        if not t: continue
        if ROLE_LINE_RE.search(t):
            segs = [s.strip() for s in t.split(",")]
            company_guess = segs[1] if len(segs) >= 2 else ""
            headers.append((i, t, company_guess))
    return headers

def map_experience_to_positions(experience: List[Dict], headers: List[Tuple[int,str,str]]) -> List[Tuple[str,int]]:
    used = set()
    positions: List[Tuple[str,int]] = []
    for e in experience:
        comp = sanitize(e.get("company",""))
        role = sanitize(e.get("role",""))
        aliases = company_aliases(comp)
        best = None; best_score = -1
        for (idx, line, company_guess) in headers:
            if idx in used: continue
            lc = sanitize(line).lower()
            score = 0
            if any(a and a in lc for a in aliases): score += 2
            rtoks = set(norm_tokens(role))
            ltoks = set(norm_tokens(line))
            if rtoks and (len(rtoks & ltoks) > 0): score += 1
            cg = " ".join(norm_tokens(company_guess))
            if cg and any(a and a in cg for a in aliases): score += 1
            if score > best_score:
                best_score = score; best = idx
        if best is not None:
            used.add(best); positions.append((comp, best))
        else:
            app.logger.info("No header match for '%s' / role '%s'", comp, role)
    positions.sort(key=lambda x: x[1])
    return positions

def next_heading_between(doc: Document, a: int, b: int) -> Optional[int]:
    for j in range(max(a,0), min(b, len(doc.paragraphs)) + 1):
        if is_heading(doc.paragraphs[j].text):
            return j
    return None

# Fallback: if a role header wasn’t detected, anchor by role+company text (no dates required)
def find_role_anchor_by_text(doc: Document, exp_start: int, exp_end: int, company: str, role: str) -> Optional[int]:
    aliases = company_aliases(company)
    role_toks = set(norm_tokens(role))
    best_idx, best_hits = None, -1
    for i in range(exp_start+1, exp_end+1):
        line = sanitize(doc.paragraphs[i].text).lower()
        if not line: continue
        if any(a and a in line for a in aliases):
            hits = len(role_toks & set(norm_tokens(line)))
            if hits > best_hits:
                best_hits = hits
                best_idx = i
    return best_idx

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
# Writers: summary & skills
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

def inject_skills(doc: Document, new_skills: Dict[str, list]):
    if not new_skills: return
    sec = find_section_bounds(doc, ["SKILLS","CORE SKILLS","TECHNICAL SKILLS","SKILLS & TOOLS","SKILLS AND TOOLS"])
    if sec is None:
        h = doc.add_paragraph(); h.add_run("SKILLS & TOOLS").bold = True
        sec = find_section_bounds(doc, ["SKILLS & TOOLS","SKILLS AND TOOLS"])
    s,e = sec

    def norm_cat_name(name: str) -> str: return _norm_heading(name).replace("  "," ").strip()
    first_idx: Dict[str, Dict[str, object]] = {}
    to_delete: List[Tuple[int,int]] = []
    current_cat = None

    for idx in range(s+1, e+1):
        line = sanitize(doc.paragraphs[idx].text)
        if not line: continue
        # Inline "Category: items"
        if ":" in line and not line.endswith(":"):
            head, items_line = line.split(":", 1)
            key = norm_cat_name(head.strip())
            items = [x.strip() for x in items_line.split(",") if x.strip()]
            if key not in first_idx:
                header = doc.paragraphs[idx]
                header.text = f"{head.strip()}:"
                insert_paragraph_after(header, ", ".join(items))
                first_idx[key] = {"header": idx, "items_idx": idx+1, "items": set(items)}
                e += 1
            else:
                first_idx[key]["items"].update(items)
                to_delete.append((idx, idx))
            continue
        # Header line
        if line.endswith(":"):
            key = norm_cat_name(line[:-1].strip())
            if key not in first_idx:
                first_idx[key] = {"header": idx, "items_idx": None, "items": set()}
                current_cat = key
            else:
                dup_start, dup_end = idx, idx
                if idx+1 <= e:
                    nxt = sanitize(doc.paragraphs[idx+1].text)
                    if nxt and not nxt.endswith(":") and not is_heading(nxt):
                        dup_end = idx+1
                        first_idx[key]["items"].update([x.strip() for x in nxt.split(",") if x.strip()])
                to_delete.append((dup_start, dup_end))
                current_cat = None
            continue
        # Items line
        if current_cat is not None:
            first_idx[current_cat]["items"].update([x.strip() for x in line.split(",") if x.strip()])
            if first_idx[current_cat]["items_idx"] is None:
                first_idx[current_cat]["items_idx"] = idx

    # Delete duplicates bottom-up
    for a,b in sorted(to_delete, key=lambda x: x[0], reverse=True):
        delete_range(doc, a, b); e -= (b-a+1)

    # Write back merged items for existing categories
    for key,meta in first_idx.items():
        h = int(meta["header"]); it = meta["items_idx"]
        items_sorted = sorted(set(meta["items"]), key=lambda x: x.lower())
        if items_sorted:
            if it is not None and it < len(doc.paragraphs):
                doc.paragraphs[it].text = ", ".join(items_sorted)
            else:
                insert_paragraph_after(doc.paragraphs[h], ", ".join(items_sorted))

    # Append new incoming skills into existing or create new headers
    for cat, additions in new_skills.items():
        additions = [x for x in additions if x]
        if not additions: continue
        key = norm_cat_name(cat)
        if key in first_idx:
            h = int(first_idx[key]["header"])
            it = first_idx[key]["items_idx"]
            if it is not None and it < len(doc.paragraphs):
                para = doc.paragraphs[it]
                existing = [x.strip() for x in sanitize(para.text).split(",") if x.strip()]
                add = [x for x in additions if x not in existing]
                if add:
                    joiner = (", " if para.text and not para.text.strip().endswith(",") else "")
                    para.text = f"{sanitize(para.text)}{joiner}{', '.join(add)}"
            else:
                insert_paragraph_after(doc.paragraphs[h], ", ".join(additions))
        else:
            s2,e2 = (find_section_bounds(doc, ["SKILLS","CORE SKILLS","TECHNICAL SKILLS","SKILLS & TOOLS","SKILLS AND TOOLS"]) or (0, len(doc.paragraphs)-1))
            anchor = doc.paragraphs[e2]
            header = insert_paragraph_after(anchor, f"{cat}:")
            insert_paragraph_after(header, ", ".join(additions))

# ----------------------------
# Bullets: generate & inject
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
        "Prefer the provided metrics; use them naturally (%, $, counts, time)",
        "Do NOT invent numbers; if none apply, leave unnumbered",
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

def inject_bullets_strict(doc: Document,
                          positions: List[Tuple[str,int]],
                          bullets_by_company: Dict[str, List[str]],
                          exp_start: int,
                          exp_end: int):
    """
    Strict replace:
      - process bottom-up,
      - delete everything after each role line up to the next role line or next heading,
      - insert new bullets immediately AFTER the matched role line.
    """
    if not positions: return

    for idx in range(len(positions)-1, -1, -1):
        comp, line_i = positions[idx]
        next_line = positions[idx+1][1] if idx+1 < len(positions) else exp_end+1
        h = next_heading_between(doc, line_i+1, next_line-1)
        boundary_end = (h-1) if h is not None else (next_line-1)

        # delete full body (not the header itself)
        if boundary_end >= line_i + 1:
            delete_range(doc, line_i + 1, boundary_end)

        # insert bullets right below header
        anchor = doc.paragraphs[line_i]
        last = anchor
        for b in bullets_by_company.get(comp, []):
            last = insert_paragraph_after(last, sanitize(b), style="List Bullet")

# ----------------------------
# Routes
# ----------------------------
@app.route("/")
def index():
    return jsonify({"ok": True, "service": "resume-tailor-server", "endpoints": ["/health", "/tailor"]])

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

    # Refresh toggles per request (default strict replace = True)
    global SHOW_KPI_PLACEHOLDER, BULLETS_STRICT_REPLACE
    SHOW_KPI_PLACEHOLDER = _env_bool("SHOW_KPI_PLACEHOLDER", True)
    BULLETS_STRICT_REPLACE = _env_bool("BULLETS_STRICT_REPLACE", True)
    opts = (data.get("options") or {})
    if "show_kpi_placeholder" in opts: SHOW_KPI_PLACEHOLDER = bool(opts["show_kpi_placeholder"])
    if "strict_replace" in opts: BULLETS_STRICT_REPLACE = bool(opts["strict_replace"])
    app.logger.info("Toggles -> strict_replace=%s, show_kpi_placeholder=%s",
                    BULLETS_STRICT_REPLACE, SHOW_KPI_PLACEHOLDER)

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

    # Experience bounds (accept variants)
    exp_sec = (find_section_bounds(base_doc, ["PROFESSIONAL EXPERIENCE"])
               or find_section_bounds(base_doc, ["WORK EXPERIENCE"])
               or find_section_bounds(base_doc, ["EXPERIENCE"]))
    if not exp_sec:
        exp_start, exp_end = 0, len(base_doc.paragraphs)-1
    else:
        exp_start, exp_end = exp_sec

    # Detect role headers and map to your experience entries
    headers = scan_role_headers(base_doc, exp_start, exp_end)
    positions = map_experience_to_positions(experience, headers)

    # Fallback 1: anchor by role+company text (no date needed)
    if experience:
        claimed = {idx for _, idx in positions}
        for e in experience:
            comp = sanitize(e.get("company","")); role = sanitize(e.get("role",""))
            if any(p[0] == comp for p in positions):
                continue
            anchor = find_role_anchor_by_text(base_doc, exp_start, exp_end, comp, role)
            if anchor is not None and anchor not in claimed:
                positions.append((comp, anchor)); claimed.add(anchor)

    # Fallback 2: pair remaining experiences by order to remaining detected headers
    if headers:
        used_idxs = {i for _, i in positions}
        free_headers = [h for h in headers if h[0] not in used_idxs]
        need = [e for e in experience if sanitize(e.get("company","")) not in {c for c,_ in positions}]
        fill = min(len(need), len(free_headers))
        for k in range(fill):
            comp_name = sanitize(need[k].get("company",""))
            positions.append((comp_name, free_headers[k][0]))
    positions.sort(key=lambda x: x[1])
    app.logger.info("Final mapped positions: %s", positions)

    # Build metrics per mapped role (resume block + JD)
    metrics_by_company: Dict[str, List[str]] = {}
    for i, (comp, line_i) in enumerate(positions):
        next_line = positions[i+1][1] if i+1 < len(positions) else exp_end+1
        h = next_heading_between(base_doc, line_i+1, next_line-1)
        boundary_end = (h-1) if h is not None else (next_line-1)
        role_metrics = harvest_metrics_from_block(base_doc, line_i, boundary_end)
        jd_metrics = extract_numeric_phrases(job_desc)
        seen=set()
        merged=[x for x in role_metrics + jd_metrics if (x not in seen and not seen.add(x))][:20]
        metrics_by_company[comp] = merged
    for e in experience:
        comp = sanitize(e.get("company",""))
        metrics_by_company.setdefault(comp, extract_numeric_phrases(job_desc))

    # Generate bullets (single call)
    bullets_by_company = gpt_bullets_batch(experience, job_desc, style_rules, metrics_by_company)

    # Skills map (JD-driven)
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
    skills_map = pick_skills(job_desc, {c: SKILL_BANK.get(c, []) for c in skills_categories}) if skills_categories else {}

    # Edit DOCX
    doc = base_doc
    write_summary(doc, summary)
    inject_skills(doc, skills_map)

    if positions:
        # Strict replace is ON by default
        if BULLETS_STRICT_REPLACE:
            inject_bullets_strict(doc, positions, bullets_by_company, exp_start, exp_end)
        else:
            # non-strict: clear contiguous bullets only
            for comp, line_i in reversed(positions):
                j = line_i + 1
                while j < len(doc.paragraphs) and is_bullet_para(doc.paragraphs[j]):
                    j += 1
                if j-1 >= line_i+1:
                    delete_range(doc, line_i+1, j-1)
                anchor = doc.paragraphs[line_i]
                last = anchor
                for b in bullets_by_company.get(comp, []):
                    last = insert_paragraph_after(last, sanitize(b), style="List Bullet")
    else:
        app.logger.warning("No role headers detected or mapped; bullets not injected.")

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
