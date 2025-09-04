import os, io, json, re, logging
from flask import Flask, request, send_file, jsonify
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from typing import List, Tuple, Optional, Dict

try:
    from openai import OpenAI
    _openai_available = True
except Exception:
    _openai_available = False

MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

app = Flask(__name__)
app.logger.setLevel(logging.INFO)

def sanitize(text: str) -> str:
    if not text: return ""
    text = text.replace("—", "-").replace("–", "-")
    text = re.sub(r"\u201c|\u201d", '"', text)
    text = re.sub(r"\u2018|\u2019", "'", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()

def find_section_bounds(doc: Document, titles: List[str]) -> Optional[Tuple[int, int]]:
    txt = [(i, sanitize(p.text)) for i,p in enumerate(doc.paragraphs) if sanitize(p.text)]
    start = next((i for i,t in txt if t.upper() in [s.upper() for s in titles]), None)
    if start is None: return None
    KNOWN = set(["PROFESSIONAL SUMMARY","SUMMARY","EXPERIENCE","WORK EXPERIENCE",
                 "SKILLS","CORE SKILLS","TECHNICAL SKILLS","EDUCATION","PROJECTS",
                 "CERTIFICATIONS","PUBLICATIONS","ACHIEVEMENTS"] + [s.upper() for s in titles])
    end = len(doc.paragraphs)-1
    for i,t in txt:
        if i<=start: continue
        if (t.isupper() and len(t)<=40) or (t.upper() in KNOWN):
            end = i-1; break
    return (start, end)

def delete_paragraph_range(doc: Document, start: int, end: int):
    for i in range(end, start-1, -1):
        p = doc.paragraphs[i]._element
        p.getparent().remove(p)

def add_heading(doc: Document, text: str):
    p = doc.add_paragraph()
    run = p.add_run(text.upper()); run.bold = True
    p.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
    return p

def add_bullet(doc: Document, text: str):
    return doc.add_paragraph(text, style="List Bullet")

def trim_to_word_range(sentence: str, min_w: int = 18, max_w: int = 30) -> str:
    words = sentence.split()
    if len(words) > max_w: words = words[:max_w]
    return " ".join(words)

def _get_client():
    if not OPENAI_API_KEY or not _openai_available: return None
    return OpenAI(api_key=OPENAI_API_KEY)

def gpt_generate(prompt: str, system: str = "You are a helpful writing assistant.") -> str:
    client = _get_client()
    if client is None:
        return "Placeholder output because OPENAI_API_KEY is not set."
    resp = client.chat.completions.create(
        model=MODEL,
        messages=[{"role":"system","content":system},{"role":"user","content":prompt}],
        temperature=0.2,
    )
    return resp.choices[0].message.content.strip()

def build_summary_prompt(jd: str, n: int, rules: List[str]) -> str:
    r = "\n".join(f"- {x}" for x in rules)
    return f"""Write a professional resume summary in {n} sentences.
Follow these rules:
{r}
Focus tightly on this job description:
---
{jd[:4000]}
---
Output only {n} sentence(s)."""

def build_bullets_prompt(company: str, role: str, k: int, jd: str, rules: List[str]) -> str:
    r = "\n".join(f"- {x}" for x in rules)
    return f"""Write exactly {k} resume bullets for "{role}" at "{company}".
Each bullet 20–28 words, specific, metricized when truthful. Do not include company names in bullets.
Rules:
{r}
Tailor to this JD:
---
{jd[:4000]}
---
Return as a numbered list 1..{k} with no extra text."""

def build_skills_prompt(cats: List[str], jd: str) -> str:
    c = "\n".join(f"- {x}" for x in cats)
    return f"""Regroup skills under these exact categories (5–8 per category):
{c}
Use concise comma-separated skills. Prioritize alignment to:
---
{jd[:4000]}
---
Return JSON: {{"Category": ["Skill1","Skill2"] ...}}"""

def parse_numbered_list(text: str, expected: int) -> List[str]:
    lines = [sanitize(l) for l in text.splitlines() if sanitize(l)]
    items = []
    for l in lines:
        m = re.match(r"^\s*(\d+)[\.\)]\s+(.*)$", l)
        if m: items.append(m.group(2).strip())
        elif l.startswith("- "): items.append(l[2:].strip())
    if not items and lines: items = lines
    return [trim_to_word_range(x, 18, 30) for x in items][:expected]

@app.get("/health")
def health():
    return jsonify({"status": "ok", "model": MODEL})

@app.post("/tailor")
def tailor():
    if "base_resume" not in request.files:
        return jsonify({"error":"missing base_resume file"}), 400
    payload_part = request.form.get("payload")
    if not payload_part:
        return jsonify({"error":"missing payload JSON"}), 400
    try:
        payload = json.loads(payload_part)
    except Exception as e:
        return jsonify({"error": f"invalid payload JSON: {e}"}), 400

    jd = sanitize(payload.get("job_description",""))
    cfg = payload.get("resume_config",{}) or {}
    sentences = int(cfg.get("summary_sentences", 2))
    experience = cfg.get("experience",[]) or []
    skills_categories = cfg.get("skills_categories",[]) or []
    style_rules = (payload.get("options",{}) or {}).get("style_rules",[]) or []

    summary = sanitize(gpt_generate(build_summary_prompt(jd, sentences, style_rules)))
    bullets_by_company: Dict[str, List[str]] = {}
    for e in experience:
        comp, role = e.get("company",""), e.get("role","")
        k = int(e.get("bullets", 0))
        if k <= 0: bullets_by_company[comp] = []; continue
        raw = gpt_generate(build_bullets_prompt(comp, role, k, jd, style_rules))
        bullets_by_company[comp] = parse_numbered_list(raw, k)

    skills_raw = gpt_generate(build_skills_prompt(skills_categories, jd))
    try:
        skills_map = json.loads(skills_raw)
    except Exception:
        skills_map = {c: [] for c in skills_categories}
        if skills_categories:
            skills_map[skills_categories[0]] = ["SQL","Python","Stakeholder management"]

    in_bytes = io.BytesIO(request.files["base_resume"].read())
    doc = Document(in_bytes)

    # Summary
    sec = find_section_bounds(doc, ["PROFESSIONAL SUMMARY","SUMMARY"])
    if sec:
        s, eidx = sec
        delete_paragraph_range(doc, s+1, eidx)
        p = doc.paragraphs[s+1] if s+1 < len(doc.paragraphs) else doc.add_paragraph("")
        p.text = sanitize(summary)
    else:
        add_heading(doc, "Professional Summary")
        doc.add_paragraph(sanitize(summary))

    # Skills
    sec = find_section_bounds(doc, ["SKILLS","CORE SKILLS","TECHNICAL SKILLS"])
    if sec:
        s, eidx = sec
        delete_paragraph_range(doc, s+1, eidx)
        for cat in skills_categories:
            doc.add_paragraph(cat + ":")
            skills = skills_map.get(cat, [])
            if skills: doc.add_paragraph(", ".join(skills))
    else:
        add_heading(doc, "Skills")
        for cat in skills_categories:
            doc.add_paragraph(cat + ":")
            skills = skills_map.get(cat, [])
            if skills: doc.add_paragraph(", ".join(skills))

    # Experience
    sec = find_section_bounds(doc, ["EXPERIENCE","WORK EXPERIENCE"])
    if sec:
        s, eidx = sec
        company_hits = {}
        for i in range(s, eidx+1):
            t = sanitize(doc.paragraphs[i].text)
            for comp in bullets_by_company.keys():
                if comp and comp.lower() in t.lower():
                    company_hits.setdefault(comp, []).append(i)
        for comp, idxs in company_hits.items():
            for ci in idxs:
                j = ci + 1
                while j <= eidx and sanitize(doc.paragraphs[j].text) and not sanitize(doc.paragraphs[j].text).isupper():
                    j += 1
                delete_paragraph_range(doc, ci+1, j-1)
                for b in bullets_by_company.get(comp, []):
                    add_bullet(doc, sanitize(b))
    else:
        add_heading(doc, "Experience")
        for comp, bullets in bullets_by_company.items():
            doc.add_paragraph(comp)
            for b in bullets: add_bullet(doc, sanitize(b))

    out = io.BytesIO()
    doc.save(out); out.seek(0)
    return send_file(out,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True, download_name="Rijul_Chaturvedi_Tailored.docx")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.getenv("PORT","8000")))
