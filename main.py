import io
import json
import os
import re
from copy import deepcopy
from typing import Dict, Iterable, List, Optional, Tuple

from fastapi import FastAPI, File, Form, Header, HTTPException, UploadFile
from fastapi.responses import StreamingResponse
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

app = FastAPI()

PLACEHOLDER = "{{CONTENT}}"


# -------------------------
# Utilities: docx traversal
# -------------------------

def iter_paragraphs_in_cell(cell) -> Iterable[Paragraph]:
    for p in cell.paragraphs:
        yield p
    for t in cell.tables:
        for row in t.rows:
            for c in row.cells:
                yield from iter_paragraphs_in_cell(c)

def iter_all_paragraphs(doc: Document) -> Iterable[Paragraph]:
    # Top-level paragraphs
    for p in doc.paragraphs:
        yield p
    # Paragraphs inside tables
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                yield from iter_paragraphs_in_cell(cell)


# -------------------------
# Utilities: styles
# -------------------------

def list_style_ids(doc: Document) -> Dict[str, str]:
    """Return mapping style_name -> style_id for styles in doc."""
    out = {}
    for s in doc.styles:
        try:
            # Some styles may not have name/id accessible cleanly
            if getattr(s, "name", None) and getattr(s, "style_id", None):
                out[s.name] = s.style_id
        except Exception:
            continue
    return out

def pick_style_id(template_doc: Document, candidates: List[str]) -> Optional[str]:
    """Pick first existing style_id in template by name candidates."""
    style_map = list_style_ids(template_doc)
    for name in candidates:
        if name in style_map:
            return style_map[name]
    return None

def set_paragraph_style_id(p: Paragraph, style_id: str) -> None:
    """Set paragraph's w:pStyle to style_id (no need style exist in source doc)."""
    if not style_id:
        return
    p_elm = p._p
    pPr = p_elm.get_or_add_pPr()
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        pStyle = OxmlElement("w:pStyle")
        pPr.insert(0, pStyle)
    pStyle.set(qn("w:val"), style_id)


# -------------------------
# Rules parsing & matching
# -------------------------

def _compile_detect(detect: str):
    """
    detect supports:
    - "regex:<pattern>"  -> regex search
    - otherwise          -> contains match (supports "|" as OR)
    """
    detect = (detect or "").strip()
    if not detect:
        return None

    if detect.lower().startswith("regex:"):
        pat = detect[6:].strip()
        try:
            rx = re.compile(pat)
            return ("regex", rx)
        except re.error:
            return None
    else:
        keys = [k.strip() for k in detect.split("|") if k.strip()]
        if not keys:
            return None
        return ("contains", keys)

def _match(compiled, text: str) -> bool:
    if compiled is None:
        return False
    kind, obj = compiled
    if kind == "regex":
        return obj.search(text) is not None
    return any(k in text for k in obj)

def normalize_rules(rules_obj: dict) -> dict:
    """
    Accept multiple shapes:
    1) {"formatting_intent": {"heading_rules":[{level, detect, style_name|style_key}], "body_style_name":...}}
    2) {"rules": {...}}  (your own future schema)
    """
    if not isinstance(rules_obj, dict):
        return {}

    if "formatting_intent" in rules_obj and isinstance(rules_obj["formatting_intent"], dict):
        return rules_obj

    # You can extend here if you later change LLM output format.
    if "rules" in rules_obj and isinstance(rules_obj["rules"], dict):
        # Convert to formatting_intent-like structure if present
        out = {"formatting_intent": {"heading_rules": []}}
        heading = rules_obj["rules"].get("heading") or []
        for item in heading:
            try:
                lvl = int(item.get("level", 1))
                detect = item.get("detect", "")
                out["formatting_intent"]["heading_rules"].append(
                    {"level": lvl, "detect": detect, "style_key": f"h{lvl}"}
                )
            except Exception:
                continue
        out["formatting_intent"]["body_style_key"] = "body"
        # Optional sections:
        for k in ["references_start", "figure_caption", "table_caption"]:
            if k in rules_obj["rules"]:
                out["formatting_intent"][k] = rules_obj["rules"][k]
        return out

    return {}

def build_matchers_from_rules(rules_obj: dict) -> List[Tuple[int, object]]:
    """
    Returns list of (level, compiled_detector) sorted by level desc (3->1)
    """
    intent = (rules_obj or {}).get("formatting_intent") or {}
    heading_rules = intent.get("heading_rules") or []
    matchers = []
    for r in heading_rules:
        try:
            level = int(r.get("level"))
        except Exception:
            continue
        compiled = _compile_detect(r.get("detect", ""))
        if compiled is None:
            continue
        matchers.append((level, compiled))
    # Prefer deeper levels first
    matchers.sort(key=lambda x: x[0], reverse=True)
    return matchers


# -------------------------
# MVP Heuristics (no LLM)
# -------------------------

RX_H1 = re.compile(r"^(第[一二三四五六七八九十]+章)\b|^\d+\s*[\.、]\s+\S")
RX_H2 = re.compile(r"^\d+\.\d+\s+\S")
RX_H3 = re.compile(r"^\d+\.\d+\.\d+\s+\S")

RX_REF_TITLE = re.compile(r"^(参考文献|REFERENCES)\b", re.IGNORECASE)
RX_FIG_CAP = re.compile(r"^图\s*\d+|^Figure\s*\d+", re.IGNORECASE)
RX_TAB_CAP = re.compile(r"^表\s*\d+|^Table\s*\d+", re.IGNORECASE)


# -------------------------
# Template style mapping
# -------------------------

def get_template_style_ids(template_doc: Document) -> Dict[str, Optional[str]]:
    """
    Returns style_id map by semantic keys.
    Prefer CM* styles if present, else fall back to default Word names.
    """
    return {
        "body": pick_style_id(template_doc, ["CMcontent", "Normal", "Body Text"]),
        "h1": pick_style_id(template_doc, ["CMheading1", "Heading 1", "标题 1", "标题1", "一级标题"]),
        "h2": pick_style_id(template_doc, ["CMheading2", "Heading 2", "标题 2", "标题2", "二级标题"]),
        "h3": pick_style_id(template_doc, ["CMheading3", "Heading 3", "标题 3", "标题3", "三级标题"]),
        "fig_caption": pick_style_id(template_doc, ["CMcaption", "Caption"]),
        "tab_caption": pick_style_id(template_doc, ["CMcaptionTable", "Caption"]),
        "ref_item": pick_style_id(template_doc, ["CMreflist", "Bibliography", "References"]),
    }


# -------------------------
# Apply styles (content unchanged)
# -------------------------

def apply_styles_in_place(
    template_doc: Document,
    source_doc: Document,
    rules_obj: Optional[dict] = None
) -> None:
    style_ids = get_template_style_ids(template_doc)

    normalized = normalize_rules(rules_obj or {})
    matchers = build_matchers_from_rules(normalized)

    # Optional section detectors from rules (if provided)
    intent = (normalized or {}).get("formatting_intent") or {}
    ref_start_det = _compile_detect(intent.get("references_start")) if intent.get("references_start") else None
    fig_cap_det = _compile_detect(intent.get("figure_caption")) if intent.get("figure_caption") else None
    tab_cap_det = _compile_detect(intent.get("table_caption")) if intent.get("table_caption") else None

    in_references = False

    for p in iter_all_paragraphs(source_doc):
        text = (p.text or "").strip()
        if not text:
            continue

        # Detect references section start
        if ref_start_det and _match(ref_start_det, text):
            in_references = True
            # Title line itself can be treated as H1 if available
            if style_ids.get("h1"):
                set_paragraph_style_id(p, style_ids["h1"])
            continue
        elif (not ref_start_det) and RX_REF_TITLE.search(text):
            in_references = True
            if style_ids.get("h1"):
                set_paragraph_style_id(p, style_ids["h1"])
            continue

        # If we're in references section, style all subsequent paragraphs as reference items (if style exists)
        if in_references and style_ids.get("ref_item"):
            set_paragraph_style_id(p, style_ids["ref_item"])
            continue

        # Captions (optional rule-based, else heuristic)
        if style_ids.get("fig_caption"):
            if (fig_cap_det and _match(fig_cap_det, text)) or ((not fig_cap_det) and RX_FIG_CAP.search(text)):
                set_paragraph_style_id(p, style_ids["fig_caption"])
                continue

        if style_ids.get("tab_caption"):
            if (tab_cap_det and _match(tab_cap_det, text)) or ((not tab_cap_det) and RX_TAB_CAP.search(text)):
                set_paragraph_style_id(p, style_ids["tab_caption"])
                continue

        # Headings (rule-based first, else heuristic)
        applied = False
        if matchers:
            # Rule-based levels: choose the first match (levels are sorted desc)
            for level, compiled in matchers:
                if _match(compiled, text):
                    key = f"h{level}"
                    if style_ids.get(key):
                        set_paragraph_style_id(p, style_ids[key])
                        applied = True
                    break

        if not applied:
            # MVP heuristic
            if style_ids.get("h3") and RX_H3.search(text):
                set_paragraph_style_id(p, style_ids["h3"])
            elif style_ids.get("h2") and RX_H2.search(text):
                set_paragraph_style_id(p, style_ids["h2"])
            elif style_ids.get("h1") and RX_H1.search(text):
                set_paragraph_style_id(p, style_ids["h1"])
            else:
                # Default body
                if style_ids.get("body"):
                    set_paragraph_style_id(p, style_ids["body"])


# -------------------------
# Insert source into template
# -------------------------

def find_placeholder_paragraph(doc: Document, placeholder: str = PLACEHOLDER) -> Optional[Paragraph]:
    for p in doc.paragraphs:
        if (p.text or "").strip() == placeholder:
            return p
    return None

def insert_doc_body_after_paragraph(anchor_p: Paragraph, src_doc: Document) -> None:
    """
    Deep-copy src_doc body elements (paragraphs, tables, pictures) after anchor.
    Skip sectPr to avoid section property conflicts.
    """
    anchor_elm = anchor_p._p
    parent = anchor_elm.getparent()
    idx = parent.index(anchor_elm)

    for child in src_doc.element.body.iterchildren():
        if child.tag.endswith("}sectPr"):
            continue
        parent.insert(idx + 1, deepcopy(child))
        idx += 1


# -------------------------
# Endpoints
# -------------------------

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/format")
async def format_docx(
    template: UploadFile = File(...),
    source: UploadFile = File(...),
    rules: str = Form(None),            # Optional JSON rules from LLM
    x_api_key: str = Header(None),      # Optional auth header
):
    # Optional API key auth
    api_key = os.getenv("API_KEY")
    if api_key and x_api_key != api_key:
        raise HTTPException(status_code=401, detail="Invalid API key")

    t_bytes = await template.read()
    s_bytes = await source.read()

    tdoc = Document(io.BytesIO(t_bytes))
    sdoc = Document(io.BytesIO(s_bytes))

    rules_obj = {}
    if rules:
        try:
            rules_obj = json.loads(rules)
        except Exception:
            rules_obj = {}

    # 1) Apply template styles in-place on source (content unchanged)
    apply_styles_in_place(tdoc, sdoc, rules_obj if rules_obj else None)

    # 2) Find placeholder in template (or append)
    anchor = find_placeholder_paragraph(tdoc, PLACEHOLDER)
    if anchor is None:
        # If no placeholder, append an empty paragraph at end as anchor
        anchor = tdoc.add_paragraph("")

    # If placeholder exists as text, clear it
    try:
        if (anchor.text or "").strip() == PLACEHOLDER:
            anchor.text = ""
    except Exception:
        pass

    # 3) Insert the full styled source body (paragraphs/tables/images)
    insert_doc_body_after_paragraph(anchor, sdoc)

    # 4) Return output docx
    out = io.BytesIO()
    tdoc.save(out)
    out.seek(0)

    headers = {"Content-Disposition": 'attachment; filename="output.docx"'}
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )
