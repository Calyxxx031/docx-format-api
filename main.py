import io
import os
import re
import json
from copy import deepcopy
from typing import Any, Dict, List, Optional, Tuple, Union

from fastapi import FastAPI, UploadFile, File, Form, Header, HTTPException
from fastapi.responses import StreamingResponse

from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table


app = FastAPI(title="docx-format-api", version="0.1.0")

# 可选：简单鉴权。Render/Railway 配环境变量 API_KEY，然后 Dify HTTP 节点 Header 传 X-API-Key
API_KEY = os.getenv("API_KEY", "").strip() or None


# -----------------------------
# Helpers: blocks iteration
# -----------------------------
def iter_block_items(doc: Document):
    """Yield Paragraph and Table items in document order."""
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def clear_document_body(doc: Document) -> None:
    """Remove all body content except section properties (sectPr)."""
    body = doc.element.body
    sectPr = body.sectPr
    # Remove everything
    for child in list(body):
        body.remove(child)
    # Re-add sectPr if present
    if sectPr is not None:
        body.append(sectPr)


def get_style_names(doc: Document) -> set:
    names = set()
    try:
        for s in doc.styles:
            # s.name is localized name in many templates (e.g., "标题 1"/"正文")
            if getattr(s, "name", None):
                names.add(s.name)
    except Exception:
        pass
    return names


def pick_existing_style(preferred: Optional[str], template_styles: set, fallbacks: List[str]) -> str:
    if preferred and preferred in template_styles:
        return preferred
    for fb in fallbacks:
        if fb in template_styles:
            return fb
    return "Normal"


# -----------------------------
# Rules parsing (方案 A)
# -----------------------------
def parse_rules(raw: Optional[str]) -> Optional[Dict[str, Any]]:
    """
    Accepts:
      - None/"" -> None
      - A JSON object string -> dict
      - A JSON string that itself contains JSON (your current LLM output) -> dict (double loads)
      - A wrapper object like {"text": "..."} -> unwrap then parse
    """
    if not raw:
        return None
    raw = raw.strip()

    # Sometimes callers accidentally pass the whole LLM node output object as JSON
    # e.g. {"text":"{...}", "usage":...}
    try:
        maybe_wrapper = json.loads(raw)
        if isinstance(maybe_wrapper, dict) and "text" in maybe_wrapper and isinstance(maybe_wrapper["text"], str):
            raw = maybe_wrapper["text"].strip()
    except Exception:
        pass

    # First parse attempt
    try:
        obj = json.loads(raw)
    except Exception:
        # Fallback: extract first {...} block
        l = raw.find("{")
        r = raw.rfind("}")
        if l != -1 and r != -1 and r > l:
            obj = json.loads(raw[l:r + 1])
        else:
            raise

    # If first load gives a string (your case), load again
    if isinstance(obj, str):
        obj = json.loads(obj)

    if not isinstance(obj, dict):
        return None
    return obj


# -----------------------------
# Compile detection rules
# -----------------------------
def compile_detect(detect: str) -> Tuple[str, Any]:
    """
    detect formats:
      - "regex:<pattern>"
      - "contains:a|b|c"
    Returns (kind, compiled)
    """
    detect = (detect or "").strip()
    if detect.lower().startswith("regex:"):
        pat = detect[6:].strip()
        return ("regex", re.compile(pat))
    if detect.lower().startswith("contains:"):
        keys = [k.strip() for k in detect[9:].split("|") if k.strip()]
        return ("contains", keys)
    # fallback: treat as contains single keyword
    return ("contains", [detect])


def match_detect(compiled: Tuple[str, Any], text: str) -> bool:
    kind, obj = compiled
    if kind == "regex":
        return bool(obj.search(text))
    # contains
    t = text or ""
    return any(k in t for k in obj)


def normalize_rules(rules_obj: Optional[Dict[str, Any]]) -> Dict[str, Any]:
    """Make rules robust even if missing fields."""
    out = {
        "formatting_intent": {
            "body_style_name": None,
            "heading_rules": [],
            "special_rules": [],
        },
        "notes": "",
    }
    if not rules_obj:
        return out

    fi = rules_obj.get("formatting_intent") or rules_obj.get("formattingIntent") or {}
    out["formatting_intent"]["body_style_name"] = fi.get("body_style_name") or fi.get("bodyStyleName")
    out["formatting_intent"]["heading_rules"] = fi.get("heading_rules") or fi.get("headingRules") or []
    out["formatting_intent"]["special_rules"] = fi.get("special_rules") or fi.get("specialRules") or []
    out["notes"] = rules_obj.get("notes") or ""
    return out


def build_matchers(rules: Dict[str, Any]):
    fi = rules["formatting_intent"]
    heading_rules = fi.get("heading_rules") or []
    special_rules = fi.get("special_rules") or []

    compiled_heading = []
    for r in heading_rules:
        try:
            level = int(r.get("level"))
            detect = r.get("detect") or ""
            style_name = r.get("style_name")
            compiled_heading.append({
                "level": level,
                "detect_raw": detect,
                "detect": compile_detect(detect),
                "style_name": style_name,
            })
        except Exception:
            continue

    # Avoid "1." swallowing "1.1": check deeper levels first
    compiled_heading.sort(key=lambda x: x["level"], reverse=True)

    compiled_special = []
    for r in special_rules:
        try:
            typ = r.get("type")
            detect = r.get("detect") or ""
            style_name = r.get("style_name")
            compiled_special.append({
                "type": typ,
                "detect_raw": detect,
                "detect": compile_detect(detect),
                "style_name": style_name,
            })
        except Exception:
            continue

    return compiled_special, compiled_heading


# -----------------------------
# Apply formatting
# -----------------------------
def decide_style_for_paragraph(
    text: str,
    compiled_special: List[Dict[str, Any]],
    compiled_heading: List[Dict[str, Any]],
    template_styles: set,
    body_style: str,
) -> str:
    t = (text or "").strip()

    # Empty paragraph: keep body
    if not t:
        return body_style

    # Special rules first
    for r in compiled_special:
        if match_detect(r["detect"], t):
            return pick_existing_style(r.get("style_name"), template_styles, [body_style, "Normal"])

    # Heading rules
    for r in compiled_heading:
        if match_detect(r["detect"], t):
            # If the rule provides a style_name, use it; else fallback by level
            provided = r.get("style_name")
            if provided and provided in template_styles:
                return provided
            if r["level"] == 1:
                return pick_existing_style(None, template_styles, ["Heading 1", "标题 1", "标题1", "Title", body_style, "Normal"])
            if r["level"] == 2:
                return pick_existing_style(None, template_styles, ["Heading 2", "标题 2", "标题2", body_style, "Normal"])
            if r["level"] == 3:
                return pick_existing_style(None, template_styles, ["Heading 3", "标题 3", "标题3", body_style, "Normal"])
            return body_style

    return body_style


def copy_table_into_doc(target_doc: Document, table: Table):
    """
    Best-effort table preservation: deep-copy table XML into target doc body.
    """
    tbl = table._tbl
    target_doc.element.body.append(deepcopy(tbl))


def rebuild_doc_with_styles(
    template_doc: Document,
    source_doc: Document,
    rules_obj: Optional[Dict[str, Any]],
) -> Document:
    # Prepare template as output base
    out_doc = template_doc
    template_styles = get_style_names(out_doc)
    clear_document_body(out_doc)

    rules = normalize_rules(rules_obj)
    compiled_special, compiled_heading = build_matchers(rules)

    body_style = pick_existing_style(
        rules["formatting_intent"].get("body_style_name"),
        template_styles,
        ["正文", "Normal", "Body Text"]
    )

    for block in iter_block_items(source_doc):
        if isinstance(block, Paragraph):
            text = block.text or ""

            style_to_apply = decide_style_for_paragraph(
                text=text,
                compiled_special=compiled_special,
                compiled_heading=compiled_heading,
                template_styles=template_styles,
                body_style=body_style,
            )

            # Create paragraph and copy plain text (MVP: keep text identical; inline run formats are not preserved)
            p = out_doc.add_paragraph(text)
            try:
                p.style = style_to_apply
            except Exception:
                # If style missing or invalid, fallback
                try:
                    p.style = body_style
                except Exception:
                    pass

        elif isinstance(block, Table):
            copy_table_into_doc(out_doc, block)

    return out_doc


# -----------------------------
# API
# -----------------------------
@app.get("/health")
def health():
    return {"ok": True}


@app.post("/format")
async def format_docx(
    template: UploadFile = File(...),
    source: UploadFile = File(...),
    rules: Optional[str] = Form(None),
    x_api_key: Optional[str] = Header(None, convert_underscores=False),
):
    # Optional auth
    if API_KEY:
        if not x_api_key or x_api_key.strip() != API_KEY:
            raise HTTPException(status_code=401, detail="Unauthorized")

    # Basic validation
    if not template.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="template must be .docx")
    if not source.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="source must be .docx")

    template_bytes = await template.read()
    source_bytes = await source.read()

    try:
        template_doc = Document(io.BytesIO(template_bytes))
        source_doc = Document(io.BytesIO(source_bytes))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to open docx: {e}")

    # Parse rules (supports double-json)
    rules_obj = None
    if rules:
        try:
            rules_obj = parse_rules(rules)
        except Exception as e:
            # If rules parsing fails, fall back to no-rules MVP (all body)
            rules_obj = None

    # Build output doc
    try:
        out_doc = rebuild_doc_with_styles(template_doc, source_doc, rules_obj)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to format docx: {e}")

    buf = io.BytesIO()
    out_doc.save(buf)
    buf.seek(0)

    headers = {
        "Content-Disposition": 'attachment; filename="result.docx"'
    }
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )
