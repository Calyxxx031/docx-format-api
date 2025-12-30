import io
import json
import re
from copy import deepcopy
from typing import Any, Dict, Optional, Tuple

from fastapi import FastAPI, UploadFile, File, Form, Header, HTTPException
from fastapi.responses import StreamingResponse
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table

# 用于按“文档真实顺序”遍历段落+表格
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


app = FastAPI()

# 可选：简单鉴权（Render / Railway 环境变量里配）
# 不想要鉴权就保持 None
API_KEY = None  # 例如 "xxxx"


# --------------------------
# Helpers
# --------------------------
def iter_block_items(doc: Document):
    """
    Yield Paragraph and Table objects in document order.
    """
    parent_elm = doc.element.body
    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def clear_body_keep_sectpr(doc: Document):
    """
    Remove all paragraphs & tables in body but keep section properties (sectPr).
    This makes the output contain NO template content, while preserving page setup/styles.
    """
    body = doc._element.body
    for child in list(body):
        # remove <w:p> and <w:tbl>, keep <w:sectPr>
        tag = child.tag.lower()
        if tag.endswith("}p") or tag.endswith("}tbl"):
            body.remove(child)


def style_exists(doc: Document, name: str) -> bool:
    try:
        _ = doc.styles[name]
        return True
    except Exception:
        return False


def pick_style(doc: Document, *candidates: Optional[str]) -> Optional[str]:
    """
    Return first existing style name in doc.styles.
    """
    for c in candidates:
        if c and style_exists(doc, c):
            return c
    return None


def parse_rules(rules_str: Optional[str]) -> Dict[str, Any]:
    """
    Accept:
    - plain JSON dict
    - JSON string that is itself a dict
    - Dify LLM output wrapper like {"text": "{...json...}", ...}
    """
    if not rules_str:
        return {}

    # Some clients may send already-json-like strings with leading/trailing spaces
    s = rules_str.strip()
    if not s:
        return {}

    def _loads_maybe(x: str):
        return json.loads(x)

    try:
        obj = _loads_maybe(s)
    except Exception:
        # not JSON, ignore
        return {}

    # If it is a JSON string that itself contains JSON dict
    if isinstance(obj, str):
        try:
            obj2 = _loads_maybe(obj)
            obj = obj2
        except Exception:
            return {}

    # If it is Dify wrapper dict {"text": "..."}
    if isinstance(obj, dict) and isinstance(obj.get("text"), str):
        inner = obj["text"].strip()
        # inner might itself be JSON dict string
        try:
            inner_obj = _loads_maybe(inner)
            if isinstance(inner_obj, dict):
                return inner_obj
        except Exception:
            # Sometimes the inner is a JSON string again
            try:
                inner_obj2 = _loads_maybe(_loads_maybe(inner))
                if isinstance(inner_obj2, dict):
                    return inner_obj2
            except Exception:
                pass
        # fallback to wrapper dict
        return obj

    return obj if isinstance(obj, dict) else {}


def eval_detect(text: str, detect: str) -> bool:
    """
    detect formats:
      - "regex:...."
      - "contains:a|b|c"
    """
    if not detect:
        return False
    d = detect.strip()
    dl = d.lower()
    if dl.startswith("regex:"):
        pat = d[6:].strip()
        try:
            return re.search(pat, text) is not None
        except re.error:
            return False
    if dl.startswith("contains:"):
        parts = [p.strip() for p in d[len("contains:"):].split("|") if p.strip()]
        return any(p in text for p in parts)
    # fallback: treat as contains with "|" support
    parts = [p.strip() for p in d.split("|") if p.strip()]
    return any(p in text for p in parts)


def build_default_rules(template_doc: Document) -> Dict[str, Any]:
    """
    Default MVP rules:
    - heading1: 第X章 / 1. / 2. / 3. ...
    - heading2: 1.1
    - heading3: 1.1.1
    - captions: 图/表
    - special headings: 摘要/关键词/参考文献/致谢/附录
    """
    body_style = pick_style(template_doc, "CMcontent", "正文", "Normal") or "Normal"
    h1 = pick_style(template_doc, "CMheading1", "标题1", "Heading 1") or "Heading 1"
    h2 = pick_style(template_doc, "CMheading2", "标题2", "Heading 2") or "Heading 2"
    h3 = pick_style(template_doc, "CMheading3", "标题3", "Heading 3") or "Heading 3"
    cap_fig = pick_style(template_doc, "CMcaption", "题注", "Caption") or body_style
    cap_tbl = pick_style(template_doc, "CMcaptionTable", "题注", "Caption") or body_style
    reflist = pick_style(template_doc, "CMreflist", "参考文献", "Normal") or body_style
    appendix_h1 = pick_style(template_doc, "附录标题1", "CMheading1", "标题1", "Heading 1") or h1

    return {
        "formatting_intent": {
            "body_style_name": body_style,
            "heading_rules": [
                {"level": 1, "detect": r"regex:^(第[零一二三四五六七八九十\d]+章|\d+\.)\s*", "style_name": h1},
                {"level": 2, "detect": r"regex:^\d+\.\d+\s*", "style_name": h2},
                {"level": 3, "detect": r"regex:^\d+\.\d+\.\d+\s*", "style_name": h3},
            ],
            "special_rules": [
                {"type": "abstract", "detect": "contains:摘要|摘 要|ABSTRACT", "style_name": h1},
                {"type": "keywords", "detect": "contains:关键词|Key words|Keywords", "style_name": body_style},
                {"type": "references", "detect": "contains:参考文献|References", "style_name": h1},
                {"type": "acknowledgement", "detect": "contains:致谢|Acknowledgement", "style_name": h1},
                {"type": "caption_figure", "detect": r"regex:^(图|Figure)\s*\d+", "style_name": cap_fig},
                {"type": "caption_table", "detect": r"regex:^(表|Table)\s*\d+", "style_name": cap_tbl},
                {"type": "appendix", "detect": r"regex:^附录\s*[A-Z]", "style_name": appendix_h1},
            ],
            "reference_list_style_name": reflist,
        }
    }


def classify_paragraph(
    text: str,
    rules: Dict[str, Any],
    in_references: bool
) -> Tuple[str, bool]:
    """
    Return (style_name, in_references_next)
    """
    intent = (rules or {}).get("formatting_intent") or {}
    body_style = intent.get("body_style_name") or "Normal"
    heading_rules = intent.get("heading_rules") or []
    special_rules = intent.get("special_rules") or []
    reflist_style = intent.get("reference_list_style_name") or body_style

    t = text or ""
    ts = t.strip()

    # Blank lines: keep as body style, and if we were in references, we can stop references on first blank
    if ts == "":
        return body_style, False if in_references else False

    # If already in references: apply reference list style when looks like a reference line
    if in_references:
        # common patterns: [1] ...  or  1. ...
        if re.match(r"^\[\d+\]\s*", ts) or re.match(r"^\d+\.\s+", ts):
            return reflist_style, True
        # if a new big section starts, stop references
        if re.match(r"^(附录|致谢|第[零一二三四五六七八九十\d]+章|\d+\.)", ts):
            # fall through to normal detection below
            in_references = False
        else:
            # otherwise still treat as reference paragraph
            return reflist_style, True

    # 1) Special rules first
    for r in special_rules:
        detect = r.get("detect") or ""
        style_name = r.get("style_name") or body_style
        rtype = (r.get("type") or "").lower()

        if eval_detect(ts, detect):
            # entering references section
            if rtype == "references":
                return style_name, True
            return style_name, False

    # 2) Heading rules
    for r in heading_rules:
        detect = r.get("detect") or ""
        style_name = r.get("style_name") or body_style
        if eval_detect(ts, detect):
            return style_name, False

    # 3) Fallback
    return body_style, False


def copy_paragraph_runs(src_p: Paragraph, dst_p: Paragraph):
    """
    Copy runs text exactly, without carrying over manual formatting.
    (Styles control formatting in output.)
    """
    if src_p.runs:
        for r in src_p.runs:
            dst_p.add_run(r.text)
    else:
        # keep exact paragraph text
        dst_p.add_run(src_p.text)


# --------------------------
# API
# --------------------------
@app.get("/health")
def health():
    return {"ok": True}


@app.post("/format")
async def format_docx(
    template: UploadFile = File(...),
    source: UploadFile = File(...),
    rules: Optional[str] = Form(None),
    x_api_key: Optional[str] = Header(None),
):
    if API_KEY is not None and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")

    template_bytes = await template.read()
    source_bytes = await source.read()

    try:
        template_doc = Document(io.BytesIO(template_bytes))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid template docx: {e}")

    try:
        source_doc = Document(io.BytesIO(source_bytes))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid source docx: {e}")

    # Parse rules; if not provided, use defaults
    parsed_rules = parse_rules(rules)
    if not parsed_rules or "formatting_intent" not in parsed_rules:
        parsed_rules = build_default_rules(template_doc)

    # Output doc: start from template, BUT remove all template body content
    out_doc = template_doc
    clear_body_keep_sectpr(out_doc)

    # Rebuild with ONLY source content formatted
    in_refs = False
    for block in iter_block_items(source_doc):
        if isinstance(block, Paragraph):
            src_p: Paragraph = block
            text = src_p.text  # used for classification only
            style_name, in_refs = classify_paragraph(text, parsed_rules, in_refs)

            dst_p = out_doc.add_paragraph()
            # set style if exists; otherwise keep default
            if style_name and style_exists(out_doc, style_name):
                dst_p.style = style_name
            copy_paragraph_runs(src_p, dst_p)

        elif isinstance(block, Table):
            # Deep copy table XML into output (keeps table content)
            out_doc._element.body.append(deepcopy(block._element))

    # If any placeholder accidentally exists in output (e.g. source includes it), remove it
    # (Optional safety)
    for p in list(out_doc.paragraphs):
        if p.text.strip() == "{{CONTENT}}":
            # remove this paragraph element
            p._element.getparent().remove(p._element)

    buf = io.BytesIO()
    out_doc.save(buf)
    buf.seek(0)

    filename = "formatted.docx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )
