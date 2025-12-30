import io
import re
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Pt, Cm

app = FastAPI(title="docx-format-api", version="0.2.0")


# ----------------------------
# Utilities: keep sectPr, clear body
# ----------------------------
def clear_body_keep_sectPr(doc: Document) -> None:
    body = doc._element.body
    for child in list(body):
        if child.tag.endswith("}sectPr"):
            continue
        body.remove(child)


def get_or_create_paragraph_style(doc: Document, name: str, base: str = "Normal"):
    styles = doc.styles
    for s in styles:
        if s.type == WD_STYLE_TYPE.PARAGRAPH and s.name == name:
            return s
    style = styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    try:
        style.base_style = styles[base]
    except Exception:
        pass
    return style


def set_style_font(style, font_name: str, font_size_pt: float, bold: bool):
    font = style.font
    font.name = font_name
    font.size = Pt(font_size_pt)
    font.bold = bold

    # ensure East Asian font works in Word/WPS
    rPr = style._element.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), font_name)


def set_para_format(
    style,
    *,
    line_spacing: float,
    space_before_pt: float,
    space_after_pt: float,
    first_line_indent_chars: int = 0,
    alignment=None,
    hanging_indent_chars: int | None = None,
):
    pf = style.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_before = Pt(space_before_pt)
    pf.space_after = Pt(space_after_pt)

    # rough char→cm mapping for Chinese in 12pt: 2 chars ≈ 0.74cm
    char_cm = 0.37

    if hanging_indent_chars is not None:
        pf.left_indent = Cm(char_cm * hanging_indent_chars)
        pf.first_line_indent = Cm(-char_cm * hanging_indent_chars)
    else:
        pf.left_indent = None
        pf.first_line_indent = Cm(char_cm * first_line_indent_chars) if first_line_indent_chars else None

    if alignment is not None:
        pf.alignment = alignment


def ensure_mvp_styles(doc: Document):
    """
    If template doesn't contain custom styles, create them.
    This makes formatting visibly change even with a "plain" template.
    """
    # Body
    body = get_or_create_paragraph_style(doc, "CMcontent")
    set_style_font(body, "宋体", 12, False)
    set_para_format(
        body,
        line_spacing=1.5,
        space_before_pt=0,
        space_after_pt=0,
        first_line_indent_chars=2,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
    )

    # Heading 1/2/3
    h1 = get_or_create_paragraph_style(doc, "CMheading1")
    set_style_font(h1, "黑体", 16, True)
    set_para_format(
        h1,
        line_spacing=1.5,
        space_before_pt=12,
        space_after_pt=6,
        first_line_indent_chars=0,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    h2 = get_or_create_paragraph_style(doc, "CMheading2")
    set_style_font(h2, "黑体", 14, True)
    set_para_format(
        h2,
        line_spacing=1.5,
        space_before_pt=10,
        space_after_pt=4,
        first_line_indent_chars=0,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
    )

    h3 = get_or_create_paragraph_style(doc, "CMheading3")
    set_style_font(h3, "黑体", 12, True)
    set_para_format(
        h3,
        line_spacing=1.5,
        space_before_pt=8,
        space_after_pt=3,
        first_line_indent_chars=0,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
    )

    # Captions
    cap_fig = get_or_create_paragraph_style(doc, "CMcaption")
    set_style_font(cap_fig, "宋体", 10.5, False)
    set_para_format(
        cap_fig,
        line_spacing=1.2,
        space_before_pt=6,
        space_after_pt=6,
        first_line_indent_chars=0,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    cap_tbl = get_or_create_paragraph_style(doc, "CMcaptionTable")
    set_style_font(cap_tbl, "宋体", 10.5, False)
    set_para_format(
        cap_tbl,
        line_spacing=1.2,
        space_before_pt=6,
        space_after_pt=6,
        first_line_indent_chars=0,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
    )

    # References
    ref_h = get_or_create_paragraph_style(doc, "CMrefHeading")
    set_style_font(ref_h, "黑体", 12, True)
    set_para_format(
        ref_h,
        line_spacing=1.5,
        space_before_pt=12,
        space_after_pt=6,
        first_line_indent_chars=0,
        alignment=WD_ALIGN_PARAGRAPH.LEFT,
    )

    ref_i = get_or_create_paragraph_style(doc, "CMrefItem")
    set_style_font(ref_i, "宋体", 12, False)
    set_para_format(
        ref_i,
        line_spacing=1.5,
        space_before_pt=0,
        space_after_pt=0,
        first_line_indent_chars=0,
        alignment=WD_ALIGN_PARAGRAPH.JUSTIFY,
        hanging_indent_chars=2,  # 悬挂缩进 2 字符
    )

    return {
        "body": "CMcontent",
        "h1": "CMheading1",
        "h2": "CMheading2",
        "h3": "CMheading3",
        "cap_fig": "CMcaption",
        "cap_tbl": "CMcaptionTable",
        "ref_h": "CMrefHeading",
        "ref_i": "CMrefItem",
    }


def classify_paragraph(text: str, in_refs: bool) -> str:
    t = text.strip()
    if not t:
        return "blank"

    # section markers
    if re.search(r"^(参考文献|References)\b", t):
        return "ref_heading"
    if re.search(r"^(致谢|Acknowledg(e)?ment)\b", t):
        return "h1"
    if re.search(r"^(摘\s*要|摘要|ABSTRACT)\b", t):
        return "h1"
    if re.search(r"^(关键词|关键字|Keywords|Key\s*words)\b", t):
        return "h2"

    # references items
    if in_refs and re.match(r"^\[\d+\]", t):
        return "ref_item"

    # captions
    if re.match(r"^(图|Figure)\s*\d+", t):
        return "cap_fig"
    if re.match(r"^(表|Table)\s*\d+", t):
        return "cap_tbl"

    # headings by numbering
    if re.match(r"^(第\s*[0-9一二三四五六七八九十]+\s*章)\b", t):
        return "h1"
    if re.match(r"^\d+\.\d+\.\d+\s", t):
        return "h3"
    if re.match(r"^\d+\.\d+\s", t):
        return "h2"
    if re.match(r"^\d+\.\s", t):
        return "h1"

    # keyword headings (MVP)
    if t in {"绪论", "引言", "相关工作", "实验", "实验与结果", "讨论", "结论", "附录"}:
        return "h1"

    return "body"


def format_docx(template_bytes: bytes, source_bytes: bytes) -> bytes:
    template_doc = Document(io.BytesIO(template_bytes))
    source_doc = Document(io.BytesIO(source_bytes))

    # clear template content but keep page setup + styles
    clear_body_keep_sectPr(template_doc)

    styles = ensure_mvp_styles(template_doc)

    in_refs = False
    for p in source_doc.paragraphs:
        text = p.text
        kind = classify_paragraph(text, in_refs)

        if kind == "blank":
            template_doc.add_paragraph("", style=styles["body"])
            continue

        if kind == "ref_heading":
            in_refs = True
            template_doc.add_paragraph(text.strip(), style=styles["ref_h"])
            continue

        if kind == "ref_item":
            template_doc.add_paragraph(text.strip(), style=styles["ref_i"])
            continue

        if kind == "h1":
            template_doc.add_paragraph(text.strip(), style=styles["h1"])
            continue

        if kind == "h2":
            template_doc.add_paragraph(text.strip(), style=styles["h2"])
            continue

        if kind == "h3":
            template_doc.add_paragraph(text.strip(), style=styles["h3"])
            continue

        if kind == "cap_fig":
            template_doc.add_paragraph(text.strip(), style=styles["cap_fig"])
            continue

        if kind == "cap_tbl":
            template_doc.add_paragraph(text.strip(), style=styles["cap_tbl"])
            continue

        template_doc.add_paragraph(text, style=styles["body"])

    out = io.BytesIO()
    template_doc.save(out)
    return out.getvalue()


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/format")
async def format_endpoint(
    template: UploadFile = File(...),
    source: UploadFile = File(...),
):
    if not template.filename.lower().endswith(".docx") or not source.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Only .docx is supported (not .doc).")

    tpl_bytes = await template.read()
    src_bytes = await source.read()

    try:
        out_bytes = format_docx(tpl_bytes, src_bytes)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Format failed: {e}")

    return StreamingResponse(
        io.BytesIO(out_bytes),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="formatted.docx"'},
    )
