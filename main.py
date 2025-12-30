import io
import re
from copy import deepcopy
from typing import Optional, Tuple

from fastapi import FastAPI, UploadFile, File, Header, HTTPException
from fastapi.responses import StreamingResponse

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl


app = FastAPI(title="docx-format-api", version="0.2.0")

# 可选：鉴权。需要就填一个字符串，并在 Dify HTTP 节点 Header 里传 X-API-Key
API_KEY = None


# ---------------------------
# Helpers: iterate blocks
# ---------------------------
def iter_block_items(doc: Document):
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, doc)
        elif isinstance(child, CT_Tbl):
            yield Table(child, doc)


def clear_body_keep_sectpr(doc: Document):
    """清空模板正文内容（不输出任何模板文字），保留页面/分节设置。"""
    body = doc._element.body
    for child in list(body):
        tag = child.tag.lower()
        if tag.endswith("}p") or tag.endswith("}tbl"):
            body.remove(child)


# ---------------------------
# Style creation (even if template is blank)
# ---------------------------
def style_exists(doc: Document, name: str) -> bool:
    try:
        _ = doc.styles[name]
        return True
    except Exception:
        return False


def ensure_eastasia_font(style, font_name: str):
    st = style._element
    rPr = st.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        st.append(rPr)

    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.append(rFonts)

    rFonts.set(qn("w:eastAsia"), font_name)
    rFonts.set(qn("w:ascii"), font_name)
    rFonts.set(qn("w:hAnsi"), font_name)
    style.font.name = font_name


def ensure_paragraph_style(doc: Document, name: str):
    if style_exists(doc, name):
        return doc.styles[name]
    return doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)


def ensure_required_styles(doc: Document):
    """
    确保我们需要的样式都存在；模板没有也会创建。
    这套参数就是你要的“通用论文排版 MVP”：
    - 正文：宋体 小四 1.5 倍行距 首行缩进2字
    - 标题1/2/3：黑体 三号/四号/小四，加粗（1居中）
    - 题注：宋体 五号 居中
    - 参考文献：宋体 五号 悬挂缩进 0.74cm
    """
    # 正文
    s = ensure_paragraph_style(doc, "CMcontent")
    ensure_eastasia_font(s, "宋体")
    s.font.size = Pt(12)
    pf = s.paragraph_format
    pf.first_line_indent = Cm(0.74)
    pf.line_spacing = 1.5
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # 一级标题
    s = ensure_paragraph_style(doc, "CMheading1")
    ensure_eastasia_font(s, "黑体")
    s.font.size = Pt(16)
    s.font.bold = True
    pf = s.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.2
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 二级标题
    s = ensure_paragraph_style(doc, "CMheading2")
    ensure_eastasia_font(s, "黑体")
    s.font.size = Pt(14)
    s.font.bold = True
    pf = s.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.2
    pf.space_before = Pt(6)
    pf.space_after = Pt(3)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 三级标题
    s = ensure_paragraph_style(doc, "CMheading3")
    ensure_eastasia_font(s, "黑体")
    s.font.size = Pt(12)
    s.font.bold = True
    pf = s.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.2
    pf.space_before = Pt(3)
    pf.space_after = Pt(3)
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # 图题注
    s = ensure_paragraph_style(doc, "CMcaption")
    ensure_eastasia_font(s, "宋体")
    s.font.size = Pt(10.5)
    pf = s.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.0
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 表题注
    s = ensure_paragraph_style(doc, "CMcaptionTable")
    ensure_eastasia_font(s, "宋体")
    s.font.size = Pt(10.5)
    pf = s.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.0
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 参考文献条目
    s = ensure_paragraph_style(doc, "CMreflist")
    ensure_eastasia_font(s, "宋体")
    s.font.size = Pt(10.5)
    pf = s.paragraph_format
    pf.line_spacing = 1.15
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.left_indent = Cm(0.74)
    pf.first_line_indent = Cm(-0.74)  # 悬挂缩进


# ---------------------------
# Classification rules (no LLM)
# ---------------------------
RE_H3 = re.compile(r"^\s*\d+\.\d+\.\d+\s+\S")
RE_H2 = re.compile(r"^\s*\d+\.\d+\s+\S")
RE_H1_NUM = re.compile(r"^\s*\d+\.\s+\S")
RE_H1_CHAPTER = re.compile(r"^\s*第\s*[0-9一二三四五六七八九十]+\s*章\b")

RE_CAP_FIG = re.compile(r"^\s*(图|Figure)\s*\d+")
RE_CAP_TBL = re.compile(r"^\s*(表|Table)\s*\d+")

RE_REF_ITEM = re.compile(r"^\s*(\[\d+\]|\d+\.)\s+")

SECTION_KEYWORDS_AS_H1 = {"绪论", "引言", "相关工作", "实验", "讨论", "结论", "致谢", "参考文献", "附录"}


def is_short_heading_like(s: str) -> bool:
    s2 = s.strip()
    if len(s2) == 0:
        return False
    # 太长一般不是“纯标题”
    if len(s2) > 25:
        return False
    # 含明显句号/分号/逗号/冒号通常不是纯标题（可按需放松）
    if any(ch in s2 for ch in ["。", "；", "，", "：", ":", ".", "、"]):
        return False
    return True


def classify_paragraph(text: str, in_refs: bool) -> Tuple[str, bool]:
    """
    返回 (style_name, in_refs_next)
    """
    t = (text or "").strip()
    if t == "":
        return "CMcontent", False

    # 参考文献段落区：遇到“参考文献”后，条目套 CMreflist，直到碰到新章节/空行
    if in_refs:
        if RE_REF_ITEM.match(t):
            return "CMreflist", True
        # 碰到新章节/附录/致谢等，退出 references
        if RE_H1_CHAPTER.match(t) or RE_H1_NUM.match(t) or t in SECTION_KEYWORDS_AS_H1 or t.startswith("附录"):
            in_refs = False
        else:
            # 仍当作参考文献条目（容错）
            return "CMreflist", True

    # 特殊块：参考文献/致谢/附录/摘要/关键词等（标题行短）
    if ("参考文献" in t) and is_short_heading_like(t):
        return "CMheading1", True  # 进入 references 区
    if ("致谢" in t or "Acknowledgement" in t) and is_short_heading_like(t):
        return "CMheading1", False
    if (t.startswith("附录") and is_short_heading_like(t)):
        return "CMheading1", False

    # 摘要/关键词标题（短）
    if (("摘要" in t or "摘 要" in t or t.upper() == "ABSTRACT") and is_short_heading_like(t)):
        return "CMheading1", False
    if (("关键词" in t or "关键字" in t or "Keywords" in t or "Key words" in t) and is_short_heading_like(t)):
        return "CMheading1", False

    # 题注
    if RE_CAP_FIG.match(t):
        return "CMcaption", False
    if RE_CAP_TBL.match(t):
        return "CMcaptionTable", False

    # 标题层级（先 3 再 2 再 1，避免 1. 吃掉 1.1）
    if RE_H3.match(t):
        return "CMheading3", False
    if RE_H2.match(t):
        return "CMheading2", False
    if RE_H1_CHAPTER.match(t) or RE_H1_NUM.match(t):
        return "CMheading1", False

    # 无编号但很像标题（比如“讨论”“结论”）
    if t in SECTION_KEYWORDS_AS_H1:
        return "CMheading1", False

    return "CMcontent", False


def copy_paragraph_text(src: Paragraph, dst: Paragraph):
    """只复制文本，不带原手工格式（格式由样式统一控制）。"""
    if src.runs:
        for run in src.runs:
            dst.add_run(run.text)
    else:
        dst.add_run(src.text)


# ---------------------------
# API
# ---------------------------
@app.get("/health")
def health():
    return {"ok": True}


@app.post("/format")
async def format_docx(
    template: UploadFile = File(...),
    source: UploadFile = File(...),
    x_api_key: Optional[str] = Header(None),
):
    if API_KEY is not None and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")

    if not template.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="template must be .docx")
    if not source.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="source must be .docx")

    template_bytes = await template.read()
    source_bytes = await source.read()

    try:
        out_doc = Document(io.BytesIO(template_bytes))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid template docx: {e}")

    try:
        src_doc = Document(io.BytesIO(source_bytes))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Invalid source docx: {e}")

    # 确保样式一定存在（模板没有也会创建）
    ensure_required_styles(out_doc)

    # 清空模板正文，只输出源内容
    clear_body_keep_sectpr(out_doc)

    # 重建：按固定规则给 source 套样式
    in_refs = False
    for block in iter_block_items(src_doc):
        if isinstance(block, Paragraph):
            style_name, in_refs = classify_paragraph(block.text, in_refs)
            p = out_doc.add_paragraph()
            if style_exists(out_doc, style_name):
                p.style = style_name
            copy_paragraph_text(block, p)

        elif isinstance(block, Table):
            # 表格直接拷贝（保留内容）
            out_doc._element.body.append(deepcopy(block._element))

    # 防止源文档里意外带了占位符
    for p in list(out_doc.paragraphs):
        if p.text.strip() == "{{CONTENT}}":
            p._element.getparent().remove(p._element)

    buf = io.BytesIO()
    out_doc.save(buf)
    buf.seek(0)

    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="formatted.docx"'},
    )
