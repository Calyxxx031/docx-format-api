import io, json, re
from copy import deepcopy
from fastapi import FastAPI, UploadFile, File, Form, Header, HTTPException
from fastapi.responses import StreamingResponse
from docx import Document

app = FastAPI()

# 可选：设置环境变量 API_KEY 后开启鉴权（Render/Railway 都能配）
API_KEY = None  # 你也可以写死，但不建议
PLACEHOLDER = "{{CONTENT}}"

def apply_rules_to_source(sdoc: Document, rules: dict):
    """
    简化版：根据 LLM 的 rules 给段落打样式（只处理段落；表格/图片保持原样）
    rules.formatting_intent.heading_rules: [{level, detect, style_name}]
    rules.formatting_intent.body_style_name
    """
    intent = (rules or {}).get("formatting_intent") or {}
    heading_rules = intent.get("heading_rules") or []
    body_style = intent.get("body_style_name") or None

    compiled = []
    for r in heading_rules:
        detect = (r.get("detect") or "").strip()
        style = (r.get("style_name") or "").strip()
        if not detect or not style:
            continue
        # 约定：detect 以 "regex:" 开头则按正则，否则按包含匹配
        if detect.lower().startswith("regex:"):
            pat = detect[6:].strip()
            try:
                compiled.append(("regex", re.compile(pat), style))
            except re.error:
                continue
        else:
            # 允许用 | 分隔多个关键词
            keys = [k.strip() for k in detect.split("|") if k.strip()]
            compiled.append(("contains", keys, style))

    for p in sdoc.paragraphs:
        text = (p.text or "").strip()
        if not text:
            continue

        matched = False
        for kind, cond, style in compiled:
            if kind == "regex" and cond.search(text):
                try:
                    p.style = style
                except Exception:
                    pass
                matched = True
                break
            if kind == "contains" and any(k in text for k in cond):
                try:
                    p.style = style
                except Exception:
                    pass
                matched = True
                break

        if (not matched) and body_style:
            # 没命中标题规则就套正文样式（可选）
            try:
                p.style = body_style
            except Exception:
                pass

def find_placeholder_paragraph(doc: Document, placeholder: str):
    for p in doc.paragraphs:
        if (p.text or "").strip() == placeholder:
            return p
    return None

def insert_doc_body_after_paragraph(anchor_p, src_doc: Document):
    """
    把 src_doc 的 body 元素（段落/表格/图片等）深拷贝插入到 anchor 段落之后
    """
    anchor_elm = anchor_p._p
    parent = anchor_elm.getparent()
    idx = parent.index(anchor_elm)

    for child in src_doc.element.body.iterchildren():
        # 跳过末尾的 sectPr，避免节属性冲突
        if child.tag.endswith("}sectPr"):
            continue
        parent.insert(idx + 1, deepcopy(child))
        idx += 1

@app.post("/format")
async def format_docx(
    template: UploadFile = File(...),
    source: UploadFile = File(...),
    rules: str = Form(None),
    x_api_key: str = Header(None),
):
    global API_KEY
    if API_KEY is None:
        # 如果你在环境变量里配了 API_KEY 就启用鉴权
        import os
        API_KEY = os.getenv("API_KEY")

    if API_KEY and x_api_key != API_KEY:
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

    # 可选：按规则给 source 段落打样式（方案B 用）
    if rules_obj:
        apply_rules_to_source(sdoc, rules_obj)

    anchor = find_placeholder_paragraph(tdoc, PLACEHOLDER)
    if anchor is None:
        # 找不到占位符就追加到末尾
        anchor = tdoc.add_paragraph(PLACEHOLDER)

    # 清空占位符文本
    anchor.text = ""

    # 插入 source 的内容（尽量保真：段落/表格/图片都复制）
    insert_doc_body_after_paragraph(anchor, sdoc)

    out = io.BytesIO()
    tdoc.save(out)
    out.seek(0)

    headers = {"Content-Disposition": 'attachment; filename="output.docx"'}
    return StreamingResponse(
        out,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )
