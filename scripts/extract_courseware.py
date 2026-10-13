#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
课件内容提取器（高职教案生成技能专用）

把 .pptx / .docx / .pdf 课件的全部文字与结构提取成纯文本，作为生成教案时
唯一的、可核对的内容依据。教案里写的知识点必须能在提取结果里找到出处，
因此不要凭印象跳过这一步。

用法：
    python extract_courseware.py <课件路径> [更多课件路径...] [-o 输出文件]

输出格式：
    ===== 文件：xxx.pptx =====
    ---- 幻灯片 3 ----
    标题：1.1.1 物联网的定义
    · 要点文字
    · 嵌套要点
    表格：表头 | 表头
          单元格 | 单元格

依赖：python-pptx, python-docx, pypdf
"""
import argparse
import os
import sys


def _clean(s):
    if s is None:
        return ""
    return " ".join(str(s).split())


def extract_pptx(path):
    """提取 PPT 每页的标题、正文（按缩进层级）、表格与备注。"""
    from pptx import Presentation

    prs = Presentation(path)
    lines = []
    for idx, slide in enumerate(prs.slides, 1):
        lines.append(f"---- 幻灯片 {idx} ----")
        title = ""
        body = []
        tables = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = _clean(shape.text_frame.text)
                if not txt:
                    continue
                # 优先把占位符标题识别为标题
                if not title and shape.has_text_frame and shape == slide.shapes.title:
                    title = txt
                else:
                    body.append(txt)
            elif getattr(shape, "has_table", False) and shape.has_table:
                rows = []
                for row in shape.table.rows:
                    rows.append(" | ".join(_clean(c.text) for c in row.cells))
                tables.append(rows)
        if title:
            lines.append(f"标题：{title}")
        # 用换行保留层级，方便判断知识点归属
        for b in body:
            for sub in str(b).split("\n"):
                sub = _clean(sub)
                if sub:
                    lines.append(f"· {sub}")
        for rows in tables:
            lines.append("表格：" + ("\n      ".join(rows)))
        notes = ""
        if slide.has_notes_slide:
            notes = _clean(slide.notes_slide.notes_text_frame.text)
        if notes:
            lines.append(f"〔备注〕{notes}")
        lines.append("")
    return "\n".join(lines)


def extract_docx(path):
    """提取 Word 的段落与表格，保持文档顺序。"""
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    d = docx.Document(path)
    lines = []
    body = d.element.body
    for child in body.iterchildren():
        tag = child.tag.split("}")[-1]
        if tag == "p":
            txt = _clean(Paragraph(child, d).text)
            if txt:
                lines.append(txt)
        elif tag == "tbl":
            t = Table(child, d)
            lines.append("表格：")
            for row in t.rows:
                cells = [_clean(c.text) for c in row.cells]
                lines.append("  " + " | ".join(cells))
            lines.append("")
    return "\n".join(lines)


def extract_pdf(path):
    from pypdf import PdfReader

    reader = PdfReader(path)
    lines = []
    for idx, page in enumerate(reader.pages, 1):
        lines.append(f"---- 第 {idx} 页 ----")
        lines.append(page.extract_text() or "")
        lines.append("")
    return "\n".join(lines)


EXTRACTORS = {
    ".pptx": extract_pptx,
    ".ppt": extract_pptx,
    ".docx": extract_docx,
    ".doc": extract_docx,
    ".pdf": extract_pdf,
}


def extract(path):
    ext = os.path.splitext(path)[1].lower()
    fn = EXTRACTORS.get(ext)
    if fn is None:
        raise SystemExit(
            f"不支持的格式：{ext}。当前支持 {', '.join(sorted(EXTRACTORS))}。"
            "遇到 .ppt/.doc 老格式请先用 Office 另存为 .pptx/.docx，"
            "pdf 如需提取图片内文字请先 OCR。"
        )
    return fn(path)


def main():
    ap = argparse.ArgumentParser(description="提取课件文字内容，作为教案生成的内容依据")
    ap.add_argument("paths", nargs="+", help="课件文件路径（可多个）")
    ap.add_argument("-o", "--out", help="输出文件；不指定则打印到屏幕")
    args = ap.parse_args()

    chunks = []
    for p in args.paths:
        if not os.path.exists(p):
            print(f"[跳过] 文件不存在：{p}", file=sys.stderr)
            continue
        chunks.append(f"===== 文件：{os.path.basename(p)} =====")
        try:
            chunks.append(extract(p))
        except Exception as e:  # noqa: BLE001 - 单个文件失败不应中断整批
            chunks.append(f"[提取失败] {type(e).__name__}: {e}")
        chunks.append("")

    text = "\n".join(chunks)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text)
        print(f"已写出 {args.out}（{len(text)} 字符）")
    else:
        print(text)


if __name__ == "__main__":
    main()
