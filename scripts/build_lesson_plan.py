#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
高职教案 Word 生成器（高职教案生成技能专用）

以院校官方教案模版（assets/教案模版.docx）为底版做无损填充：
  · 保留原模版的页面设置、页眉、字体、边框、列宽
  · 填满「基本信息表」，教学过程表的环节行按课时数推算的环节数生成
  · 删除模版里所有红色注释文字与页尾说明
  · 板书设计/教学反思表原样保留（教学反思留空给教师手写）

用法：
    python build_lesson_plan.py 内容.json -o 输出.docx [--template 模版.docx]

内容 JSON 结构见 references/template-spec.md。脚本会做三项自检：
  1. 环节数是否与课时数对应（1→5、2→7、3→8、4→9，更多课时依次递增）
  2. 各环节时长之和是否等于 学时 × 45 分钟
  3. 必填字段是否齐全
自检不通过会打印警告；加 --strict 则直接报错退出。

另有两条不拦截的提醒：四列篇幅上限、课程导入 5-10 分钟。

生成成功后会把输出路径追加到 state/recent_outputs.json，供下一次生成时
用 extract_prev_meta.py 回读「授课班级 / 所用教材 / 授课地点」等字段。
"""
import argparse
import copy
import json
import os
import re
import sys
from datetime import datetime

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

FONT = "宋体"
SZ = "21"          # 半磅值，21 = 10.5pt（五号），与院校参考案例一致
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
MINUTES_PER_HOUR = 45
INTRO_RANGE = (5, 10)   # 第一个环节（课程导入）的时长区间

# 环节数随课时数递增。课时数多说明内容多，需要更多的段落来承载，
# 但对数增长——加得比课时快就会把课切碎，每段不足 20 分钟就不成其为
# 一段完整的教学组织。
STAGE_COUNT_BY_HOURS = {1: 5, 2: 7, 3: 8, 4: 9}
STAGE_COUNT_MAX = 12

# 每次生成都把输出路径记下来，供下一次提问时回读「授课班级 / 所用教材 /
# 授课地点」等变动不大的字段（见 extract_prev_meta.py）。同一门课的教案
# 通常连着做几次，字段靠记忆重复填，很容易前后写法不一致。
STATE_REL = os.path.join("..", "state", "recent_outputs.json")
STATE_KEEP = 20

# 学情分析历史记忆：按「专业 + 课程」作标签分组，每个标签保留最近 6 次，
# 供新教案生成时参考（学情里的学生特点对同一专业同一门课是稳定的）。
ANALYSIS_STATE_REL = os.path.join("..", "state", "analysis_history.json")
ANALYSIS_KEEP = 6


def expected_stage_count(hours):
    """按课时数推算应有的环节数。5 课时及以上按每课时加 1 个外推，上限 12。"""
    if hours in STAGE_COUNT_BY_HOURS:
        return STAGE_COUNT_BY_HOURS[hours]
    return min(STAGE_COUNT_MAX, STAGE_COUNT_BY_HOURS[4] + (hours - 4))

# 四列篇幅上限（字符数）。教案是给评委和教研检查看的，
# 一栏里字越多越看不出教学组织的清晰度，所以宁可短。
COLUMN_BUDGET = {
    "content": 100,   # 教学内容：只写知识点与案例，不抄教材
    "teacher": 55,    # 教师活动：动作序列，不写目的
    "student": 45,    # 学生活动：可观察的行为
    "intent": 30,     # 设计意图：只说为什么这么设计
}
COLUMN_LABEL = {
    "content": "教学内容", "teacher": "教师活动",
    "student": "学生活动", "intent": "设计意图",
}


# ---------------------------------------------------------------- 底层构造

def _make_run(text, sz=SZ, bold=False):
    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:hint"), "eastAsia")
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(attr), FONT)
    rPr.append(fonts)
    if bold:
        rPr.append(OxmlElement("w:b"))
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    rPr.append(color)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), sz)
        rPr.append(el)
    r.append(rPr)
    t = OxmlElement("w:t")
    t.set(XML_SPACE, "preserve")
    t.text = text
    r.append(t)
    return r


def _make_para(text, jc="left", sz=SZ, bold=False):
    """按 Word 规范的 pPr 子元素顺序构造段落。"""
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    spacing = OxmlElement("w:spacing")
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), "240")
    spacing.set(qn("w:lineRule"), "auto")
    pPr.append(spacing)
    j = OxmlElement("w:jc")
    j.set(qn("w:val"), jc)
    pPr.append(j)
    rPr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:hint"), "eastAsia")
    for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(attr), FONT)
    rPr.append(fonts)
    if bold:
        rPr.append(OxmlElement("w:b"))
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "000000")
    rPr.append(color)
    for tag in ("w:sz", "w:szCs"):
        el = OxmlElement(tag)
        el.set(qn("w:val"), sz)
        rPr.append(el)
    pPr.append(rPr)
    p.append(pPr)
    if text:
        p.append(_make_run(text, sz=sz, bold=bold))
    return p


def _as_lines(value):
    """把 str / list / dict 统一成段落文本列表。"""
    if value is None:
        return []
    if isinstance(value, str):
        return [ln.strip() for ln in value.split("\n") if ln.strip()]
    if isinstance(value, (list, tuple)):
        out = []
        for v in value:
            out.extend(_as_lines(v))
        return out
    return [str(value)]


def set_cell(tc, lines, jc="left", sz=SZ, bold=False):
    """清空单元格原有段落（连同红色注释）并写入新内容。保留 tcPr。

    默认按正文格式（五号 10.5pt、不加粗）重写；标题等需要保留模板大字号
    加粗的单元格，由调用方显式传 sz / bold 覆盖。
    """
    for p in list(tc.findall(qn("w:p"))):
        tc.remove(p)
    lines = _as_lines(lines) or [""]
    for line in lines:
        tc.append(_make_para(line, jc=jc, sz=sz, bold=bold))


def set_cell_width(tc, width):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        tc.insert(0, tcPr)
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.insert(0, tcW)
    tcW.set(qn("w:w"), str(width))
    tcW.set(qn("w:type"), "dxa")


def row_tcs(row):
    return row._tr.findall(qn("w:tc"))


def row_cell(row, idx):
    tcs = row_tcs(row)
    if idx >= len(tcs):
        raise IndexError(f"该行只有 {len(tcs)} 个单元格，取不到第 {idx} 个")
    return tcs[idx]


# ---------------------------------------------------------------- 内容组装

def objectives_lines(obj):
    """知识/能力/素质/思政目标，每类独立成段，条目内联编号。"""
    labels = [
        ("knowledge", "知识目标："),
        ("ability", "能力目标："),
        ("quality", "素质目标："),
        ("ideology", "思政目标："),
    ]
    lines = []
    for key, label in labels:
        items = _as_lines(obj.get(key))
        lines.append(label)
        for i, item in enumerate(items, 1):
            if len(items) > 1:
                lines.append(f"{i}. {item}")
            else:
                lines.append(item)
    return lines


def design_lines(design, stages, total):
    """教学设计：课前 / 课中 / 课后。课中未提供时按环节自动生成时间轴。"""
    lines = []
    lines.append("课前：")
    lines.extend(_as_lines(design.get("pre")) or ["—"])
    lines.append("课中：")
    during = _as_lines(design.get("during"))
    if during:
        lines.extend(during)
    else:
        segs = [
            f"{s.get('name', '').replace('环节', '环节').split('：')[-1]}"
            f"（{s.get('minutes', 0)}′）"
            for s in stages
        ]
        lines.append(
            f"{len(stages)}个环节共{total}分钟：" + "→".join(segs) + "。"
        )
    lines.append("课后：")
    lines.extend(_as_lines(design.get("post")) or ["—"])
    return lines


def stage_rows(stages, after, template_tr, widths):
    """按环节数生成教学过程表的行（深拷贝模版空行，保证样式一致）。"""
    rows = []
    for st in stages:
        tr = copy.deepcopy(template_tr)
        for trHeight in tr.findall(qn("w:trPr") + "/" + qn("w:trHeight")):
            trHeight.getparent().remove(trHeight)
        tcs = tr.findall(qn("w:tc"))
        name = str(st.get("name", "")).strip()
        minutes = st.get("minutes", 0)
        set_cell(tcs[0], [name, f"（{minutes}分钟）"], jc="center")
        set_cell(tcs[1], st.get("content"), jc="left")
        set_cell(tcs[2], st.get("teacher"), jc="left")
        set_cell(tcs[3], st.get("student"), jc="left")
        set_cell(tcs[4], st.get("intent"), jc="left")
        for tc, w in zip(tcs, widths):
            set_cell_width(tc, w)
        rows.append(tr)

    tr = copy.deepcopy(template_tr)
    tcs = tr.findall(qn("w:tc"))
    set_cell(tcs[0], ["课后", "（或拓展）"], jc="center")
    set_cell(tcs[1], after.get("content"), jc="left")
    set_cell(tcs[2], after.get("teacher"), jc="left")
    set_cell(tcs[3], after.get("student"), jc="left")
    set_cell(tcs[4], after.get("intent"), jc="left")
    for tc, w in zip(tcs, widths):
        set_cell_width(tc, w)
    rows.append(tr)
    return rows


def set_header(doc, text):
    """页眉右侧分院名替换。header 只给分院名（如「万年产业学院」）时，
    自动保留模版学校名与对齐空格；含空格时视为完整页眉整体替换。"""
    if not text:
        return
    for section in doc.sections:
        for hdr in (section.header, section.first_page_header, section.even_page_header):
            try:
                paras = hdr.paragraphs
            except Exception:
                continue
            if not paras:
                continue
            p = paras[0]
            jc = p.alignment
            runs = p.runs
            if runs:
                if " " in text.strip():
                    new_text = text
                else:
                    stripped = p.text.rstrip()
                    new_text = re.sub(r"[^ ]+$", text, stripped) if stripped else text
                runs[0].text = new_text
                for r in runs[1:]:
                    r.text = ""
            else:
                p.add_run(text)
            p.alignment = jc


# ---------------------------------------------------------------- 主流程

def lint_columns(stages, after):
    """检查四列篇幅与导入时长。超限只提醒不拦截——这些是风格与节奏问题，
    不是格式错误，但需要让写的人看见具体是哪一处。"""
    rows = list(stages) + ([after] if after else [])
    over = []
    for idx, st in enumerate(rows):
        name = st.get("name") or ("课后" if idx == len(rows) - 1 else f"环节{idx + 1}")
        title = str(name).replace("\n", "")
        for key, budget in COLUMN_BUDGET.items():
            n = len("".join(_as_lines(st.get(key))))
            if n > budget:
                over.append(f"{title} 的「{COLUMN_LABEL[key]}」{n} 字，超过建议上限 {budget} 字")
    if over:
        print("篇幅提醒（不影响生成，但建议精简）：", file=sys.stderr)
        for item in over:
            print(f"  - {item}", file=sys.stderr)

    # 课程导入时长：第一个环节过长会挤压主体环节，教案一看就不合理
    if stages:
        lo, hi = INTRO_RANGE
        first = stages[0]
        try:
            mins = int(first.get("minutes", 0) or 0)
        except (TypeError, ValueError):
            mins = 0
        if mins and not (lo <= mins <= hi):
            title = str(first.get("name") or "环节1").replace("\n", "")
            print(
                f"节奏提醒：第一个环节「{title}」为 {mins} 分钟，"
                f"课程导入一般控制在 {lo}-{hi} 分钟。导入拖长会压缩主体环节，"
                f"把多出的时间挪给后面的讲授或实训环节。",
                file=sys.stderr,
            )
    return over


def validate(data, strict):
    meta = data.get("meta", {})
    stages = data.get("stages", [])
    problems = []
    total = sum(int(s.get("minutes", 0) or 0) for s in stages)

    if not _as_lines(meta.get("course")):
        problems.append("缺少课程名称 meta.course")
    if not _as_lines(meta.get("project")):
        problems.append("缺少项目名称 meta.project")
    if not _as_lines(data.get("analysis")):
        problems.append("缺少学情分析 analysis")

    n = len(stages)
    hours = meta.get("hours")
    if hours:
        hours = int(hours)
        expect_n = expected_stage_count(hours)
        if n != expect_n:
            problems.append(
                f"教学环节数为 {n}，{hours} 课时应对应 {expect_n} 个环节"
                f"（1→5、2→7、3→8、4→9，更多课时依次递增）"
            )
        expect_total = hours * MINUTES_PER_HOUR
        if total != expect_total:
            problems.append(
                f"各环节时长合计 {total} 分钟，与 学时({hours})×{MINUTES_PER_HOUR} = {expect_total} 分钟不符"
            )
    else:
        problems.append("缺少学时 meta.hours")

    for i, s in enumerate(stages, 1):
        if not _as_lines(s.get("content")):
            problems.append(f"环节{i} 缺少教学内容")
        if not _as_lines(s.get("intent")):
            problems.append(f"环节{i} 缺少设计意图")

    if problems:
        head = "内容自检未通过：" if strict else "内容自检提醒（不影响生成）："
        print(head, file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        if strict:
            raise SystemExit(1)
    return total


def build(data, template, out_path, strict=False):
    total = validate(data, strict)
    lint_columns(data.get("stages", []), data.get("after", {}) or {})
    meta = data.get("meta", {})

    doc = Document(template)
    body = doc.element.body
    tables = doc.tables
    if len(tables) < 4:
        raise SystemExit(
            f"模版结构异常：期望 4 张顶层表格（基本信息/教学过程/环节/板书），实际 {len(tables)} 张"
        )
    t_info, t_proc, t_stage, t_board = tables[0], tables[1], tables[2], tables[3]

    # ---- 1. 基本信息表 ----
    info = t_info.rows
    hours = meta.get("hours", "")
    set_cell(row_cell(info[0], 0), [f"《{meta.get('course', '')}》课程教案"], jc="center", sz="36", bold=True)
    set_cell(row_cell(info[1], 1), meta.get("project"), jc="left")
    set_cell(row_cell(info[1], 3), meta.get("task"), jc="left")
    set_cell(row_cell(info[2], 1), [str(hours)], jc="left")
    set_cell(row_cell(info[2], 3), meta.get("textbook"), jc="left")
    set_cell(row_cell(info[3], 1), meta.get("class"), jc="left")
    set_cell(row_cell(info[3], 3), meta.get("location", ""), jc="left")
    set_cell(row_cell(info[4], 1), data.get("analysis"), jc="left")
    set_cell(row_cell(info[5], 1), objectives_lines(data.get("objectives", {})), jc="left")
    set_cell(row_cell(info[6], 1), data.get("key_points"), jc="left")
    set_cell(row_cell(info[7], 1), data.get("difficulties"), jc="left")
    set_cell(row_cell(info[8], 1), data.get("resources"), jc="left")
    set_cell(row_cell(info[9], 1), data.get("methods"), jc="left")
    set_cell(row_cell(info[10], 1), data.get("ideology"), jc="left")
    set_cell(
        row_cell(info[11], 1),
        design_lines(data.get("design", {}), data.get("stages", []), total),
        jc="left",
    )

    # ---- 2. 教学过程表：保留标题行+表头行，删掉模版的示例/注释行 ----
    proc_trs = t_proc._tbl.findall(qn("w:tr"))
    for tr in proc_trs[2:]:
        t_proc._tbl.remove(tr)

    grid = t_proc._tbl.find(qn("w:tblGrid"))
    widths = [int(c.get(qn("w:w"))) for c in grid.findall(qn("w:gridCol"))]

    template_tr = t_stage._tbl.findall(qn("w:tr"))[1]   # 「环节3：XXXX」空行，作样式母版
    for tr in stage_rows(
        data.get("stages", []), data.get("after", {}) or {}, template_tr, widths
    ):
        t_proc._tbl.append(tr)

    # ---- 3. 环节表整体删除（内容已并入教学过程表）----
    t_stage._tbl.getparent().remove(t_stage._tbl)

    # ---- 4. 板书设计 / 教学反思表 ----
    board = t_board.rows
    set_cell(row_cell(board[0], 0), ["板书设计："] + _as_lines(data.get("board")), jc="left")
    set_cell(row_cell(board[1], 0), ["教学反思："], jc="left")

    # ---- 5. 清理页面：删掉模版遗留的空段与红色说明段，表格间留一个空段 ----
    sectPr = body.find(qn("w:sectPr"))
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
    for tb in [t_info._tbl, t_proc._tbl, t_board._tbl]:
        body.append(tb)
        body.append(OxmlElement("w:p"))
    if sectPr is not None:
        body.append(sectPr)

    # ---- 6. 页眉 ----
    set_header(doc, meta.get("header"))

    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    doc.save(out_path)
    return out_path, total


def record_output(out_path, data):
    """把本次输出记进 state，供下一次生成时回读字段候选值。写失败不影响出教案。

    state 是「授课班级 / 所用教材 / 授课地点」这类字段的权威记忆来源——
    它们由用户在上一次生成时确认过，直接记下来比下次再去解析 docx 可靠得多，
    也不依赖那份文件是否还留在磁盘上。每个字段都完整记录（含空值，
    空值同样是有意义的信息：表示用户上次就留空了，这次别再拿出来当选项）。
    """
    here = os.path.dirname(os.path.abspath(__file__))
    state_path = os.path.normpath(os.path.join(here, STATE_REL))
    meta = data.get("meta", {})
    try:
        records = []
        if os.path.exists(state_path):
            with open(state_path, encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, list):
                records = loaded
        entry = {
            "path": os.path.abspath(out_path),
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "course": meta.get("course") or None,
            "project": meta.get("project") or None,
            "task": meta.get("task") or None,
            "hours": meta.get("hours") or None,
            "textbook": meta.get("textbook") or None,
            "class": meta.get("class") or None,
            "location": meta.get("location") or None,
            "header": meta.get("header") or None,
        }
        records = [r for r in records
                   if not (isinstance(r, dict) and r.get("path") == entry["path"])]
        records.insert(0, entry)
        os.makedirs(os.path.dirname(state_path), exist_ok=True)
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump(records[:STATE_KEEP], f, ensure_ascii=False, indent=2)
    except Exception as exc:                      # 记录失败不该挡住交付
        print(f"（提示：未记录本次输出路径，{exc}）", file=sys.stderr)


def record_analysis(out_path, data):
    """把本次学情分析记进「专业 + 课程」标签的历史，供下次参考。

    学情里的「学生特点与已有基础」对同一专业同一门课是稳定的，按标签
    积累起来、保留最近 6 次，新教案生成时据此参考，避免每次凭空重写
    （上一次就出过「二年级」这种编造年级的事）。写失败不影响出教案。
    """
    here = os.path.dirname(os.path.abspath(__file__))
    state_path = os.path.normpath(os.path.join(here, ANALYSIS_STATE_REL))
    meta = data.get("meta", {})
    analysis = data.get("analysis") or []
    major = (meta.get("major") or "").strip()
    course = (meta.get("course") or "").strip()
    if not major and not course:
        return                                 # 没有标签，不记
    key = " + ".join(p for p in (major, course) if p)
    if not analysis:
        return                                 # 没有学情内容，不记
    try:
        history = {}
        if os.path.exists(state_path):
            with open(state_path, encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                history = loaded
        entry = {
            "analysis": list(analysis),
            "project": meta.get("project") or None,
            "task": meta.get("task") or None,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        items = history.get(key, [])
        if not isinstance(items, list):
            items = []
        items.insert(0, entry)
        history[key] = items[:ANALYSIS_KEEP]
        os.makedirs(os.path.dirname(state_path), exist_ok=True)
        with open(state_path, "w", encoding="utf-8") as f:
            json.dump(history, f, ensure_ascii=False, indent=2)
    except Exception as exc:                    # 记录失败不该挡住交付
        print(f"（提示：未记录本次学情历史，{exc}）", file=sys.stderr)


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    default_tpl = os.path.join(here, "..", "assets", "教案模版.docx")

    ap = argparse.ArgumentParser(description="按院校模版生成教案 .docx")
    ap.add_argument("content", help="内容 JSON 文件")
    ap.add_argument("-o", "--out", required=True, help="输出 .docx 路径")
    ap.add_argument("--template", default=default_tpl, help="模版 .docx（默认用技能自带模版）")
    ap.add_argument("--strict", action="store_true", help="自检不通过时报错退出")
    args = ap.parse_args()

    with open(args.content, encoding="utf-8") as f:
        data = json.load(f)

    out, total = build(data, args.template, args.out, strict=args.strict)
    record_output(out, data)
    record_analysis(out, data)
    n = len(data.get("stages", []))
    print(f"已生成：{out}")
    print(f"环节数：{n}　总时长：{total} 分钟　学时：{data.get('meta', {}).get('hours')}")


if __name__ == "__main__":
    main()
