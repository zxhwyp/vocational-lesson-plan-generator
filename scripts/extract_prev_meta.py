#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
上一次教案的字段提取器（高职教案生成技能专用）

目的只有一个：把「授课班级 / 所用教材 / 授课地点」这类每次课几乎不变的字段
拿出来，交给提问环节当选项。用户不必每次重新打字，教案里的这些字段也不会
因为重打一遍而前后不一致（同一门课同一批班级，教案交上去班级名写法不一样，
教研检查时很扎眼）。

来源优先级：
    1. 技能 state/recent_outputs.json —— 权威记忆，且**只取最新一条**（即上一次
       生成的教案）。生成教案时 build_lesson_plan.py 会把用户确认过的
       course/project/task/hours/textbook/class/location/header 原样记进去。
       这里直接读它，**不依赖那份 docx 是否还在磁盘上，也不往前翻更早的教案**。
    2. --from 显式指定一份教案 docx（仅当用户主动点明要读哪份时用）。

**不做**的事：不扫描磁盘、不去猜哪些 docx 是历史教案。「上一次」由生成流程
自己记账，靠翻文件去反推，既不可靠（文件可能被移走/删除），又会误读更早的、
别的课的教案。

输出 JSON：
    {"found": true, "source": "...", "from_state": true|false,
     "course": "...", "project": "...", "task": "...", "hours": "...",
     "textbook": "...", "class": "...", "location": "...",
     "header": "分院名", "header_full": "整行页眉"}
找不到时输出 {"found": false, "reason": "...", "searched": [...]}。
空字符串一律归一成 null——「没有候选值」和「候选值是空」在提问时是同一个意思，
都是不要给这个选项。
"""
import argparse
import json
import os
import re
from datetime import datetime

from docx import Document

# 「标签 → 输出字段」。键就是教案基本信息表里的原字，不做同义词猜测：
# 猜错的代价是把「学时」的值填到「所用教材」里，比不填更糟。
LABEL_KEYS = {
    "项目名称": "project",
    "任务名称": "task",
    "学时": "hours",
    "所用教材": "textbook",
    "授课班级": "class",
    "授课地点": "location",
}

# 至少要认出的标签数。低于这个数就当成不是教案（普通文档、课程表之类），
# 免得把不相干的 docx 里的某个单元格当成班级名。
MIN_LABEL_HITS = 3

# 模版自带的占位值。命中组合即判为模版本体而非生成出来的教案。
# 注意：这些字是写死的，若院校更换了 assets/教案模版.docx 的占位文案，
# 需要同步改这里，否则模版本体会被误当成「上一次的教案」混进来。
TEMPLATE_PLACEHOLDER = {"class": "智能班级", "project_prefix": "项目1探寻物联网世界"}

STATE_REL = os.path.join("..", "state", "recent_outputs.json")


def norm(value):
    """空白归一：空串 / 纯空白 / None 都返回 None。"""
    if value is None:
        return None
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text or None


def row_texts(row):
    """取一行里各单元格的文本，按底层 tc 去重（合并单元格会重复返回同一格）。"""
    texts, last_tc = [], None
    for cell in row.cells:
        tc = cell._tc
        if tc is last_tc:
            continue
        last_tc = tc
        texts.append(cell.text.strip())
    return texts


def parse_lesson_plan(path):
    """把一份教案 docx 解析成字段字典。不像教案就返回 None。"""
    try:
        doc = Document(path)
    except Exception:
        return None
    if not doc.tables:
        return None

    info = {}          # label -> 原始值（未归一化）
    hits = 0
    table = doc.tables[0]
    for row in table.rows:
        texts = row_texts(row)
        for i, text in enumerate(texts):
            key = LABEL_KEYS.get(text)
            if not key:
                continue
            hits += 1
            if i + 1 < len(texts):
                info.setdefault(key, texts[i + 1])
    if hits < MIN_LABEL_HITS:
        return None

    # 标题行：「《课程名》课程教案」
    course = None
    for row in table.rows[:3]:
        for text in row_texts(row):
            m = re.search(r"《(.+?)》\s*课程教案", text)
            if m:
                course = m.group(1)
                break
        if course:
            break

    # 页眉右侧是分院名，用来填 meta.header，省得每次问
    header_full = None
    try:
        for section in doc.sections:
            for hdr in (section.header, section.first_page_header,
                        section.even_page_header):
                for para in hdr.paragraphs:
                    text = norm(para.text)
                    if text:
                        header_full = text
                        break
                if header_full:
                    break
            if header_full:
                break
    except Exception:
        header_full = None
    header = header_full.split()[-1] if header_full else None

    result = {
        "course": norm(course),
        "header": header,
        "header_full": header_full,
    }
    for key in LABEL_KEYS.values():
        result[key] = norm(info.get(key))

    # 模版本体不是「上一次生成的教案」，它的占位值给出去只会误导
    if (result["class"] == TEMPLATE_PLACEHOLDER["class"]
            and (result["project"] or "").startswith(TEMPLATE_PLACEHOLDER["project_prefix"])
            and not result["task"]):
        return None

    return result


# ---------------------------------------------------------------- state 记忆

def load_state(state_path):
    """读 state 记忆，返回按时间从新到旧的记录列表。损坏/缺失返回 []。"""
    if not os.path.exists(state_path):
        return []
    try:
        with open(state_path, encoding="utf-8") as f:
            records = json.load(f)
    except Exception:
        return []
    if not isinstance(records, list):
        return []
    return [r for r in records if isinstance(r, dict)]


def state_to_meta(rec):
    """把一条 state 记录转成和 parse_lesson_plan 同构的字段字典。

    state 记录里 course/project/task/hours/textbook/class/location/header
    都是生成教案时用户确认过的原值，直接当权威候选。header 没有
    header_full，取 header 本身即可。
    """
    def pick(*keys):
        for k in keys:
            v = rec.get(k)
            if norm(v):
                return norm(v)
        return None

    return {
        "course": pick("course"),
        "project": pick("project"),
        "task": pick("task"),
        "hours": pick("hours"),
        "textbook": pick("textbook"),
        "class": pick("class"),
        "location": pick("location"),
        "header": pick("header"),
        "header_full": pick("header"),
        "source": rec.get("path"),
        "source_mtime": rec.get("generated_at"),
        "from_state": True,
    }


def find_in_state(state_path):
    """只取 state 里最新一条记录，绝不往前翻更早的教案。

    「上一次生成的教案」指的是一次、且仅那一次。往前翻会把更早的
    班级/教材/地点捞出来当候选，用户要的是「沿用上一份」，不是
    「从历史里挑一份最近的」——两者不是一回事。

    返回 (meta, reason)。meta 为 None 表示没有「上一次」可参照，
    reason 说明为什么（state 未初始化 / 最新一条无可用字段）。
    """
    records = load_state(state_path)
    if not records:
        return None, "state 尚未初始化"
    latest = records[0]
    meta = state_to_meta(latest)
    # 只看这最新一条。核心字段（班级/教材/课程/项目）一个都没有，说明
    # 上一次用户就没填这些，此时同样不算「有上一次」——不给候选值，
    # 而不是退而求其次去找更早的记录。
    if any(meta.get(k) for k in ("class", "textbook", "course", "project")):
        return meta, None
    return None, "上一次生成的教案里这些字段都为空"


# ---------------------------------------------------------------- 主入口

def extract(explicit=None, want_all=False):
    """取「上一次生成的教案」的可复用字段。

    唯一来源是技能自己的 state 记忆（最新一条），外加 --from 显式指定。
    **不扫描、不读取磁盘上任何历史 docx**——「上一次」由生成流程自己记账，
    比去翻文件可靠，也不会误读更早的、别的课的教案。

    返回 (meta_list, reason)：
      - meta_list 非空：拿到了候选；
      - meta_list 为空：reason 说明为什么没拿到，调用方据此决定怎么提示用户。
    """
    here = os.path.dirname(os.path.abspath(__file__))
    state_path = os.path.normpath(os.path.join(here, STATE_REL))

    # 1. --from 显式指定：用户点明要读哪份，只认这一份。
    if explicit:
        path = os.path.abspath(os.path.expanduser(explicit))
        if not os.path.exists(path):
            return [], f"指定的文件不存在：{path}"
        meta = parse_lesson_plan(path)
        if not meta:
            return [], f"指定的文件不是教案：{path}"
        meta["source"] = path
        meta["source_mtime"] = _mtime_str(path)
        meta["from_state"] = False
        return [meta], None

    # 2. state 记忆：只取最新一条（即上一次），取不到就如实报，不回退、不扫描。
    if want_all:
        metas = [state_to_meta(r) for r in load_state(state_path)]
        return metas, (None if metas else "state 尚未初始化")
    meta, state_reason = find_in_state(state_path)
    if meta:
        return [meta], None
    return [], state_reason


def _mtime_str(path):
    try:
        return datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M")
    except OSError:
        return None


def main():
    ap = argparse.ArgumentParser(description="提取上一次生成的教案里的可复用字段")
    ap.add_argument("-o", "--out", help="结果 JSON 输出路径（缺省打印到 stdout）")
    ap.add_argument("--from", dest="explicit", help="指定要读取的教案 .docx")
    ap.add_argument("--all", action="store_true",
                    help="返回全部记忆记录（默认只取最新一条）")
    args = ap.parse_args()

    metas, reason = extract(args.explicit, args.all)
    if metas:
        payload = metas[0] if not args.all else {"found": True, "items": metas}
        if not args.all:
            payload = dict(payload)
            payload["found"] = True
    else:
        payload = {
            "found": False,
            "reason": reason,
        }

    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text + "\n")
        print(f"已写出 {args.out}")
        if not metas:
            print(f"（{reason}，提问时按无候选值处理）")
    else:
        print(text)


if __name__ == "__main__":
    main()
