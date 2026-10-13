#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
历史学情分析提取器（高职教案生成技能专用）

把「学情分析」按「专业 + 课程」作标签记忆，生成新教案时按当前课的
「专业 + 课程」检索历史学情，作为参考——学情里的「学生特点与已有基础」
对同一专业同一门课是稳定的（同一批学生、同一基础），不该每次凭空重写。

记忆文件：state/analysis_history.json，结构为按「专业 + 课程」标签分组的
数组，每个标签保留最近 6 次（新记录在最前）。

用法：
    extract_prev_analysis.py -o /tmp/prev_analysis.json                # 不带过滤，列出全部标签
    extract_prev_analysis.py --course 物联网工程导论 -o /tmp/x.json    # 按课程过滤
    extract_prev_analysis.py --major 物联网应用技术 --course 物联网工程导论

输出 JSON：
    {"found": true, "key": "专业 + 课程", "course": "...", "major": "...",
     "history": [{"analysis": ["...","...","..."], "project": "...",
                  "task": "...", "generated_at": "..."}, ...]}
    // 或过滤后无匹配：{"found": false, "reason": "..."}

学情不落到「上一次」那条机制的 state/recent_outputs.json 里，因为：
    - recent_outputs.json 只取最新一条、且语义是「沿用上一次的班级/教材/地点」；
    - 学情要的是「同一专业+课程的历史积累」，两者生命周期不同，混在一起会互相干扰。
"""
import argparse
import json
import os

STATE_REL = os.path.join("..", "state", "analysis_history.json")
ANALYSIS_KEEP = 6          # 每个「专业 + 课程」标签最多保留最近几次


def load_history(state_path):
    """读学情历史记忆。损坏/缺失返回空 dict。"""
    if not os.path.exists(state_path):
        return {}
    try:
        with open(state_path, encoding="utf-8") as f:
            loaded = json.load(f)
    except Exception:
        return {}
    if not isinstance(loaded, dict):
        return {}
    # 每个标签下是 list，逐项过滤掉非 dict 的脏数据
    out = {}
    for key, items in loaded.items():
        if isinstance(items, list):
            out[key] = [it for it in items if isinstance(it, dict)]
    return out


def build_key(major, course):
    """「专业 + 课程」作标签。两者都缺则返回 None（无有效标签）。"""
    parts = [p for p in (major, course) if p and str(p).strip()]
    return " + ".join(parts) if parts else None


def extract(major=None, course=None):
    """按「专业 + 课程」检索历史学情。

    返回 (payload, reason)。payload 非空即命中；否则 reason 说明原因。
    - 给了 major/course：只返回该标签的历史（标签不存在则 found:false）。
    - 都没给：返回全部标签的历史。
    """
    here = os.path.dirname(os.path.abspath(__file__))
    state_path = os.path.normpath(os.path.join(here, STATE_REL))
    history = load_history(state_path)

    if major or course:
        key = build_key(major, course)
        if not key:
            return None, "未提供有效的「专业 / 课程」"
        if key not in history or not history[key]:
            return None, f"该「{key}」尚无历史学情"
        payload = {
            "found": True,
            "key": key,
            "major": major or None,
            "course": course or None,
            "history": history[key],
        }
        return payload, None

    # 不过滤：返回全部标签
    if not history:
        return None, "学情历史尚未初始化"
    payload = {
        "found": True,
        "all": True,
        "groups": [
            {"key": k, "history": v} for k, v in history.items() if v
        ],
    }
    return payload, None


def main():
    ap = argparse.ArgumentParser(description="按「专业 + 课程」检索历史学情分析")
    ap.add_argument("-o", "--out", help="结果 JSON 输出路径（缺省打印到 stdout）")
    ap.add_argument("--major", help="专业名")
    ap.add_argument("--course", help="课程名")
    args = ap.parse_args()

    payload, reason = extract(args.major, args.course)
    if payload is None:
        payload = {"found": False, "reason": reason}

    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if args.out:
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(text + "\n")
        print(f"已写出 {args.out}")
    else:
        print(text)


if __name__ == "__main__":
    main()
