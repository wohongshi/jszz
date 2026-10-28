#!/usr/bin/env python3
"""按学科盘点 U 盘「11」里的学案与课件，抽栏目骨架与页面序列。

输出 analysis/subjects.json（机器用）+ analysis/subjects.txt（人读）。
"""
import json
import re
from collections import Counter, defaultdict
from pathlib import Path

BASE = Path("/root/.hermes/workspace/jszz/analysis")
recs = [json.loads(l) for l in (BASE / "corpus.jsonl").read_text("utf-8").splitlines() if l.strip()]

SUBJECTS = ["语文", "数学", "英语", "物理", "化学", "生物"]
ALIAS = {
    "语文": ["语文", "百合花", "烛之武", "论语", "作文", "文言", "小说", "散文", "诗歌", "整本书", "红楼梦", "晨诵"],
    "数学": ["数学", "导数", "圆锥曲线", "数列", "立体几何", "概率", "三角函数", "向量", "不等式", "复数"],
    "英语": ["英语", "English", "话题", "晨诵", "晚写", "语法", "听力", "阅读", "完形", "写作", "U1", "U2", "U3", "U4", "U5", "U6"],
    "物理": ["物理", "力学", "电磁", "动量", "机械波", "光学", "热学", "原子", "牛顿", "匀变速", "电路", "磁场"],
    "化学": ["化学", "反应热", "电化学", "有机", "金属", "晶体", "原电池", "摩尔", "离子", "元素周期"],
    "生物": ["生物", "细胞", "遗传", "基因", "光合", "呼吸", "稳态", "发酵", "免疫", "种群"],
}


def subject_of(rec) -> str:
    path = rec["path"]
    name = rec["name"]
    head = path.split("/")[0]
    # 目录直接判定
    if head.startswith("w物理") or "物理" in head:
        return "物理"
    if head in ("语文", "数学", "化学", "生物", "英语"):
        return head
    if head.startswith("04 音效") or "英语" in path:
        return "英语"
    # 混合目录（gs1 / gszbjs1 / 根目录）按文件名 + 学科词判定
    text = name
    if re.search(r"物理|匀变速|动量|机械波|热学|电磁|光学|牛顿|抛体|圆周", text):
        return "物理"
    if re.search(r"化学|反应热|电化学|晶体|原电池|有机物|元素周期|离子", text):
        return "化学"
    if re.search(r"生物|细胞|遗传|基因|光合|呼吸|稳态|发酵|免疫|种群|蛋白", text):
        return "生物"
    if re.search(r"英语|English|晨诵|晚写|话题|语法|听力|U\d|unit", text, re.I):
        return "英语"
    if re.search(r"数学|导数|曲线|数列|几何|概率|三角|向量|复数", text):
        return "数学"
    if re.search(r"语文|百合花|烛之武|文言|小说|散文|诗歌|作文|红楼梦|史记|廉颇|苏武", text):
        return "语文"
    if head in ("历史", "政治", "地理"):
        return head
    return "其他"


def kind_of(rec) -> str:
    name = rec["name"]
    if re.search(r"学案|导学案|自主学习|预习案|复习案|课前预习", name):
        return "学案"
    if re.search(r"课件|上课|复习|一轮|二轮|专题|讲评|第\d+讲|课时|早读", name):
        return "课件"
    if re.search(r"答案|解析", name):
        return "答案"
    return "其他"


def text_of(rec):
    tf = rec.get("text_file")
    p = BASE / tf if tf else None
    return p.read_text("utf-8", errors="replace") if p and p.exists() else ""


out = defaultdict(lambda: {"学案": [], "课件": [], "答案": [], "其他": []})
for r in recs:
    if r["ext"] in (".png", ".jpg", ".rar", ".mp3", ".mp4"):
        continue
    sub = subject_of(r)
    out[sub][kind_of(r)].append(r)

lines = ["# 六科学案 / 课件盘点", ""]
data = {}
for sub in SUBJECTS:
    bucket = out.get(sub, {"学案": [], "课件": [], "答案": [], "其他": []})
    lines.append(f"## {sub} — 学案 {len(bucket['学案'])} / 课件 {len(bucket['课件'])} / "
                 f"答案 {len(bucket['答案'])} / 其他 {len(bucket['其他'])}")
    entry = {"counts": {k: len(v) for k, v in bucket.items()}, "学案": [], "课件": []}

    for kind in ("学案", "课件"):
        files = sorted(bucket[kind], key=lambda r: -(r.get("chars") or 0))
        for r in files[:40]:
            text = text_of(r)
            meta = r.get("meta") or {}
            if kind == "课件":
                slides = meta.get("slides") or 0
                heads = [s.get("head", "") for s in (r.get("slides") or [])]
                entry[kind].append({
                    "name": r["name"], "path": r["path"], "slides": slides,
                    "heads": [h for h in heads if h][:40],
                })
            else:
                labels = re.findall(r"【([^】]{2,14})】", text)
                seen, order = set(), []
                for l in labels:
                    if l not in seen:
                        seen.add(l)
                        order.append(l)
                entry[kind].append({
                    "name": r["name"], "path": r["path"], "chars": r.get("chars", 0),
                    "labels": order[:20],
                    "pages": (meta.get("pages") if isinstance(meta, dict) else None),
                })
    data[sub] = entry
    lines.append("")
    lines.append("### 学案")
    for it in entry["学案"][:14]:
        lines.append(f"- {it['name'][:60]}（{it['chars']} 字）栏目：{' → '.join(it['labels'][:10])}")
    lines.append("")
    lines.append("### 课件")
    for it in entry["课件"][:14]:
        lines.append(f"- {it['name'][:60]}（{it['slides']} 页）页序：{' / '.join(it['heads'][:9])}")
    lines.append("")

(BASE / "subjects.json").write_text(json.dumps(data, ensure_ascii=False, indent=1), "utf-8")
(BASE / "subjects.txt").write_text("\n".join(lines), "utf-8")
print("\n".join(lines[:80]))
print("\n... written analysis/subjects.txt")
