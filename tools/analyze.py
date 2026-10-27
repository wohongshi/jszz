#!/usr/bin/env python3
"""Corpus mining: what do the teacher's 学案 / 试卷 / 课件 actually look like?

Produces analysis/patterns.json (machine readable) and analysis/patterns.md (report).
"""
import collections
import json
import re
from pathlib import Path

BASE = Path("/root/.hermes/workspace/jszz/analysis")
recs = [json.loads(l) for l in (BASE / "corpus.jsonl").read_text("utf-8").splitlines() if l.strip()]


def text_of(r):
    tf = r.get("text_file")
    if not tf:
        return ""
    p = BASE / tf
    return p.read_text("utf-8", errors="replace") if p.exists() else ""


docs = [(r, text_of(r)) for r in recs]
docs = [(r, t) for r, t in docs if t]


def name_kind(r):
    n = r["name"]
    if re.search(r"学案|自主学习|导学|预习案|复习案", n):
        return "学案"
    if re.search(r"试卷|试题|考试|月考|期中|期末|检测|评测|调研|模拟|真题|单元", n):
        return "试卷"
    if re.search(r"课件|讲评|复习|一轮|二轮|专题", n):
        return "课件"
    return r["kind"]


buckets = collections.defaultdict(list)
for r, t in docs:
    buckets[name_kind(r)].append((r, t))

# ---------------- 学案 ----------------
xuean = buckets["学案"]
bracket = collections.Counter()
seq_samples = []
for r, t in xuean:
    labels = re.findall(r"【([^】]{2,14})】", t)
    bracket.update(labels)
    if len(labels) >= 4 and len(seq_samples) < 8:
        seen, order = set(), []
        for l in labels:
            if l not in seen:
                seen.add(l)
                order.append(l)
        seq_samples.append({"path": r["path"], "order": order[:12]})

# ---------------- 试卷 ----------------
papers = [(r, t) for r, t in buckets["试卷"]
          if r["ext"] in (".docx", ".doc", ".pdf")
          and re.search(r"满分|考试时间|答题卡|选择题|本试卷", t)]
paper_heads = collections.Counter()
paper_markers = collections.Counter()
score_patterns = collections.Counter()
for r, t in papers:
    for line in t.split("\n"):
        s = line.strip()
        if re.match(r"^[一二三四五六七八九十]{1,2}\s*[、.]\s*\S", s) and len(s) < 60 and "题" in s:
            paper_heads[re.sub(r"^\s*[一二三四五六七八九十]{1,2}\s*[、.]\s*", "", s)[:24]] += 1
    for m in re.findall(r"本(?:大题|题|部分)?共\s*\d+\s*小题[^。\n]{0,40}", t):
        score_patterns[re.sub(r"\d+", "N", m)[:48]] += 1
    for k in ("注意事项", "【答案】", "【详解】", "【解析】", "单选题", "多选题",
              "非选择题", "第Ⅰ卷", "第Ⅱ卷", "答题卡", "考试时间", "满分", "选择题"):
        if k in t:
            paper_markers[k] += 1

# ---------------- 课件 ----------------
decks = buckets["课件"]
slide_heads = collections.Counter()
deck_stats = {"slides": [], "chars_per_slide": [], "pics_per_slide": []}
for r, t in decks:
    meta = r.get("meta") or {}
    n = meta.get("slides") or 0
    if n:
        deck_stats["slides"].append(n)
        deck_stats["chars_per_slide"].append(r.get("chars", 0) / n)
        deck_stats["pics_per_slide"].append((meta.get("pictures") or 0) / n)
    for s in (r.get("slides") or [])[:60]:
        h = re.sub(r"\s+", "", s.get("head") or "")
        if 2 <= len(h) <= 18 and not re.search(r"[，。；：？！]$", h):
            slide_heads[h] += 1

# ---------------- 讲义/答案 ----------------
ans = buckets["答案讲评"]
ans_markers = collections.Counter()
for r, t in ans:
    for k in ("【答案】", "【解析】", "【详解】", "【分析】", "【点睛】", "答案", "解析"):
        ans_markers[k] += t.count(k)


def median(xs):
    xs = sorted(xs)
    return xs[len(xs) // 2] if xs else None


report = {
    "counts": {k: len(v) for k, v in buckets.items()},
    "xuean": {
        "docs": len(xuean),
        "labels": bracket.most_common(24),
        "sample_orders": seq_samples[:6],
    },
    "paper": {
        "docs": len(papers),
        "heads": paper_heads.most_common(20),
        "markers": paper_markers.most_common(),
        "score_patterns": score_patterns.most_common(8),
    },
    "deck": {
        "docs": len(decks),
        "slides_median": median(deck_stats["slides"]),
        "chars_per_slide_median": round(median(deck_stats["chars_per_slide"]) or 0, 1),
        "pics_per_slide_median": round(median(deck_stats["pics_per_slide"]) or 0, 2),
        "heads": slide_heads.most_common(20),
    },
    "answer": {"docs": len(ans), "markers": ans_markers.most_common()},
}
(BASE / "patterns.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), "utf-8")

lines = ["# 素材库结构分析（自动统计）", ""]
lines.append(f"素材总量：{len(docs)} 个可读文件；分类：{report['counts']}")
lines.append("")
lines.append("## 学案")
lines.append(f"样本 {len(xuean)} 份。最常见的栏目标签（出现次数）：")
for k, v in report["xuean"]["labels"]:
    lines.append(f"- 【{k}】 {v}")
lines.append("")
lines.append("典型栏目顺序（按首次出现）：")
for s in report["xuean"]["sample_orders"]:
    lines.append(f"- {s['path'].split('/')[-1][:48]} → " + " → ".join(s["order"]))
lines.append("")
lines.append("## 试卷")
lines.append(f"样本 {len(papers)} 份。大题标题 TOP：")
for k, v in report["paper"]["heads"]:
    lines.append(f"- {k}：{v}")
lines.append("")
lines.append("题型/分值表述模板：")
for k, v in report["paper"]["score_patterns"]:
    lines.append(f"- {k}（{v}）")
lines.append("")
lines.append("结构要素命中：")
lines.append(", ".join(f"{k}={v}" for k, v in report["paper"]["markers"]))
lines.append("")
lines.append("## 课件")
lines.append(f"样本 {len(decks)} 份；页数中位数 {report['deck']['slides_median']}；"
             f"每页字数中位数 {report['deck']['chars_per_slide_median']}；"
             f"每页图片数中位数 {report['deck']['pics_per_slide_median']}")
lines.append("")
lines.append("页面标题 TOP：")
for k, v in report["deck"]["heads"]:
    lines.append(f"- {k}：{v}")
lines.append("")
lines.append("## 答案/讲评")
lines.append(f"样本 {len(ans)} 份；标记频次：" + ", ".join(f"{k}={v}" for k, v in report["answer"]["markers"]))
(BASE / "patterns.md").write_text("\n".join(lines), "utf-8")
print("\n".join(lines[:60]))
print("...\nwritten", BASE / "patterns.md")
