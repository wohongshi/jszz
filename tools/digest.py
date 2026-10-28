#!/usr/bin/env python3
"""六科内容摘要：每个文件的页面骨架 + 学案正文要点，输出 analysis/digest_<科目>.txt"""
import json
import re
from pathlib import Path

BASE = Path("/root/.hermes/workspace/jszz/analysis")
OUT = BASE / "digest"
OUT.mkdir(exist_ok=True)
recs = [json.loads(l) for l in (BASE / "corpus.jsonl").read_text("utf-8").splitlines() if l.strip()]

SUB_DIRS = {
    "语文": ["语文"], "数学": ["数学"], "生物": ["生物"], "化学": ["化学"],
    "物理": ["w物理", "20.化学反应速率及影响因素"], "英语": ["04 音效"],
}
MIXED = ["gs1", "gszbjs1", "jzszgsx", "ge1"]
KEY = {
    "语文": r"语文|百合花|鸿门宴|烛之武|劝学|师说|六国论|文言|小说|散文|诗歌|作文|早读|早自习|整本书",
    "数学": r"数学|导数|函数|数列|几何|概率|二项式|随机变量|回归|计数|不等式|集合",
    "英语": r"英语|English|晨诵|晚写|话题|语法|听力|完形|七选五|U\d|unit|讲评",
    "物理": r"物理|匀变速|牛顿|动量|机械波|热学|光学|电磁|抛体|圆周|磁场|原子|实验",
    "化学": r"化学|反应热|电化学|晶体|原电池|有机物|元素周期|离子|金属|结构",
    "生物": r"生物|细胞|遗传|基因|光合|呼吸|稳态|发酵|免疫|种群|蛋白|核酸|工程",
}


def pick(rec, sub):
    p = rec["path"]
    if any(p.startswith(d + "/") for d in SUB_DIRS.get(sub, [])):
        return True
    head = p.split("/")[0]
    if head in MIXED and re.search(KEY[sub], rec["name"], re.I):
        # 混合目录按关键词，但要排除别的学科词
        others = [s for s in KEY if s != sub and re.search(KEY[s], rec["name"], re.I)]
        return not others
    return False


def text_of(rec):
    tf = rec.get("text_file")
    p = BASE / tf if tf else None
    return p.read_text("utf-8", errors="replace") if p and p.exists() else ""


for sub in ["语文", "数学", "英语", "物理", "化学", "生物"]:
    files = [r for r in recs if pick(r, sub) and r["ext"] in (".pptx", ".ppt", ".docx", ".doc", ".pdf")]
    decks = [r for r in files if r["ext"] in (".pptx", ".ppt") and (r.get("meta") or {}).get("slides")]
    docs = [r for r in files if r["ext"] not in (".pptx", ".ppt") or not (r.get("meta") or {}).get("slides")]
    lines = [f"# {sub}：课件 {len(decks)} 份 / 文档 {len(docs)} 份", ""]

    lines.append("## 课件（页数 / 页面骨架）")
    for r in sorted(decks, key=lambda r: -(r["meta"].get("slides") or 0))[:26]:
        heads = []
        for s in (r.get("slides") or []):
            h = re.sub(r"\s+", " ", (s.get("head") or "")).strip()
            if h and h != "[TABLE]" and h not in heads:
                heads.append(h)
        lines.append(f"\n### {r['name'][:70]}  （{r['meta']['slides']} 页）")
        lines.append("  " + " / ".join(heads[:22]))

    lines.append("\n## 文档（正文开头 / 栏目）")
    for r in sorted(docs, key=lambda r: -(r.get("chars") or 0))[:16]:
        t = text_of(r)
        labels = []
        for l in re.findall(r"【([^】]{2,14})】", t):
            if l not in labels:
                labels.append(l)
        lines.append(f"\n### {r['name'][:70]}（{r.get('chars', 0)} 字）")
        if labels:
            lines.append("  栏目：" + " → ".join(labels[:14]))
        snippet = re.sub(r"\n{2,}", "\n", t.strip())[:700]
        lines.append("  " + snippet.replace("\n", "\n  "))

    (OUT / f"{sub}.txt").write_text("\n".join(lines), "utf-8")
    print(sub, "课件", len(decks), "文档", len(docs), "->", OUT / f"{sub}.txt")
