#!/usr/bin/env python3
"""从真实文件名里挖各科「特征词」（本学科高频、他学科罕见），供 ss.guess() 兜底猜学科。

用法：python3 tools/mine_subject_terms.py    → analysis/full4/subject-terms.json
"""
import json
import re
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "analysis/full4/subject-terms.json"
SUBJECTS = ["语文", "数学", "英语", "物理", "化学", "生物", "政治", "历史", "地理"]
STOP = {"一轮", "复习", "课件", "讲评", "学案", "检测", "试题", "答案", "高三", "高二", "高一",
        "单元", "期末", "期初", "月考", "专题", "微专题", "编号", "课件1", "副本"}


def main() -> None:
    names: dict[str, list[str]] = {s: [] for s in SUBJECTS}
    for line in (ROOT / "analysis/full4/index.jsonl").open(encoding="utf-8"):
        rel = json.loads(line)["rel"]
        for s in SUBJECTS:
            if f"/{s}/" in rel or rel.startswith(s + "/"):
                names[s].append(Path(rel).name)
    # 候选词：中文 2–4 字
    cand: dict[str, Counter] = {s: Counter() for s in SUBJECTS}
    for s, files in names.items():
        for f in files:
            for seg in re.findall(r"[\u4e00-\u9fa5]{2,6}", f):
                for n in (2, 3, 4):
                    for i in range(len(seg) - n + 1):
                        w = seg[i:i + n]
                        if w not in STOP:
                            cand[s][w] += 1
    out: dict[str, list[list]] = {}
    for s in SUBJECTS:
        others = Counter()
        for o in SUBJECTS:
            if o != s:
                others.update(cand[o])
        scored = []
        for w, c in cand[s].items():
            if c < 3:
                continue
            if others[w] > c * 0.34:      # 他科也常用 → 不算特征
                continue
            scored.append((c / (others[w] + 1), c, w))
        scored.sort(reverse=True)
        out[s] = [[w, c] for _, c, w in scored[:18]]
        print(f"{s}（{len(names[s])} 份）：" + " ".join(w for w, _ in out[s][:14]))
    OUT.write_text(json.dumps({"files": {s: len(v) for s, v in names.items()},
                               "terms": out}, ensure_ascii=False, indent=1), "utf-8")
    print("→", OUT)


if __name__ == "__main__":
    main()
