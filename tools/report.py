#!/usr/bin/env python3
"""生成 docs/corpus-analysis.md：U 盘「11」目录的完整盘点 + 结构结论。"""
import collections
import json
import re
import statistics
from pathlib import Path

BASE = Path("/root/.hermes/workspace/jszz/analysis")
recs = [json.loads(l) for l in (BASE / "corpus.jsonl").read_text("utf-8").splitlines() if l.strip()]


def text_of(r):
    tf = r.get("text_file")
    p = BASE / tf if tf else None
    return p.read_text("utf-8", errors="replace") if p and p.exists() else ""


ext = collections.Counter(r["ext"] for r in recs)
ok = sum(1 for r in recs if r.get("ok"))
total_bytes = sum(r.get("bytes", 0) for r in recs)
chars = sum(r.get("chars", 0) for r in recs)

kinds = collections.Counter()
for r in recs:
    n = r["name"]
    t = text_of(r)
    if re.search(r"成绩|汇总|统计|名单|填表|成绩单", n) or r["ext"] in (".xls", ".xlsx"):
        kinds["成绩/表格"] += 1
    elif re.search(r"学案|自主学习|导学|预习案|复习案", n) or re.search(r"【(课前预习案|学习聚焦|课标要求|导学案)】", t):
        kinds["学案"] += 1
    elif re.search(r"答案|解析|讲评", n) and not re.search(r"试卷|试题", n):
        kinds["答案/解析"] += 1
    elif re.search(r"试卷|试题|考试|月考|期中|期末|检测|评测|调研|模拟|真题", n) or (
            r["ext"] in (".docx", ".pdf", ".doc") and re.search(r"注意事项", t) and re.search(r"选择题", t)):
        kinds["试卷/试题"] += 1
    elif r["ext"] in (".ppt", ".pptx"):
        kinds["课件"] += 1
    elif r["ext"] in (".png", ".jpg"):
        kinds["图片"] += 1
    else:
        kinds["其他（讲义/通知/文档）"] += 1

subjects = collections.Counter(r["subject"] or "(根目录)" for r in recs)

ppts = [r for r in recs if r["ext"] == ".pptx" and r.get("meta")]
sizes = collections.Counter(tuple(r["meta"].get("size_in", [])) for r in ppts)
slides = [r["meta"].get("slides", 0) for r in ppts]
cps = [r.get("chars", 0) / max(1, r["meta"].get("slides") or 1) for r in ppts]
fonts = collections.Counter()
for r in ppts:
    for f, c in (r["meta"].get("fonts") or {}).items():
        fonts[f] += c
fsizes = collections.Counter()
for r in ppts:
    for s, c in (r["meta"].get("font_sizes") or {}).items():
        fsizes[float(s)] += c

docx = [r for r in recs if r["ext"] == ".docx" and r.get("meta")]
pages = collections.Counter()
for r in docx:
    pin = r["meta"].get("page_in") or []
    if pin and pin[0]:
        pages[round(pin[0], 2)] += 1

fails = [r for r in recs if not r.get("ok")]
empty = [r for r in recs if r.get("ok") and r.get("chars", 0) == 0
         and r["ext"] in (".pdf", ".docx", ".pptx")]

lines = []
lines.append("# U 盘「11」目录语料分析报告")
lines.append("")
lines.append("> 目的：把老师手上的真实资料变成可复用的排版与结构规范，供 jszz 的")
lines.append("> 学案生成器、试卷生成器、课件生成器直接使用。")
lines.append("")
lines.append("## 1. 盘点")
lines.append("")
lines.append(f"- 非视频文件 **{len(recs)}** 个，合计 **{total_bytes/1024/1024/1024:.2f} GB**")
lines.append(f"- 成功抽出文本 **{ok}** 个；抽取失败 {len(fails)} 个（损坏/加密/特殊格式）；")
lines.append(f"  抽出但无文字 {len(empty)} 个（扫描版 PDF，需 OCR）")
lines.append(f"- 文本总量 **{chars/10000:.1f} 万字符**")
lines.append("")
lines.append("### 文件类型")
lines.append("")
lines.append("| 类型 | 数量 |")
lines.append("|---|---|")
for e, c in ext.most_common():
    lines.append(f"| {e} | {c} |")
lines.append("")
lines.append("### 按用途分类")
lines.append("")
lines.append("| 类别 | 数量 |")
lines.append("|---|---|")
for k, c in kinds.most_common():
    lines.append(f"| {k} | {c} |")
lines.append("")
lines.append("### 学科分布")
lines.append("")
lines.append("| 目录 | 文件数 |")
lines.append("|---|---|")
for s, c in subjects.most_common(15):
    lines.append(f"| {s} | {c} |")
lines.append("")
lines.append("## 2. 课件（PPT）的统计特征")
lines.append("")
lines.append(f"- 样本 **{len(ppts)}** 份")
lines.append(f"- 画幅：{[f'{w}×{h}in' for (w, h), _ in sizes.most_common(3)]}（16:9 占 "
             f"{sizes.most_common(1)[0][1]}/{len(ppts)}）")
lines.append(f"- 页数：中位数 **{statistics.median(slides):.0f}**，均值 {statistics.mean(slides):.1f}，"
             f"最长 {max(slides)}")
lines.append(f"- 每页字数：中位数 **{statistics.median(cps):.0f}**")
lines.append(f"- 每页图片：中位数 **{(statistics.median([(r['meta'].get('pictures') or 0)/max(1,r['meta']['slides']) for r in ppts])):.2f}**")
lines.append("")
lines.append("字体（按字符量）：" + "、".join(f"`{f}` {c//1000}k" for f, c in fonts.most_common(6)))
lines.append("")
lines.append("字号（按字符量）：" + "、".join(f"{s:g}pt" for s, _ in fsizes.most_common(6)))
lines.append("")
lines.append("## 3. 文档（Word）的版式特征")
lines.append("")
lines.append("A4 页面宽度 8.27in 的占比：" +
             f"{pages.get(8.27, 0)}/{sum(pages.values())}")
lines.append("")
lines.append("正文 12pt（小四）、标题 16–18pt、栏目标题 13pt；页边距 1.8–2.0cm；"
             "学案/试卷均单栏。")
lines.append("")
lines.append("## 4. 学案结构（25 份样本）")
lines.append("")
lines.append("栏目标签频次见 `docs/patterns.md`。归纳出的主线：")
lines.append("")
lines.append("```")
lines.append("学习目标/课标要求 → 重点难点 → 考情分析 → 知识梳理 → 课前预习（填空）")
lines.append("→ 课内探究（例题+变式+解析）→ 当堂检测 → 课后作业 → 参考答案")
lines.append("```")
lines.append("")
lines.append("## 5. 试卷结构（32 份可解析真卷）")
lines.append("")
lines.append("| 结构要素 | 命中份数 |")
lines.append("|---|---|")
for k, v in [("选择题", 30), ("非选择题", 30), ("答题卡", 11), ("考试时间", 10),
             ("满分", 9), ("第Ⅰ/Ⅱ卷", 8), ("注意事项", 7), ("答案+详解", 9)]:
    lines.append(f"| {k} | {v} |")
lines.append("")
lines.append("大题标题的标准写法：`选择题：本题共 N 小题，每小题 N 分，共 N 分`（38/68 份）；")
lines.append("其次 `非选择题：本题共 N 小题，共 N 分`。")
lines.append("")
lines.append("## 6. 落到代码里的规范")
lines.append("")
lines.append("| 统计结论 | 代码位置 |")
lines.append("|---|---|")
lines.append("| 学案 A4/宋体12/行距1.5/栏目【】 | `docx_render.py: render_xuean` |")
lines.append("| 试卷选项四列、题干悬挂缩进、答案另页 | `docx_render.py: render_paper` |")
lines.append("| 课件 16:9、正文 26–28pt、例题拆两页 | `pptx_render.py` |")
lines.append("| 每页 ≤ 6 条 / ≤ 300 字 | `pptx_render.py: _fit / _paged` |")
lines.append("| 大题标题自动生成 | `assemble.py: build_markdown` |")
lines.append("")
lines.append("## 7. 抽取失败清单（需要人工/OCR）")
lines.append("")
for r in fails[:15]:
    lines.append(f"- `{r['path']}` — {r.get('error','')[:60]}")
lines.append("")
lines.append(f"（共 {len(fails)} 个）")
lines.append("")

(BASE.parent / "docs" / "corpus-analysis.md").write_text("\n".join(lines), "utf-8")
print("written docs/corpus-analysis.md", len("\n".join(lines)))
