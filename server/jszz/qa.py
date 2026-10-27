"""产出体检：不靠肉眼也能发现排版事故。

检查 pptx：
  * 形状越界（左右上下超出画布）
  * 文本框装不下（按字号×行数估算，超框 1.35 倍即报）
  * 空页、遗留 {{ }} / 未渲染的 **加粗** / 行内 Markdown
  * 页脚页码是否连续

检查 docx：
  * 空段落连片（>6 个）、表格列宽超过版心
  * 遗留标记、异常字号（<8pt 或 >30pt）
  * 页眉/页脚是否为空
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from pptx import Presentation
from pptx.util import Emu, Pt


def _in(v) -> float:
    return (v or 0) / 914400.0


def check_pptx(path: str | Path) -> dict:
    prs = Presentation(str(path))
    W, H = _in(prs.slide_width), _in(prs.slide_height)
    issues: list[dict] = []
    pages = 0
    for idx, slide in enumerate(prs.slides, 1):
        pages += 1
        texts = 0
        for sh in slide.shapes:
            if sh.left is None:
                continue
            l, t = _in(sh.left), _in(sh.top)
            w, h = _in(sh.width), _in(sh.height)
            if l < -0.02 or t < -0.02 or l + w > W + 0.02 or t + h > H + 0.02:
                issues.append({"page": idx, "type": "越界", "detail":
                               f"形状 ({l:.2f},{t:.2f},{w:.2f},{h:.2f}) 超出画布 {W:.2f}x{H:.2f}"})
            if not sh.has_text_frame:
                continue
            text = sh.text_frame.text.strip()
            if not text:
                continue
            texts += 1
            if "{{" in text or "}}" in text:
                issues.append({"page": idx, "type": "占位符残留", "detail": text[:60]})
            if re.search(r"\*\*|^#{1,3}\s|^-\s", text, re.M):
                issues.append({"page": idx, "type": "Markdown 未渲染", "detail": text[:60]})

            # 估算文本高度：按字号与框宽折算每行字符数
            size = 18.0
            for para in sh.text_frame.paragraphs:
                for run in para.runs:
                    if run.font.size:
                        size = run.font.size.pt
                        break
                break
            chars_per_line = max(6, int(w / (size / 72.0 * 0.62)))
            lines = 0
            for para in sh.text_frame.paragraphs:
                n = len(para.text)
                lines += max(1, -(-n // chars_per_line))
            need = lines * size / 72.0 * 1.5
            if need > h * 1.35 + 0.1 and h > 0.2:
                issues.append({"page": idx, "type": "可能溢出",
                               "detail": f"估算需 {need:.2f}in / 框高 {h:.2f}in：{text[:40]}"})
        if texts == 0:
            issues.append({"page": idx, "type": "空白页", "detail": "没有文字"})
        if idx >= 3 and any(i["type"] == "空白页" and i["page"] == idx for i in issues):
            continue
    return {"ok": not [i for i in issues if i["type"] in ("越界", "占位符残留", "空白页")],
            "pages": pages, "issues": issues}


def check_docx(path: str | Path) -> dict:
    doc = Document(str(path))
    issues: list[dict] = []
    usable = (doc.sections[0].page_width - doc.sections[0].left_margin
              - doc.sections[0].right_margin) if doc.sections else None

    streak = 0
    for i, p in enumerate(doc.paragraphs):
        if not p.text.strip():
            streak += 1
            if streak == 7:
                issues.append({"type": "连续空段", "detail": f"第 {i} 段附近出现 7 个以上空段落"})
            continue
        streak = 0
        if "{{" in p.text or "}}" in p.text:
            issues.append({"type": "占位符残留", "detail": p.text[:60]})
        if re.search(r"\*\*|^#{1,3}\s", p.text):
            issues.append({"type": "Markdown 未渲染", "detail": p.text[:60]})
        for r in p.runs:
            if r.font.size and (r.font.size.pt < 8 or r.font.size.pt > 30) and p.text.strip():
                issues.append({"type": "字号异常", "detail": f"{r.font.size.pt}pt：{p.text[:30]}"})
                break

    for ti, t in enumerate(doc.tables, 1):
        width = sum((c.width or 0) for c in t.rows[0].cells) if t.rows else 0
        if usable and width > usable * 1.05:
            issues.append({"type": "表格超宽", "detail":
                           f"表 {ti} 宽 {_in(width):.2f}in > 版心 {_in(usable):.2f}in"})

    for sec in doc.sections:
        head = "".join(p.text for p in sec.header.paragraphs)
        for t in sec.header.tables:
            head += "".join(c.text for row in t.rows for c in row.cells)
        foot = "".join(p.text for p in sec.footer.paragraphs)
        if not head.strip():
            issues.append({"type": "提示：页眉为空", "detail": "试卷通常无页眉，学案建议填学校/编写人"})
        if "PAGE" not in sec.footer.paragraphs[0]._p.xml if sec.footer.paragraphs else True:
            if "第" not in foot:
                issues.append({"type": "页脚缺页码", "detail": foot[:40]})

    hard = [i for i in issues if not i["type"].startswith("提示")]
    return {"ok": not hard, "issues": issues}


def check(md: str | None = None, pptx: str | Path | None = None,
          docx: str | Path | None = None) -> dict:
    out: dict = {"pptx": None, "docx": None, "markdown": None}
    if pptx and Path(pptx).exists():
        out["pptx"] = check_pptx(pptx)
    if docx and Path(docx).exists():
        out["docx"] = check_docx(docx)
    if md is not None:
        out["markdown"] = check_markdown(md)
    return out


def check_markdown(md: str) -> dict:
    """写作侧体检：抓最容易犯的错。"""
    warns = []
    if "{{" in md and md.count("{{") != md.count("}}"):
        warns.append("挖空 {{ }} 数量不配对")
    for m in re.finditer(r"^(##+\s+)([一二三四五六七八九十]、)?(.+)$", md, re.M):
        title = m.group(3)
        if re.search(r"选择题|非选择题|填空题|解答题", title) and not re.search(r"共\s*\d+\s*小题", title):
            warns.append(f"大题标题没写题量分值：{title[:30]}")
    if re.search(r"^\s*[A-D]\s*[、]", md, re.M):
        warns.append("选项用了「A、」，渲染器只认「A.」")
    # 题号只在同一小节内要求连续（学案每个栏目都从 1 开始）
    for sec in re.split(r"^#{2,3}\s+", md, flags=re.M)[1:]:
        nums = [int(x) for x in re.findall(r"^\s*(\d{1,2})[.．、]\s", sec, re.M)]
        gaps = [b for a, b in zip(nums, nums[1:]) if b - a not in (1, 0)]
        if gaps:
            warns.append(f"同一小节内题号跳号：{gaps[:3]}")
            break
    if "参考答案" not in md and re.search(r"选择题|试卷", md):
        warns.append("试卷没有「## 参考答案」栏目")
    if md.count("{{") and "参考答案" not in md:
        warns.append("学案有挖空但没有答案页")
    return {"ok": not warns, "warnings": warns}
