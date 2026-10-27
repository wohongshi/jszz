"""组卷：把结构化题目（多来源）拼成一份可直接打印的试卷。

输入是一个题目数组，字段宽松：
  {stem, options?, answer?, analysis?, score?, type?(选择/多选/填空/解答/判断), source?}

输出：Markdown（可再次编辑）+ docx（默认）+ 可选 PPT 讲评。
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

from .docx_render import render_paper
from .pptx_render import build_paper_deck

TYPE_ORDER = {"选择": 1, "单选": 1, "多选": 2, "填空": 3, "判断": 4, "解答": 5, "其他": 6}

DEFAULT_NOTE = {
    "选择": "在每小题给出的四个选项中，只有一项是符合题目要求的。",
    "单选": "在每小题给出的四个选项中，只有一项是符合题目要求的。",
    "多选": "在每小题给出的四个选项中，有多项符合题目要求。全部选对的得满分，选对但不全的得部分分，有选错的得0分。",
    "填空": "把答案填在题中横线上。",
    "判断": "判断下列说法是否正确。",
    "解答": "写出必要的文字说明、方程式和重要的演算步骤。",
    "其他": "",
}

CN_NUM = "一二三四五六七八九十"


def _guess_type(q: dict) -> str:
    t = (q.get("type") or "").strip()
    if t in TYPE_ORDER:
        return "选择" if t == "单选" else t
    if q.get("options"):
        return "多选" if len(q.get("answer") or "") > 1 and re.fullmatch(r"[A-D]{2,4}", (q.get("answer") or "").strip()) else "选择"
    if q.get("answer") and len(str(q["answer"])) <= 6:
        return "填空"
    return "解答"


def _clean_stem(q: dict) -> str:
    stem = str(q.get("stem") or "").strip()
    stem = re.sub(r"^\s*\d{1,3}\s*[.．、]\s*", "", stem)
    return stem


def build_markdown(questions: list[dict], *, title: str, duration: int | str = 90,
                   total_score: int | None = None, school: str = "",
                   with_answers: bool = True, notes: str = "",
                   subject: str = "", grade: str = "") -> str:
    groups: dict[str, list[dict]] = {}
    for q in questions:
        groups.setdefault(_guess_type(q), []).append(q)

    if total_score is None:
        total_score = sum(int(q.get("score") or 0) for q in questions) or 100

    lines: list[str] = [f"# {title}", ""]
    lines.append(f"> 时间：{duration}分钟 | 满分：{total_score}分")
    if subject:
        lines.append(f"> 学科：{subject}")
    if grade:
        lines.append(f"> 年级：{grade}")
    lines += ["", "## 注意事项", ""]
    lines.append(notes or "1. 答卷前，请将姓名、班级、考号填写在答题卡指定位置。\n"
                          "2. 选择题答案用2B铅笔涂黑，非选择题用0.5毫米黑色签字笔作答。\n"
                          "3. 考试结束后，将试卷和答题卡一并交回。")
    lines.append("")

    idx = 0
    for gi, t in enumerate(sorted(groups, key=lambda k: TYPE_ORDER.get(k, 9))):
        qs = groups[t]
        scores = [int(q.get("score") or 0) for q in qs]
        if not any(scores):
            per = 3 if t in ("选择", "单选", "多选") else 10
            scores = [per] * len(qs)
        label = {"选择": "单项选择题", "单选": "单项选择题", "多选": "多项选择题"}.get(t, f"{t}题")
        head = f"## {CN_NUM[gi]}、{label}：本题共{len(qs)}小题"
        if t in ("选择", "单选", "多选"):
            head += f"，每小题{scores[0]}分，共{sum(scores)}分"
        else:
            head += f"，共{sum(scores)}分"
        lines.append(head)
        lines.append("")
        lines.append(DEFAULT_NOTE.get(t, ""))
        lines.append("")
        for q, sc in zip(qs, scores):
            idx += 1
            stem = _clean_stem(q)
            lines.append(f"{idx}. {stem}")
            lines.append("")
            for j, opt in enumerate(q.get("options") or []):
                letter = "ABCD"[j] if j < 4 else "?"
                text = re.sub(r"^\s*[A-DＡ-Ｄ]\s*[.．、]\s*", "", str(opt)).strip()
                lines.append(f"{letter}．{text}")
            lines.append("")
            if q.get("source"):
                lines.append(f"> 来源：{q['source']}")
                lines.append("")
        lines.append("")

    if with_answers:
        lines += ["---", "", "## 参考答案", ""]
        idx = 0
        for t in sorted(groups, key=lambda k: TYPE_ORDER.get(k, 9)):
            for q in groups[t]:
                idx += 1
                ans = str(q.get("answer") or "").strip()
                ana = str(q.get("analysis") or "").strip()
                lines.append(f"{idx}. 【答案】{ans or '略'}")
                if ana:
                    lines.append(f"【解析】{ana}")
                lines.append("")
    return "\n".join(lines)


def make_paper(questions: list[dict], *, title: str, out_dir: str | Path,
               duration: int | str = 90, total_score: int | None = None,
               school: str = "", with_answers: bool = True, notes: str = "",
               subject: str = "", grade: str = "", formats=("md", "docx", "pptx"),
               pptx_mode: str = "review", name: str | None = None,
               theme: str = "ink") -> dict:
    from .docx_render import render_paper as _rp  # 兼容旧名

    md = build_markdown(questions, title=title, duration=duration,
                        total_score=total_score, school=school,
                        with_answers=with_answers, notes=notes,
                        subject=subject, grade=grade)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    base = name or re.sub(r"[\\/:*?\"<>|\s]+", "_", title)[:40]
    written = {}
    if "md" in formats:
        p = out / f"{base}.md"
        p.write_text(md, "utf-8")
        written["md"] = str(p)
    if "docx" in formats:
        written["docx"] = str(render_paper(md, out / f"{base}.docx", school=school,
                                           answer_page=with_answers))
    if "pptx" in formats:
        written["pptx"] = str(build_paper_deck(md, out / f"{base}.pptx", theme=theme,
                                               school=school, mode=pptx_mode))
    return {"ok": True, "markdown": md, "files": written,
            "count": len(questions), "total_score": total_score}
