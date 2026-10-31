"""写作工艺体检 / 优化：把 12-写作工艺.md 的规律变成可调用检查。

规则来源：全 788 份教师资料语言层取证（tools/writecraft*.py）。
核心：短句原则（要点独立给分）、禁口语、详解逐项排除、口诀、题源标注、考情/素养。
"""
from __future__ import annotations

import re

# 口语/残句（真实阅卷「不得分」写法，禁止出现在答案/详解里）
_COLLOQ = [
    "把渣子", "多加点", "搞", "弄", "差不多", "挺好的", "不好", "好了", "就行了",
    "看一下", "搞掉", "去掉就行", "差不多就行", "随便", "大概",
]

# 短句/采分：主观题答案里一句过长 = 未按采分点拆分
_LONG_SENT = re.compile(r"[^。；！？\n]{40,}[。；！？]")


def _sections(md: str) -> list[tuple[str, str]]:
    """切 ## 栏目 → [(标题, 正文)]。"""
    parts = re.split(r"^##\s+", md, flags=re.M)
    out = []
    for p in parts[1:]:
        lines = p.splitlines()
        title = lines[0].strip()
        body = "\n".join(lines[1:])
        out.append((title, body))
    return out


def _questions(md: str) -> list[dict]:
    """抽题目块：题号行 → 题干/选项/【答案】/【详解】。"""
    qs = []
    for m in re.finditer(r"^\s*(\d{1,2})[.．、]\s*(.+?)(?=^\s*\d{1,2}[.．、]\s|\Z)",
                         md, re.M | re.S):
        block = m.group(2)
        ans = re.search(r"【答案】\s*([^\n]+)", block)
        exp = re.search(r"【(详解|解析)】\s*(.+?)(?=【|$)", block, re.S)
        qs.append({
            "no": m.group(1),
            "stem": block.split("【")[0].strip()[:200],
            "answer": ans.group(1).strip() if ans else "",
            "explain": exp.group(2).strip() if exp else "",
            "has_exp": bool(exp),
        })
    return qs


def check_writing(md: str) -> dict:
    """写作工艺体检：返回 issues/warnings/hints 三级 + 各维度计数。"""
    issues, warnings, hints = [], [], []
    qs = _questions(md)

    # 1. 详解体例：选择题应逐项排除 + 「故选」收口
    for q in qs:
        if q["answer"] and q["has_exp"]:
            exp = q["explain"]
            n_opts = len(re.findall(r"[ABCD][.．：:]", q["stem"]))
            n_elim = len(re.findall(r"(排除|错误|不正确|不符|不选)", exp))
            if n_opts >= 3 and n_elim < n_opts - 1 and "故选" not in exp:
                warnings.append(
                    f"第{q['no']}题：选择题详解未逐项排除（{n_opts}选项仅{n_elim}处判定），"
                    f"建议「X：理由，排除/正确」逐项 + 「故选」收口")
            if "故选" not in exp and q["answer"] in "ABCDEFGH":
                hints.append(f"第{q['no']}题：详解末尾建议加「故选 {q['answer']}。」")
        elif q["answer"] and not q["has_exp"]:
            hints.append(f"第{q['no']}题：有答案无详解，建议补【详解】")

    # 2. 短句原则：主观题答案长句未按采分点拆
    for sec_title, body in _sections(md):
        if re.search(r"答案|详解|解析|采分", sec_title):
            for sent in _LONG_SENT.findall(body):
                if len(sent) > 60:
                    warnings.append(
                        f"「{sec_title}」有一句 {len(sent)} 字未拆采分点（阅卷按要点独立给分），"
                        f"建议按「要点N：主题：展开（X分）」分行")
                    break

    # 3. 禁口语：扫答案/详解区（栏目正文 + 题目详解），题干允许口语情境
    _answer_zones = []
    for sec_title, body in _sections(md):
        if re.search(r"答案|详解|解析|规范|表述", sec_title):
            _answer_zones.append((sec_title, body))
    if not _answer_zones:
        # 无 ## 栏目时，退而扫题目详解/答案本身
        for q in qs:
            if q["explain"] or q["answer"]:
                _answer_zones.append((f"第{q['no']}题详解",
                                      q["explain"] + " " + q["answer"]))
    for kw in _COLLOQ:
        for sec_title, body in _answer_zones:
            if kw in body:
                issues.append(f"「{sec_title}」出现口语「{kw}」，阅卷不得分，改用学科规范表述")
                break

    # 4. 口诀/归纳：好的学案/课件应有归纳或口诀
    if not re.search(r"口诀|巧记|归纳|总结|记忆", md):
        hints.append("未见归纳/口诀，建议加「> 口诀：…」或归纳总结（真实资料 459/788 有口诀）")

    # 5. 题源标注：真题建议标注（学年·地区·考试）
    n_src = len(re.findall(r"[（(]\d{2}-\d{2}[^）)]*[）)]", md))
    n_choice = len(re.findall(r"【答案】", md))
    if n_choice >= 3 and n_src == 0:
        hints.append("多道题无题源标注，真题建议照学科网原文标注「（26-27高二上·河北石家庄·阶段检测）」")

    # 6. 考情/素养（政史地化生复习课标配）
    return {
        "ok": not issues,
        "issues": issues,
        "warnings": warnings[:12],
        "hints": hints[:12],
        "counts": {
            "questions": len(qs),
            "with_explain": sum(1 for q in qs if q["has_exp"]),
            "with_source": n_src,
            "has_mnemonic": bool(re.search(r"口诀|巧记|记忆", md)),
            "has_scoring": bool(re.search(r"采分|要点\d", md)),
            "has_craft": bool(re.search(r"归纳|总结|特别提醒|易错", md)),
        },
    }


def suggest_scoring(answer: str) -> list[str]:
    """把一段主观题答案按短句原则拆成采分点草稿（启发式，供人工核对）。"""
    # 按分号/句号/换行切，去掉过短残句
    segs = [s.strip() for s in re.split(r"[；;。\n]", answer) if len(s.strip()) >= 6]
    out = []
    for i, seg in enumerate(segs, 1):
        # 去掉开头连接词，保证主谓宾
        seg = re.sub(r"^(并且|而且|同时|另外|此外|其中|因此|所以|因为|由于)[，,]?", "", seg)
        out.append(f"要点{i}：{seg}")
    return out
