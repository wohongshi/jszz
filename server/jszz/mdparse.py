"""A small, forgiving parser for teacher-style Markdown used by jszz.

Agents write lesson plans / papers as Markdown; this module turns that into a
document model the docx and pptx renderers can share.

Supported:
  # 一级标题         -> doc title / slide cover
  ## 二级标题        -> section (学案栏目 / 试卷大题)
  ### 三级标题       -> subsection (知识点 / 探究)
  > 元信息 key: value
  - 列表 / 1. 列表
  | a | b |          -> table (first row = header)
  ![alt](path)       -> image
  ---                -> page break (docx) / new section (pptx)
  {{答案句}}         -> 挖空（下划线空格），供学案填空
  【解析】...        -> answer block
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any


META_KEYS = {
    "学科", "年级", "课", "课时", "编写", "审核", "时间", "满分", "日期", "说明",
    "密封线", "命题", "学校", "班级", "姓名", "考号", "来源", "版本", "教材",
    "满分值", "考试时间", "适用", "类型",
}


@dataclass
class Block:
    kind: str                     # text | bullet | num | table | image | answer | note
    text: str = ""
    level: int = 0
    rows: list[list[str]] = field(default_factory=list)
    src: str = ""
    meta: dict[str, Any] = field(default_factory=dict)


@dataclass
class Section:
    title: str
    level: int                    # 2 or 3
    blocks: list[Block] = field(default_factory=list)


@dataclass
class Doc:
    title: str = ""
    meta: dict[str, str] = field(default_factory=dict)
    sections: list[Section] = field(default_factory=list)
    preface: list[Block] = field(default_factory=list)

    def flat_blocks(self):
        yield from self.preface
        for s in self.sections:
            yield from s.blocks


LABEL_SPLIT = re.compile(r"(?=(?:【\s*(?:答案|解析|详解|分析|点睛|点评|方法|归纳|规律|"
                         r"特别提醒|易错提醒|提醒|注意|易错|警示|导语|拓展)\s*】))")


def _kind_of_label(text: str) -> str:
    m = re.match(r"^【\s*([^】]+)\s*】", text.strip())
    if not m:
        return "text"
    lab = m.group(1)
    if lab in ("特别提醒", "易错提醒", "提醒", "注意", "易错", "警示"):
        return "tip"
    if lab in ("点睛", "点评", "方法", "归纳", "规律"):
        return "insight"
    if lab in ("答案", "解析", "详解", "分析"):
        return "answer"
    return "note"


def _resplit(blocks: list[Block]) -> list[Block]:
    """一段里塞了多个【标签】（常见于复制粘贴）时，拆成多个块。"""
    out: list[Block] = []
    for b in blocks:
        if b.kind in ("text", "answer", "tip", "insight") and b.text.count("【") >= 2:
            parts = [x.strip() for x in LABEL_SPLIT.split(b.text) if x.strip()]
            if len(parts) >= 2:
                for part in parts:
                    out.append(Block(_kind_of_label(part), text=part))
                continue
        out.append(b)
    return out


def parse(md: str) -> Doc:
    doc = Doc()
    cur: Section | None = None
    lines = md.replace("\r\n", "\n").split("\n")
    i = 0
    while i < len(lines):
        raw = lines[i]
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped:
            i += 1
            continue

        # headings
        m = re.match(r"^(#{1,4})\s+(.*)$", stripped)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            if level == 1 and not doc.title:
                doc.title = text
            elif level == 2:
                cur = Section(title=text, level=2)
                doc.sections.append(cur)
            elif level >= 3:
                cur = Section(title=text, level=3)
                doc.sections.append(cur)
            i += 1
            continue

        # meta block  > key: value
        if stripped.startswith(">"):
            body = stripped.lstrip("> ").strip()
            mm_scene = re.match(r"^(引文|公式|对照|反应|概念|史料|材料|图表|模板|"
                                r"口诀|规范表述|采分|考情)\s*[:：]\s*(.+)$", body)
            if mm_scene:
                (cur.blocks if cur else doc.preface).append(
                    Block("scene", text=mm_scene.group(2).strip(),
                          meta={"scene": mm_scene.group(1)}))
                i += 1
                continue
            # 一行可写多组元信息：> 学科：物理 | 年级：高一 | 课时：1
            parts = [x.strip() for x in re.split(r"[|｜]", body) if x.strip()]
            pairs = []
            for part in parts:
                mm = re.match(r"^([\u4e00-\u9fa5A-Za-z ]{1,12})[:：]\s*(.*)$", part)
                if mm and mm.group(1).strip() in META_KEYS:
                    pairs.append((mm.group(1).strip(), mm.group(2).strip()))
            if pairs:
                for k, v in pairs:
                    doc.meta[k] = v
            elif body:
                target = cur.blocks if cur else doc.preface
                target.append(Block("note", body))
            i += 1
            continue

        # page break
        if re.match(r"^-{3,}$", stripped):
            if cur is None:
                doc.sections.append(Section(title="", level=2))
                cur = doc.sections[-1]
            cur.blocks.append(Block("break"))
            i += 1
            continue

        # table
        if stripped.startswith("|") and stripped.endswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
                if not re.match(r"^:?-{2,}:?$", cells[0] if cells else ""):
                    rows.append(cells)
                i += 1
            (cur.blocks if cur else doc.preface).append(Block("table", rows=rows))
            continue

        # image
        m = re.match(r"^!\[([^\]]*)\]\(([^)]+)\)", stripped)
        if m:
            (cur.blocks if cur else doc.preface).append(
                Block("image", text=m.group(1), src=m.group(2)))
            i += 1
            continue

        # 提醒 / 点睛 / 解析 —— 三种语义不同的块
        if re.match(r"^[【\[]\s*(特别提醒|易错提醒|提醒|注意|易错|警示)\s*[】\]]", stripped):
            (cur.blocks if cur else doc.preface).append(Block("tip", stripped))
            i += 1
            continue
        if re.match(r"^[【\[]\s*(点睛|点评|方法|归纳|规律)\s*[】\]]", stripped) or \
                re.match(r"^(点睛|点评)\s*[:：]", stripped):
            (cur.blocks if cur else doc.preface).append(Block("insight", stripped))
            i += 1
            continue
        if re.match(r"^[【\[]\s*(答案|解析|详解|分析)\s*[】\]]", stripped) or \
                re.match(r"^(答案|解析|详解|分析)\s*[:：]", stripped):
            (cur.blocks if cur else doc.preface).append(Block("answer", stripped))
            i += 1
            continue

        # bullets
        m = re.match(r"^([-*+])\s+(.*)$", stripped)
        if m:
            indent = (len(raw) - len(raw.lstrip())) // 2
            (cur.blocks if cur else doc.preface).append(
                Block("bullet", text=m.group(2).strip(), level=min(indent, 2)))
            i += 1
            continue

        m = re.match(r"^(\d{1,2})[.、)]\s*(.*)$", stripped)
        if m:
            (cur.blocks if cur else doc.preface).append(
                Block("num", text=f"{m.group(1)}. {m.group(2).strip()}",
                      meta={"num": int(m.group(1))}))
            i += 1
            continue

        # paragraph (join wrapped lines until blank / block start)
        buf = [stripped]
        i += 1
        while i < len(lines):
            nxt = lines[i].strip()
            if (not nxt or nxt.startswith(("#", ">", "|", "![", "- ", "* ", "---"))
                    or re.match(r"^\d{1,2}[.、)]\s", nxt)):
                break
            buf.append(nxt)
            i += 1
        text = buf[0]
        for nxt_line in buf[1:]:
            if text and nxt_line:
                if ord(text[-1]) > 127 and ord(nxt_line[0]) > 127:
                    text += nxt_line
                else:
                    text += " " + nxt_line
            else:
                text += nxt_line
        target = cur.blocks if cur else doc.preface
        # 紧跟【解析】之后的续段仍算解析（老师常把 (1)(2) 分成多段）
        kind = "text"
        if target and target[-1].kind == "answer" and not re.match(r"^[【\[]", text):
            kind = "answer"
        target.append(Block(kind, text=text))

    doc.preface = _resplit(doc.preface)
    for sec in doc.sections:
        sec.blocks = _resplit(sec.blocks)
    return doc


def blanks(text: str) -> list[tuple[str, str]]:
    """Split a 学案 paragraph into ('text', s) / ('blank', 答案原文) / ('strong', 词) runs.

    blank 的 val 是挖空处的答案原文（真实答案版学案把答案红粗下划线填回）。
    strong 的 val 是 `**关键词**`（真实资料：关键词红粗，如得分点/条目名）。
    """
    out = []
    for part in re.split(r"(\{\{[^}]*\}\}|\*\*[^*]+\*\*)", text):
        if not part:
            continue
        if part.startswith("{{") and part.endswith("}}"):
            out.append(("blank", part[2:-2].strip()))
        elif part.startswith("**") and part.endswith("**"):
            out.append(("strong", part[2:-2].strip()))
        else:
            out.append(("text", part))
    return out


def count_chars(text: str) -> int:
    return len(re.sub(r"\s", "", text))


_NUM_PREFIX = re.compile(r"^(?:[一二三四五六七八九十]+[、．.]|（[一二三四五六七八九十\d]+）"
                         r"|[（(][一二三四五六七八九十\d]+[）)]|\d+[、．.])\s*")


def sec_key(title: str) -> str:
    """栏目标题归一：去【】括号、去「一、/(1)/1.」序号，供栏目词表匹配。

    真实资料的栏目写法五花八门（【学习目标】/一、试题分析/（一）评分标准），
    词表匹配一律走归一后的键，别再让括号和序号把栏目拦在门外。
    """
    t = (title or "").strip()
    t = re.sub(r"^[【\[]\s*", "", t)
    t = re.sub(r"\s*[】\]]$", "", t)
    t = re.sub(r"\s*[（(][^）)]*[）)]\s*$", "", t)   # 尾注「试题分析（命题立意）」→ 试题分析
    t = _NUM_PREFIX.sub("", t)
    return t.strip()
