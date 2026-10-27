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
            mm = re.match(r"^([\u4e00-\u9fa5A-Za-z ]{1,12})[:：]\s*(.+)$", body)
            if mm:
                doc.meta[mm.group(1).strip()] = mm.group(2).strip()
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

        # answer marker
        if re.match(r"^[【\[]\s*(答案|解析|详解|分析|点睛|点评)\s*[】\]]", stripped) or \
                re.match(r"^(答案|解析|详解|分析|点睛|点评)\s*[:：]", stripped):
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
        (cur.blocks if cur else doc.preface).append(Block("text", text=text))
    return doc


def blanks(text: str) -> list[str]:
    """Split a 学案 paragraph into ('text', s) / ('blank', width) runs from {{...}}."""
    out = []
    for part in re.split(r"(\{\{[^}]*\}\})", text):
        if not part:
            continue
        if part.startswith("{{") and part.endswith("}}"):
            inner = part[2:-2].strip()
            out.append(("blank", max(len(inner), 2)))
        else:
            out.append(("text", part))
    return out


def count_chars(text: str) -> int:
    return len(re.sub(r"\s", "", text))
