"""16:9 课件渲染：学案 → PPT（新授课）/ 试卷 → PPT（讲评课）。

版式取自本机素材库的统计结果（318 份真实课件）：16:9 占 96%，正文 24–28pt，
标题 32pt+，数字用 Times New Roman，中文用 微软雅黑；平均每页 185 字。
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

from .mdparse import Block, Doc, Section, parse

W, H = Inches(13.333), Inches(7.5)
FONT_CN = "微软雅黑"
FONT_EN = "Times New Roman"

THEMES = {
    "blue": {"main": "0B3D91", "accent": "2F6FED", "light": "EAF1FD", "ink": "1B2430"},
    "teal": {"main": "0E6B5B", "accent": "12A48C", "light": "E6F5F1", "ink": "1B2B29"},
    "violet": {"main": "4B2E83", "accent": "7C5CD3", "light": "F0EBFA", "ink": "241C33"},
    "orange": {"main": "9A3412", "accent": "EA580C", "light": "FDF0E7", "ink": "2B1A12"},
    "ink": {"main": "1F2937", "accent": "4B5563", "light": "F1F3F6", "ink": "111827"},
}


def _rgb(hexstr: str) -> RGBColor:
    h = hexstr.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


@dataclass
class DeckTheme:
    name: str = "blue"

    @property
    def main(self):
        return _rgb(THEMES[self.name]["main"])

    @property
    def accent(self):
        return _rgb(THEMES[self.name]["accent"])

    @property
    def light(self):
        return _rgb(THEMES[self.name]["light"])

    @property
    def ink(self):
        return _rgb(THEMES[self.name]["ink"])


class Deck:
    def __init__(self, theme: str = "blue", footer: str = "", school: str = ""):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.t = DeckTheme(theme if theme in THEMES else "blue")
        self.footer = footer
        self.school = school
        self.n = 0

    # ---------- primitives ----------
    def _blank(self):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.n += 1
        return s

    def _rect(self, slide, x, y, w, h, color, line=None, shape=MSO_SHAPE.RECTANGLE):
        sp = slide.shapes.add_shape(shape, x, y, w, h)
        sp.fill.solid()
        sp.fill.fore_color.rgb = color
        if line is None:
            sp.line.fill.background()
        else:
            sp.line.color.rgb = line
        sp.shadow.inherit = False
        return sp

    def _text(self, slide, x, y, w, h, *, size=24, bold=False, color=None,
              align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, line_spacing=1.25,
              font_cn=FONT_CN, font_en=FONT_EN):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Inches(0.05)
        tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = align
        p.line_spacing = line_spacing
        return tb, tf, p

    @staticmethod
    def _run(p, text, *, size, bold=False, color=None, font_cn=FONT_CN, font_en=FONT_EN):
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.name = font_en
        if color is not None:
            r.font.color.rgb = color
        rpr = r._r.get_or_add_rPr()
        ea = rpr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
        if ea is None:
            from lxml import etree
            ea = etree.SubElement(rpr, "{http://schemas.openxmlformats.org/drawingml/2006/main}ea")
        ea.set("typeface", font_cn)
        return r

    def _footer(self, slide, label: str = ""):
        y = H - Inches(0.42)
        tb, tf, p = self._text(slide, Inches(0.6), y, Inches(9), Inches(0.3),
                               size=10, color=_rgb("8A94A6"))
        t = " · ".join(x for x in [self.school, label] if x)
        if t:
            self._run(p, t, size=10, color=_rgb("8A94A6"))
        tb2, tf2, p2 = self._text(slide, W - Inches(1.3), y, Inches(0.8), Inches(0.3),
                                  size=10, color=_rgb("8A94A6"), align=PP_ALIGN.RIGHT)
        self._run(p2, str(self.n), size=10, color=_rgb("8A94A6"))

    def _header(self, slide, title: str, *, kicker: str = "", size=30):
        self._rect(slide, 0, 0, W, Inches(0.12), self.t.main)
        self._rect(slide, Inches(0.6), Inches(0.55), Inches(0.13), Inches(0.62), self.t.accent)
        tb, tf, p = self._text(slide, Inches(0.86), Inches(0.5), W - Inches(1.6), Inches(0.75),
                               size=size, bold=True, color=self.t.main, anchor=MSO_ANCHOR.MIDDLE)
        self._run(p, title, size=size, bold=True, color=self.t.main)
        if kicker:
            if len(kicker) > 12:
                kicker = kicker[:11] + "…"
            chip_w = Inches(0.20 * len(kicker) + 0.34)
            self._rect(slide, W - chip_w - Inches(0.6), Inches(0.58), chip_w, Inches(0.42),
                       self.t.light, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            tb2, tf2, p2 = self._text(slide, W - chip_w - Inches(0.6), Inches(0.6), chip_w,
                                      Inches(0.4), size=12, color=self.t.accent,
                                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p2, kicker, size=12, bold=True, color=self.t.accent)

    # ---------- slides ----------
    def cover(self, title: str, subtitle: str = "", meta: str = "", tag: str = ""):
        s = self._blank()
        self._rect(s, 0, 0, W, H, _rgb("FFFFFF"))
        self._rect(s, 0, 0, Inches(0.42), H, self.t.main)
        self._rect(s, Inches(0.42), 0, Inches(0.1), H, self.t.accent)
        # soft block lower-right
        self._rect(s, W - Inches(3.1), H - Inches(2.1), Inches(3.1), Inches(2.1), self.t.light)

        if tag:
            self._rect(s, Inches(1.15), Inches(1.35), Inches(0.22 * len(tag) + 0.5), Inches(0.5),
                       self.t.accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE)
            tb, tf, p = self._text(s, Inches(1.15), Inches(1.37), Inches(0.22 * len(tag) + 0.5),
                                   Inches(0.46), size=14, bold=True, color=_rgb("FFFFFF"),
                                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p, tag, size=14, bold=True, color=_rgb("FFFFFF"))

        y = Inches(2.05)
        tsize = 40 if len(title) <= 14 else (34 if len(title) <= 20 else (28 if len(title) <= 30 else 24))
        tb, tf, p = self._text(s, Inches(1.1), y, W - Inches(2.6), Inches(2.2),
                               size=tsize, bold=True, color=self.t.main, line_spacing=1.18)
        _rich(p, title, size=tsize, color=self.t.main, accent=self.t.main, bold=True)

        if subtitle:
            tb2, tf2, p2 = self._text(s, Inches(1.15), Inches(4.35), W - Inches(2.8),
                                      Inches(0.9), size=22, color=_rgb("44506B"),
                                      line_spacing=1.3)
            self._run(p2, subtitle, size=22, color=_rgb("44506B"))
        self._rect(s, Inches(1.15), Inches(5.35), Inches(1.6), Emu(38100), self.t.accent)
        if meta:
            tb3, tf3, p3 = self._text(s, Inches(1.15), Inches(5.55), W - Inches(2.8),
                                      Inches(0.5), size=14, color=_rgb("7A8698"))
            self._run(p3, meta, size=14, color=_rgb("7A8698"))
        return s

    def divider(self, title: str, index: str = "", subtitle: str = ""):
        s = self._blank()
        self._rect(s, 0, 0, W, H, self.t.main)
        self._rect(s, 0, H - Inches(0.16), W, Inches(0.16), self.t.accent)
        if index:
            tb, tf, p = self._text(s, Inches(0.9), Inches(1.5), Inches(4), Inches(1.6),
                                   size=96, bold=True, color=self.t.accent)
            self._run(p, index, size=96, bold=True, color=self.t.accent)
        tb2, tf2, p2 = self._text(s, Inches(0.95), Inches(3.2), W - Inches(2), Inches(1.2),
                                  size=44, bold=True, color=_rgb("FFFFFF"))
        self._run(p2, title, size=44, bold=True, color=_rgb("FFFFFF"))
        if subtitle:
            tb3, tf3, p3 = self._text(s, Inches(1.0), Inches(4.35), W - Inches(2), Inches(0.8),
                                      size=20, color=_rgb("D8E2F5"))
            self._run(p3, subtitle, size=20, color=_rgb("D8E2F5"))
        return s

    def bullets(self, title: str, items: list[str], *, kicker: str = "",
                numbered: bool = False, lead: str = ""):
        s = self._blank()
        self._header(s, title, kicker=kicker)
        top = Inches(1.62)
        size, spacing = _fit(items, base=26)
        if lead:
            tb, tf, p = self._text(s, Inches(0.85), top, W - Inches(1.7), Inches(0.7),
                                   size=18, color=_rgb("5A657A"), line_spacing=1.2)
            self._run(p, lead, size=18, color=_rgb("5A657A"))
            top += Inches(0.72)
        tb, tf, p = self._text(s, Inches(0.85), top, W - Inches(1.6), H - top - Inches(0.7),
                               size=size, line_spacing=1.35)
        for i, raw in enumerate(items):
            item, level = _level_of(raw)
            if i:
                p = tf.add_paragraph()
                p.line_spacing = 1.35
                p.space_before = Pt(spacing)
            has_marker = bool(re.match(r"^(\d{1,2}[.．、)）]|[（(]\d+[）)]|[①-⑳]|[一二三四五六七八九十]+[、.])", item))
            if numbered and not has_marker:
                self._run(p, f"{i+1:02d}  ", size=size, bold=True, color=self.t.accent)
                _indent(p, 0.55 + 0.45 * level, -0.55)
            elif has_marker or level:
                _indent(p, 0.55 + 0.45 * level, -0.55)
                if not has_marker:
                    self._run(p, "▪ ", size=size, color=self.t.accent)
            else:
                self._run(p, "▪ ", size=size, color=self.t.accent)
                _indent(p, 0.55, -0.55)
            if level and size > 20:
                _rich(p, item, size=size - 2, color=self.t.ink, accent=self.t.accent)
            else:
                _rich(p, item, size=size, color=self.t.ink, accent=self.t.accent)
        self._footer(s, title)
        return s

    def table(self, title: str, rows: list[list[str]], *, kicker: str = ""):
        s = self._blank()
        self._header(s, title, kicker=kicker)
        if not rows:
            return s
        cols = max(len(r) for r in rows)
        rows = [r + [""] * (cols - len(r)) for r in rows]
        height = min(Inches(0.55) * len(rows), H - Inches(2.1))
        gt = s.shapes.add_table(len(rows), cols, Inches(0.85), Inches(1.65),
                                W - Inches(1.7), height).table
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = gt.cell(ri, ci)
                cell.text = ""
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = cell.margin_right = Inches(0.08)
                cell.fill.solid()
                cell.fill.fore_color.rgb = self.t.main if ri == 0 else (
                    self.t.light if ri % 2 else _rgb("FFFFFF"))
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER if ri == 0 else PP_ALIGN.LEFT
                self._run(p, val, size=16 if ri else 17, bold=(ri == 0),
                          color=_rgb("FFFFFF") if ri == 0 else self.t.ink)
        self._footer(s, title)
        return s

    def question(self, title: str, stem: str, options: list[str] | None = None,
                 *, kicker: str = "", tag: str = ""):
        s = self._blank()
        self._header(s, title, kicker=kicker or tag)
        top = Inches(1.7)
        if tag and not kicker:
            pass
        # stem
        stem_size = 24 if len(stem) < 150 else (22 if len(stem) < 260 else 20)
        tb, tf, p = self._text(s, Inches(0.9), top, W - Inches(1.8), Inches(2.6),
                               size=stem_size, color=self.t.ink, line_spacing=1.45)
        _rich(p, stem, size=stem_size, color=self.t.ink, accent=self.t.accent)
        if options:
            # 估算题干占几行：24pt 正文在 11.7in 宽度下一行约 26 个汉字
            cpl = max(16, int((W - Inches(1.8)) / Inches(0.36)))
            lines = max(1, -(-len(stem) // cpl))
            line_h = stem_size / 72.0 * 1.5
            oy = top + Inches(line_h * lines + 0.35)
            oy = min(oy, H - Inches(3.2))
            cols = 2 if (len(options) == 4 and max(len(o) for o in options) <= 16) else 1
            size = 22 if cols == 2 else 20
            if cols == 2:
                for i, opt in enumerate(options):
                    x = Inches(0.95) + (W - Inches(2.1)) / 2 * (i % 2)
                    y = oy + Inches(0.72) * (i // 2)
                    tbo, tfo, po = self._text(s, x, y, (W - Inches(2.1)) / 2 - Inches(0.2),
                                              Inches(0.66), size=size, color=self.t.ink)
                    self._run(po, "ABCD"[i] + "．", size=size, bold=True, color=self.t.accent)
                    self._run(po, opt.lstrip("ABCDABCD.． 、"), size=size, color=self.t.ink)
            else:
                tbo, tfo, po = self._text(s, Inches(0.95), oy, W - Inches(2.0),
                                          H - oy - Inches(0.7), size=size, line_spacing=1.4)
                for i, opt in enumerate(options):
                    if i:
                        po = tfo.add_paragraph()
                        po.line_spacing = 1.4
                    self._run(po, "ABCD"[i] + "．", size=size, bold=True, color=self.t.accent)
                    self._run(po, opt.lstrip("ABCDABCD.． 、"), size=size, color=self.t.ink)
        self._footer(s, title)
        return s

    def answer(self, title: str, answer: str, analysis: str = "", *, kicker: str = "解析"):
        s = self._blank()
        self._header(s, title, kicker=kicker)
        self._rect(s, 0, Inches(1.45), Inches(0.09), H - Inches(2.2), _rgb("16A34A"))
        y = Inches(1.66)
        if answer:
            tb, tf, p = self._text(s, Inches(0.95), y, W - Inches(1.8), Inches(0.9),
                                   size=24, color=_rgb("15803D"), line_spacing=1.2)
            self._run(p, "答案：", size=24, bold=True, color=_rgb("15803D"))
            _rich(p, answer, size=24, color=_rgb("15803D"), accent=_rgb("15803D"))
            y += Inches(0.85)
        if analysis:
            size = 22 if len(analysis) < 320 else 19
            tb2, tf2, p2 = self._text(s, Inches(0.95), y, W - Inches(1.85),
                                      H - y - Inches(0.65), size=size,
                                      color=self.t.ink, line_spacing=1.42)
            _rich(p2, analysis, size=size, color=self.t.ink, accent=self.t.accent)
        self._footer(s, title)
        return s

    def image(self, title: str, src: str, caption: str = "", *, kicker: str = ""):
        s = self._blank()
        self._header(s, title, kicker=kicker)
        path = Path(src)
        if path.exists():
            from PIL import Image
            iw, ih = Image.open(path).size
            max_w, max_h = W - Inches(2.4), H - Inches(2.5)
            scale = min(max_w / iw, max_h / ih)
            pic = s.shapes.add_picture(str(path), Inches(1.2), Inches(1.7),
                                       width=Emu(int(iw * scale)))
            pic.left = Emu(int((W - pic.width) / 2))
        if caption:
            tb, tf, p = self._text(s, Inches(0.9), H - Inches(1.2), W - Inches(1.8),
                                   Inches(0.5), size=14, color=_rgb("6B7280"),
                                   align=PP_ALIGN.CENTER)
            self._run(p, caption, size=14, color=_rgb("6B7280"))
        self._footer(s, title)
        return s

    def closing(self, title: str = "本节课小结", items: list[str] | None = None):
        s = self._blank()
        self._rect(s, 0, 0, W, H, self.t.light)
        self._rect(s, 0, 0, W, Inches(0.14), self.t.accent)
        tb, tf, p = self._text(s, Inches(0.9), Inches(0.9), W - Inches(1.8), Inches(1.0),
                               size=40, bold=True, color=self.t.main)
        self._run(p, title, size=40, bold=True, color=self.t.main)
        if items:
            size, sp = _fit(items, base=24)
            tb2, tf2, p2 = self._text(s, Inches(1.0), Inches(2.2), W - Inches(2.0),
                                      H - Inches(3.0), size=size, line_spacing=1.4)
            for i, it in enumerate(items):
                if i:
                    p2 = tf2.add_paragraph()
                    p2.line_spacing = 1.4
                    p2.space_before = Pt(sp)
                self._run(p2, "✔  ", size=size, color=self.t.accent)
                _rich(p2, it, size=size, color=self.t.ink, accent=self.t.accent)
        tb3, tf3, p3 = self._text(s, Inches(1.0), H - Inches(1.1), W - Inches(2.0),
                                  Inches(0.5), size=16, color=_rgb("64748B"))
        self._run(p3, self.school or "", size=16, color=_rgb("64748B"))
        return s

    def save(self, path: str | Path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(path))
        return path


def _fit(items: list[str], base: int = 26) -> tuple[int, int]:
    total = sum(len(i) for i in items)
    n = len(items)
    if n <= 4 and total <= 160:
        return base, 10
    if n <= 6 and total <= 300:
        return base - 2, 8
    if n <= 8 and total <= 460:
        return base - 4, 6
    return max(16, base - 8), 4


def _indent(p, marL_in: float, first_in: float = 0.0):
    """Hanging indent in EMU so wrapped lines align under the text, not the glyph."""
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(marL_in * 914400)))
    pPr.set("indent", str(int(first_in * 914400)))


def _level_of(raw: str) -> tuple[str, int]:
    t = raw.strip()
    if re.match(r"^[（(]\d+[）)]", t) or re.match(r"^[①-⑳]", t):
        return t, 1
    return t, 0


def _rich(p, text: str, *, size: int, color, accent, bold=False):
    """Render inline markup: 【…】 -> accent bold, {{…}} -> underlined blank."""
    for part in re.split(r"(【[^】]{1,14}】|\{\{[^}]*\}\})", text):
        if not part:
            continue
        if part.startswith("【") and part.endswith("】"):
            r = p.add_run()
            r.text = part
            r.font.size = Pt(size)
            r.font.bold = True
            r.font.name = FONT_EN
            r.font.color.rgb = accent
        elif part.startswith("{{") and part.endswith("}}"):
            inner = part[2:-2].strip()
            width = max(3, min(10, len(inner) + 1))
            r = p.add_run()
            r.text = " " * width
            r.font.size = Pt(size)
            r.font.name = FONT_EN
            r.font.underline = True
            r.font.color.rgb = color
        else:
            r = p.add_run()
            r.text = part
            r.font.size = Pt(size)
            r.font.bold = bold
            r.font.name = FONT_EN
            r.font.color.rgb = color


# ---------------------------------------------------------------- builders
def _chunks(text: str, limit: int = 260) -> list[str]:
    if len(text) <= limit:
        return [text]
    parts, buf = [], ""
    for sent in re.split(r"(?<=[。；！？])", text):
        if len(buf) + len(sent) > limit and buf:
            parts.append(buf)
            buf = sent
        else:
            buf += sent
    if buf:
        parts.append(buf)
    return parts


def build_xuean_deck(md: str, out_path: str | Path, *, theme: str = "blue",
                     school: str = "", author: str = "") -> Path:
    d = parse(md)
    deck = Deck(theme=theme, footer=d.title, school=school)
    subject = d.meta.get("学科", "")
    grade = d.meta.get("年级", "")
    deck.cover(d.title or "学案", subtitle=f"{grade}{subject}　{d.meta.get('课时','')}".strip(),
               meta=school or author, tag="学案")

    sections = [s for s in d.sections if s.title]
    for idx, sec in enumerate(sections, 1):
        title = _clean_heading(sec.title)
        is_quiz = bool(re.search(r"当堂检测|课堂检测|检测|练习|作业|自测", sec.title))
        bullets, items, tables, images = _split_section(sec, quiz=is_quiz)

        if re.match(r"^(课前预习|课内探究|当堂检测|课后作业)", sec.title):
            deck.divider(title, index=f"{idx:02d}")

        if bullets:
            for k, chunk in enumerate(_paged(title, bullets)):
                if not chunk:
                    continue
                deck.bullets(title + ("（续）" if k else ""), chunk,
                             kicker="目标" if "目标" in sec.title else "")
        for it in items:
            head = f"{it['name']}｜{title}" if it["name"] else title
            chunks = _chunks(it["stem"], 330) if it["stem"] else []
            for pi, part in enumerate(chunks):
                deck.question(head + ("（续）" if pi else ""), part,
                              it["options"] if pi == len(chunks) - 1 else None)
            if it["answer"] or it["analysis"]:
                deck.answer(head, it["answer"], it["analysis"])
        for rows in tables:
            deck.table(title, rows)
        for src, cap in images:
            deck.image(title, src, cap)

    deck.closing("本节课小结", _summary_points(d))
    return deck.save(out_path)


INLINE_OPT = re.compile(r"([A-DＡ-Ｄ])\s*[.．、]\s*")


def split_inline_options(text: str) -> list[str]:
    """'A.甲　B.乙　C.丙　D.丁' -> ['A.甲','B.乙','C.丙','D.丁'] (needs >=3 markers)."""
    marks = list(INLINE_OPT.finditer(text))
    if len(marks) < 3:
        return []
    out = []
    for i, m in enumerate(marks):
        end = marks[i + 1].start() if i + 1 < len(marks) else len(text)
        piece = text[m.start():end].strip()
        if piece:
            out.append(piece)
    return out if len(out) >= 3 else []


def _split_section(sec: Section, quiz: bool = False):
    """把一节的块分成 普通过渡要点 / 例题（题干+解析） / 表格 / 配图。"""
    bullets, items, tables, images = [], [], [], []
    cur = None
    for b in sec.blocks:
        text = b.text.strip() if b.text else ""
        if b.kind == "table":
            tables.append(b.rows)
            continue
        if b.kind == "image":
            images.append((b.src, b.text))
            continue
        if b.kind == "answer":
            if cur is None:
                cur = {"name": "", "stem": "", "options": [], "answer": "", "analysis": ""}
                items.append(cur)
            mm = re.match(r"^[【\[]?\s*(答案|解析|详解)\s*[】\]]?\s*[:：]\s*(.*)$", text)
            if mm:
                key = "answer" if mm.group(1) == "答案" else "analysis"
                cur[key] = (cur[key] + " " + mm.group(2)).strip()
            else:
                cur["analysis"] = (cur["analysis"] + " " + text).strip()
            continue
        if not text:
            continue
        head = re.match(r"^(例\s*\d+|探究[一二三四五六七八九十\d]*|变式[一二三四五六七八九十\d]*)[：:、.\s]*(.*)$", text)
        if head and len(text) < 400:
            cur = {"name": head.group(1).replace(" ", ""), "stem": head.group(2).strip(),
                   "options": [], "answer": "", "analysis": ""}
            items.append(cur)
            continue
        if quiz and re.match(r"^\d{1,2}[.．、]", text):
            cur = {"name": "", "stem": text, "options": [], "answer": "", "analysis": ""}
            items.append(cur)
            continue
        if cur is not None and (cur["answer"] or cur["analysis"]):
            cur["analysis"] = (cur["analysis"] + " " + text).strip()
            continue
        if cur is not None and not cur["answer"] and not cur["analysis"]:
            inline = split_inline_options(text)
            if inline and len(cur["options"]) < 4:
                cur["options"] += inline[:4]
                continue
            mo = re.match(r"^([A-DＡ-Ｄ])\s*[.．、]\s*(.+)$", text)
            if mo and len(cur["options"]) < 4 and len(text) < 60:
                cur["options"].append(text)
                continue
            cur["stem"] = (cur["stem"] + " " + text).strip()
        else:
            bullets.append(text)
    return bullets, items, tables, images


def _clean_heading(t: str) -> str:
    t = re.sub(r"^知识点[一二三四五六七八九十]+[:：]\s*", "", t)
    return t.strip()


def _paged(title: str, items: list[str], per: int = 7, limit: int = 420) -> list[list[str]]:
    pages, cur, count = [], [], 0
    for it in items:
        if cur and (count + len(it) > limit or len(cur) >= per):
            pages.append(cur)
            cur, count = [], 0
        cur.append(it)
        count += len(it)
    if cur:
        pages.append(cur)
    return pages or [[]]


def _summary_points(d: Doc) -> list[str]:
    pts = []
    for s in d.sections:
        if re.search(r"学习目标|小结|归纳", s.title):
            pts += [b.text.strip() for b in s.blocks if b.kind in ("bullet", "num", "text")]
    if not pts:
        pts = [s.title for s in d.sections if s.title][:5]
    return pts[:6]


def build_paper_deck(md: str, out_path: str | Path, *, theme: str = "ink",
                     school: str = "", mode: str = "review") -> Path:
    """试卷 → 讲评课件：每题一页题干（含选项），答案与解析单独一页。"""
    d = parse(md)
    deck = Deck(theme=theme, footer=d.title, school=school)
    meta = "　".join(x for x in [d.meta.get("时间"), d.meta.get("满分")] if x)
    deck.cover(d.title or "试卷讲评", subtitle="试卷讲评" if mode == "review" else "试卷",
               meta=meta, tag="试卷讲评")

    # 把 "1. 题干 ... A. .. 【答案】.. 【解析】.." 聚合成题
    questions = _group_questions(d)
    _merge_answer_section(d, questions)
    if questions:
        deck.bullets("试卷结构", [
            f"共 {len(questions)} 题，满分 {d.meta.get('满分', '—')}，考试时间 {d.meta.get('时间', '—')}",
            "先看错题分布，再按题型归因：概念不清 / 审题失误 / 计算错误 / 表达不规范",
        ], kicker="总览")

    for qi, q in enumerate(questions, 1):
        head = f"第 {qi} 题" if not q["number"] else f"第 {q['number']} 题"
        stem_pages = _chunks(q["stem"], 330)
        for pi, part in enumerate(stem_pages):
            deck.question(head + ("（续）" if pi else ""), part,
                          q["options"] if pi == len(stem_pages) - 1 else None,
                          kicker=q["section"][:12])
        if q["answer"] or q["analysis"]:
            deck.answer(head, q["answer"], q["analysis"])

    deck.closing("易错点回顾", [
        "审题：圈出限定词（正确的是 / 不正确的是 / 最多 / 至少）",
        "规范：写出依据公式，再代入数据，最后带单位",
        "复盘：把错题按知识点归档，三天后重做一遍",
    ])
    return deck.save(out_path)


def _merge_answer_section(d: Doc, questions: list[dict]) -> None:
    """把「参考答案」栏目里的答案按题号补回题目。"""
    ans_sec = [s for s in d.sections if re.match(r"^(参考答案|答案与解析|答案|解析)", (s.title or "").strip())]
    if not ans_sec:
        return
    by_num = {str(q["number"]): q for q in questions if q["number"]}
    pending = None
    for s in ans_sec:
        for b in s.blocks:
            text = (b.text or "").strip()
            if not text:
                continue
            m = re.match(r"^(\d{1,2})\s*[.．、]\s*(.*)$", text)
            if m:
                pending = by_num.get(m.group(1))
                text = m.group(2).strip()
            if pending is None:
                continue
            mm = re.match(r"^[【\[]?\s*(答案|解析|详解)\s*[】\]]?\s*[:：]?\s*(.*)$", text)
            if mm:
                key = "answer" if mm.group(1) == "答案" else "analysis"
                pending[key] = (pending[key] + " " + mm.group(2)).strip()
            elif not pending["answer"] and len(text) <= 40:
                pending["answer"] = text
            else:
                pending["analysis"] = (pending["analysis"] + " " + text).strip()


NUM_RE = re.compile(r"^\s*(\d{1,2})\s*[.．、]\s*(.+)$")
OPT_RE = re.compile(r"^\s*([A-DＡ-Ｄ])\s*[.．、]\s*(.+)$")


def _group_questions(d: Doc) -> list[dict]:
    out: list[dict] = []
    cur = None
    section = ""
    for s in d.sections:
        if s.title:
            section = s.title
        if re.search(r"注意事项|考试说明|考生须知|评分标准|答题卡|考试要求|命题说明", s.title or ""):
            continue
        if s.title and not re.search(r"题|部分|卷|[一二三四五六七八九十]、", s.title):
            continue
        for b in s.blocks:
            if b.kind == "table" or b.kind == "image":
                if cur is not None:
                    cur["images"].append(b)
                continue
            text = b.text.strip()
            if b.kind.startswith("answer"):
                pass
            m = NUM_RE.match(text) if b.kind in ("text", "num") else None
            if m and len(text) < 500:
                cur = {"number": m.group(1), "stem": m.group(2), "options": [],
                       "answer": "", "analysis": "", "section": section, "images": []}
                out.append(cur)
                continue
            if cur is None:
                continue
            mo = OPT_RE.match(text)
            if mo and len(cur["options"]) < 4 and len(text) < 40:
                cur["options"].append(mo.group(1) + "．" + mo.group(2))
                continue
            if b.kind == "answer":
                mm = re.match(r"^[【\[]?\s*(答案|解析|详解)\s*[】\]]?\s*[:：]\s*(.*)$", text)
                if mm:
                    key = "answer" if mm.group(1) == "答案" else "analysis"
                    cur[key] = (cur[key] + " " + mm.group(2)).strip()
                else:
                    cur["analysis"] = (cur["analysis"] + " " + text).strip()
                continue
            if cur["answer"] or cur["analysis"]:
                cur["analysis"] = (cur["analysis"] + " " + text).strip()
            else:
                cur["stem"] = (cur["stem"] + " " + text).strip()
    return [q for q in out if q["stem"]]
