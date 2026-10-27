"""16:9 课件渲染：学案 → 新授课 PPT，试卷 → 讲评 PPT。

版式来自 318 份真实课件的统计 + 6 份逐页拆解（物理一轮 / 历史一轮 / 化学 / 生物）：
  * 16:9；正文 24–28pt；中文 微软雅黑，西文 Times New Roman
  * 三色分工：深蓝=框架/页头，青蓝=连接/提示，**红=易错/高频**（红只用于警示）
  * 内页两套页头：栏目式（主题 · 子主题）与通栏式（考点一 …，深蓝通栏 + 白字）
  * 固定底部色带 + 页脚「学校 · 课题 | 页码」
  * 例题 = 序号块 + 题源小字 + 选项网格 + 右侧红色批注框
"""
from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from .mdparse import Block, Doc, Section, parse

W, H = Inches(13.333), Inches(7.5)
FONT_CN = "微软雅黑"
FONT_EN = "Times New Roman"

# 三色体系：main 结构 / accent 连接 / warn 警示 / gold 封面点缀
THEMES = {
    "blue": {"main": "0F2B55", "accent": "2E86D6", "warn": "C81E1E",
             "light": "EAF2FB", "ink": "1B2430", "gold": "E8B21A"},
    "teal": {"main": "0B4F4A", "accent": "12A48C", "warn": "C2410C",
             "light": "E6F5F1", "ink": "16262A", "gold": "D9A31B"},
    "violet": {"main": "3B2A6B", "accent": "7C5CD3", "warn": "BE123C",
               "light": "F0EBFA", "ink": "241C33", "gold": "D4A017"},
    "orange": {"main": "7C2D12", "accent": "EA580C", "warn": "B91C1C",
               "light": "FDF0E7", "ink": "2B1A12", "gold": "C9941A"},
    "ink": {"main": "1F2937", "accent": "475569", "warn": "B91C1C",
            "light": "F1F3F6", "ink": "111827", "gold": "C9A227"},
}


def _rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


@dataclass
class DeckTheme:
    name: str = "blue"

    def __post_init__(self):
        self.d = THEMES.get(self.name, THEMES["blue"])

    @property
    def main(self):
        return _rgb(self.d["main"])

    @property
    def accent(self):
        return _rgb(self.d["accent"])

    @property
    def warn(self):
        return _rgb(self.d["warn"])

    @property
    def light(self):
        return _rgb(self.d["light"])

    @property
    def ink(self):
        return _rgb(self.d["ink"])

    @property
    def gold(self):
        return _rgb(self.d["gold"])

    @property
    def white(self):
        return _rgb("FFFFFF")


class Deck:
    def __init__(self, theme: str = "blue", footer: str = "", school: str = ""):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.t = DeckTheme(theme if theme in THEMES else "blue")
        self.footer = footer
        self.school = school
        self.n = 0

    # ------------------------------------------------------------ primitives
    def _blank(self):
        s = self.prs.slides.add_slide(self.prs.slide_layouts[6])
        self.n += 1
        return s

    def _rect(self, slide, x, y, w, h, fill=None, line=None, shape=MSO_SHAPE.RECTANGLE,
              radius=None, line_w=1.0):
        sp = slide.shapes.add_shape(shape, x, y, w, h)
        if fill is None:
            sp.fill.background()
        else:
            sp.fill.solid()
            sp.fill.fore_color.rgb = fill
        if line is None:
            sp.line.fill.background()
        else:
            sp.line.color.rgb = line
            sp.line.width = Pt(line_w)
        sp.shadow.inherit = False
        if radius is not None and shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            try:
                sp.adjustments[0] = radius
            except Exception:
                pass
        return sp

    def _text(self, slide, x, y, w, h, *, size=24, align=PP_ALIGN.LEFT,
              anchor=MSO_ANCHOR.TOP, line_spacing=1.3, vertical=False):
        tb = slide.shapes.add_textbox(x, y, w, h)
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = anchor
        tf.margin_left = tf.margin_right = Inches(0.06)
        tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.alignment = align
        p.line_spacing = line_spacing
        p.font.size = Pt(size)
        if vertical:
            bodyPr = tf._txBody.find(qn("a:bodyPr"))
            bodyPr.set("vert", "eaVert")
        return tb, tf, p

    @staticmethod
    def _run(p, text, *, size, bold=False, color=None, cn=FONT_CN, en=FONT_EN,
             underline=False):
        r = p.add_run()
        r.text = text
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.underline = underline
        r.font.name = en
        if color is not None:
            r.font.color.rgb = color
        rpr = r._r.get_or_add_rPr()
        ea = rpr.find(qn("a:ea"))
        if ea is None:
            from lxml import etree
            ea = etree.SubElement(rpr, qn("a:ea"))
        ea.set("typeface", cn)
        return r

    def _p(self, tf, *, first=False, align=PP_ALIGN.LEFT, line_spacing=1.35,
           space_before=0):
        if first:
            p = tf.paragraphs[0]
            p.alignment = align
            p.line_spacing = line_spacing
        else:
            p = tf.add_paragraph()
            p.alignment = align
            p.line_spacing = line_spacing
        if space_before:
            p.space_before = Pt(space_before)
        return p

    def _rich(self, p, text: str, *, size: int, color=None, bold=False):
        """【…】→警示红加粗；{{…}}→下划线空格；其余正文色。"""
        color = color or self.t.ink
        text = _nb(text)
        for part in re.split(r"(【[^】]{1,16}】|\{\{[^}]*\}\})", text):
            if not part:
                continue
            if part.startswith("【") and part.endswith("】"):
                self._run(p, part, size=size, bold=True, color=self.t.warn)
            elif part.startswith("{{") and part.endswith("}}"):
                inner = part[2:-2].strip()
                self._run(p, "\u3000" * max(2, min(10, len(inner) + 1)),
                          size=size, color=color, underline=True)
            else:
                self._run(p, part, size=size, color=color, bold=bold)

    # ------------------------------------------------------------ chrome
    def _chrome(self, slide, *, band: bool = True):
        if band:
            self._rect(slide, 0, H - Inches(0.13), W, Inches(0.13), self.t.accent)
        y = H - Inches(0.56)
        tb, tf, p = self._text(slide, Inches(0.6), y, Inches(9.4), Inches(0.3), size=10)
        school = "" if (self.school and self.school in (self.footer or "")) else self.school
        label = " · ".join(x for x in [school, self.footer] if x)
        if label:
            self._run(p, label, size=10, color=_rgb("8A94A6"))
        tb2, tf2, p2 = self._text(slide, W - Inches(1.25), y, Inches(0.75), Inches(0.3),
                                  size=10, align=PP_ALIGN.RIGHT)
        self._run(p2, str(self.n), size=10, color=_rgb("8A94A6"))

    def _chip(self, slide, text: str, *, x=None, y=Inches(0.5), dark=False, max_chars=14):
        if len(text) > max_chars:
            text = text[: max_chars - 1] + "…"
        cw = Inches(0.185 * len(text) + 0.34)
        x = x if x is not None else W - cw - Inches(0.55)
        self._rect(slide, x, y, cw, Inches(0.4),
                   _rgb("2A3F63") if dark else self.t.light,
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.35)
        tb, tf, p = self._text(slide, x, y + Inches(0.02), cw, Inches(0.36), size=11,
                               align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        self._run(p, text, size=11, bold=True,
                  color=self.t.white if dark else self.t.accent)

    def _header(self, slide, title: str, *, kicker: str = ""):
        """栏目式页头：顶部细线 + 左侧竖条 + 标题（主题 · 子主题）+ 右上 chip。"""
        self._rect(slide, 0, 0, W, Inches(0.1), self.t.main)
        self._rect(slide, Inches(0.58), Inches(0.48), Inches(0.12), Inches(0.58),
                   self.t.accent)
        size = 28 if len(title) <= 18 else (24 if len(title) <= 26 else 21)
        tb, tf, p = self._text(slide, Inches(0.82), Inches(0.42), W - Inches(2.7),
                               Inches(0.7), size=size, anchor=MSO_ANCHOR.MIDDLE)
        self._rich(p, title, size=size, color=self.t.main, bold=True)
        if kicker and kicker.strip() != title.strip():
            self._chip(slide, kicker)

    def _header_band(self, slide, title: str, *, tags: list[str] | None = None):
        """通栏式页头：深蓝色块通栏 + 白字标题 + 右上小标签（考点/知识点/例题页）。"""
        self._rect(slide, 0, 0, W, Inches(0.9), self.t.main)
        self._rect(slide, 0, Inches(0.9), W, Inches(0.05), self.t.accent)
        size = 26 if len(title) <= 20 else (22 if len(title) <= 28 else 19)
        tb, tf, p = self._text(slide, Inches(0.6), Inches(0.15), W - Inches(4.1),
                               Inches(0.6), size=size, anchor=MSO_ANCHOR.MIDDLE)
        self._run(p, title, size=size, bold=True, color=self.t.white)
        if tags:
            tag_list, seen = [], set()
            for x in tags:
                if x and x not in seen:
                    seen.add(x)
                    tag_list.append(x)
            tag_list = tag_list[:2]
            x_cursor = W - Inches(0.6)
            for tg in tag_list:
                tg = tg if len(tg) <= 11 else tg[:10] + "…"
                cw = Inches(0.19 * len(tg) + 0.36)
                x_cursor -= cw
                self._rect(slide, x_cursor, Inches(0.24), cw, Inches(0.4),
                           _rgb("2E4A78"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tb2, tf2, p2 = self._text(slide, x_cursor, Inches(0.26), cw, Inches(0.36),
                                          size=11, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p2, tg, size=11, bold=True, color=self.t.white)
                x_cursor -= Inches(0.1)

    def _spine(self, slide, text: str, *, y=Inches(1.55), h=Inches(2.4), x=Inches(0.24)):
        """左侧竖排书脊标签（真实课件用它标页面角色）。"""
        self._rect(slide, x, y, Inches(0.4), h, self.t.main,
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.18)
        tb, tf, p = self._text(slide, x, y + Inches(0.1), Inches(0.4), h - Inches(0.2),
                               size=15, align=PP_ALIGN.CENTER,
                               anchor=MSO_ANCHOR.MIDDLE, vertical=True)
        self._run(p, text, size=15, bold=True, color=self.t.white)

    def _tip_box(self, slide, text: str, *, x=Inches(0.85), y=None, w=None, size=15,
                 label="特别提醒"):
        w = w or (W - Inches(1.7))
        y = y if y is not None else H - Inches(1.9)
        self._rect(slide, x, y, w, Inches(1.1), _rgb("FDF2F2"), line=_rgb("F3B4B4"),
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
        self._rect(slide, x, y, Inches(0.08), Inches(1.1), self.t.warn)
        tb, tf, p = self._text(slide, x + Inches(0.25), y + Inches(0.12),
                               w - Inches(0.45), Inches(0.9), size=size, line_spacing=1.4)
        self._run(p, f"{label}：", size=size, bold=True, color=self.t.warn)
        self._rich(p, text, size=size, color=_rgb("8A1B1B"))

    # ------------------------------------------------------------ slides
    def cover(self, title: str, subtitle: str = "", meta: str = "", tag: str = "",
              index: str = ""):
        s = self._blank()
        self._rect(s, W * 0.58, 0, W * 0.42, H, self.t.main)
        self._rect(s, W * 0.575, 0, Inches(0.09), H, self.t.accent)
        self._rect(s, 0, 0, Inches(0.16), H, self.t.main)

        if index:
            tb, tf, p = self._text(s, W * 0.6, Inches(0.75), W * 0.36, Inches(0.9),
                                   size=40, align=PP_ALIGN.RIGHT)
            self._run(p, index, size=40, bold=True, color=self.t.gold)
            self._rect(s, W * 0.6 + Inches(0.1), Inches(1.62), W * 0.36 - Inches(0.35),
                       Emu(28575), self.t.accent)

        title = _soft_wrap(title, 13)
        longest = max(len(x) for x in title.split("\v"))
        avail_in = (W * 0.47) / 914400.0
        tsize = int(min(42, max(20, avail_in * 72.0 / max(longest, 1) * 0.92)))
        tb, tf, p = self._text(s, Inches(0.95), Inches(2.0), W * 0.47, Inches(2.2),
                               size=tsize, line_spacing=1.2)
        for li, line in enumerate(title.split("\v")):
            if li:
                p.add_line_break()          # 真正的软换行，不是字面量
            self._rich(p, line, size=tsize, color=self.t.main, bold=True)

        if subtitle:
            tb2, tf2, p2 = self._text(s, Inches(1.0), Inches(3.95), W * 0.48, Inches(0.9),
                                      size=20, line_spacing=1.3)
            self._run(p2, subtitle, size=20, color=_rgb("44506B"))
        self._rect(s, Inches(1.0), Inches(5.05), Inches(1.5), Emu(38100), self.t.accent)
        if meta:
            tb3, tf3, p3 = self._text(s, Inches(1.0), Inches(5.25), W * 0.5, Inches(0.5),
                                      size=13)
            self._run(p3, meta, size=13, color=_rgb("7A8698"))
        if tag:
            tw = Inches(0.2 * len(tag) + 0.5)
            self._rect(s, W - tw - Inches(0.7), H - Inches(1.15), tw, Inches(0.46),
                       self.t.white, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4)
            tb4, tf4, p4 = self._text(s, W - tw - Inches(0.7), H - Inches(1.13), tw,
                                      Inches(0.42), size=13, align=PP_ALIGN.CENTER,
                                      anchor=MSO_ANCHOR.MIDDLE)
            self._run(p4, tag, size=13, bold=True, color=self.t.main)
        return s

    def divider(self, title: str, index: str = "", subtitle: str = "", spine: str = ""):
        s = self._blank()
        self._rect(s, 0, 0, W, H, self.t.main)
        self._rect(s, 0, H - Inches(0.16), W, Inches(0.16), self.t.accent)
        if index:
            tb, tf, p = self._text(s, Inches(0.85), Inches(1.3), Inches(4), Inches(1.7),
                                   size=92)
            self._run(p, index, size=92, bold=True, color=self.t.accent)
        tb2, tf2, p2 = self._text(s, Inches(0.92), Inches(3.1), W - Inches(2.2), Inches(1.2),
                                  size=40)
        self._run(p2, title, size=40, bold=True, color=self.t.white)
        if subtitle:
            tb3, tf3, p3 = self._text(s, Inches(0.98), Inches(4.25), W - Inches(2.2),
                                      Inches(0.8), size=18)
            self._run(p3, subtitle, size=18, color=_rgb("CBD8F0"))
        if spine:
            self._rect(s, W - Inches(0.55), Inches(2.2), Inches(0.4), Inches(2.6),
                       self.t.accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
            tb4, tf4, p4 = self._text(s, W - Inches(0.55), Inches(2.3), Inches(0.4),
                                      Inches(2.4), size=15, align=PP_ALIGN.CENTER,
                                      anchor=MSO_ANCHOR.MIDDLE, vertical=True)
            self._run(p4, spine, size=15, bold=True, color=self.t.white)
        return s

    def bullets(self, title: str, items: list[str], *, kicker: str = "", lead: str = "",
                numbered: bool = False, band: bool = False, spine: str = "",
                density: str = "normal"):
        s = self._blank()
        if band:
            self._header_band(s, title, tags=[kicker] if kicker else None)
            top = Inches(1.3)
        else:
            self._header(s, title, kicker=kicker)
            top = Inches(1.45)
        if spine:
            self._spine(s, spine)
            left = Inches(1.0)
        else:
            left = Inches(0.85)

        size, spacing = _fit(items, base=26 if density == "normal" else 21)
        if lead:
            tb, tf, p = self._text(s, left, top, W - left - Inches(0.85), Inches(0.7),
                                   size=17, line_spacing=1.25)
            self._run(p, lead, size=17, color=_rgb("5A657A"))
            top += Inches(0.66)
        tb, tf, p = self._text(s, left, top, W - left - Inches(0.85),
                               H - top - Inches(0.72), size=size, line_spacing=1.35)
        for i, raw in enumerate(items):
            item, level = _level_of(raw)
            p = self._p(tf, first=(i == 0), space_before=0 if i == 0 else spacing)
            has_marker = bool(re.match(r"^(\d{1,2}[.．、)）]|[（(]\d+[）)]|[①-⑳])", item))
            _indent(p, 0.52 + 0.45 * level, -0.52)
            if numbered and not has_marker:
                self._run(p, f"{i + 1:02d}　", size=size, bold=True, color=self.t.accent)
            elif not has_marker:
                self._run(p, "▪ ", size=size, color=self.t.accent)
            self._rich(p, item, size=size - (2 if level else 0))
        self._chrome(s)
        return s

    def columns(self, title: str, left_title: str, left_items: list[str],
                right_title: str = "", right_items: list[str] | None = None, *,
                kicker: str = "", note: str = "", left_ratio: float = 0.56):
        """两栏页：左文右图/右要点——真实课件最常用的图文关系。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        top = Inches(1.5)
        bottom = H - Inches(1.5 if note else 0.72)
        gap = Inches(0.28)
        lw = (W - Inches(1.6)) * left_ratio
        rw = (W - Inches(1.6)) - lw - gap

        self._rect(s, Inches(0.7), top, lw + Inches(0.2), bottom - top,
                   self.t.light, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)
        tb, tf, p = self._text(s, Inches(0.88), top + Inches(0.12), lw - Inches(0.2),
                               Inches(0.4), size=17)
        self._run(p, left_title, size=17, bold=True, color=self.t.main)
        size, sp = _fit(left_items, base=20)
        tb2, tf2, p2 = self._text(s, Inches(0.93), top + Inches(0.6), lw - Inches(0.3),
                                  bottom - top - Inches(0.5), size=size, line_spacing=1.35)
        for i, item in enumerate(left_items):
            pp = self._p(tf2, first=(i == 0), space_before=0 if i == 0 else sp)
            _indent(pp, 0.42, -0.42)
            self._run(pp, "▪ ", size=size, color=self.t.accent)
            self._rich(pp, item, size=size)

        if right_items:
            rx = Inches(0.7) + lw + Inches(0.2) + gap
            self._rect(s, rx, top, rw, bottom - top, self.t.light,
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.03)
            if right_title:
                tb3, tf3, p3 = self._text(s, rx + Inches(0.18), top + Inches(0.12),
                                          rw - Inches(0.36), Inches(0.4), size=17)
                self._run(p3, right_title, size=17, bold=True, color=self.t.accent)
            size2, sp2 = _fit(right_items, base=19)
            tb4, tf4, p4 = self._text(s, rx + Inches(0.22), top + Inches(0.6),
                                      rw - Inches(0.44), bottom - top - Inches(0.5),
                                      size=size2, line_spacing=1.35)
            for i, item in enumerate(right_items):
                pp = self._p(tf4, first=(i == 0), space_before=0 if i == 0 else sp2)
                _indent(pp, 0.4, -0.4)
                self._run(pp, "· ", size=size2, color=self.t.accent)
                self._rich(pp, item, size=size2)
        if note:
            self._tip_box(s, note, y=H - Inches(1.45), size=13)
        self._chrome(s)
        return s

    def table(self, title: str, rows: list[list[str]], *, kicker: str = "",
              highlight: str = ""):
        s = self._blank()
        self._header(s, title, kicker=kicker)
        if not rows:
            return s
        cols = max(len(r) for r in rows)
        rows = [r + [""] * (cols - len(r)) for r in rows]
        row_h = min(Inches(0.52), (H - Inches(2.3)) / max(1, len(rows)))
        gt = s.shapes.add_table(len(rows), cols, Inches(0.8), Inches(1.5),
                                W - Inches(1.6), row_h * len(rows)).table
        for ri, row in enumerate(rows):
            for ci, val in enumerate(row):
                cell = gt.cell(ri, ci)
                cell.text = ""
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = cell.margin_right = Inches(0.08)
                cell.fill.solid()
                if ri == 0:
                    cell.fill.fore_color.rgb = self.t.main
                elif highlight and highlight in val:
                    cell.fill.fore_color.rgb = _rgb("FDECEC")
                else:
                    cell.fill.fore_color.rgb = self.t.light if ri % 2 else self.t.white
                p = cell.text_frame.paragraphs[0]
                p.alignment = PP_ALIGN.CENTER if (ri == 0 or len(val) <= 8) else PP_ALIGN.LEFT
                color = self.t.white if ri == 0 else self.t.ink
                if ri and highlight and highlight in val:
                    color = self.t.warn
                self._run(p, val, size=14 if ri else 15, bold=(ri == 0), color=color)
        self._chrome(s)
        return s

    def question(self, title: str, stem: str, options: list[str] | None = None, *,
                 source: str = "", note: str = "", tag: str = "", kicker: str = "",
                 index: str = ""):
        """例题页：序号块 + 题源 + 题干 + 选项网格 + 右侧红色批注框。"""
        s = self._blank()
        tags = [x for x in [tag, kicker] if x]
        self._header_band(s, title, tags=tags)
        left = Inches(0.6)
        top = Inches(1.18)
        note_w = Inches(3.0) if note else 0
        body_w = W - left - Inches(0.6) - (note_w + Inches(0.25) if note else 0)

        if index:
            bw = Inches(0.34 * len(index) + 0.5)
            self._rect(s, left, top, bw, Inches(0.44), self.t.accent,
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
            tb, tf, p = self._text(s, left, top + Inches(0.02), bw, Inches(0.4), size=15,
                                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p, index, size=15, bold=True, color=self.t.white)
            if source:
                tb2, tf2, p2 = self._text(s, left + bw + Inches(0.15), top + Inches(0.06),
                                          body_w - bw, Inches(0.34), size=12)
                self._run(p2, source, size=12, color=_rgb("7A8698"))
            top += Inches(0.62)
        elif source:
            tb2, tf2, p2 = self._text(s, left, top, body_w, Inches(0.32), size=12)
            self._run(p2, source, size=12, color=_rgb("7A8698"))
            top += Inches(0.34)

        stem_size = 22 if len(stem) < 130 else (20 if len(stem) < 240 else 18)
        tb, tf, p = self._text(s, left, top, body_w, Inches(2.3), size=stem_size,
                               line_spacing=1.45)
        self._rich(p, stem, size=stem_size)

        if options:
            cpl = max(14, int(body_w / Inches(0.4)))
            lines = max(1, -(-len(stem) // cpl))
            oy = top + Inches((stem_size / 72.0) * 1.55 * lines + 0.26)
            oy = min(oy, H - Inches(2.6))
            self._option_grid(s, options, left, oy, body_w)

        if note:
            nx = W - Inches(0.6) - note_w
            hgt = H - top - Inches(0.7)
            self._rect(s, nx, top - Inches(0.12), note_w, hgt, _rgb("FDFBF7"),
                       line=_rgb("E5C9A8"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
            self._rect(s, nx, top - Inches(0.12), Inches(0.07), hgt, self.t.warn)
            tb2, tf2, p2 = self._text(s, nx + Inches(0.2), top, note_w - Inches(0.36),
                                      hgt - Inches(0.2), size=14, line_spacing=1.45)
            self._run(p2, "判定要点", size=14, bold=True, color=self.t.warn)
            for i, line in enumerate(_sentences(note)):
                pp = self._p(tf2, space_before=6)
                _indent(pp, 0.32, -0.32)
                self._run(pp, f"{i + 1}. ", size=13, bold=True, color=self.t.warn)
                self._rich(pp, line, size=13, color=_rgb("6B2B2B"))
        self._chrome(s)
        return s

    def _option_grid(self, slide, options: list[str], left, top, width):
        """4 选项：够短就 4 列，否则 2×2；每格悬挂缩进，字母加粗着色。"""
        def cw(t):
            return sum(1.0 if ord(c) > 127 else 0.55 for c in t)

        n = len(options)
        per = 4 if (n == 4 and sum(cw(o) for o in options) <= 46) else 2
        rows = [options[i:i + per] for i in range(0, n, per)]
        size = 19 if per == 4 else 18
        cell_h = Inches(0.6)
        for ri, row in enumerate(rows):
            for ci, opt in enumerate(row):
                x = left + (width / per) * ci
                y = top + (cell_h + Inches(0.05)) * ri
                tb, tf, p = self._text(slide, x, y, width / per - Inches(0.12), cell_h,
                                       size=size, line_spacing=1.25)
                _indent(p, 0.4, -0.4)
                letter = "ABCD"[ri * per + ci] if ri * per + ci < 4 else "?"
                body = re.sub(r"^\s*[A-DＡ-Ｄ]\s*[.．、]\s*", "", opt).strip()
                self._run(p, f"{letter}．", size=size, bold=True, color=self.t.accent)
                self._rich(p, body, size=size)

    def answer(self, title: str, answer: str, analysis: str = "", *, kicker: str = "",
               tag: str = "", index: str = ""):
        s = self._blank()
        tags = [x for x in [tag, kicker] if x]
        self._header_band(s, title, tags=tags)
        y = Inches(1.2)
        if index:
            bw = Inches(0.34 * len(index) + 0.5)
            self._rect(s, Inches(0.6), y, bw, Inches(0.42), _rgb("15803D"),
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
            tb, tf, p = self._text(s, Inches(0.6), y + Inches(0.02), bw, Inches(0.38),
                                   size=14, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p, index, size=14, bold=True, color=self.t.white)
            y += Inches(0.6)
        if answer:
            asize = 22 if len(answer) <= 36 else (19 if len(answer) <= 90 else 16)
            ah = Inches(0.62 if len(answer) <= 36 else (1.05 if len(answer) <= 90 else 1.6))
            self._rect(s, Inches(0.6), y, Inches(0.07), ah, _rgb("16A34A"))
            tb, tf, p = self._text(s, Inches(0.8), y, W - Inches(1.4), ah, size=asize,
                                   line_spacing=1.35)
            self._run(p, "答案：", size=asize, bold=True, color=_rgb("15803D"))
            self._rich(p, answer, size=asize, color=_rgb("15803D"))
            y += ah + Inches(0.18)
        if analysis:
            analysis = re.sub(r"^(解析|详解|分析)\s*[:：]\s*", "", analysis.strip())
            size = 20 if len(analysis) < 300 else 18
            tb2, tf2, p2 = self._text(s, Inches(0.62), y, W - Inches(1.25),
                                      H - y - Inches(0.7), size=size, line_spacing=1.45)
            if "【" not in analysis:
                self._run(p2, "解析　", size=size, bold=True, color=self.t.accent)
            self._rich(p2, analysis, size=size)
        self._chrome(s)
        return s

    def tip(self, title: str, text: str, *, kicker: str = "易错提醒"):
        s = self._blank()
        self._header(s, title, kicker=kicker)
        self._rect(s, Inches(0.8), Inches(1.55), W - Inches(1.6), Inches(3.3),
                   _rgb("FDF2F2"), line=_rgb("F3B4B4"),
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        self._rect(s, Inches(0.8), Inches(1.55), Inches(0.1), Inches(3.3), self.t.warn)
        size = 22 if len(text) < 160 else 19
        tb, tf, p = self._text(s, Inches(1.1), Inches(1.8), W - Inches(2.3), Inches(2.8),
                               size=size, line_spacing=1.5)
        self._run(p, "⚠ ", size=size, color=self.t.warn)
        self._rich(p, text, size=size, color=_rgb("8A1B1B"))
        self._chrome(s)
        return s

    def image(self, title: str, src: str, caption: str = "", *, kicker: str = "",
              band: bool = False):
        s = self._blank()
        if band:
            self._header_band(s, title, tags=[kicker] if kicker else None)
            top = Inches(1.18)
        else:
            self._header(s, title, kicker=kicker)
            top = Inches(1.45)
        path = Path(src)
        if path.exists():
            from PIL import Image
            iw, ih = Image.open(path).size
            max_w = W - Inches(2.2)
            max_h = H - top - Inches(1.0)
            scale = min(max_w / iw, max_h / ih)
            pic = s.shapes.add_picture(str(path), Inches(1.1), top,
                                       width=Emu(int(iw * scale)))
            if pic.height > max_h:
                pic.height = Emu(int(max_h))
                pic.width = Emu(int(iw * (max_h / ih)))
            pic.left = Emu(int((W - pic.width) / 2))
        if caption:
            tb, tf, p = self._text(s, Inches(0.9), H - Inches(1.0), W - Inches(1.8),
                                   Inches(0.38), size=13, align=PP_ALIGN.CENTER)
            self._run(p, caption, size=13, color=_rgb("6B7280"))
        self._chrome(s)
        return s

    def closing(self, title: str = "本节课小结", items: list[str] | None = None):
        s = self._blank()
        self._rect(s, 0, 0, W, H, self.t.light)
        self._rect(s, 0, 0, W, Inches(0.12), self.t.main)
        self._rect(s, 0, H - Inches(0.12), W, Inches(0.12), self.t.accent)
        tb, tf, p = self._text(s, Inches(0.85), Inches(0.8), W - Inches(1.7), Inches(0.9),
                               size=38)
        self._run(p, title, size=38, bold=True, color=self.t.main)
        if items:
            size, sp = _fit(items, base=23)
            tb2, tf2, p2 = self._text(s, Inches(0.95), Inches(1.95), W - Inches(1.9),
                                      H - Inches(2.85), size=size, line_spacing=1.45)
            for i, it in enumerate(items):
                pp = self._p(tf2, first=(i == 0), space_before=0 if i == 0 else sp)
                _indent(pp, 0.45, -0.45)
                self._run(pp, "✔ ", size=size, color=self.t.accent)
                self._rich(pp, it, size=size)
        tb3, tf3, p3 = self._text(s, Inches(0.95), H - Inches(0.95), W - Inches(1.9),
                                  Inches(0.38), size=14)
        self._run(p3, self.school or "", size=14, color=_rgb("64748B"))
        tb4, tf4, p4 = self._text(s, W - Inches(1.25), H - Inches(0.95), Inches(0.75),
                                  Inches(0.38), size=14, align=PP_ALIGN.RIGHT)
        self._run(p4, str(self.n), size=14, color=_rgb("64748B"))
        return s

    def save(self, path: str | Path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        self.prs.save(str(path))
        return path


# ---------------------------------------------------------------- utilities
def _indent(p, mar_l_in: float, first_in: float = 0.0):
    """Hanging indent in EMU so wrapped lines align under the text, not the glyph."""
    pPr = p._p.get_or_add_pPr()
    pPr.set("marL", str(int(mar_l_in * 914400)))
    pPr.set("indent", str(int(first_in * 914400)))


def _fit(items: list[str], base: int = 26) -> tuple[int, int]:
    total = sum(len(i) for i in items)
    n = len(items)
    if n <= 4 and total <= 150:
        return base, 10
    if n <= 6 and total <= 280:
        return base - 2, 8
    if n <= 8 and total <= 430:
        return base - 4, 6
    return max(15, base - 8), 4


def _level_of(raw: str) -> tuple[str, int]:
    t = raw.strip()
    if re.match(r"^[（(]\d+[）)]", t) or re.match(r"^[①-⑳]", t):
        return t, 1
    return t, 0


def _nb(text: str) -> str:
    units = ("m/s²", "m/s", "km/h", "kg", "mol", "cm", "mm", "km", "mL", "Pa", "℃",
             "m", "s", "g", "L", "N", "J", "W", "V", "A", "%")
    pat = re.compile(r"(?<=\d)\s+(" + "|".join(re.escape(u) for u in units) + r")(?![\w/²])")
    return pat.sub(lambda m: "\u00a0" + m.group(1), text)


def _soft_wrap(title: str, limit: int) -> str:
    """封面标题在语义边界换行（括号前 / 顿号前），避免把「规律」拆成两行。"""
    if len(title) <= limit or "\v" in title:
        return title
    for mark in ("（", "(", "：", ":", "——"):
        idx = title.find(mark)
        if 0 < idx <= limit + 4:
            head, tail = title[:idx].strip(), title[idx:].strip()
            if head and tail:
                return head + "\v" + tail
    cut = title.rfind("、", 0, limit) if "、" in title[:limit] else -1
    if cut > 4:
        return title[:cut + 1] + "\v" + title[cut + 1:].lstrip()
    return title


def _sentences(text: str) -> list[str]:
    parts = [p.strip() for p in re.split(r"(?<=[。；;])", text) if p.strip()]
    return parts or [text]


def _chunks(text: str, limit: int = 280) -> list[str]:
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


def _clean_heading(t: str) -> str:
    t = re.sub(r"^知识点[一二三四五六七八九十]+[:：]\s*", "", t)
    return t.strip()


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


def _new_item(name: str = "", stem: str = "") -> dict:
    return {"name": name, "stem": stem, "options": [], "answer": "", "analysis": "",
            "note": "", "source": ""}


def _split_section(sec: Section, quiz: bool = False):
    """把一节的块分成：要点 / 例题（题干+选项+答案+批注）/ 表格 / 配图 / 提醒。"""
    bullets, items, tables, images, tips = [], [], [], [], []
    cur = None
    for b in sec.blocks:
        text = b.text.strip() if b.text else ""
        if b.kind == "table":
            tables.append(b.rows)
            continue
        if b.kind == "image":
            images.append((b.src, b.text))
            continue
        if not text:
            continue
        if b.kind == "tip":
            tips.append(re.sub(r"^【(特别提醒|易错提醒|提醒|注意|易错|警示)】\s*", "", text))
            continue
        if b.kind == "insight":
            if cur is None:
                cur = _new_item()
                items.append(cur)
            cur["note"] = (cur["note"] + " " + re.sub(r"^【(点睛|点评|方法|归纳|规律)】\s*", "", text)).strip()
            continue
        if b.kind == "answer":
            if cur is None:
                cur = _new_item()
                items.append(cur)
            mm = re.match(r"^[【\[]?\s*(答案|解析|详解|点睛|点评|分析)\s*[】\]]?\s*[:：]?\s*(.*)$",
                          text)
            if mm:
                kind, body = mm.group(1), mm.group(2)
                if kind == "答案":
                    cur["answer"] = (cur["answer"] + " " + body).strip()
                elif kind in ("点睛", "点评"):
                    cur["note"] = (cur["note"] + " " + body).strip()
                else:
                    cur["analysis"] = (cur["analysis"] + " " + body).strip()
            else:
                cur["analysis"] = (cur["analysis"] + " " + text).strip()
            continue
        if b.kind == "note":
            src = re.sub(r"^来源[:：]\s*", "", text)
            if cur is not None:
                cur["source"] = src
            else:
                bullets.append(text)
            continue
        head = re.match(r"^(例\s*\d+|探究[一二三四五六七八九十\d]*|变式[一二三四五六七八九十\d]*)"
                        r"[：:、.．\s]*(.*)$", text)
        if head and len(text) < 400:
            cur = _new_item(head.group(1).replace(" ", ""), head.group(2).strip())
            items.append(cur)
            continue
        if quiz and re.match(r"^\d{1,2}[.．、]", text):
            cur = _new_item("", text)
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
    return bullets, items, tables, images, tips


def _paged(title: str, items: list[str], per: int = 7, limit: int = 400) -> list[list[str]]:
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
        if re.search(r"学习目标|小结|归纳", s.title or ""):
            pts += [b.text.strip() for b in s.blocks if b.kind in ("bullet", "num", "text")]
    if not pts:
        pts = [s.title for s in d.sections if s.title][:5]
    return pts[:6]


# ---------------------------------------------------------------- builders
BIG = re.compile(r"^(课前预习|课内探究|当堂检测|课堂检测|课后作业|课后巩固|学习目标|"
                 r"重点难点|考情分析|参考答案|答案)")


def _cover_index(hour: str) -> str:
    m = re.search(r"\d+", hour or "")
    return f"第 {int(m.group()):02d} 讲" if m else ""


def build_xuean_deck(md: str, out_path: str | Path, *, theme: str = "blue",
                     school: str = "", author: str = "") -> Path:
    d = parse(md)
    deck = Deck(theme=theme, footer=d.title, school=school)
    subject, grade = d.meta.get("学科", ""), d.meta.get("年级", "")
    hour = d.meta.get("课时", "")
    sub = f"{grade}{subject}" if (grade or subject) else ""
    deck.cover(d.title or "学案", subtitle=sub, meta=school or author, tag="学案",
               index=_cover_index(hour))

    for idx, sec in enumerate([x for x in d.sections if x.title], 1):
        title = _clean_heading(sec.title)
        is_quiz = bool(re.search(r"当堂检测|课堂检测|检测|练习|作业|自测", sec.title))
        bullets, items, tables, images, tips = _split_section(sec, quiz=is_quiz)

        if BIG.match(sec.title):
            deck.divider(title, index=f"{idx:02d}",
                         spine="导入" if "目标" in title else "")

        if bullets:
            for k, chunk in enumerate(_paged(title, bullets)):
                if not chunk:
                    continue
                deck.bullets(title + ("（续）" if k else ""), chunk,
                             kicker=("学习目标" if "目标" in title else title[:10]),
                             band=bool(items or tables),
                             spine="" if (items or tables) else "知识梳理")
        for it in items:
            chunks = _chunks(it["stem"], 260) if it["stem"] else []
            for pi, part in enumerate(chunks):
                last = pi == len(chunks) - 1
                deck.question(title, part,
                              it["options"] if last else None,
                              source=it["source"], note=it["note"] if last else "",
                              index=it["name"], tag="练习" if is_quiz else "例题")
            if it["answer"] or it["analysis"]:
                deck.answer(title, it["answer"], it["analysis"], index=it["name"],
                            tag="练习" if is_quiz else "例题")
        for rows in tables:
            deck.table(title, rows, kicker=title[:10])
        for src, cap in images:
            deck.image(title, src, cap, kicker=title[:10], band=True)
        for tp in tips:
            deck.tip("特别提醒", tp)

    deck.closing("本节课小结", _summary_points(d))
    return deck.save(out_path)


def build_paper_deck(md: str, out_path: str | Path, *, theme: str = "ink",
                     school: str = "", mode: str = "review") -> Path:
    """试卷 → 讲评课件：每题题干页（留白给学生） + 答案解析页。"""
    d = parse(md)
    deck = Deck(theme=theme, footer=d.title, school=school)
    meta = "　".join(x for x in [d.meta.get("时间"), d.meta.get("满分")] if x)
    deck.cover(d.title or "试卷讲评", subtitle="试卷讲评" if mode == "review" else "试卷",
               meta=meta, tag="试卷讲评")

    questions = _group_questions(d)
    _merge_answer_section(d, questions)
    if questions:
        total = sum(int(q.get("score") or 0) for q in questions) or d.meta.get("满分", "—")
        deck.bullets("试卷结构", [
            f"共 {len(questions)} 题，满分 {total}，考试时间 {d.meta.get('时间', '—')}",
            "先看错题分布，再按题型归因：概念不清 / 审题失误 / 计算错误 / 表达不规范",
        ], kicker="总览")

    for qi, q in enumerate(questions, 1):
        head = f"第 {q['number']} 题" if q["number"] else f"第 {qi} 题"
        q["section"] = re.split(r"[:：]", q["section"] or "")[0].strip() or q["section"]
        stem_pages = _chunks(q["stem"], 300)
        for pi, part in enumerate(stem_pages):
            last = pi == len(stem_pages) - 1
            deck.question(head, part, q["options"] if last else None,
                          source=q.get("source", ""), note=q.get("note", "") if last else "",
                          tag=q["section"][:10])
        if q["answer"] or q["analysis"]:
            deck.answer(head, q["answer"], q["analysis"], tag=q["section"][:10])

    deck.closing("易错点回顾", [
        "审题：圈出限定词（正确的是 / 不正确的是 / 最多 / 至少）",
        "规范：写出依据公式，再代入数据，最后带单位",
        "复盘：把错题按知识点归档，三天后重做一遍",
    ])
    return deck.save(out_path)


NUM_RE = re.compile(r"^\s*(\d{1,2})\s*[.．、]\s*(.+)$")
OPT_RE = re.compile(r"^\s*([A-DＡ-Ｄ])\s*[.．、]\s*(.+)$")


def _group_questions(d: Doc) -> list[dict]:
    out: list[dict] = []
    cur = None
    section = ""
    for s in d.sections:
        if s.title:
            section = s.title
        if re.search(r"注意事项|考试说明|考生须知|评分标准|答题卡|考试要求|命题说明", section):
            continue
        if section and not re.search(r"题|部分|卷|[一二三四五六七八九十]、", section):
            continue
        for b in s.blocks:
            if b.kind in ("table", "image"):
                continue
            text = (b.text or "").strip()
            if not text:
                continue
            m = NUM_RE.match(text) if b.kind in ("text", "num") else None
            if m and len(text) < 500:
                cur = {"number": m.group(1), "stem": m.group(2), "options": [],
                       "answer": "", "analysis": "", "note": "", "source": "",
                       "section": section}
                out.append(cur)
                continue
            if cur is None:
                continue
            mo = OPT_RE.match(text)
            if mo and len(cur["options"]) < 4 and len(text) < 40:
                cur["options"].append(text)
                continue
            if b.kind == "answer":
                mm = re.match(r"^[【\[]?\s*(答案|解析|详解|点睛|点评)\s*[】\]]?\s*[:：]?\s*(.*)$",
                              text)
                if mm:
                    kind, body = mm.group(1), mm.group(2)
                    if kind == "答案":
                        cur["answer"] = (cur["answer"] + " " + body).strip()
                    elif kind in ("点睛", "点评"):
                        cur["note"] = (cur["note"] + " " + body).strip()
                    else:
                        cur["analysis"] = (cur["analysis"] + " " + body).strip()
                else:
                    cur["analysis"] = (cur["analysis"] + " " + text).strip()
                continue
            if text.startswith("来源"):
                cur["source"] = re.sub(r"^来源[:：]\s*", "", text)
                continue
            if cur["answer"] or cur["analysis"]:
                cur["analysis"] = (cur["analysis"] + " " + text).strip()
            else:
                cur["stem"] = (cur["stem"] + " " + text).strip()
    return [q for q in out if q["stem"]]


def _merge_answer_section(d: Doc, questions: list[dict]) -> None:
    """把「参考答案」栏目里的答案按题号补回题目。"""
    ans_sec = [s for s in d.sections
               if re.match(r"^(参考答案|答案与解析|答案|解析)", (s.title or "").strip())]
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
            # 【解析】4. … 这种把题号写在括号后面的，按里面的题号归属
            inner = re.match(r"^[【\[]?\s*(答案|解析|详解|点睛)\s*[】\]]?\s*(\d{1,2})\s*[.．、]\s*(.*)$",
                             text)
            if inner:
                pending = by_num.get(inner.group(2)) or pending
                text = f"{inner.group(1)}：{inner.group(3)}"
            mm = re.match(r"^[【\[]?\s*(答案|解析|详解|点睛)\s*[】\]]?\s*[:：]?\s*(.*)$", text)
            if mm:
                kind, body = mm.group(1), mm.group(2)
                if kind == "答案":
                    pending["answer"] = (pending["answer"] + " " + body).strip()
                elif kind == "点睛":
                    pending["note"] = (pending["note"] + " " + body).strip()
                else:
                    pending["analysis"] = (pending["analysis"] + " " + text).strip()
            elif not pending["answer"] and (len(text) <= 60 or not pending.get("options")):
                pending["answer"] = text
            else:
                pending["analysis"] = (pending["analysis"] + " " + text).strip()
