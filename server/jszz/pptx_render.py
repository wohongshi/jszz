"""16:9 课件渲染：学案 → 新授课 PPT，试卷 → 讲评 PPT。

版式来自 U 盘 17 份真实课件（每科 2–3 份：早读/常规课/习题讲评）的转图逐页目视
+ slide XML 全量统计（837 份语料的配色/字号/形状密度/页头形态）：
  * 16:9，白底；正文 24–28pt（真实课件 28pt 一统）
  * 红=答案/纠错/易错（六科通吃），蓝=标题/要点/解析，黄=高亮标签
  * 页头六式按学科取证：语文提示横条 / 数学细线框 / 英语青色块 / 物理红标签 /
    化学蓝胶囊 / 生物绿块栏目
  * 例题页：红「例N」标签 + 题源小字 + 选项网格 + 右侧红色批注框
  * 物理/化学页码方块导航（真实课件的页码按钮条）
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

from . import subject_style as ss
from .mdparse import Block, Doc, Section, parse, sec_key

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
    """六科各一套视觉语言：传 subject=学科名 自动选型，theme= 显式指定主题名覆盖。"""

    def __init__(self, theme: str = "", footer: str = "", school: str = "",
                 subject: str = ""):
        self.prs = Presentation()
        self.prs.slide_width, self.prs.slide_height = W, H
        self.s = ss.get(theme or subject)
        self.t = self.s              # 兼容旧调用（.main/.accent/... 都是 RGBColor）
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

    def _run(self, p, text, *, size, bold=False, color=None, cn=None, en=FONT_EN,
             underline=False):
        cn = ss.pick_font(cn or self.s.font_cn)
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
        """【…】→警示红加粗；{{…}}→下划线空格；**…**→红粗关键词；其余正文色。"""
        color = color or self.t.ink
        text = _nb(text)
        for part in re.split(r"(【[^】]{1,16}】|\{\{[^}]*\}\}|\*\*[^*]+\*\*)", text):
            if not part:
                continue
            if part.startswith("【") and part.endswith("】"):
                self._run(p, part, size=size, bold=True, color=self.t.warn)
            elif part.startswith("{{") and part.endswith("}}"):
                inner = part[2:-2].strip()
                self._run(p, "\u3000" * max(2, min(10, len(inner) + 1)),
                          size=size, color=color, underline=True)
            elif part.startswith("**") and part.endswith("**"):
                # **关键词**：真实资料的红粗强调（得分点，史轮2-24 取证）
                self._run(p, part[2:-2], size=size, bold=True, color=self.t.warn)
            else:
                self._run(p, part, size=size, color=color, bold=bold)

    # ------------------------------------------------------------ chrome
    def _chrome(self, slide, *, band: bool = True):
        # 真实课件页脚取证：物理/化学用页码方块导航条；英语用右下双胶囊标签
        if self.s.key in ("physics", "chemistry") and band:
            # 页码方块（当前页高亮蓝块）——总数未知，只放当前页 + �后退/前进占位，保持每页一致
            bw = Inches(0.34)
            yb = H - Inches(0.52)
            x0 = W - Inches(0.6) - bw * 3 - Inches(0.12)
            for i, lb in enumerate(["‹", str(self.n), "›"]):
                x = x0 + (bw + Inches(0.06)) * i
                cur = (i == 1)
                self._rect(slide, x, yb, bw, Inches(0.32),
                           self.s.accent if cur else _rgb("E4E9F0"),
                           shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.18)
                tb, tf, p = self._text(slide, x, yb + Inches(0.02), bw, Inches(0.3),
                                       size=10, align=PP_ALIGN.CENTER,
                                       anchor=MSO_ANCHOR.MIDDLE)
                self._run(p, lb, size=10, bold=cur,
                          color=self.s.white if cur else _rgb("6B7785"))
            tb, tf, p = self._text(slide, Inches(0.6), yb, Inches(8.2), Inches(0.3),
                                   size=10, anchor=MSO_ANCHOR.MIDDLE)
            school = "" if (self.school and self.school in (self.footer or "")) else self.school
            label = " · ".join(x for x in [school, self.footer] if x)
            if label:
                self._run(p, label, size=10, color=_rgb("8A94A6"))
            return

        if band:
            self._rect(slide, 0, H - Inches(0.13), W, Inches(0.13), self.t.accent)
        y = H - Inches(0.56)
        tb, tf, p = self._text(slide, Inches(0.6), y, Inches(8.2), Inches(0.3), size=10)
        school = "" if (self.school and self.school in (self.footer or "")) else self.school
        label = " · ".join(x for x in [school, self.footer] if x)
        if label:
            self._run(p, label, size=10, color=_rgb("8A94A6"))
        if self.s.key == "english":
            # 右下双胶囊标签（真实英语课件每页固定）
            x = W - Inches(3.55)
            for lb in ("核心词汇突破", "重点句型透视"):
                cw = Inches(1.62)
                self._rect(slide, x, y - Inches(0.04), cw, Inches(0.34),
                           _rgb("F2F4F8"), shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                           radius=0.4)
                tb3, tf3, p3 = self._text(slide, x, y - Inches(0.03), cw, Inches(0.32),
                                          size=9, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p3, lb, size=9, bold=True, color=_rgb("6B7280"))
                x += cw + Inches(0.12)
            tb2, tf2, p2 = self._text(slide, W - Inches(0.55), y, Inches(0.35),
                                      Inches(0.3), size=10, align=PP_ALIGN.RIGHT)
            self._run(p2, str(self.n), size=10, color=_rgb("8A94A6"))
            return
        tb2, tf2, p2 = self._text(slide, W - Inches(1.25), y, Inches(0.75), Inches(0.3),
                                  size=10, align=PP_ALIGN.RIGHT)
        self._run(p2, str(self.n), size=10, color=_rgb("8A94A6"))

    def _chip(self, slide, text: str, *, x=None, y=Inches(0.5), dark=False, max_chars=22):
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
        """六种页头（按真实课件取证）：banner 提示横条 / frame 细线框 / cyanblock 青色块
        / nav 红标签行 / capsule 蓝胶囊 / greenblock 绿块栏目。"""
        kind = self.s.header
        size = 28 if len(title) <= 18 else (24 if len(title) <= 26 else 21)
        ss.draw_deco(slide, self.s)
        # kicker 与标题互为子串时不再重复显示（避免"考点二 xxx"标两次）
        if kicker and (kicker.strip() in title.strip() or title.strip() in kicker.strip()):
            kicker = ""

        if kind == "banner":
            # 语文：顶部提示横条（红字提示 + 黄底高亮），下接大标题
            self._rect(slide, 0, 0, W, Inches(0.5), self.s.light)
            self._rect(slide, 0, Inches(0.5), W, Emu(12700), self.s.accent)
            tb0, tf0, p0 = self._text(slide, Inches(0.6), Inches(0.08), W - Inches(1.2),
                                      Inches(0.36), size=13)
            hint = kicker if (kicker and kicker.strip() != title.strip()) else self.footer
            self._run(p0, "温馨提示：", size=13, bold=True, color=self.s.warn)
            self._run(p0, (hint or "全体齐声朗诵，落实默写")[:30], size=13, color=self.s.ink)
            tb, tf, p = self._text(slide, Inches(0.6), Inches(0.66), W - Inches(3.2),
                                   Inches(0.78), size=size, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p, title, size=size, bold=True, color=self.s.ink,
                      cn=self.s.font_head)
            return

        if kind == "frame":
            # 数学：圆角细线框栏目名 + 红色小节标签
            self._rect(slide, Inches(0.55), Inches(0.42), W - Inches(1.1), Inches(0.82),
                       self.s.white, line=self.s.accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                       radius=0.16, line_w=1.25)
            tb, tf, p = self._text(slide, Inches(0.95), Inches(0.5), W - Inches(4.2),
                                   Inches(0.66), size=size, anchor=MSO_ANCHOR.MIDDLE)
            self._rich(p, title, size=size, color=self.s.main, bold=True)
            if kicker and kicker.strip() != title.strip():
                self._chip(slide, kicker, y=Inches(0.62))
            return

        if kind == "cyanblock":
            # 英语：青色块栏目标题（【课前背记】式），标题前置方括号
            tw = Inches(0.34 * (len(title) + 2) + 0.42)
            self._rect(slide, Inches(0.6), Inches(0.4), min(tw, W - Inches(3.6)),
                       Inches(0.62), self.s.accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                       radius=0.12)
            tb, tf, p = self._text(slide, Inches(0.6), Inches(0.42),
                                   min(tw, W - Inches(3.6)), Inches(0.58), size=min(size, 23),
                                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p, f"【{title}】", size=min(size, 23), bold=True, color=self.s.white)
            if kicker and kicker.strip() != title.strip():
                tb2, tf2, p2 = self._text(slide, Inches(0.6) + min(tw, W - Inches(3.6)) + Inches(0.2),
                                          Inches(0.5), Inches(4.2), Inches(0.44), size=13)
                self._run(p2, kicker, size=13, bold=True, color=self.s.warn)
            return

        if kind == "capsule":
            # 化学：蓝胶囊标题 + 红字讲次小标
            tw = Inches(0.33 * len(title) + 0.72)
            self._rect(slide, Inches(0.6), Inches(0.38), min(tw, W - Inches(3.8)),
                       Inches(0.64), self.s.accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                       radius=0.32)
            tb, tf, p = self._text(slide, Inches(0.6), Inches(0.4),
                                   min(tw, W - Inches(3.8)), Inches(0.6), size=min(size, 23),
                                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p, title, size=min(size, 23), bold=True, color=self.s.white)
            if kicker and kicker.strip() != title.strip():
                tb2, tf2, p2 = self._text(slide, Inches(0.6) + min(tw, W - Inches(3.8)) + Inches(0.24),
                                          Inches(0.5), Inches(4.4), Inches(0.44), size=13)
                self._run(p2, kicker, size=13, bold=True, color=self.s.warn)
            return

        if kind == "greenblock":
            # 生物：绿色块栏目名（从社会中来 / 练一练 式）+ 黄荧光提示
            tw = Inches(0.34 * len(title) + 0.6)
            self._rect(slide, Inches(0.6), Inches(0.4), min(tw, W - Inches(3.8)),
                       Inches(0.62), self.s.main, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                       radius=0.1)
            tb, tf, p = self._text(slide, Inches(0.6), Inches(0.42),
                                   min(tw, W - Inches(3.8)), Inches(0.58), size=min(size, 23),
                                   align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p, title, size=min(size, 23), bold=True, color=self.s.white)
            if kicker and kicker.strip() != title.strip():
                kw = Inches(0.17 * len(kicker) + 0.5)
                self._rect(slide, Inches(0.6) + min(tw, W - Inches(3.8)) + Inches(0.22),
                           Inches(0.5), kw, Inches(0.42), _rgb("FFF2CC"),
                           shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.28)
                tb2, tf2, p2 = self._text(slide, Inches(0.6) + min(tw, W - Inches(3.8)) + Inches(0.22),
                                          Inches(0.52), kw, Inches(0.38), size=12,
                                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
                self._run(p2, kicker, size=12, bold=True, color=_rgb("8A6D1B"))
            return

        if kind == "scrollbar":
            # 历史：宣纸横条 + 竖排卷轴条 + 宋体栏目名（史主2/史轮2 取证）
            self._rect(slide, 0, 0, W, Inches(0.62), self.s.light)
            self._rect(slide, 0, Inches(0.62), W, Emu(11430), self.s.main)
            self._rect(slide, Inches(0.5), Inches(0.06), Inches(0.16), Inches(0.5),
                       self.s.main)
            tb, tf, p = self._text(slide, Inches(0.85), Inches(0.06), W - Inches(3.4),
                                   Inches(0.5), size=size, anchor=MSO_ANCHOR.MIDDLE)
            self._rich(p, title, size=size, color=self.s.ink, bold=True)
            if kicker and kicker.strip() != title.strip():
                tb2, tf2, p2 = self._text(slide, W - Inches(3.3), Inches(0.1),
                                          Inches(2.7), Inches(0.42), size=13,
                                          align=PP_ALIGN.RIGHT)
                self._run(p2, kicker, size=13, bold=True, color=self.s.gold)
            return

        if kind == "skyline":
            # 地理：深蓝左块栏目名 + 右上斜切三角（地课2 封面取证）
            tw = Inches(0.33 * len(title) + 0.8)
            self._rect(slide, 0, 0, min(tw, W - Inches(4.2)), Inches(0.72), self.s.main)
            tb, tf, p = self._text(slide, Inches(0.34), Inches(0.05),
                                   min(tw, W - Inches(4.2)) - Inches(0.4), Inches(0.62),
                                   size=min(size, 24), anchor=MSO_ANCHOR.MIDDLE)
            self._run(p, title, size=min(size, 24), bold=True, color=self.s.white)
            tri = slide.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE,
                                         W - Inches(2.3), 0, Inches(2.3), Inches(0.72))
            tri.fill.solid()
            tri.fill.fore_color.rgb = self.s.accent
            tri.line.fill.background()
            tri.shadow.inherit = False
            if kicker and kicker.strip() != title.strip():
                kw = Inches(0.18 * len(kicker) + 0.5)
                self._rect(slide, min(tw, W - Inches(4.2)) + Inches(0.24), Inches(0.14),
                           kw, Inches(0.42), self.s.gold,
                           shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tb2, tf2, p2 = self._text(slide, min(tw, W - Inches(4.2)) + Inches(0.24),
                                          Inches(0.16), kw, Inches(0.38), size=12,
                                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
                self._run(p2, kicker, size=12, bold=True, color=self.s.white)
            return

        if kind == "redbar":
            # 政治：红粗竖条 + 栏目名 + 蓝细线（党政红蓝配色取证）
            self._rect(slide, Inches(0.55), Inches(0.34), Inches(0.14), Inches(0.74),
                       self.s.main)
            tb, tf, p = self._text(slide, Inches(0.88), Inches(0.34), W - Inches(3.6),
                                   Inches(0.74), size=size, anchor=MSO_ANCHOR.MIDDLE)
            self._rich(p, title, size=size, color=self.s.ink, bold=True)
            self._rect(slide, Inches(0.88), Inches(1.14), W - Inches(1.2), Emu(9525),
                       self.s.accent)
            if kicker and kicker.strip() != title.strip():
                self._chip(slide, kicker, y=Inches(0.48))
            return

        # nav（物理/通用）：标题左 + 红「标签」行 + 细分隔线（真实物理课件页头）
        tb, tf, p = self._text(slide, Inches(0.6), Inches(0.38), W - Inches(3.6),
                               Inches(0.72), size=size, anchor=MSO_ANCHOR.MIDDLE)
        self._rich(p, title, size=size, color=self.s.main, bold=True)
        if kicker and kicker.strip() != title.strip():
            self._chip(slide, kicker, y=Inches(0.52))
        self._rect(slide, Inches(0.6), Inches(1.22), W - Inches(1.2), Emu(9525),
                   _rgb("D8DEE8"))

    def _header_band(self, slide, title: str, *, tags: list[str] | None = None):
        """例题/解析页页头（真实课件取证：白底 + 学科色标题 + 红标签 + 细线，不做深色通栏）。"""
        self._rect(slide, 0, 0, W, Inches(0.09), self.t.main)
        size = 26 if len(title) <= 20 else (22 if len(title) <= 28 else 19)
        tb, tf, p = self._text(slide, Inches(0.6), Inches(0.32), W - Inches(4.6),
                               Inches(0.62), size=size, anchor=MSO_ANCHOR.MIDDLE)
        self._run(p, title, size=size, bold=True, color=self.t.main)
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
                self._rect(slide, x_cursor, Inches(0.42), cw, Inches(0.4),
                           self.t.warn, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tb2, tf2, p2 = self._text(slide, x_cursor, Inches(0.44), cw, Inches(0.36),
                                          size=11, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p2, tg, size=11, bold=True, color=self.t.white)
                x_cursor -= Inches(0.1)
        self._rect(slide, Inches(0.6), Inches(1.04), W - Inches(1.2), Emu(9525),
                   _rgb("D8DEE8"))

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
              hint: str | None = None, index: str = ""):
        """六种封面版式：paper / tags / aims / orbit / photo / radial（+ split 兜底）。"""
        s = self._blank()
        kind = self.s.cover
        title = _soft_wrap(title, 13)
        longest = max(len(x) for x in title.split("\v"))
        tsize = int(min(42, max(20, (W * 0.47) / 914400.0 * 72.0 / max(longest, 1) * 0.92)))

        if kind == "paper":
            # 语文：象牙白纸感 + 顶部提示横条 + 黄底篇目标签 + 红字讲次（早读课件取证）
            self._rect(s, 0, 0, W, H, self.s.paper)
            self._rect(s, 0, 0, W, Inches(0.5), self.s.light)
            self._rect(s, 0, Inches(0.5), W, Emu(12700), self.s.accent)
            tb0, tf0, p0 = self._text(s, Inches(0.62), Inches(0.08), W - Inches(1.24),
                                      Inches(0.36), size=13)
            msg = "全体齐声朗诵，落实背诵默写" if hint is None else hint
            if msg:
                self._run(p0, "温馨提示：", size=13, bold=True, color=self.s.warn)
                self._run(p0, msg[:30], size=13, color=self.s.ink)
            if tag:
                tw = Inches(0.32 * len(tag) + 0.5)
                self._rect(s, Inches(0.95), Inches(1.2), tw, Inches(0.52),
                           _rgb("FFE699"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
                tb1, tf1, p1 = self._text(s, Inches(0.95), Inches(1.23), tw, Inches(0.46),
                                          size=15, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p1, tag, size=15, bold=True, color=self.s.ink)
            tb, tf, p = self._text(s, Inches(0.95), Inches(2.0), W * 0.66, Inches(2.3),
                                   size=tsize, line_spacing=1.28)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=tsize, bold=True, color=self.s.ink,
                          cn=self.s.font_head)
            self._rect(s, Inches(0.98), Inches(4.42), Inches(1.1), Emu(28575), self.s.accent)
            if index:
                tb2, tf2, p2 = self._text(s, Inches(0.95), Inches(4.62), Inches(6.2),
                                          Inches(0.55), size=18)
                self._run(p2, index, size=18, bold=True, color=self.s.warn)
            if subtitle:
                tb3, tf3, p3 = self._text(s, Inches(0.95), Inches(5.22), Inches(6.4),
                                          Inches(0.6), size=19)
                self._run(p3, subtitle, size=19, color=_rgb("4A4034"))
            if meta:
                tb4, tf4, p4 = self._text(s, Inches(0.95), Inches(5.92), Inches(6.4),
                                          Inches(0.5), size=13)
                self._run(p4, meta, size=13, color=_rgb("6B6257"))
            # 班级/主备人/日期 三栏（真实学案封面信息栏）
            fy = Inches(6.45)
            for i, lb in enumerate(["班级：", "主备人：", "日期："]):
                fx = Inches(0.95) + Inches(2.35) * i
                tb5, tf5, p5 = self._text(s, fx, fy, Inches(2.1), Inches(0.34), size=12)
                self._run(p5, lb, size=12, color=self.s.accent)
                self._run(p5, "　" * 6, size=12, color=self.s.ink, underline=True)

        elif kind == "tags":
            # 数学：白底 + 悬挂式标题牌（细线框）+ 彩色圆角标签组（真实课件封面取证）
            self._rect(s, 0, 0, W, H, self.s.white)
            ss.draw_deco(s, self.s)
            self._rect(s, 0, 0, W, Inches(0.1), self.s.main)
            # 标题牌：圆角细线框
            self._rect(s, Inches(1.4), Inches(1.35), W - Inches(2.8), Inches(2.5),
                       self.s.white, line=self.s.accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                       radius=0.06, line_w=1.5)
            if index:
                tb, tf, p = self._text(s, Inches(1.4), Inches(1.62), W - Inches(2.8),
                                       Inches(0.55), size=17, align=PP_ALIGN.CENTER)
                self._run(p, index, size=17, bold=True, color=self.s.accent)
            tb, tf, p = self._text(s, Inches(1.7), Inches(2.25), W - Inches(3.4),
                                   Inches(1.5), size=tsize, align=PP_ALIGN.CENTER,
                                   line_spacing=1.22, anchor=MSO_ANCHOR.MIDDLE)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=tsize, bold=True, color=self.s.main)
            if subtitle:
                tb2, tf2, p2 = self._text(s, Inches(1.4), Inches(4.15), W - Inches(2.8),
                                          Inches(0.7), size=20, align=PP_ALIGN.CENTER)
                self._run(p2, subtitle, size=20, color=_rgb("4A5568"))
            # 彩色圆角标签组（5 个，真实课件封面的章节标签云）
            labels = ["知识梳理", "典型例题", "变式训练", "当堂检测", "课堂小结"]
            lw = Inches(1.72)
            total = lw * 5 + Inches(0.24) * 4
            x0 = (W - total) / 2
            for i, lb in enumerate(labels):
                x = x0 + (lw + Inches(0.24)) * i
                self._rect(s, x, Inches(5.05), lw, Inches(0.56), self.s.white,
                           line=[self.s.main, self.s.accent, self.s.warn,
                                 self.s.gold, _rgb("2E8B57")][i],
                           shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.4, line_w=1.5)
                tb3, tf3, p3 = self._text(s, x, Inches(5.08), lw, Inches(0.5), size=13,
                                          align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
                self._run(p3, lb, size=13, bold=True,
                          color=[self.s.main, self.s.accent, self.s.warn,
                                 self.s.gold, _rgb("2E8B57")][i])
            if meta:
                tb4, tf4, p4 = self._text(s, Inches(1.4), Inches(6.1), W - Inches(2.8),
                                          Inches(0.5), size=13, align=PP_ALIGN.CENTER)
                self._run(p4, meta, size=13, color=_rgb("7A8698"))

        elif kind == "aims":
            # 英语：白底 + 红色大标题 + 黄底 Learning aims 块 + 青色标签（真实课件封面取证）
            self._rect(s, 0, 0, W, H, self.s.white)
            ss.draw_deco(s, self.s)
            self._rect(s, 0, 0, W, Inches(0.12), self.s.accent)
            if index:
                tb, tf, p = self._text(s, Inches(0.95), Inches(0.95), Inches(5.4),
                                       Inches(0.55), size=16)
                self._run(p, index, size=16, bold=True, color=self.s.accent)
            tb, tf, p = self._text(s, Inches(0.95), Inches(1.6), W * 0.8, Inches(1.9),
                                   size=tsize, line_spacing=1.2)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=tsize, bold=True, color=self.s.warn)
            # 黄底 Learning aims 块（真实课件封面最显眼的元素）
            self._rect(s, Inches(0.95), Inches(3.75), Inches(3.4), Inches(0.6),
                       _rgb("FFE699"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
            tb0, tf0, p0 = self._text(s, Inches(0.95), Inches(3.78), Inches(3.4),
                                      Inches(0.54), size=17, align=PP_ALIGN.CENTER,
                                      anchor=MSO_ANCHOR.MIDDLE)
            self._run(p0, "Learning aims:", size=17, bold=True, color=self.s.ink)
            if subtitle:
                # 学习目标条目：蓝色关键词式
                tb2, tf2, p2 = self._text(s, Inches(1.15), Inches(4.6), W - Inches(2.6),
                                          Inches(1.5), size=18, line_spacing=1.55)
                for i, line in enumerate(subtitle.split("\v")):
                    pp = self._p(tf2, first=(i == 0))
                    _indent(pp, 0.42, -0.42)
                    self._run(pp, f"{i + 1}. ", size=18, bold=True, color=self.s.main)
                    self._run(pp, line, size=18, color=self.s.ink)
            if meta:
                tb3, tf3, p3 = self._text(s, Inches(1.15), Inches(6.15), W - Inches(2.6),
                                          Inches(0.5), size=13)
                self._run(p3, meta, size=13, color=_rgb("7A8698"))
            if tag:
                tw = Inches(0.22 * len(tag) + 0.5)
                self._rect(s, W - tw - Inches(0.8), Inches(0.92), tw, Inches(0.46),
                           self.s.accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tb4, tf4, p4 = self._text(s, W - tw - Inches(0.8), Inches(0.95), tw,
                                          Inches(0.4), size=13, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p4, tag, size=13, bold=True, color=self.s.white)

        elif kind == "split":
            self._rect(s, W * 0.5, 0, W * 0.5, H, self.s.main)
            self._rect(s, W * 0.5 - Inches(0.06), 0, Inches(0.1), H, self.s.accent)
            ss.draw_deco(s, self.s)
            if index:
                tb, tf, p = self._text(s, Inches(0.95), Inches(1.1), Inches(5.4),
                                       Inches(0.6), size=15)
                self._run(p, index, size=15, bold=True, color=self.s.accent)
            tb, tf, p = self._text(s, Inches(0.95), Inches(2.0), W * 0.44, Inches(2.4),
                                   size=int(tsize * 0.86), line_spacing=1.2)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=int(tsize * 0.86), bold=True, color=self.s.main)
            if subtitle:
                tb2, tf2, p2 = self._text(s, W * 0.5 + Inches(0.55), Inches(1.9),
                                          W * 0.42, Inches(2.4), size=19,
                                          line_spacing=1.5)
                self._run(p2, subtitle, size=19, color=self.s.white)
            if meta:
                tb3, tf3, p3 = self._text(s, W * 0.5 + Inches(0.55), Inches(4.5),
                                          W * 0.42, Inches(0.9), size=12)
                self._run(p3, meta, size=12, color=_rgb("C9CEE8"))
            if tag:
                tb4, tf4, p4 = self._text(s, W * 0.5 + Inches(0.6), Inches(5.9),
                                          W * 0.4, Inches(0.5), size=14,
                                          align=PP_ALIGN.RIGHT)
                self._run(p4, tag.upper(), size=14, bold=True, color=self.s.accent)

        elif kind == "orbit":
            # 物理：深蓝渐变底 + 轨道弧线（真实课件封面取证：深蓝 + 白色椭圆轨道）
            self._rect(s, 0, 0, W, H, self.s.main)
            self._rect(s, 0, 0, W, Inches(0.16), self.s.gold)
            # 轨道弧线（右下大椭圆 + 小圆点）
            ring = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(6.9), Inches(2.3),
                                      Inches(6.2), Inches(5.15))
            ring.fill.background()
            ring.line.color.rgb = _rgb("3A5575")
            ring.line.width = Pt(2.0)
            ring.shadow.inherit = False
            ring2 = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(7.9), Inches(3.35),
                                       Inches(4.4), Inches(3.1))
            ring2.fill.background()
            ring2.line.color.rgb = _rgb("2E4A72")
            ring2.line.width = Pt(1.25)
            ring2.shadow.inherit = False
            dot = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(10.75), Inches(3.05),
                                     Inches(0.24), Inches(0.24))
            dot.fill.solid()
            dot.fill.fore_color.rgb = self.s.gold
            dot.line.fill.background()
            dot.shadow.inherit = False
            if index:
                tb, tf, p = self._text(s, Inches(0.9), Inches(1.05), Inches(5), Inches(0.6),
                                       size=16)
                self._run(p, index, size=16, bold=True, color=self.s.gold)
            tb, tf, p = self._text(s, Inches(0.9), Inches(1.95), W - Inches(5.2),
                                   Inches(2.1), size=tsize, line_spacing=1.2)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=tsize, bold=True, color=self.s.white)
            self._rect(s, Inches(0.93), Inches(4.15), Inches(1.1), Emu(28575), self.s.gold)
            if subtitle:
                tb2, tf2, p2 = self._text(s, Inches(0.95), Inches(4.45), W - Inches(5.2),
                                          Inches(0.8), size=19)
                self._run(p2, subtitle, size=19, color=_rgb("C7D3E4"))
            if meta:
                tb3, tf3, p3 = self._text(s, Inches(0.95), Inches(5.35), W - Inches(5.2),
                                          Inches(0.5), size=12)
                self._run(p3, meta, size=12, color=_rgb("8FA2BC"))
            if tag:
                tw = Inches(0.2 * len(tag) + 0.6)
                self._rect(s, W - tw - Inches(0.8), Inches(0.82), tw, Inches(0.44),
                           self.s.white, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tb4, tf4, p4 = self._text(s, W - tw - Inches(0.8), Inches(0.85), tw,
                                          Inches(0.4), size=13, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p4, tag, size=13, bold=True, color=self.s.main)

        elif kind == "scroll":
            # 历史：宣纸底 + 卷轴竖条 + 居中墨字标题 + 朱红印章（史主2/史轮2 取证）
            self._rect(s, 0, 0, W, H, self.s.light)
            ss.draw_deco(s, self.s)
            for x in (Inches(0.55), W - Inches(0.75)):
                self._rect(s, x, 0, Inches(0.2), H, self.s.main)
            if tag:
                tw = Inches(0.32 * len(tag) + 0.5)
                self._rect(s, (W - tw) / 2, Inches(1.0), tw, Inches(0.46),
                           self.s.main, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.18)
                tb1, tf1, p1 = self._text(s, (W - tw) / 2, Inches(1.03), tw, Inches(0.4),
                                          size=14, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p1, tag, size=14, bold=True, color=self.s.light)
            tb, tf, p = self._text(s, Inches(1.6), Inches(1.95), W - Inches(3.2),
                                   Inches(2.2), size=tsize, align=PP_ALIGN.CENTER,
                                   line_spacing=1.24)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=tsize, bold=True, color=self.s.ink,
                          cn=self.s.font_head)
            self._rect(s, (W - Inches(1.6)) / 2, Inches(4.35), Inches(1.6), Emu(28575),
                       self.s.gold)
            if subtitle:
                tb2, tf2, p2 = self._text(s, Inches(1.6), Inches(4.62), W - Inches(3.2),
                                          Inches(0.6), size=19, align=PP_ALIGN.CENTER)
                self._run(p2, subtitle, size=19, color=self.s.main)
            if meta:
                tb3, tf3, p3 = self._text(s, Inches(1.6), Inches(5.35), W - Inches(3.2),
                                          Inches(0.5), size=12, align=PP_ALIGN.CENTER)
                self._run(p3, meta, size=12, color=_rgb("7A6A55"))

        elif kind == "atmos":
            # 地理：白底 + 右侧斜切云天块 + 底部通栏深蓝条 + 橙胶囊专题（地课2 封面取证）
            self._rect(s, 0, 0, W, H, self.s.white)
            tri = s.shapes.add_shape(MSO_SHAPE.RIGHT_TRIANGLE, W * 0.55, 0,
                                     W * 0.45, H - Inches(0.5))
            tri.fill.solid()
            tri.fill.fore_color.rgb = self.s.light
            tri.line.fill.background()
            tri.shadow.inherit = False
            tri.rotation = 0.0
            self._rect(s, 0, H - Inches(0.5), W, Inches(0.5), self.s.main)
            if tag:
                tw = Inches(0.28 * len(tag) + 0.55)
                self._rect(s, Inches(0.95), Inches(0.85), tw, Inches(0.42),
                           self.s.gold, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tb1, tf1, p1 = self._text(s, Inches(0.95), Inches(0.88), tw, Inches(0.36),
                                          size=13, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p1, tag, size=13, bold=True, color=self.s.white)
            tb, tf, p = self._text(s, Inches(0.9), Inches(1.6), W * 0.5, Inches(1.6),
                                   size=tsize, line_spacing=1.2)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=tsize, bold=True, color=self.s.main,
                          cn=self.s.font_head)
            if subtitle:
                tb2, tf2, p2 = self._text(s, Inches(0.95), Inches(3.35), W * 0.5,
                                          Inches(0.7), size=19)
                self._run(p2, subtitle, size=19, color=_rgb("5A6B7B"))
            if meta:
                tb3, tf3, p3 = self._text(s, Inches(0.95), Inches(4.3), W * 0.5,
                                          Inches(0.5), size=13)
                self._run(p3, meta, size=13, color=_rgb("7A8896"))

        elif kind == "podium":
            # 政治：左侧红粗竖条 + 顶部深蓝条 + 金色细线（党旗红/政务蓝/金 取证）
            self._rect(s, 0, 0, W, H, self.s.light)
            self._rect(s, 0, 0, W, Inches(0.14), self.s.accent)
            self._rect(s, Inches(0.85), Inches(0.9), Inches(0.28), Inches(3.6),
                       self.s.main)
            if tag:
                tw = Inches(0.3 * len(tag) + 0.55)
                self._rect(s, W - tw - Inches(0.8), Inches(0.62), tw, Inches(0.42),
                           self.s.accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tb1, tf1, p1 = self._text(s, W - tw - Inches(0.8), Inches(0.65), tw,
                                          Inches(0.36), size=13, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p1, tag, size=13, bold=True, color=self.s.white)
            tb, tf, p = self._text(s, Inches(1.45), Inches(1.45), W - Inches(2.6),
                                   Inches(2.1), size=tsize, line_spacing=1.22)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=tsize, bold=True, color=self.s.ink,
                          cn=self.s.font_head)
            self._rect(s, Inches(1.48), Inches(3.75), Inches(2.2), Emu(28575), self.s.gold)
            if subtitle:
                tb2, tf2, p2 = self._text(s, Inches(1.45), Inches(3.72), W - Inches(2.6),
                                          Inches(0.7), size=19)
                self._run(p2, subtitle, size=19, color=self.s.main)
            if meta:
                tb3, tf3, p3 = self._text(s, Inches(1.45), Inches(4.85), W - Inches(2.6),
                                          Inches(0.5), size=12)
                self._run(p3, meta, size=12, color=_rgb("7A6A6A"))

        elif kind == "photo":
            # 化学：浅灰白旋涡 + 深蓝粗标题（左）+ 圆角实验卡片（右）（真实课件封面取证）
            self._rect(s, 0, 0, W, H, self.s.light)
            ss.draw_deco(s, self.s)
            self._rect(s, 0, 0, W, Inches(0.12), self.s.main)
            # 右侧圆角卡片（模拟实物照片卡：青灰渐层 + 器皿线条示意）
            cx, cy, cw, chh = W * 0.63, Inches(1.6), Inches(4.15), Inches(3.9)
            self._rect(s, cx, cy, cw, chh, _rgb("DCE6EE"),
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
            self._rect(s, cx, cy + chh - Inches(0.55), cw, Inches(0.55),
                       _rgb("C7D4E0"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.08)
            # 试管示意（三根圆角竖条）
            for i, col in enumerate([_rgb("7FB3D9"), _rgb("E8A87C"), _rgb("82C9A6")]):
                tx = cx + Inches(0.95 + 1.05 * i)
                tube = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, tx,
                                          cy + Inches(0.9), Inches(0.52), Inches(2.2))
                tube.fill.solid()
                tube.fill.fore_color.rgb = col
                tube.line.color.rgb = _rgb("9FB2C4")
                tube.line.width = Pt(1.0)
                tube.shadow.inherit = False
            if index:
                tb, tf, p = self._text(s, Inches(0.9), Inches(1.15), Inches(5.4),
                                       Inches(0.55), size=15)
                self._run(p, index, size=15, bold=True, color=self.s.warn)
            tb, tf, p = self._text(s, Inches(0.9), Inches(1.8), W * 0.56, Inches(2.4),
                                   size=tsize, line_spacing=1.2)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=tsize, bold=True, color=self.s.main)
            self._rect(s, Inches(0.93), Inches(3.95), Inches(1.2), Inches(0.09),
                       self.s.accent)
            if subtitle:
                tb2, tf2, p2 = self._text(s, Inches(0.93), Inches(4.25), W * 0.55,
                                          Inches(0.8), size=19)
                self._run(p2, subtitle, size=19, color=_rgb("2F4F5C"))
            if meta:
                tb3, tf3, p3 = self._text(s, Inches(0.93), Inches(5.2), W * 0.55,
                                          Inches(0.5), size=12)
                self._run(p3, meta, size=12, color=_rgb("5E808C"))
            if tag:
                tw = Inches(0.2 * len(tag) + 0.5)
                self._rect(s, Inches(0.93), H - Inches(1.25), tw, Inches(0.44),
                           self.s.main, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tb4, tf4, p4 = self._text(s, Inches(0.93), H - Inches(1.23), tw,
                                          Inches(0.4), size=13, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p4, tag, size=13, bold=True, color=self.s.white)

        else:  # radial
            # 生物：白底 + 绿色块标题 + 放射概念图（真实课件封面取证：绿块白字 + 彩色标签放射）
            self._rect(s, 0, 0, W, H, self.s.white)
            ss.draw_deco(s, self.s)
            self._rect(s, 0, 0, W, Inches(0.14), self.s.main)
            self._rect(s, 0, H - Inches(0.14), W, Inches(0.14), self.s.accent)
            # 右侧放射概念图：中心绿块 + 4 彩色小标签
            hub_x, hub_y = W * 0.72, Inches(3.3)
            self._rect(s, hub_x - Inches(0.85), hub_y - Inches(0.32), Inches(1.7),
                       Inches(0.64), self.s.main, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                       radius=0.18)
            tbh, tfh, ph = self._text(s, hub_x - Inches(0.85), hub_y - Inches(0.3),
                                      Inches(1.7), Inches(0.6), size=15,
                                      align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._run(ph, self.s.label, size=15, bold=True, color=self.s.white)
            for i, (lb, col, ang) in enumerate([
                    ("结构", self.s.accent, 210), ("功能", _rgb("5B9BD5"), 150),
                    ("特点", _rgb("E8A87C"), 270), ("应用", _rgb("B57EDC"), 330)]):
                rad = math.radians(ang)
                lx = int(hub_x + Inches(2.35) * math.cos(rad) - Inches(0.55))
                ly = int(hub_y + Inches(1.95) * math.sin(rad) - Inches(0.26))
                self._rect(s, lx, ly, Inches(1.1), Inches(0.52), col,
                           shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tbl_, tfl, pl = self._text(s, lx, ly + Inches(0.02), Inches(1.1),
                                           Inches(0.48), size=12, align=PP_ALIGN.CENTER,
                                           anchor=MSO_ANCHOR.MIDDLE)
                self._run(pl, lb, size=12, bold=True, color=self.s.white)
                # 连线
                ln = s.shapes.add_connector(1, int(hub_x), int(hub_y),
                                            int(lx + Inches(0.55)), int(ly + Inches(0.26)))
                ln.line.color.rgb = _rgb("C9D6CC")
                ln.line.width = Pt(1.0)
            if index:
                tb, tf, p = self._text(s, Inches(0.95), Inches(1.0), Inches(6), Inches(0.6),
                                       size=15)
                self._run(p, index, size=15, bold=True, color=self.s.accent)
            tb, tf, p = self._text(s, Inches(0.95), Inches(1.8), W * 0.5, Inches(2.4),
                                   size=int(tsize * 0.94), line_spacing=1.2)
            for i, line in enumerate(title.split("\v")):
                if i:
                    p.add_line_break()
                self._run(p, line, size=int(tsize * 0.94), bold=True, color=self.s.main)
            if subtitle:
                # 黄荧光条（真实生物课件的「必背」高亮）
                self._rect(s, Inches(0.98), Inches(4.35), W * 0.48, Inches(0.62),
                           _rgb("FFF2CC"))
                tb2, tf2, p2 = self._text(s, Inches(1.12), Inches(4.38), W * 0.45,
                                          Inches(0.56), size=17, anchor=MSO_ANCHOR.MIDDLE)
                self._run(p2, subtitle, size=17, bold=True, color=self.s.ink)
            if meta:
                tb3, tf3, p3 = self._text(s, Inches(0.98), Inches(5.35), W * 0.48,
                                          Inches(0.5), size=12)
                self._run(p3, meta, size=12, color=_rgb("6B8C85"))
            if tag:
                tw = Inches(0.2 * len(tag) + 0.5)
                self._rect(s, Inches(0.98), H - Inches(1.2), tw, Inches(0.44),
                           self.s.main, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
                tb4, tf4, p4 = self._text(s, Inches(0.98), H - Inches(1.18), tw,
                                          Inches(0.4), size=13, align=PP_ALIGN.CENTER,
                                          anchor=MSO_ANCHOR.MIDDLE)
                self._run(p4, tag, size=13, bold=True, color=self.s.white)
        return s

    def divider(self, title: str, index: str = "", subtitle: str = "", spine: str = ""):
        s = self._blank()
        dark = self.s.key in ("physics", "chemistry")   # 真实课件：物理/化学有深蓝章节页
        if dark:
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
        else:
            # 白底章节页（语数英生物真实课件没有满版深色过渡页）
            self._rect(s, 0, 0, W, H, self.t.white)
            self._rect(s, 0, 0, W, Inches(0.12), self.t.main)
            self._rect(s, 0, H - Inches(0.12), W, Inches(0.12), self.t.accent)
            self._rect(s, Inches(0.85), Inches(2.55), Inches(0.14), Inches(1.5),
                       self.t.accent)
            if index:
                tb, tf, p = self._text(s, Inches(1.25), Inches(1.35), Inches(4), Inches(1.6),
                                       size=88)
                self._run(p, index, size=88, bold=True, color=self.t.light)
            tb2, tf2, p2 = self._text(s, Inches(1.25), Inches(2.6), W - Inches(2.5),
                                      Inches(1.1), size=40)
            self._run(p2, title, size=40, bold=True, color=self.t.main)
            if subtitle:
                tb3, tf3, p3 = self._text(s, Inches(1.32), Inches(3.85), W - Inches(2.5),
                                          Inches(0.8), size=18)
                self._run(p3, subtitle, size=18, color=_rgb("5A657A"))
        if spine:
            # 竖排标签放不稳（LibreOffice 裁字），改成小竖条 + 横排小标签
            bar_col = self.t.gold if dark else self.t.accent
            self._rect(s, Inches(0.85), Inches(2.2), Inches(0.42), Inches(0.95),
                       bar_col, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.2)
            tb4, tf4, p4 = self._text(s, Inches(1.4), Inches(2.2), Inches(2.2),
                                      Inches(0.42), size=15)
            self._run(p4, spine, size=15, bold=True,
                      color=self.t.white if dark else self.t.accent)
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
        if self.s.key == "chinese":
            spacing = max(spacing, 14)
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
        row_h = int(min(Inches(0.52), (H - Inches(2.3)) / max(1, len(rows))))
        gt = s.shapes.add_table(len(rows), cols, Inches(0.8), Inches(1.5),
                                W - Inches(1.6), int(row_h * len(rows))).table
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
                # 两列表：左列=条目名，蓝色加粗居中（史轮2-24 答案表格取证）
                entry_col = (ri > 0 and cols == 2 and ci == 0 and len(val) <= 12)
                if entry_col:
                    color = self.t.main
                # 走 _rich：cell 里的 **关键词**（红粗）和 {{挖空}}（下划线）都要生效
                self._rich(p, val, size=14 if ri else 15, color=color,
                           bold=(ri == 0 or entry_col))
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
            # 真实课件：红「例N」圆角标签（六科通吃）
            self._rect(s, left, top, bw, Inches(0.44), self.t.warn,
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
            sents = _sentences(note)
            hgt = min(H - top - Inches(0.7),
                      Inches(0.78 + 0.34 * min(len(sents), 6)))
            self._rect(s, nx, top - Inches(0.12), note_w, hgt, _rgb("FDFBF7"),
                       line=_rgb("E5C9A8"), shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
            self._rect(s, nx, top - Inches(0.12), Inches(0.07), hgt, self.t.warn)
            tb2, tf2, p2 = self._text(s, nx + Inches(0.2), top, note_w - Inches(0.36),
                                      hgt - Inches(0.2), size=13, line_spacing=1.32)
            self._run(p2, "判定要点", size=13, bold=True, color=self.t.warn)
            for i, line in enumerate(_sentences(note)):
                pp = self._p(tf2, space_before=4)
                _indent(pp, 0.32, -0.32)
                self._run(pp, f"{i + 1}. ", size=12, bold=True, color=self.t.warn)
                self._rich(pp, line, size=12, color=_rgb("6B2B2B"))
        self._chrome(s)
        return s

    def _option_grid(self, slide, options: list[str], left, top, width):
        """4 选项：够短就 4 列，否则 2×2，长选项竖排单列；每格悬挂缩进，字母加粗着色。"""
        def cw(t):
            return sum(1.0 if ord(c) > 127 else 0.55 for c in t)

        n = len(options)
        body_ws = [cw(re.sub(r"^\s*[A-DＡ-Ｄ]\s*[.．、]\s*", "", o).strip()) for o in options]
        maxw = max(body_ws) if body_ws else 0
        # 按可用格宽反推列数：文本宽 + 悬挂缩进必须放得下，否则减列
        per = 1
        for cand, sz in ((4, 19), (2, 18), (1, 17)):
            if n < cand:
                continue
            cell_w = width / cand - Inches(0.12)
            needed = Inches(maxw * sz / 72.0 * 1.02 + 0.42)
            if needed <= cell_w:
                per = cand
                break
        rows = [options[i:i + per] for i in range(0, n, per)]
        size = 19 if per == 4 else (18 if per == 2 else 17)
        cell_h = Inches(0.6 if per >= 2 else 0.52)
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
            # 真实课件：解析页左上是蓝色「解析」竖标/蓝标签
            self._rect(s, Inches(0.6), y, bw, Inches(0.42), self.t.accent,
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.25)
            tb, tf, p = self._text(s, Inches(0.6), y + Inches(0.02), bw, Inches(0.38),
                                   size=14, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p, index, size=14, bold=True, color=self.t.white)
            y += Inches(0.6)
        if answer:
            asize = 22 if len(answer) <= 36 else (19 if len(answer) <= 90 else 16)
            ah = Inches(0.62 if len(answer) <= 36 else (1.05 if len(answer) <= 90 else 1.6))
            # 真实课件：红=答案（六科通吃）
            self._rect(s, Inches(0.6), y, Inches(0.07), ah, self.t.warn)
            tb, tf, p = self._text(s, Inches(0.8), y, W - Inches(1.4), ah, size=asize,
                                   line_spacing=1.35)
            self._run(p, "答案：", size=asize, bold=True, color=self.t.warn)
            self._rich(p, answer, size=asize, color=self.t.warn)
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

    # ------------------------------------------------ 学科专属版式
    def template(self, title: str, rows: list[str], *, mode: str = "",
                 kicker: str = "答题模板"):
        """历史/政治：答题模板页——顶部模式行 + 青蓝色块内「标签：×××」行（史主2-08 取证）。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        y = Inches(1.5)
        if mode:
            tb, tf, p = self._text(s, Inches(0.9), y, W - Inches(1.8), Inches(0.5), size=19)
            self._run(p, "答题模式：", size=19, bold=True, color=self.t.ink)
            self._rich(p, mode, size=19, bold=True)
            y += Inches(0.72)
        bh = min(Inches(0.5) * max(3, len(rows)), H - y - Inches(1.1))
        self._rect(s, Inches(0.65), y, W - Inches(1.3), bh, _rgb("00B0F0"))
        tb2, tf2, p2 = self._text(s, Inches(1.05), y + Inches(0.12), W - Inches(2.1),
                                  bh - Inches(0.24), size=15, line_spacing=1.5)
        for i, row in enumerate(rows):
            if i:
                p2.add_line_break()
            self._run(p2, row, size=15, bold=True, color=_rgb("111111"))
        self._chrome(s)
        return s

    def quote(self, title: str, text: str, *, source: str = "", note: str = "",
              kicker: str = "原文"):
        """语文：引文页——大段原文居中，右侧朱砂竖条，出处小字。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        self._rect(s, Inches(0.9), Inches(1.6), W - Inches(1.8), H - Inches(2.9),
                   self.s.light, shape=MSO_SHAPE.RECTANGLE)
        self._rect(s, Inches(0.9), Inches(1.6), Inches(0.08), H - Inches(2.9),
                   self.s.accent)
        size = 26 if len(text) < 90 else (22 if len(text) < 180 else 19)
        tb, tf, p = self._text(s, Inches(1.35), Inches(1.75), W - Inches(2.9),
                               H - Inches(3.2), size=size, line_spacing=1.85,
                               anchor=MSO_ANCHOR.MIDDLE)
        tf.word_wrap = True
        p.paragraph_format if False else None
        tb.width = W - Inches(2.9)
        self._rich(p, text, size=size, color=self.s.ink)
        if source:
            tb2, tf2, p2 = self._text(s, Inches(1.3), H - Inches(1.55), W - Inches(3.4),
                                      Inches(0.4), size=13, align=PP_ALIGN.RIGHT)
            self._run(p2, f"—— {source}", size=13, color=_rgb("6B6257"),
                      cn=self.s.font_head)
        if note:
            tb3, tf3, p3 = self._text(s, Inches(1.3), H - Inches(1.1), W - Inches(2.7),
                                      Inches(0.4), size=12)
            self._run(p3, note, size=12, color=self.s.accent)
        self._chrome(s)
        return s

    def mnemonic(self, title: str, text: str, *, kicker: str = "口诀"):
        """口诀/记忆术页：居中大字对仗短句 + 强调色底，朗朗上口便于记忆。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        self._rect(s, Inches(0.9), Inches(1.7), W - Inches(1.8), H - Inches(3.1),
                   self.s.light, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        # 对仗口诀按「，」断行，居中竖排节奏
        segs = [x.strip() for x in re.split(r"[，,；;]", text) if x.strip()]
        size = 34 if len(text) < 30 else (28 if len(text) < 60 else 23)
        tb, tf, p = self._text(s, Inches(1.3), Inches(1.95), W - Inches(2.6),
                               H - Inches(3.6), size=size,
                               align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                               line_spacing=1.55)
        tf.word_wrap = True
        for i, seg in enumerate(segs):
            if i:
                p.add_line_break()
            self._run(p, seg, size=size, bold=True, color=self.s.main)
        self._chrome(s)
        return s

    def misconceptions(self, title: str, rows: list[tuple[str, str]], *,
                       kicker: str = "易错警示"):
        """规范表述双栏失分对照页：左=错误写法（不得分），右=规范表述（得分）。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        half = (W - Inches(1.9)) / 2
        top = Inches(1.5)
        heads = [("错误写法（不得分）", _rgb("C81E1E"), Inches(0.95)),
                 ("规范表述（得分）", _rgb("2E7D32"), Inches(0.95) + half + Inches(0.3))]
        for tt, col, x in heads:
            self._rect(s, x, top, half, Inches(0.5),
                       _rgb("FDECEC") if col == _rgb("C81E1E") else _rgb("EAF5EC"),
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.12)
            tb, tf, p = self._text(s, x + Inches(0.18), top + Inches(0.04),
                                   half - Inches(0.36), Inches(0.42), size=15)
            self._run(p, tt, size=15, bold=True, color=col)
        y = top + Inches(0.62)
        row_h = Inches(0.62)
        for wrong, right in rows[:6]:
            for i, (txt, col, x) in enumerate([
                    (wrong, _rgb("8A1B1B"), Inches(0.95)),
                    (right, _rgb("1B5E20"), Inches(0.95) + half + Inches(0.3))]):
                self._rect(s, x, y, half, row_h - Inches(0.08),
                           _rgb("FBF1F1") if i == 0 else _rgb("F1F7F2"),
                           shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.1)
                tb, tf, p = self._text(s, x + Inches(0.18), y + Inches(0.03),
                                       half - Inches(0.36), row_h - Inches(0.12),
                                       size=13, anchor=MSO_ANCHOR.MIDDLE)
                tf.word_wrap = True
                self._run(p, txt, size=13, color=col)
            y += row_h
        self._chrome(s)
        return s

    def scoring(self, title: str, rows: list[str], *, kicker: str = "采分点"):
        """采分点页：每行一个得分点「要点N：主题词：展开（X分）」，分行利于抓取。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        items = []
        for r in rows[:8]:
            m = re.match(r"^(要点\s*\d+[:：])", r)
            if m:
                items.append((m.group(1), r[m.end():].strip()))
            else:
                items.append(("", r))
        n = max(1, len(items))
        row_h = min(Inches(0.82), (H - Inches(2.4)) / n)
        y = Inches(1.55)
        for i, (lab, body) in enumerate(items):
            self._rect(s, Inches(0.9), y, W - Inches(1.8), row_h - Inches(0.1),
                       self.s.light if i % 2 == 0 else self.s.white,
                       line=_rgb("E2E6EE"), shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                       radius=0.08)
            if lab:
                self._rect(s, Inches(1.05), y + Inches(0.12), Inches(0.14),
                           row_h - Inches(0.34), self.s.accent)
                tb, tf, p = self._text(s, Inches(1.32), y + Inches(0.06),
                                       W - Inches(2.4), row_h - Inches(0.12),
                                       size=15, anchor=MSO_ANCHOR.MIDDLE)
                tf.word_wrap = True
                self._run(p, lab + " ", size=15, bold=True, color=self.s.accent)
                self._rich(p, body, size=14)
            else:
                tb, tf, p = self._text(s, Inches(1.15), y + Inches(0.06),
                                       W - Inches(2.3), row_h - Inches(0.12),
                                       size=14, anchor=MSO_ANCHOR.MIDDLE)
                tf.word_wrap = True
                self._rich(p, body, size=14)
            y += row_h
        self._chrome(s)
        return s

    def formula(self, title: str, formula: str, *, note: str = "", legs: list[str] | None = None,
                kicker: str = ""):
        """数学/物理/化学：公式页——公式居中大字，下面两栏「条件 / 结论」或分步说明。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        self._rect(s, Inches(0.8), Inches(1.5), W - Inches(1.6), Inches(1.75),
                   self.s.light, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.04)
        fsize = 40 if len(formula) < 26 else (32 if len(formula) < 48 else 26)
        tb, tf, p = self._text(s, Inches(1.1), Inches(1.6), W - Inches(2.2), Inches(1.55),
                               size=fsize, align=PP_ALIGN.CENTER,
                               anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.2)
        self._run(p, formula, size=fsize, bold=True, color=self.s.main)
        if note:
            tb2, tf2, p2 = self._text(s, Inches(1.1), Inches(3.35), W - Inches(2.2),
                                      Inches(0.5), size=13, align=PP_ALIGN.CENTER)
            self._run(p2, note, size=13, color=_rgb("6B7785"))
        if legs:
            n = min(2, len(legs))
            gap = Inches(0.3)
            cw = (W - Inches(1.6) - gap * (n - 1)) / n
            for i, leg in enumerate(legs[:2]):
                x = Inches(0.8) + (cw + gap) * i
                self._rect(s, x, Inches(4.05), cw, H - Inches(5.15), self.s.white,
                           line=_rgb("E2E6EE"), shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                           radius=0.04)
                tb3, tf3, p3 = self._text(s, x + Inches(0.25), Inches(4.2),
                                          cw - Inches(0.5), Inches(0.45), size=15)
                self._run(p3, "条件" if i == 0 else "结论", size=15, bold=True,
                          color=self.s.accent)
                tb4, tf4, p4 = self._text(s, x + Inches(0.25), Inches(4.7),
                                          cw - Inches(0.5), H - Inches(5.9), size=16,
                                          line_spacing=1.5)
                self._rich(p4, leg, size=16)
        self._chrome(s)
        return s

    def compare(self, title: str, left_title: str, left_text: str,
                right_title: str, right_text: str, *, kicker: str = "",
                words: list[str] | None = None):
        """英语：英汉对照页——左英文右中文，中间细线；底部可挂词块。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        top = Inches(1.45)
        bottom = H - Inches(1.15 if words else 0.7)
        half = (W - Inches(1.9)) / 2
        # 中缝：竖直细线，颜色要看得见
        self._rect(s, W / 2 - Emu(6350), top + Inches(0.1), Emu(12700),
                   bottom - top - Inches(0.2), _rgb("DCE0EE"))
        for i, (tt, body, col) in enumerate([(left_title, left_text, self.s.main),
                                             (right_title, right_text, self.s.accent)]):
            x = Inches(0.95) + (half + Inches(0.3)) * i
            tb, tf, p = self._text(s, x, top, half, Inches(0.44), size=14)
            self._run(p, tt, size=14, bold=True, color=col)
            size = 19 if len(body) < 260 else 17
            tb2, tf2, p2 = self._text(s, x, top + Inches(0.5), half,
                                      bottom - top - Inches(0.55), size=size,
                                      line_spacing=1.45)
            self._rich(p2, body, size=size)
        if words:
            self._rect(s, Inches(0.62), H - Inches(1.05), W - Inches(1.24),
                       Inches(0.62), self.s.light, shape=MSO_SHAPE.ROUNDED_RECTANGLE,
                       radius=0.2)
            tb3, tf3, p3 = self._text(s, Inches(0.85), H - Inches(0.98), W - Inches(1.7),
                                      Inches(0.5), size=13, anchor=MSO_ANCHOR.MIDDLE)
            self._run(p3, "词块　", size=13, bold=True, color=self.s.accent)
            self._run(p3, "　·　".join(words[:6]), size=13, color=self.s.ink)
        self._chrome(s)
        return s

    def reaction(self, title: str, equation: str, *, condition: str = "",
                 note: str = "", points: list[str] | None = None,
                 kicker: str = ""):
        """化学：反应式页——条件写在箭头（或等号）上方，下方列现象/要点。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        # 方程式框（只装式子）
        self._rect(s, Inches(0.8), Inches(1.4), W - Inches(1.6), Inches(1.25),
                   self.s.light, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
        fsize = 32 if len(equation) < 34 else 26
        tb, tf, p = self._text(s, Inches(1.0), Inches(1.45), W - Inches(2.0), Inches(1.15),
                               size=fsize, align=PP_ALIGN.CENTER,
                               anchor=MSO_ANCHOR.MIDDLE)
        self._run(p, equation, size=fsize, bold=True, color=self.s.main)
        y = Inches(2.78)
        if condition:
            # 反应条件单独一行（真实卷子写在等号上方，PPT 里放在式子下方最稳）
            cw = Inches(0.19 * len(condition) + 0.9)
            self._rect(s, W / 2 - cw / 2, y, cw, Inches(0.4), self.s.white,
                       line=self.s.accent, shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.3)
            tb2, tf2, p2 = self._text(s, W / 2 - cw / 2, y + Inches(0.02), cw, Inches(0.36),
                                      size=13, align=PP_ALIGN.CENTER,
                                      anchor=MSO_ANCHOR.MIDDLE)
            self._run(p2, f"条件：{condition}", size=13, bold=True, color=self.s.accent)
            y += Inches(0.62)
        if note:
            tb3, tf3, p3 = self._text(s, Inches(1.0), y, W - Inches(2.0), Inches(0.5),
                                      size=14, align=PP_ALIGN.CENTER)
            self._rich(p3, note, size=14, color=_rgb("4A5A55"))
            y += Inches(0.6)
        if points:
            size, sp = _fit(points, base=20)
            tb4, tf4, p4 = self._text(s, Inches(0.85), max(y, Inches(3.5)),
                                      W - Inches(1.7), H - Inches(4.5), size=size,
                                      line_spacing=1.5)
            for i, it in enumerate(points):
                pp = self._p(tf4, first=(i == 0), space_before=0 if i == 0 else sp)
                _indent(pp, 0.45, -0.45)
                self._run(pp, "▪ ", size=size, color=self.s.accent)
                self._rich(pp, it, size=size)
        self._chrome(s)
        return s

    def concept(self, title: str, center: str, branches: list[tuple[str, list[str]]],
                *, kicker: str = "归纳"):
        """生物：概念图页——中心词 + 若干分支卡片（2×2）。"""
        s = self._blank()
        self._header(s, title, kicker=kicker)
        cw = Inches(2.5)
        ch = Inches(0.78)
        cx = (W - cw) / 2
        self._rect(s, cx, Inches(1.45), cw, ch, self.s.main,
                   shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.18)
        tb, tf, p = self._text(s, cx, Inches(1.45), cw, ch, size=20,
                               align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        self._run(p, center, size=20, bold=True, color=self.s.white)
        self._rect(s, cx + cw / 2 - Emu(6350), Inches(2.23), Emu(12700), Inches(0.28),
                   self.s.accent)
        self._rect(s, Inches(1.6), Inches(2.51), W - Inches(3.2), Emu(12700),
                   self.s.accent)
        n = min(4, len(branches))
        if n == 0:
            self._chrome(s)
            return s
        gap = Inches(0.3)
        bw = (W - Inches(1.6) - gap * (n - 1)) / n
        for i, (label, items) in enumerate(branches[:n]):
            x = Inches(0.8) + (bw + gap) * i
            y = Inches(2.85)
            self._rect(s, x + bw / 2 - Emu(6350), Inches(2.51), Emu(12700),
                       Inches(0.34), self.s.accent)
            self._rect(s, x, y, bw, H - Inches(3.9), self.s.light,
                       shape=MSO_SHAPE.ROUNDED_RECTANGLE, radius=0.05)
            tb2, tf2, p2 = self._text(s, x + Inches(0.18), y + Inches(0.14),
                                      bw - Inches(0.36), Inches(0.4), size=15)
            self._run(p2, label, size=15, bold=True, color=self.s.main)
            size, sp = _fit(items, base=14)
            tb3, tf3, p3 = self._text(s, x + Inches(0.2), y + Inches(0.62),
                                      bw - Inches(0.4), H - Inches(4.7), size=size,
                                      line_spacing=1.4)
            for j, it in enumerate(items):
                pp = self._p(tf3, first=(j == 0), space_before=0 if j == 0 else sp)
                _indent(pp, 0.28, -0.28)
                self._run(pp, "· ", size=size, color=self.s.accent)
                self._rich(pp, it, size=size)
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
        tb, tf, p = self._text(s, Inches(0.85), Inches(1.1), W - Inches(1.7), Inches(0.9),
                               size=38)
        self._run(p, title, size=38, bold=True, color=self.t.main)
        if items:
            size, sp = _fit(items, base=23)
            # 垂直居中排要点，避免页面下半空置
            n = len(items)
            block_h = Inches(0.52 * n + 0.1 * max(0, n - 1))
            top = max(Inches(2.3), (H - block_h) / 2 + Inches(0.25))
            tb2, tf2, p2 = self._text(s, Inches(0.95), top, W - Inches(1.9),
                                      H - top - Inches(1.1), size=size, line_spacing=1.45)
            for i, it in enumerate(items):
                pp = self._p(tf2, first=(i == 0), space_before=0 if i == 0 else sp)
                _indent(pp, 0.45, -0.45)
                self._run(pp, "✔ ", size=size, color=self.t.accent)
                self._rich(pp, it, size=size)
        if self.s.key in ("physics", "chemistry"):
            self._chrome(s)      # 页码方块 + 页脚（含学校），不再单排一行
        else:
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
    scenes: list[Block] = []
    cur = None
    for b in sec.blocks:
        if b.kind == "scene":
            scenes.append(b)
            continue
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
    return bullets, items, tables, images, tips, scenes


def _render_scene(deck, block, sec_title: str) -> None:
    """渲染 > 引文/公式/对照/反应/概念 这几类学科版式。"""
    kind = (block.meta or {}).get("scene", "")
    raw = block.text.strip()
    head = f"{sec_title}｜{kind}" if sec_title else kind
    if kind == "引文":
        src = ""
        m = re.search(r"[（(]?\s*(?:——|--)\s*(.+?)[）)]?$", raw)
        body = raw
        if m and len(m.group(1)) < 40:
            src = m.group(1).strip()
            body = raw[: m.start()].strip()
        deck.quote(head, body, source=src)
    elif kind == "公式":
        parts = [x.strip() for x in raw.split("|")]
        formula = parts[0]
        legs, note = [], ""
        for extra in parts[1:]:
            if extra.startswith(("条件", "前提")):
                legs.append(extra.split("：", 1)[-1].split(":", 1)[-1].strip())
            elif extra.startswith(("结论", "适用")):
                legs.append(extra.split("：", 1)[-1].split(":", 1)[-1].strip())
            else:
                note = extra
        deck.formula(head, formula, note=note,
                     legs=legs if len(legs) >= 2 else (legs + [""])[:2] if legs else None)
    elif kind == "对照":
        parts = [x.strip() for x in raw.split("|")]
        if len(parts) >= 4:
            deck.compare(head, parts[0], parts[1], parts[2], parts[3])
        elif len(parts) == 2:
            deck.compare(head, "English", parts[0], "中文", parts[1])
        else:
            deck.bullets(head, [raw])
    elif kind == "反应":
        parts = [x.strip() for x in raw.split("|")]
        eq = parts[0]
        cond = next((p.split("：", 1)[-1] for p in parts[1:] if p.startswith(("条件", "加热", "催化"))), "")
        pts = [p for p in parts[1:] if not p.startswith(("条件", "加热", "催化"))]
        deck.reaction(head, eq, condition=cond, points=pts or None)
    elif kind == "史料":
        # 历史史料页：材料 + 出处落款（> 史料：…——摘编自《…》）
        src = ""
        m = re.search(r"(?:——|--)\s*(.+)$", raw)
        body = raw
        if m and len(m.group(1)) < 50:
            src = m.group(1).strip()
            body = raw[: m.start()].strip()
        deck.quote(head, body, source=src, kicker="史料")
    elif kind == "材料":
        # 政治时政材料页：材料卡 + 设问（> 材料：… | 设问：…）
        parts = [x.strip() for x in raw.split("|")]
        body = parts[0]
        ask = next((p.split("：", 1)[-1].strip()
                    for p in parts[1:] if p.startswith(("设问", "问题"))), "")
        deck.quote(head, body, note=f"设问：{ask}" if ask else "", kicker="时政材料")
    elif kind == "图表":
        # 地理图表页：图例 + 判读步骤（> 图表：… | 判读：①…②…）
        parts = [x.strip() for x in raw.split("|")]
        legend = parts[0]
        steps = next((p.split("：", 1)[-1].strip()
                      for p in parts[1:] if p.startswith(("判读", "步骤", "读图"))), "")
        deck.bullets(head, [f"图例：{legend}"] +
                     [f"判读 {s.strip()}" for s in re.split(r"[；;]", steps) if s.strip()],
                     kicker="图表判读")
    elif kind == "模板":
        # 答题模板页：模式行 + 青蓝色块「标签：×××」行（史主2-08 取证）
        parts = [x.strip() for x in raw.split("|") if x.strip()]
        mode = ""
        rows = []
        for p in parts:
            if p.startswith(("模式", "答题模式")):
                mode = p.split("：", 1)[-1].strip()
            else:
                rows.append(p)
        deck.template(head, rows or [raw], mode=mode)
    elif kind == "口诀":
        deck.mnemonic(head, raw)
    elif kind == "规范表述":
        # 失分对照：错误写法｜规范表述，每行一对
        pairs = []
        for seg in raw.split("|"):
            seg = seg.strip()
            if not seg:
                continue
            if "｜" in seg:
                w, _, r = seg.partition("｜")
                pairs.append((w.strip(), r.strip()))
            elif "→" in seg:
                w, _, r = seg.partition("→")
                pairs.append((w.strip(), r.strip()))
            else:
                pairs.append((seg, ""))
        deck.misconceptions(head, pairs or [(raw, "")])
    elif kind == "采分":
        rows = [x.strip() for x in raw.split("|") if x.strip()]
        deck.scoring(head, rows or [raw])
    elif kind == "考情":
        # 考情分析表：表头固定「年份｜卷别题号分值｜命题情境和角度」
        rows = []
        first = [x.strip() for x in raw.split("|")[0].split("｜") if x.strip()]
        if first and not re.match(r"^\d{4}", first[0]):
            rows.append(first)
        for seg in raw.split("|"):
            seg = seg.strip()
            if not seg:
                continue
            cells = [x.strip() for x in seg.split("｜") if x.strip()]
            if cells:
                rows.append(cells)
        if not rows:
            rows = [[raw]]
        if rows and not re.match(r"^(年份|时间|考点)", rows[0][0]):
            rows.insert(0, ["年份", "卷别·题号·分值", "命题情境和角度"])
        deck.table(head, rows, kicker="考情分析")
    elif kind == "概念":
        segs = [x.strip() for x in raw.split(">") if x.strip()]
        center = segs[0]
        branches = []
        for seg in segs[1:]:
            if "：" in seg or ":" in seg:
                lab, body = re.split(r"[:：]", seg, 1)
                branches.append((lab.strip(), [x.strip() for x in re.split(r"[;；、]", body) if x.strip()]))
            else:
                branches.append((seg, []))
        deck.concept(head, center, branches)
    else:
        deck.bullets(head, [raw])


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
                 r"重点难点|考情分析|参考答案|答案|课标要求|时空定位|重要概念解析|"
                 r"知识梳理|要点回顾|课程标准|素养提升|学以致用|精读教材|教考衔接|"
                 r"练透考法|归纳拓展|课堂任务|写作提升|晨诵|暮读|晚写|"
                 r"目标要求|内容索引|知识拓展|实战训练|精讲核心考点|疑难解析|归纳总结|"
                 r"原理及装置图|实验思路|数据处理|注意事项|误差分析|回忆与展望|合作探究|"
                 r"开头段|论证段|结尾段|中心论点|分论点|辩证分析|"
                 r"导语|短语背记|写作万能|写作高分句|N题详解|"
                 r"以题悟道|易错易混点|点睛|思维建模|总结提升|"
                 r"课后巩固区|课堂巩固区|目标具体分解|评分标准|方法技巧|特别提醒|"
                 r"知识建构|简记|视角|核心考点|历年对照|本题特点|命题分析|命题趋势|"
                 r"官方答案|官方解析|深化拓展|课前背记|词汇整理|句式结构|核心词块|话题微写作|"
                 r"试题分析|命题立意|命题特点|思路解析|学生答题优点|存在的突出问题|备考提示|"
                 r"原题呈现|审题|关键词解读|立意|行文思路|审题常见误区|"
                 r"强基础|研考点|课时精练|必备知识|关键能力|突破题型|掌握解法|高考真题|"
                 r"典例|解题思路|我会做|我能做对|我的反思|课堂互动区|易错清零|情境导入|"
                 r"命题透视|考情前瞻|核心突破|重难全析|解题赋能|素养提升|体系构建|思维导图|知识解构|考向破译|考向探究|考点突破|对标素养|法条链接|错题重做|典题|课前阅读区|重要历史概念|必备知识|阶段特征|考察目标|难点突破|中外对比|概念阐释|大事年表|知识导学|情境探究|时空坐标|背诵任务|知识链接|拓展深化|答题模板|易错清单|微专题|类题巩固|变式训练|评分细则|温故知新|简图绘制|考点提取|读图观变|旁栏思考|探究·实践)")


def _cover_index(hour: str) -> str:
    m = re.search(r"\d+", hour or "")
    return f"第 {int(m.group()):02d} 讲" if m else ""


def build_xuean_deck(md: str, out_path: str | Path, *, theme: str = "",
                     school: str = "", author: str = "", subject: str = "") -> Path:
    d = parse(md)
    # 学科判定：显式参数 > md 元信息 > 标题正文猜测
    subject = subject or d.meta.get("学科", "") or ss.guess(d.title or "")
    deck = Deck(theme=theme, subject=subject, footer=d.title, school=school)
    subject = d.meta.get("学科", "") or (deck.s.label if deck.s.key != "general" else "")
    grade = d.meta.get("年级", "")
    hour = d.meta.get("课时", "") if deck.s.key in ("physics", "chemistry", "biology") else ""
    if re.search(r"第\s*\d+\s*[讲节课]|^\s*\d+\s*[.．]", d.title or ""):
        hour = ""            # 标题自带编号，就不要再叠「第 NN 讲」
    sub = f"{grade}{subject}" if (grade or subject) else ""
    deck.cover(d.title or "学案", subtitle=sub, meta=school or author, tag="学案",
               index=_cover_index(hour))

    for idx, sec in enumerate([x for x in d.sections if x.title], 1):
        title = _clean_heading(sec.title)
        if re.match(r"^(参考答案|答案与解析|答案及解析|答案)$", sec_key(sec.title)):
            deck.divider(title, index=f"{idx:02d}")
            plain = [b.text.strip() for b in sec.blocks
                     if b.kind in ("num", "text", "answer", "bullet") and b.text.strip()]
            for chunk in _paged(title, plain, per=6, limit=340):
                deck.bullets(title, chunk, kicker="答案")
            continue
        is_quiz = bool(re.search(r"当堂检测|课堂检测|检测|练习|作业|自测", sec.title))
        bullets, items, tables, images, tips, scenes = _split_section(sec, quiz=is_quiz)

        if BIG.match(sec_key(sec.title)):
            deck.divider(title, index=f"{idx:02d}",
                         spine="导入" if "目标" in title else "")

        if bullets:
            for k, chunk in enumerate(_paged(title, bullets)):
                if not chunk:
                    continue
                deck.bullets(title + ("（续）" if k else ""), chunk,
                             kicker=("学习目标" if "目标" in title else title[:10]),
                             band=bool(items or tables),
                             spine="知识梳理" if (not items and not tables
                                                  and deck.s.key in ("math", "physics"))
                             else "")
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
        for sc in scenes:
            _render_scene(deck, sc, title)
        for tp in tips:
            deck.tip("特别提醒", tp)

    deck.closing("本节课小结", _summary_points(d))
    return deck.save(out_path)


def build_paper_deck(md: str, out_path: str | Path, *, theme: str = "",
                     school: str = "", mode: str = "review", subject: str = "") -> Path:
    """试卷 → 讲评课件：每题题干页（留白给学生） + 答案解析页（视觉跟随学科）。"""
    d = parse(md)
    # 学科判定：显式参数 > md 元信息 > 标题正文猜测（别让物理卷戴语文口令）
    subject = subject or d.meta.get("学科", "") or ss.guess(d.title or "")
    deck = Deck(theme=theme, subject=subject or "general", footer=d.title, school=school)
    meta = "　".join(f"{k}：{v}" for k, v in
                    [("时间", d.meta.get("时间")), ("满分", d.meta.get("满分"))] if v)
    hint = None if deck.s.key == "chinese" else ""
    sub = (deck.s.label if deck.s.key != "general" else "") if mode == "review" else "试卷"
    deck.cover(d.title or "试卷讲评", subtitle=sub,
               meta=meta, tag="试卷讲评", hint=hint)

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


def build_grade_deck(analysis: dict, out_path: str | Path, *, theme: str = "",
                     school: str = "", title: str = "成绩分析",
                     subject: str = "", notes: str = "") -> Path:
    """成绩分析 → 汇报 PPT（照真实成绩分析会课件结构）。

    取证：U 盘 139 页真实《2026年高考山东卷分析暨2027年高考展望》——
    封面 → 目录（壹/贰/叁章）→ 章节页（大数字）→ 数据统计表（多年对比）→
    特点/趋势分析 → 备考建议（分阶段）。照此编排 grade 数据。
    """
    from .grades import analyze as _an  # noqa: F401  （analysis 已算好，此处仅类型提示）
    deck = Deck(theme=theme, subject=subject or "general", footer=title, school=school)
    n = analysis.get("n_rows", 0)
    tot = analysis.get("total", {})
    exams = (analysis.get("trend") or {}).get("exams", [])
    meta = f"共 {n} 人"
    if exams:
        meta += f"　考次：{'、'.join(exams[:4])}"
    deck.cover(title, subtitle="成绩分析报告", meta=meta, tag="质量分析",
               hint=None if deck.s.key == "chinese" else "")

    # 目录（壹/贰/叁）
    deck.bullets("目录", [
        "壹　成绩概况与分数段分布",
        "贰　各科分析与班级对比",
        "叁　问题诊断与教学建议",
    ], kicker="总览")

    # 壹 成绩概况
    deck.divider("成绩概况与分数段分布", index="壹", subtitle="总分/均分/分布")
    if tot.get("n"):
        deck.bullets("总分概况", [
            f"参考 {tot['n']} 人　均分 {tot['mean']}　中位 {tot['median']}　标准差 {tot['std']}",
            f"最高 {tot['max']}　最低 {tot['min']}　及格率 {tot['pass_rate']}%　优秀率 {tot['excellence_rate']}%",
        ], kicker="总览")
    if tot.get("bands"):
        band_rows = [["分数段"] + list(tot["bands"].keys()),
                     ["人数"] + [str(v) for v in tot["bands"].values()]]
        deck.table("分数段分布", band_rows, kicker="数据")
    if analysis.get("top"):
        rk = analysis.get("rank_key", "总分")
        top_rows = [["名次", "姓名", "班级", rk]]
        for i, t in enumerate(analysis["top"][:8], 1):
            top_rows.append([str(i), t["name"], t.get("class_", ""), str(t["score"])])
        deck.table("总分名次榜（前 8）", top_rows, kicker="数据")

    # 贰 各科分析与班级对比
    deck.divider("各科分析与班级对比", index="贰", subtitle="均分/及格率/班级差异")
    subj = analysis.get("subjects", {})
    if subj:
        srows = [["科目", "均分", "中位", "最高", "最低", "及格率%", "优秀率%"]]
        for s, st in subj.items():
            if st.get("n"):
                srows.append([s, str(st["mean"]), str(st["median"]), str(st["max"]),
                              str(st["min"]), str(st["pass_rate"]), str(st["excellence_rate"])])
        deck.table("各科成绩统计", srows, kicker="数据")
    if analysis.get("subject_tops"):
        st_rows = [["科目", "第一名", "第二名", "第三名"]]
        for s, tops in analysis["subject_tops"].items():
            cells = [s] + [f"{t['name']}（{t['score']}）" for t in tops[:3]]
            while len(cells) < 4:
                cells.append("")
            st_rows.append(cells)
        deck.table("单科榜首", st_rows, kicker="数据")
    if analysis.get("class_stats"):
        cs_rows = [["班级", "人数", "均分", "中位", "最高", "最低", "及格率%"]]
        for c, st in analysis["class_stats"].items():
            if st.get("n"):
                cs_rows.append([c, str(st["n"]), str(st["mean"]), str(st["median"]),
                                str(st["max"]), str(st["min"]), str(st["pass_rate"])])
        deck.table("班级对比", cs_rows, kicker="数据")
    tr = analysis.get("trend")
    if tr and tr.get("improved"):
        imp_rows = [["姓名", "首次", "最近", "提升"]]
        for t in tr["improved"][:5]:
            imp_rows.append([t["name"], str(t["first"]), str(t["last"]),
                             f"+{t['delta']}" if t["delta"] >= 0 else str(t["delta"])])
        deck.table("进步榜", imp_rows, kicker="数据")

    # 叁 弱科分析与提升推荐（用分析器的 weak_subjects/recommendations，与 docx 一致）
    deck.divider("弱科分析与提升推荐", index="叁", subtitle="弱科归因/提升措施")
    weak = analysis.get("weak_subjects") or []
    recs = analysis.get("recommendations") or []
    if weak:
        wrows = [["科目", "均分率%", "均分/满分", "及格率%", "优秀率%", "弱科原因"]]
        for w in weak[:5]:
            wrows.append([w["subject"], str(w["mean_rate"]), f"{w['mean']}/{w['full']}",
                          str(w["pass_rate"]), str(w["excellence_rate"]), w["reason"]])
        deck.table("弱科识别（按均分率/及格率/优秀率综合判定）", wrows, kicker="诊断")
    diag = []
    for w in weak[:2]:
        diag.append(f"{w['subject']}均分率仅 {w['mean_rate']}%（{w['mean']}/{w['full']}），{w['reason']}")
    if tot.get("n") and tot.get("pass_rate", 100) < 80:
        diag.append(f"及格率 {tot['pass_rate']}% 偏低，关注低分段学生的基础落实")
    if tr and tr.get("declined"):
        diag.append(f"{len(tr['declined'])} 名学生退步明显，需个别谈话与错题跟踪")
    if not diag:
        diag = ["整体成绩稳定，继续保持分层训练", "关注学科均衡，避免偏科拉分"]
    deck.bullets("问题诊断", diag, kicker="归因", numbered=True)
    if recs:
        deck.bullets("提升推荐", recs, kicker="措施", numbered=True)
    else:
        advice = [ln.strip() for ln in (notes or "").split("\n") if ln.strip()]
        if not advice:
            advice = ["分层布置作业：基础题全员过关，提高题面向尖子",
                      "错题重做：建立错题本，三天后重做一遍",
                      "规范答题：按采分点分行书写，杜绝口语化表达"]
        deck.bullets("教学建议", advice, kicker="措施", numbered=True)
    deck.closing("小结", [
        f"本次参考 {n} 人，均分 {tot.get('mean', '—')}",
        "数据背后找问题，问题背后定措施",
        "下一阶段：抓基础、抓规范、抓落实",
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
               if re.match(r"^(参考答案|答案与解析|答案|解析)", sec_key(s.title))]
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
