"""A4 print renderers for 学案 and 试卷 (python-docx).

版式来自 25 份真实学案 / 32 份可解析真卷的逐份测量：
  * A4；正文 宋体 12pt、行距 1.5；填空=下划线全角空格
  * 页眉：学校+科目+文档名　编写人／审核／时间（9pt 黑体，制表位分栏）
  * 栏目【…】黑体 12pt **左对齐**；「课前预习案 / 当堂检测」这类大栏目居中 15pt
  * 试卷：标题与日期同行、班级姓名考号三栏底框线、选项 2×2、答案另页 + 速查表
  * 页脚：学校+文档名　　第 X 页（共 Y 页）
"""
from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

from .mdparse import Block, Doc, blanks, parse, sec_key

CN_BODY = "宋体"
CN_HEAD = "黑体"
CN_TITLE = "华文中宋"
EN_BODY = "Times New Roman"
GREY = RGBColor(0x59, 0x59, 0x59)
ANSWER_INK = RGBColor(0x6B, 0x12, 0x12)

PAGE_W, PAGE_H = Cm(21.0), Cm(29.7)


# ------------------------------------------------------------------ helpers
def _style_run(run, cn=CN_BODY, en=EN_BODY, size: float = 12, bold=False, color=None,
               underline=False):
    run.font.name = en
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.underline = underline
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), cn)
    rfonts.set(qn("w:ascii"), en)
    rfonts.set(qn("w:hAnsi"), en)


def _para(doc, text="", *, cn=CN_BODY, en=EN_BODY, size=12, bold=False,
          align=None, first_indent=None, left_indent=None, hanging=None,
          space_before=0, space_after=0, line=1.5, color=None, style=None,
          keep_next=False):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    pf.line_spacing = line
    pf.keep_with_next = keep_next
    if left_indent is not None:
        pf.left_indent = Cm(left_indent)
    if hanging is not None:
        pf.first_line_indent = Cm(-hanging)
    elif first_indent is not None:
        pf.first_line_indent = Cm(first_indent)
    if align is not None:
        p.alignment = align
    if text:
        _style_run(p.add_run(text), cn, en, size, bold, color)
    return p


def _blank_run(p, chars: int, size: float = 12):
    # 全角空格，1 个 ≈ 1 个汉字宽，下划线连续不断
    r = p.add_run("\u3000" * max(1, chars))
    _style_run(r, CN_BODY, EN_BODY, size, underline=True)


LABEL_RE = re.compile(r"^(重点|难点|考点|例\s*\d+|探究[一二三四五六七八九十]?|变式[一二三四五六七八九十]?|"
                      r"思考|提示|归纳|答案|解析|详解|分析|导语|学习聚焦|知识链接|易错警示|方法指导|"
                      r"点睛|本题特点|命题分析|方法技巧|特别提醒|短语背记|评分标准|易错易混点)\s*[:：]")


UNITS = ("m/s²", "m/s", "km/h", "kg", "mol", "cm", "mm", "km", "mL", "Pa", "℃",
         "m", "s", "g", "L", "N", "J", "W", "V", "A", "%")
UNIT_RE = re.compile(r"(?<=\d)\s+(" + "|".join(re.escape(u) for u in UNITS) + r")(?![\w/²])")


def _nb(text: str) -> str:
    """数字与单位之间换成不换行空格，防止「10 m/s」被拆到两行。"""
    return UNIT_RE.sub(lambda m: "\u00a0" + m.group(1), text)


def _write_runs(p, text: str, size: float, color=None, bold=False, *, answers=False):
    """Write a paragraph body, turning {{答案}} into underlined blanks.

    answers=True 出答案版：答案原文红粗下划线直接填回（真实答案版学案样式）。
    """
    text = _tidy(_nb(text))
    m = LABEL_RE.match(text)
    if m:
        _style_run(p.add_run(m.group(0)), CN_HEAD, EN_BODY, size, True, color)
        text = text[m.end():]
    for kind, val in blanks(text):
        if kind == "blank":
            if answers and val:
                r = p.add_run(val)
                # 真实答案版：红字压在下划线上，不加粗
                _style_run(r, CN_BODY, EN_BODY, size, False,
                           RGBColor(0xC0, 0x00, 0x00), underline=True)
            else:
                _blank_run(p, max(4, min(14, len(val) + 2)), size)
        else:
            _style_run(p.add_run(val), CN_BODY, EN_BODY, size, bold, color)


def _field(p, instr: str, size=9):
    run = p.add_run()
    _style_run(run, CN_BODY, EN_BODY, size, color=GREY)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(it)
    run._r.append(end)


def _footer(section, school: str = "", title: str = "", keep_school: bool = True):
    """页脚居中：学校+文档名　第 X 页（共 Y 页）——真卷的写法。"""
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(0)
    short = re.sub(r"[（(][^）)]*[）)]\s*$", "", title or "").strip()
    if len(short) > 20:
        short = short[:19] + "…"
    head = "".join(x for x in [(school if keep_school else ""), short] if x)
    if head:
        _style_run(p.add_run(head + "\u3000\u3000"), CN_BODY, EN_BODY, 9, color=GREY)
    _style_run(p.add_run("第 "), CN_BODY, EN_BODY, 9, color=GREY)
    _field(p, "PAGE", 9)
    _style_run(p.add_run(" 页（共 "), CN_BODY, EN_BODY, 9, color=GREY)
    _field(p, "NUMPAGES", 9)
    _style_run(p.add_run(" 页）"), CN_BODY, EN_BODY, 9, color=GREY)


def _header(section, *, school="", subject="", grade="", title="",
            author="", reviewer="", date="", doc_kind="学案"):
    """页眉：左「学校+学段学科+学案」，右「编写人／审核／时间」。

    用无框两列表格而不是制表位——制表位会命中 Header 样式自带的居中/右制表位，
    在 LibreOffice/WPS 里跳位不一致。
    """
    head = section.header
    p0 = head.paragraphs[0]
    p0.paragraph_format.space_after = Pt(0)
    for r in list(p0.runs):
        r._element.getparent().remove(r._element)

    usable = section.page_width - section.left_margin - section.right_margin
    table = head.add_table(rows=1, cols=2, width=usable)
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)

    left = "".join(x for x in [school, grade, subject] if x)
    left = (left + doc_kind) if left else doc_kind
    right = "　".join(f"{k}：{v}" for k, v in
                      (("编写人", author), ("审核", reviewer), ("时间", date)) if v)

    c0, c1 = table.cell(0, 0), table.cell(0, 1)
    c0.width = int(usable * 0.42)
    c1.width = int(usable * 0.58)
    pa = c0.paragraphs[0]
    pa.paragraph_format.space_after = Pt(0)
    _style_run(pa.add_run(left), CN_HEAD, EN_BODY, 9, bold=True)
    pb = c1.paragraphs[0]
    pb.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pb.paragraph_format.space_after = Pt(0)
    _style_run(pb.add_run(right), CN_HEAD, EN_BODY, 9, bold=True)

    # 去掉表格边框（只留一张隐形网格），但底边保留细分隔线（真实学案页眉下的横线）
    tbl_borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        tbl_borders.append(el)
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:color"), "808080")
    tbl_borders.append(bottom)
    tblPr.append(tbl_borders)

    # 表格后必须跟一个段落，否则 Word/LibreOffice 会把整个页眉判为非法而不显示
    tail = head.add_paragraph()
    tail.paragraph_format.space_before = Pt(0)
    tail.paragraph_format.space_after = Pt(0)
    tail.paragraph_format.line_spacing = Pt(1)
    _style_run(tail.add_run(""), CN_BODY, EN_BODY, 1)


def _table_borders(table):
    tblPr = table._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "4")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "999999")
        borders.append(el)
    tblPr.insert_element_before(
        borders, "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        "w:tblCaption", "w:tblDescription", "w:tblPrChange")


def _shade(p, fill="F4F1F7"):
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    pPr.append(shd)


def _left_bar(p, color="C81E1E", size=18):
    """段左侧竖条（用段落边框模拟，比插形状稳）。"""
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    left = OxmlElement("w:left")
    left.set(qn("w:val"), "single")
    left.set(qn("w:sz"), str(size))
    left.set(qn("w:space"), "6")
    left.set(qn("w:color"), color)
    pbdr.append(left)
    pPr.append(pbdr)


def _rule(p, color="0B4FA8"):
    pPr = p._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "8")
    bottom.set(qn("w:space"), "2")
    bottom.set(qn("w:color"), color)
    pbdr.append(bottom)
    pPr.append(pbdr)


def new_doc(margins=(Cm(1.9), Cm(1.9), Cm(2.0), Cm(2.0))):
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = PAGE_W, PAGE_H
    s.left_margin, s.right_margin = margins[0], margins[1]
    s.top_margin, s.bottom_margin = margins[2], margins[3]
    normal = doc.styles["Normal"]
    normal.font.name = EN_BODY
    normal.font.size = Pt(12)
    normal.element.rPr.rFonts.set(qn("w:eastAsia"), CN_BODY)
    for name, size, cn in (("Heading 1", 18, CN_TITLE), ("Heading 2", 13, CN_HEAD),
                           ("Heading 3", 12, CN_HEAD)):
        st = doc.styles[name]
        st.font.name = EN_BODY
        st.font.size = Pt(size)
        st.font.bold = True
        st.font.color.rgb = RGBColor(0x1A, 0x1A, 0x1A)
        rpr = st.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        rf.set(qn("w:eastAsia"), cn)
    return doc


BIG_SECTION = re.compile(r"^(课前预习案?|课内探究|当堂检测|课堂检测|课后作业|"
                         r"课后巩固案?|参考答案|答案与解析|答案及解析|答案|"
                         r"学习目标|重点难点|考情分析|评分标准|"
                         r"课标要求|时空定位|重要概念解析|知识梳理|要点回顾|课程标准|"
                         r"素养提升|学以致用|精读教材|教考衔接|练透考法|归纳拓展|"
                         r"考情解读|考点提取|写作提升|晨诵|暮读|晚写|"
                         r"开头段|论证段|结尾段|中心论点|分论点|导语|短语背记|"
                         r"以题悟道|易错易混点|思维建模|总结提升|课后巩固区|课堂巩固区|"
                         r"目标具体分解[一二三1-3]?|方法技巧|特别提醒|知识建构|"
                         r"核心考点|历年对照|本题特点|命题分析|命题趋势|官方答案|深化拓展|"
                         r"试题分析|命题立意|命题特点|思路解析|学生答题优点|存在的突出问题|"
                         r"备考提示|原题呈现|审题指导|关键词解读|立意|行文思路|审题常见误区|"
                         r"强基础|研考点|课时精练|必备知识|关键能力|突破题型|掌握解法|"
                         r"高考真题|典例|解题思路|我会做|我能做对|我的反思|课堂互动区)"
                         r"(?:[·•・].*)?$")   # 允许「强基础·教考衔接」式中点复合词


def _heading(doc, text, level, *, center=False, size=None, rule=False, big=False,
             color=None):
    """栏目标题：普通栏目黑体 12pt 左对齐；大栏目（课前预习案…）黑体 15pt 居中。"""
    p = doc.add_paragraph(style=f"Heading {level}")
    pf = p.paragraph_format
    pf.space_before = Pt(14 if big else (10 if level == 2 else 8))
    pf.space_after = Pt(6 if big else 4)
    pf.line_spacing = 1.25
    pf.keep_with_next = True
    if center or big:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if level == 2 and not big and not text.startswith("【"):
        text = f"【{text}】"
    _style_run(p.add_run(text), CN_HEAD, EN_BODY,
               size or (15 if big else 12), bold=True, color=color)
    if rule:
        _rule(p)
    return p


# ------------------------------------------------------------------ blocks
def _render_blocks(doc, blocks, *, body_size=12, indent=0.74, line=1.5,
                   answers: bool = False):
    for b in blocks:
        if b.kind == "text":
            p = _para(doc, "", size=body_size, first_indent=indent, line=line,
                      space_after=1)
            _write_runs(p, b.text, body_size, answers=answers)
        elif b.kind == "note":
            _para(doc, b.text, size=body_size - 1.5, color=GREY, first_indent=indent,
                  line=1.3, space_after=2)
        elif b.kind in ("bullet", "num"):
            p = _para(doc, "", size=body_size, left_indent=indent + 0.6 * b.level,
                      hanging=indent, line=line, space_after=1)
            if b.kind == "bullet":
                _style_run(p.add_run("▪ "), CN_BODY, EN_BODY, body_size - 1)
            _write_runs(p, b.text, body_size, answers=answers)
        elif b.kind == "table":
            if not b.rows:
                continue
            cols = max(len(r) for r in b.rows)
            t = doc.add_table(rows=len(b.rows), cols=cols)
            t.alignment = WD_TABLE_ALIGNMENT.CENTER
            t.autofit = False
            usable = (doc.sections[0].page_width - doc.sections[0].left_margin
                      - doc.sections[0].right_margin)
            tblPr = t._tbl.tblPr
            tw = OxmlElement("w:tblW")
            tw.set(qn("w:type"), "dxa")
            tw.set(qn("w:w"), str(int(usable / 635)))
            tblPr.append(tw)
            for ci in range(cols):
                for cell in [t.cell(r, ci) for r in range(len(b.rows))]:
                    cell.width = int(usable / cols)
            _table_borders(t)
            for ri, row in enumerate(b.rows):
                for ci in range(cols):
                    cell = t.cell(ri, ci)
                    cell.text = ""
                    p = cell.paragraphs[0]
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.line_spacing = 1.15
                    # 表格里的 {{挖空}} 也要渲染成下划线/答案（曾原样显示 {{}}）
                    _write_runs(p, row[ci] if ci < len(row) else "",
                                body_size - 1.5, bold=(ri == 0), answers=answers)
            doc.add_paragraph().paragraph_format.space_after = Pt(2)
        elif b.kind == "image":
            path = Path(b.src)
            if path.exists():
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.add_run().add_picture(str(path), width=Cm(9))
                if b.text:
                    _para(doc, b.text, size=9, align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY)
        elif b.kind == "answer":
            p = _para(doc, "", size=body_size, left_indent=0.55, line=line, space_after=2)
            _shade(p, "F4F5F7")
            _left_bar(p, "9AA3AD", 12)
            _write_runs(p, b.text, body_size, color=ANSWER_INK)
        elif b.kind == "insight":
            # 点睛/归纳：蓝色竖条 + 浅蓝底
            p = _para(doc, "", size=body_size, left_indent=0.55, line=line, space_after=3)
            _shade(p, "EEF4FB")
            _left_bar(p, "2E86D6")
            _write_runs(p, b.text, body_size, color=RGBColor(0x1B, 0x4F, 0x8A))
        elif b.kind == "tip":
            # 特别提醒：红色竖条 + 浅红底 + 深红字（灰度打印也能靠底纹+竖条区分）
            p = _para(doc, "", size=body_size, left_indent=0.55, line=line, space_after=4)
            _shade(p, "FDECEC")
            _left_bar(p, "C81E1E")
            _write_runs(p, b.text, body_size, color=RGBColor(0x8A, 0x1B, 0x1B))
        elif b.kind == "scene":
            # PPT 专属版式（公式/引文/对照/反应/概念）在 docx 里落成可读文本，
            # 不能丢——曾导致「知识点2 三个基本公式」标题下整节空白。
            meta = b.meta or {}
            st = meta.get("scene", "")
            text = b.text or ""
            if st == "公式" and "|" in text:
                parts = [x.strip() for x in text.split("|")]
                formula = parts[0]
                cond = next((x.split("：", 1)[1] for x in parts if x.startswith("条件：")), "")
                concl = next((x.split("：", 1)[1] for x in parts
                              if x.startswith(("结论：", "要点："))), "")
                p = _para(doc, "", size=body_size + 4, align=WD_ALIGN_PARAGRAPH.CENTER,
                          line=1.3, space_after=2)
                pf = p.paragraph_format
                pf.keep_with_next = True
                _style_run(p.add_run(formula), CN_BODY, EN_BODY, body_size + 4, bold=True)
                if cond or concl:
                    bits = "　".join(x for x in
                                    ([f"条件：{cond}"] if cond else []) +
                                    ([f"结论：{concl}"] if concl else []) if x)
                    _para(doc, bits, size=body_size - 1, color=GREY,
                          align=WD_ALIGN_PARAGRAPH.CENTER, line=1.3, space_after=6)
            elif st == "引文" and "|" in text:
                body, _, src = text.rpartition("|")
                p = _para(doc, body.strip(), size=body_size + 1, first_indent=indent,
                          line=line, space_after=2)
                if src.strip():
                    _para(doc, "—— " + src.strip(" ——"), size=body_size - 1,
                          align=WD_ALIGN_PARAGRAPH.RIGHT, color=GREY, space_after=6)
            elif st == "对照" and "|" in text:
                left, _, right = text.partition("|")
                for lab, seg in (("EN", left), ("CN", right)):
                    p = _para(doc, "", size=body_size, first_indent=indent, line=line,
                              space_after=2)
                    _style_run(p.add_run(f"[{lab}] "), CN_HEAD, EN_BODY, body_size - 1,
                               bold=True, color=RGBColor(0x2E, 0x75, 0xB6))
                    _style_run(p.add_run(seg.strip()), CN_BODY, EN_BODY, body_size)
            else:
                # 反应/概念/其他：按「key：value | key：value」拆成短行
                p = _para(doc, "", size=body_size, first_indent=indent, line=line,
                          space_after=2)
                for i, seg in enumerate(text.split("|")):
                    seg = seg.strip()
                    if not seg:
                        continue
                    if i:
                        p.add_run().add_break()
                    if "：" in seg[:4]:
                        k, _, v = seg.partition("：")
                        _style_run(p.add_run(k + "："), CN_BODY, EN_BODY, body_size,
                                   bold=True, color=RGBColor(0x2E, 0x75, 0xB6))
                        _style_run(p.add_run(v), CN_BODY, EN_BODY, body_size)
                    else:
                        _style_run(p.add_run(seg), CN_BODY, EN_BODY, body_size)
        elif b.kind == "break":
            doc.add_page_break()


# ------------------------------------------------------------------ 学案
def render_xuean(md: str, out_path: str | Path, *, school: str = "",
                 answers: bool = False) -> Path:
    """answers=True 出答案版：{{挖空}} 红粗下划线填回答案（真实答案版学案样式）。"""
    d = parse(md)
    doc = new_doc()
    title = d.title or "学案"
    if answers and "答案" not in title:
        title += "（答案版）"
    _header(doc.sections[0], school=school, subject=d.meta.get("学科", ""),
            grade=d.meta.get("年级", ""), title=d.meta.get("课") or "",
            author=d.meta.get("编写", ""), reviewer=d.meta.get("审核", ""),
            date=d.meta.get("时间", ""))
    _footer(doc.sections[0], school, title, keep_school=False)

    _para(doc, title, cn=CN_TITLE, size=16, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, line=1.2)
    sub = "　".join(x for x in [
        d.meta.get("年级"), d.meta.get("学科"),
        f"第{d.meta['课时']}课时" if d.meta.get("课时") else "",
    ] if x)
    if sub:
        _para(doc, sub, size=10.5, align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY,
              space_after=6, line=1.1)

    if d.preface:
        _render_blocks(doc, d.preface, answers=answers, line=1.3)

    for s in d.sections:
        if not s.title:
            _render_blocks(doc, s.blocks, answers=answers, line=1.3)
            continue
        if s.level == 2:
            if BIG_SECTION.match(sec_key(s.title)):
                _heading(doc, s.title, 2, big=True, size=15)
            else:
                _heading(doc, s.title, 2, center=False)
        else:
            # 真实学案「知识点一：…」用蓝色加粗小标题
            if re.match(r"^知识点", s.title):
                _heading(doc, s.title, 3, color=RGBColor(0x00, 0x00, 0xC0))
            else:
                _heading(doc, s.title, 3)
        _render_blocks(doc, s.blocks, answers=answers, line=1.3)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path



def _seal_line(doc):
    """班级 / 姓名 / 考号 三栏，用表格底框线做填写线（中间留空隙，三段独立）。"""
    usable = doc.sections[0].page_width - doc.sections[0].left_margin - doc.sections[0].right_margin
    widths = [0.30, 0.034, 0.30, 0.034, 0.30]
    labels = {0: "班级：", 2: "姓名：", 4: "考号："}
    table = doc.add_table(rows=1, cols=5)
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for ci, ratio in enumerate(widths):
        cell = table.cell(0, ci)
        cell.width = int(usable * ratio)
        para = cell.paragraphs[0]
        para.paragraph_format.space_before = Pt(2)
        para.paragraph_format.space_after = Pt(6)
        if ci not in labels:
            continue
        tcPr = cell._tc.get_or_add_tcPr()
        borders = OxmlElement("w:tcBorders")
        bottom = OxmlElement("w:bottom")
        bottom.set(qn("w:val"), "single")
        bottom.set(qn("w:sz"), "6")
        bottom.set(qn("w:color"), "808080")
        borders.append(bottom)
        tcPr.append(borders)
        _style_run(para.add_run(labels[ci]), CN_BODY, EN_BODY, 10.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def _expand_options(blocks: list[Block]) -> list[Block]:
    """把 'A.甲　B.乙　C.丙　D.丁' 这种一行多选项拆成独立块。

    前缀（题干「1．…（　　）」）保留为独立块——曾把题干整段丢掉。
    """
    out = []
    for b in blocks:
        if b.kind in ("text", "bullet") and b.text:
            t = b.text.strip()
            marks = list(INLINE_OPT.finditer(t))
            if len(marks) >= 3:
                head = t[:marks[0].start()].strip()
                if head:
                    out.append(Block(b.kind, text=head))
                for i, m in enumerate(marks):
                    end = marks[i + 1].start() if i + 1 < len(marks) else len(t)
                    piece = t[m.start():end].strip()
                    if piece:
                        out.append(Block(b.kind, text=piece))
                continue
        out.append(b)
    return out

# ------------------------------------------------------------------ 试卷
OPTION_RE = re.compile(r"^([A-DＡ-Ｄ])\s*[.．、]\s*(.*)$")
INLINE_OPT = re.compile(r"([A-DＡ-Ｄ])\s*[.．、]\s*")


def render_paper(md: str, out_path: str | Path, *, school: str = "",
                 answer_page: bool = True, seal_line: bool = True) -> Path:
    d = parse(md)
    doc = new_doc(margins=(Cm(1.8), Cm(1.8), Cm(2.0), Cm(1.9)))
    title = d.title or "试卷"
    _footer(doc.sections[0], school, title, keep_school=False)

    # 标题与日期同一行（真卷写法）
    head = title + (f"　{d.meta['日期']}" if d.meta.get("日期") else "")
    _para(doc, head, cn=CN_TITLE, size=16, bold=True,
          align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, line=1.2)
    meta_bits = [f"{k}：{d.meta[k]}" for k in ("时间", "满分") if d.meta.get(k)]
    if d.meta.get("说明"):
        _para(doc, d.meta["说明"], size=11, align=WD_ALIGN_PARAGRAPH.LEFT,
              space_after=4, line=1.5)
    elif meta_bits:
        _para(doc, "（" + "　".join(meta_bits) + "）", size=10.5,
              align=WD_ALIGN_PARAGRAPH.CENTER, color=GREY, space_after=4)
    if seal_line:
        _seal_line(doc)

    answer_key = _collect_answer_key(d)
    in_answers = False
    for s in d.sections:
        if re.match(r"^(参考答案|答案与解析|答案及解析|答案|解析)", sec_key(s.title)):
            in_answers = True
            if answer_page:
                doc.add_page_break()
            _para(doc, s.title, cn=CN_HEAD, size=14, bold=True,
                  align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
            if answer_key:
                _answer_key_table(doc, answer_key)
            _render_blocks(doc, s.blocks, body_size=11)
            continue
        if in_answers:
            _render_blocks(doc, s.blocks, body_size=11)
            continue
        if s.title:
            _heading(doc, s.title, 2, center=False, size=12)
        blocks = _expand_options(s.blocks)
        i = 0
        while i < len(blocks):
            b = blocks[i]
            if b.kind in ("text", "bullet") and OPTION_RE.match(b.text.strip()):
                row = []
                while i < len(blocks) and blocks[i].kind in ("text", "bullet") \
                        and OPTION_RE.match(blocks[i].text.strip()):
                    row.append(blocks[i])
                    i += 1
                _option_row(doc, row)
                continue
            _paper_block(doc, b)
            i += 1

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path



def _collect_answer_key(d: Doc) -> list[tuple[str, str]]:
    """从参考答案里抽出「题号 → 简短答案」，用于答案页顶部的速查表。"""
    key: list[tuple[str, str]] = []
    in_ans = False
    for s in d.sections:
        if re.match(r"^(参考答案|答案与解析|答案|解析)", sec_key(s.title or "")):
            in_ans = True
            continue
        if not in_ans:
            continue
        for b in s.blocks:
            text = (b.text or "").strip()
            m = re.match(r"^(\d{1,2})\s*[.．、]\s*(.*)$", text)
            if not m:
                continue
            num, rest = m.group(1), m.group(2)
            ans = rest
            mm = re.match(r"^[【\[]?\s*答案\s*[】\]]?\s*[:：]?\s*(.*)$", rest)
            if mm:
                ans = mm.group(1)
            elif "【" in rest:
                ans = rest.split("【")[0]
            ans = ans.strip("　 。；;")
            if ans and len(ans) <= 24:
                key.append((num, ans))
    return key


def _answer_key_table(doc, key: list[tuple[str, str]], per_row: int = 8):
    """答案速查表：题号行 + 答案行，两行一组（真卷答案页常见）。"""
    for start in range(0, len(key), per_row):
        chunk = key[start:start + per_row]
        cols = len(chunk)
        table = doc.add_table(rows=2, cols=cols)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _table_borders(table)
        for ci, (num, ans) in enumerate(chunk):
            for ri, val in enumerate((num, ans)):
                cell = table.cell(ri, ci)
                cell.text = ""
                para = cell.paragraphs[0]
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                para.paragraph_format.line_spacing = 1.1
                para.paragraph_format.space_before = Pt(1)
                para.paragraph_format.space_after = Pt(1)
                _style_run(para.add_run(val), CN_BODY, EN_BODY, 11,
                           bold=(ri == 0))
        doc.add_paragraph().paragraph_format.space_after = Pt(2)


def _tidy(text: str) -> str:
    """避免 '（　　）' 孤行：括号内用不换行空格，并与前一个汉字绑定。"""
    text = re.sub(r"（[\s\u3000]*）", "（\u00a0\u00a0）", text)
    text = re.sub(r"(?<=[\u4e00-\u9fa5])\s*（\u00a0", "\u00a0（\u00a0", text)
    return text


def _option_row(doc, row: list[Block]):
    """同行 4 项用无框表格排，宽度不够就 2 项一行，保证竖向对齐、不断词。"""
    def cw(t: str) -> float:
        return sum(1.0 if ord(c) > 127 else 0.55 for c in t)

    total = sum(cw(b.text.strip()) for b in row) + 2.0 * (len(row) - 1)
    # 真卷实测：短选项 2×2；含公式/长选项单列竖排（物理月考卷每项一行）
    maxw = max((cw(b.text.strip()) for b in row), default=0)
    if len(row) <= 2 or maxw > 22:
        per = 1
    elif total <= 38 and len(row) == 4:
        per = 4
    else:
        per = 2
    usable = doc.sections[0].page_width - doc.sections[0].left_margin - doc.sections[0].right_margin
    rows = [row[i:i + per] for i in range(0, len(row), per)]
    table = doc.add_table(rows=len(rows), cols=per)
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblPr.append(layout)
    for ri, r in enumerate(rows):
        for ci in range(per):
            cell = table.cell(ri, ci)
            cell.width = int(usable / per)
            para = cell.paragraphs[0]
            para.paragraph_format.space_after = Pt(2)
            para.paragraph_format.line_spacing = 1.3
            if ci < len(r):
                _style_run(para.add_run(r[ci].text.strip()), CN_BODY, EN_BODY, 11.5)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def _paper_block(doc, b: Block, keep_next: bool = False):
    if b.kind == "text":
        text = _tidy(b.text.strip())
        if re.match(r"^\d{1,2}[.．、]", text):
            p = _para(doc, "", size=11.5, left_indent=0.72, hanging=0.72, line=1.45,
                      space_after=3, keep_next=keep_next)
        elif OPTION_RE.match(text):
            p = _para(doc, "", size=11.5, left_indent=0.74, hanging=0.74, line=1.35,
                      space_after=2, keep_next=keep_next)
        elif text.startswith("（") and text.endswith("）") and len(text) < 40:
            _para(doc, text, size=11, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
            return
        else:
            p = _para(doc, "", size=12, first_indent=0.74, line=1.5, space_after=3)
        _write_runs(p, text, 12 if not OPTION_RE.match(text) else 11.5)
    elif b.kind == "num":
        p = _para(doc, "", size=11.5, left_indent=0.72, hanging=0.72, line=1.45,
                  space_after=3, keep_next=keep_next)
        _write_runs(p, b.text, 11.5)
    elif b.kind == "answer":
        p = _para(doc, "", size=11.5, left_indent=0.4, line=1.5, space_after=2)
        _shade(p)
        _write_runs(p, b.text, 11.5, color=ANSWER_INK)
    elif b.kind in ("tip", "insight", "answer"):
        _render_blocks(doc, [b], body_size=11.5)
    elif b.kind == "bullet":
        p = _para(doc, "", size=11.5, left_indent=0.74, hanging=0.74, line=1.45, space_after=2)
        _style_run(p.add_run("▪ "), CN_BODY, EN_BODY, 10.5)
        _write_runs(p, b.text, 11.5)
    elif b.kind in ("table", "image"):
        _render_blocks(doc, [b], body_size=11.5)
    elif b.kind == "break":
        doc.add_page_break()
    else:
        _render_blocks(doc, [b], body_size=11)
