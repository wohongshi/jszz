"""把学科网预览页（SVG）还原成可编辑 docx / 图片。

原理：预览页是原始 docx 的矢量渲染——
  * 文字是 <text x="..." y="..." font-size="...">…</text>
  * 公式、几何图是 <image x= y= width= height= xlink:href="data:image/png;base64,…">

两个坑（都踩过）：
  1. `x` 常常是**逐字形坐标列表**（"172.3 182.9 193.4 …"），只取第一个数字才是行首；
  2. `font-size` 往往挂在父级 `<g>` 上，`<text>` 自己没写 —— 必须向下继承。

这不是原文件——原始 docx 在学科网要登录+点数下载。这里产出的是**还原版**：
文字可编辑、公式与图形是位图，适合老师直接用或当讲评课素材。
"""
from __future__ import annotations

import base64
import io
import re
from dataclasses import dataclass, field
from pathlib import Path

SVG_W, SVG_H = 793.0, 1122.0     # A4 @96dpi
INLINE_MAX_H = 26.0              # 低于这个高度（SVG 单位）当行内公式

TAG_RE = re.compile(r"<(/?)(g|text|image|svg)\b([^>]*?)(/?)>", re.S)


@dataclass
class Run:
    x: float
    y: float
    size: float
    text: str


@dataclass
class Pic:
    x: float
    y: float
    w: float
    h: float
    data: bytes
    fmt: str = "png"


@dataclass
class Page:
    runs: list[Run] = field(default_factory=list)
    pics: list[Pic] = field(default_factory=list)

    def lines(self, tol: float = 6.0) -> list[dict]:
        rows: list[list[Run]] = []
        for r in sorted(self.runs, key=lambda r: (r.y, r.x)):
            if r.y <= 1.0:                      # 无坐标的隐藏层
                continue
            if "学科网（北京）" in r.text:        # 页脚版权水印
                continue
            if rows and abs(r.y - rows[-1][0].y) <= tol:
                rows[-1].append(r)
            else:
                rows.append([r])
        out = []
        for row in rows:
            row.sort(key=lambda r: r.x)
            text = "".join(r.text for r in row)
            if re.fullmatch(r"[\d\s页/共第]+", text or ""):   # "1 页 / 共 5 页"
                continue
            out.append({"y": row[0].y, "x": min(r.x for r in row),
                        "size": max(r.size for r in row), "text": text})
        return out


def _first_num(value: str, default: float = 0.0) -> float:
    m = re.search(r"-?\d+(?:\.\d+)?", value or "")
    return float(m.group(0)) if m else default


def _attr(tag_attrs: str, name: str) -> str:
    m = re.search(rf'\b{name}="([^"]*)"', tag_attrs)
    return m.group(1) if m else ""


def _translate(transform: str) -> tuple[float, float]:
    m = re.search(r"translate\(\s*(-?[\d.]+)[,\s]+(-?[\d.]+)", transform or "")
    return (float(m.group(1)), float(m.group(2))) if m else (0.0, 0.0)


def _decode_text(raw: str) -> str:
    t = re.sub(r"<[^>]+>", "", raw)
    return (t.replace("&#160;", " ").replace("&nbsp;", " ")
             .replace("&#10;", "").replace("&#13;", "")
             .replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&#39;", "'").replace("&amp;", "&"))


def _decode_img(raw: str):
    if not raw.startswith("data:"):
        return None
    head, _, b64 = raw.partition(",")
    fmt = "jpg" if ("jpeg" in head or "jpg" in head) else "png"
    b64 = b64.replace("\n", "").replace("\r", "").replace("&#10;", "").replace("&#13;", "")
    try:
        return base64.b64decode(b64), fmt
    except Exception:
        return None


def parse_svg(svg: str) -> Page:
    """带继承栈的解析：font-size / transform 从 <g> 往下传。"""
    page = Page()
    stack: list[dict] = [{"size": 12.0, "tx": 0.0, "ty": 0.0}]
    for m in TAG_RE.finditer(svg):
        closing, tag, attrs, selfclose = m.group(1), m.group(2), m.group(3), m.group(4)
        top = stack[-1]

        if tag == "g":
            if closing:
                if len(stack) > 1:
                    stack.pop()
            elif not selfclose:
                tx, ty = _translate(_attr(attrs, "transform"))
                size = _first_num(_attr(attrs, "font-size"), 0.0) or top["size"]
                stack.append({"size": size, "tx": top["tx"] + tx, "ty": top["ty"] + ty})
            continue

        if tag not in ("text", "image"):
            continue

        if tag == "text":
            end = svg.find("</text>", m.end())
            inner = svg[m.end():end] if end > 0 else ""
            text = _decode_text(inner)
            if not text.strip():
                continue
            page.runs.append(Run(
                _first_num(_attr(attrs, "x")) + top["tx"],
                _first_num(_attr(attrs, "y")) + top["ty"],
                _first_num(_attr(attrs, "font-size"), 0.0) or top["size"],
                text,
            ))
            continue

        href = re.search(r'(?:xlink:)?href="([^"]+)"', attrs)
        if not href:
            continue
        got = _decode_img(href.group(1))
        if not got:
            continue
        data, fmt = got
        page.pics.append(Pic(
            _first_num(_attr(attrs, "x")) + top["tx"],
            _first_num(_attr(attrs, "y")) + top["ty"],
            _first_num(_attr(attrs, "width"), 10),
            _first_num(_attr(attrs, "height"), 10),
            data, fmt,
        ))
    return page


def _flow(page: Page) -> list[dict]:
    """一页压成「行」序列：文字与行内小图按 y 归行，行内按 x 排序。"""
    entries: list[tuple[float, float, str, object]] = []
    for r in page.runs:
        if r.y <= 1.0:      # y=0 的是一层无坐标的隐藏文本（含页脚水印），丢掉
            continue
        entries.append((r.y, r.x, "text", r))
    for pic in page.pics:
        if pic.h <= INLINE_MAX_H and pic.y > 1.0:
            entries.append((pic.y + pic.h / 2, pic.x, "pic", pic))
    entries.sort(key=lambda e: (e[0], e[1]))
    bands: list[dict] = []
    tol = 10.0        # 行距约 31px、字号 12–21px，10px 既能合并上下标又不会跨行
    for y, x, kind, obj in entries:
        if bands and abs(y - bands[-1]["y"]) <= tol:
            bands[-1]["items"].append((kind, obj))
            bands[-1]["x"] = min(bands[-1]["x"], x)
            if kind == "text":
                bands[-1]["size"] = max(bands[-1]["size"], obj.size)
        else:
            bands.append({"y": y, "x": x,
                          "size": obj.size if kind == "text" else 12.0,
                          "items": [(kind, obj)]})
    for b in bands:
        b["items"].sort(key=lambda it: it[1].x)
    return bands


def build_docx(pages: list[Page], out_path: str | Path, *, title: str = "",
               footer: str = "") -> Path:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.5)
    sec.top_margin = sec.bottom_margin = Cm(1.5)
    usable_cm = (sec.page_width - sec.left_margin - sec.right_margin) / 360000.0
    px2cm = 2.54 / 96.0                     # 1 SVG px = 1/96 inch
    fit = min(1.0, usable_cm / (SVG_W * px2cm))     # 页宽不够时整体缩一点

    base = doc.styles["Normal"]
    base.font.name = "Times New Roman"
    base.font.size = Pt(10.5)
    base.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")

    if title:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.font.size = Pt(14)
        r.font.bold = True
        r.font.name = "Times New Roman"
        r._element.get_or_add_rPr().append(_ea("黑体"))

    for pi, page in enumerate(pages):
        if pi:
            doc.add_page_break()
        last_y = 0.0
        arts = []
        for band in _flow(page):
            if band["y"] < 12:
                arts.append(band)
                continue
            gap_cm = max(0.0, (band["y"] - last_y - band["size"]) * px2cm * fit)
            p = doc.add_paragraph()
            pf = p.paragraph_format
            pf.space_before = Pt(min(10.0, gap_cm * 28.35 * 0.9))
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            indent_cm = band["x"] * px2cm * fit
            if indent_cm > 0.08:
                pf.left_indent = Cm(min(indent_cm, 2.6))
            for kind, obj in band["items"]:
                if kind == "text":
                    run = p.add_run(obj.text)
                    run.font.size = Pt(max(7.0, min(20.0, obj.size * 0.75)))
                else:
                    run_pt = (obj.h * px2cm * fit * 28.35)
                    h_cm = max(0.14, min(0.42, run_pt / 28.35))
                    try:
                        p.add_run().add_picture(io.BytesIO(obj.data), height=Cm(h_cm))
                    except Exception:
                        pass
            last_y = band["y"]

        for band in arts:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(4)
            r = p.add_run("".join(o.text for k, o in band["items"] if k == "text").strip())
            r.font.size = Pt(7.5)
            r.font.color.rgb = RGBColor(0x99, 0x99, 0x99)

        for pic in sorted(page.pics, key=lambda q: q.y):
            if pic.y <= 1.0 or pic.h <= INLINE_MAX_H or pic.w < 140:   # 窄图多为公式碎片，已按行内处理
                continue
            w_cm = max(0.6, min(usable_cm, pic.w * px2cm * fit))
            try:
                p = doc.add_paragraph()
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(1)
                p.add_run().add_picture(io.BytesIO(pic.data), width=Cm(w_cm))
            except Exception:
                continue

    if footer:
        p = doc.sections[0].footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(footer)
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def _ea(font: str):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    el = OxmlElement("w:rFonts")
    el.set(qn("w:eastAsia"), font)
    return el


def save_images(pages: list[Page], out_dir: str | Path, prefix: str = "p") -> list[str]:
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    saved: list[str] = []
    for pi, page in enumerate(pages, 1):
        for k, pic in enumerate(page.pics, 1):
            if len(pic.data) < 200:
                continue
            f = out / f"{prefix}{pi:02d}_{k:02d}.{pic.fmt}"
            f.write_bytes(pic.data)
            saved.append(str(f))
    return saved


def svg_text(svg: str) -> str:
    """纯文本提取（保持行序），供 zxxk.preview_text 使用。"""
    return "\n".join(ln["text"].strip() for ln in parse_svg(svg).lines() if ln["text"].strip())
