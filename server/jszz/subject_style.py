"""六科各一套视觉语言：配色 / 字体 / 封面 / 页头 / 装饰母题 / 专属版式。

设计依据（来自 U 盘里各科真实课件的逐页拆解）：
  语文：宣纸米白 + 墨黑 + 朱砂；竖排标题、引文居中、朱批式强调
  数学：蓝灰 + 细网格；编号块、考点 chip、公式居中大字
  英语：靛青 + 洋红；左右英汉对照、词块表
  物理：深海蓝 + 琥珀；刻度尺装饰、已知/求/解三段、矢量箭头
  化学：墨绿 + 橙；分子点阵、反应式（条件写在箭头上）
  生物：青绿 + 珊瑚；四栏栏目页头（精读教材/教考衔接/练透考法/归纳拓展）、概念图
"""
from __future__ import annotations

from dataclasses import dataclass

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Emu, Inches, Pt

W, H = Inches(13.333), Inches(7.5)


def rgb(h: str) -> RGBColor:
    h = h.lstrip("#")
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


@dataclass
class Style:
    key: str
    label: str
    main_hex: str
    accent_hex: str
    warn_hex: str
    light_hex: str
    ink_hex: str
    gold_hex: str
    font_cn: str
    font_head: str
    cover: str          # vertical | grid | split | ruler | molecule | cell
    header: str         # tick | bar | bracket | band | ribbon
    deco: str           # ink | grid | dual | ruler | molecule | cell
    numerals: str = ""  # 语文用「壹贰叁」

    # ---- 以 RGBColor 形式暴露给渲染器 ----
    @property
    def main(self):
        return rgb(self.main_hex)

    @property
    def accent(self):
        return rgb(self.accent_hex)

    @property
    def warn(self):
        return rgb(self.warn_hex)

    @property
    def light(self):
        return rgb(self.light_hex)

    @property
    def ink(self):
        return rgb(self.ink_hex)

    @property
    def gold(self):
        return rgb(self.gold_hex)

    @property
    def white(self):
        return rgb("FFFFFF")

    @property
    def paper(self):
        return rgb("FBF7F0") if self.key == "chinese" else rgb("FFFFFF")


STYLES: dict[str, Style] = {
    "chinese": Style(
        key="chinese", label="语文", main_hex="1C1B19", accent_hex="A6322A", warn_hex="A6322A",
        light_hex="F3EDE3", ink_hex="23211E", gold_hex="B8893A",
        font_cn="宋体", font_head="宋体", cover="vertical", header="bracket",
        deco="ink", numerals="壹贰叁肆伍陆柒捌玖拾",
    ),
    "math": Style(
        key="math", label="数学", main_hex="0B1F3A", accent_hex="2F6FED", warn_hex="C2410C",
        light_hex="EAF0FB", ink_hex="16202E", gold_hex="E8B21A",
        font_cn="微软雅黑", font_head="微软雅黑", cover="grid", header="tick",
        deco="grid",
    ),
    "english": Style(
        key="english", label="英语", main_hex="1B2559", accent_hex="D6336C", warn_hex="B42318",
        light_hex="F1F3FF", ink_hex="1A1F36", gold_hex="E8B21A",
        font_cn="微软雅黑", font_head="微软雅黑", cover="split", header="ribbon",
        deco="dual",
    ),
    "physics": Style(
        key="physics", label="物理", main_hex="10233F", accent_hex="E8A33D", warn_hex="B42318",
        light_hex="EDF1F7", ink_hex="16212F", gold_hex="E8A33D",
        font_cn="微软雅黑", font_head="微软雅黑", cover="ruler", header="bar",
        deco="ruler",
    ),
    "chemistry": Style(
        key="chemistry", label="化学", main_hex="0E4F45", accent_hex="E8710A", warn_hex="B42318",
        light_hex="E9F0F2", ink_hex="122A26", gold_hex="E8B21A",
        font_cn="微软雅黑", font_head="微软雅黑", cover="molecule", header="bar",
        deco="molecule",
    ),
    "biology": Style(
        key="biology", label="生物", main_hex="0B6B5F", accent_hex="12A48C", warn_hex="C2410C",
        light_hex="E9F6F2", ink_hex="102622", gold_hex="E8B21A",
        font_cn="微软雅黑", font_head="微软雅黑", cover="cell", header="band",
        deco="cell",
    ),
    # 通用（无学科信息时的默认）
    "general": Style(
        key="general", label="通用", main_hex="0F2B55", accent_hex="2E86D6", warn_hex="C81E1E",
        light_hex="EAF2FB", ink_hex="1B2430", gold_hex="E8B21A",
        font_cn="微软雅黑", font_head="微软雅黑", cover="split", header="tick",
        deco="grid",
    ),
}

# 显式 --theme 参数仍可用，映射到风格
ALIASES = {
    "blue": "general", "teal": "biology", "violet": "english",
    "orange": "chemistry", "ink": "chinese",
}

SUBJECT_KEYS = {
    "语文": "chinese", "数学": "math", "英语": "english",
    "物理": "physics", "化学": "chemistry", "生物": "biology",
    "chinese": "chinese", "math": "math", "english": "english",
    "physics": "physics", "chemistry": "chemistry", "biology": "biology",
}


def get(name: str | None) -> Style:
    """按学科名或主题名取风格；认不出就用通用。"""
    if not name:
        return STYLES["general"]
    key = SUBJECT_KEYS.get(name.strip()) or ALIASES.get(name.strip()) or name.strip()
    return STYLES.get(key, STYLES["general"])


# ------------------------------------------------------------------ 装饰母题
def draw_deco(slide, s: Style) -> None:
    """按学科画角落装饰（很淡，不抢内容）。"""
    kind = s.deco
    if kind == "grid":
        for i in range(5):
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          W - Inches(2.1) + Inches(0.42) * i,
                                          Inches(0.0), Emu(12700), Inches(1.55))
            line.fill.solid()
            line.fill.fore_color.rgb = rgb("E4EBF7")
            line.line.fill.background()
            line.shadow.inherit = False
        for i in range(4):
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          W - Inches(2.1), Inches(0.42) * i,
                                          Inches(2.1), Emu(12700))
            line.fill.solid()
            line.fill.fore_color.rgb = rgb("E4EBF7")
            line.line.fill.background()
            line.shadow.inherit = False
    elif kind == "ruler":
        # 右上角小刻度（内页也放得下，不跟底部色带打架）
        y = Inches(1.06) if False else Inches(0.0)
        x0, x1 = W - Inches(2.3), W - Inches(0.5)
        y = Inches(1.02)
        base = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, x0, y, x1 - x0, Emu(9525))
        base.fill.solid()
        base.fill.fore_color.rgb = rgb("D8DEE8")
        base.line.fill.background()
        base.shadow.inherit = False
        for i in range(12):
            h = Inches(0.11) if i % 3 == 0 else Inches(0.055)
            t = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                       x0 + (x1 - x0) / 12 * i, y - h, Emu(9525), h)
            t.fill.solid()
            t.fill.fore_color.rgb = rgb("C6CEDA")
            t.line.fill.background()
            t.shadow.inherit = False
    elif kind == "molecule":
        pts = [(11.6, 1.0), (12.2, 1.55), (11.75, 2.15), (12.35, 2.55), (11.35, 1.85)]
        for i, (x, y) in enumerate(pts):
            r = Inches(0.17 if i % 2 == 0 else 0.11)
            dot = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y), r, r)
            dot.fill.solid()
            dot.fill.fore_color.rgb = rgb("CBE7DF")
            dot.line.fill.background()
            dot.shadow.inherit = False
        for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
            ln = slide.shapes.add_connector(1, Inches(x1 + 0.09), Inches(y1 + 0.09),
                                            Inches(x2 + 0.09), Inches(y2 + 0.09))
            ln.line.color.rgb = rgb("D7EAE4")
            ln.line.width = Pt(1.0)
    elif kind == "cell":
        for i, (d, a) in enumerate([(1.7, 0.55), (1.25, 0.72), (0.8, 0.85)]):
            ring = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                          W - Inches(d + 0.55), Inches(1.35),
                                          Inches(d), Inches(d))
            ring.fill.background()
            ring.line.color.rgb = rgb("D5EDE6")
            ring.line.width = Pt(1.25)
            ring.shadow.inherit = False
    elif kind == "ink":
        band = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, W - Inches(0.55),
                                      Inches(0.5), Inches(0.08), Inches(1.6))
        band.fill.solid()
        band.fill.fore_color.rgb = rgb("D9CDB8")
        band.line.fill.background()
        band.shadow.inherit = False
    elif kind == "dual":
        blk = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                     Inches(0.22), H)
        blk.fill.solid()
        blk.fill.fore_color.rgb = rgb("E6E9FB")
        blk.line.fill.background()
        blk.shadow.inherit = False


# ------------------------------------------------------------------ 字体兜底
_SERIF_PREF = ["宋体", "SimSun", "华文中宋", "STSong", "Noto Serif CJK SC", "Source Han Serif SC"]
_SANS_PREF = ["微软雅黑", "Microsoft YaHei", "黑体", "SimHei", "Noto Sans CJK SC",
              "Source Han Sans SC"]
_AVAIL: set[str] | None = None
_PICKED: dict[str, str] = {}


def _available() -> set[str]:
    global _AVAIL
    if _AVAIL is None:
        try:
            import subprocess
            out = subprocess.run(["fc-list", ":lang=zh", "family"], capture_output=True,
                                 text=True, timeout=8).stdout
            _AVAIL = {x.strip() for part in out.splitlines() for x in part.split(",")}
        except Exception:
            _AVAIL = set()
    return _AVAIL


def pick_font(name: str) -> str:
    """本机没有这个字体就退到同类字形（宋→宋体类，黑→黑体类），保证渲染不塌。"""
    if not name:
        return "Noto Sans CJK SC"
    if name in _PICKED:
        return _PICKED[name]
    avail = _available()
    chosen = name
    if avail and name not in avail:
        serif_like = any(k in name for k in ("宋", "楷", "姚", "Song", "Serif"))
        chain = _SERIF_PREF if serif_like else _SANS_PREF
        chosen = next((f for f in chain if f in avail), name)
    _PICKED[name] = chosen
    return chosen
