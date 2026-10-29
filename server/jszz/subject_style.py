"""六科各一套视觉语言（v2：按 U 盘 17 份真实课件逐页取证重建）。

取证方法：每科 2–3 份真实课件（早读/常规课/习题讲评）转 PDF→PNG 逐页目视
+ slide XML 全量统计（配色/字号/形状密度/页头形态）。

真实课件的公共约定（17/17 白底、正文 28pt 一统、红=答案/纠错、蓝=标题/要点），
各科差异化的部分（封面/页头/装饰/强调件）按取证复刻：

  语文：象牙白纸感 + 顶部提示横条 + 黄底篇目标签；红=篇目/时间，蓝=答案（下划线）；楷体
  数学：白底 + 圆角细线框栏目 + 黄底公式卡 + 红色步骤标签（设事件/写概率/代公式）；宋体+Times
  英语：青色块栏目标题【课前背记】+ 黄色 Learning aims 块 + 右下双胶囊标签；红=答案/译文
  物理：深蓝渐变封面 + 轨道弧线；页码方块导航 + 红「例N」标签 + 蓝「解析」竖标签
  化学：浅灰旋涡封面 + 深蓝粗标题 + 实物照片卡；蓝色胶囊标题 + 红手写批注/√×
  生物：绿色块栏目名（从社会中来/练一练）+ 放射概念图 + 黄荧光必背 + 红字解析框
"""
from __future__ import annotations

from dataclasses import dataclass

from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
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
    cover: str          # paper | tags | aims | orbit | photo | radial | split
    header: str         # banner | frame | cyanblock | nav | capsule | greenblock
    deco: str           # leaf | grid | duo | arc | swirl | rings
    numerals: str = ""

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
        return rgb("FBF8F2") if self.key == "chinese" else rgb("FFFFFF")


STYLES: dict[str, Style] = {
    # 语文：象牙白 + 赭红（9F5647 是真实课件里的实测色）+ 黄底标签；楷体正文
    "chinese": Style(
        key="chinese", label="语文", main_hex="3B3226", accent_hex="9F5647", warn_hex="C00000",
        light_hex="F7F2E8", ink_hex="2B2620", gold_hex="B8893A",
        font_cn="楷体", font_head="黑体", cover="paper", header="banner",
        deco="leaf", numerals="壹贰叁肆伍陆柒捌玖拾",
    ),
    # 数学：深蓝 + 亮蓝 + 红步骤标签 + 黄底公式卡；宋体 + Times New Roman
    "math": Style(
        key="math", label="数学", main_hex="1F3864", accent_hex="2F6FED", warn_hex="C00000",
        light_hex="EEF3FB", ink_hex="1A2233", gold_hex="E8B21A",
        font_cn="宋体", font_head="黑体", cover="tags", header="frame",
        deco="grid",
    ),
    # 英语：青色栏目块 + 黄色高亮 + 讲评紫；微软雅黑
    "english": Style(
        key="english", label="英语", main_hex="0B6E75", accent_hex="00A5B5", warn_hex="C00000",
        light_hex="E8F6F7", ink_hex="16323A", gold_hex="5C52BA",
        font_cn="微软雅黑", font_head="微软雅黑", cover="aims", header="cyanblock",
        deco="duo",
    ),
    # 物理：深海蓝 + 琥珀 + 深蓝/深红标注；宋体 + Times New Roman
    "physics": Style(
        key="physics", label="物理", main_hex="1B3A6B", accent_hex="2F5597", warn_hex="C00000",
        light_hex="EDF1F7", ink_hex="16212F", gold_hex="E8A33D",
        font_cn="宋体", font_head="黑体", cover="orbit", header="nav",
        deco="arc",
    ),
    # 化学：深蓝粗标题 + 蓝胶囊 + 红手写批注；宋体 + Times New Roman
    "chemistry": Style(
        key="chemistry", label="化学", main_hex="1F4E79", accent_hex="2E75B6", warn_hex="C00000",
        light_hex="F1F5F9", ink_hex="1C2B3A", gold_hex="B7791F",
        font_cn="宋体", font_head="黑体", cover="photo", header="capsule",
        deco="swirl",
    ),
    # 生物：绿块栏目 + 黄荧光必背 + 红字解析框；微软雅黑
    "biology": Style(
        key="biology", label="生物", main_hex="3F7A52", accent_hex="70AD47", warn_hex="C00000",
        light_hex="EAF3EC", ink_hex="1D2B22", gold_hex="B7791F",
        font_cn="微软雅黑", font_head="微软雅黑", cover="radial", header="greenblock",
        deco="rings",
    ),
    # 通用（无学科信息时的默认）
    "general": Style(
        key="general", label="通用", main_hex="1F4E79", accent_hex="2E75B6", warn_hex="C00000",
        light_hex="EEF3F8", ink_hex="1B2430", gold_hex="E8B21A",
        font_cn="微软雅黑", font_head="微软雅黑", cover="split", header="nav",
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


# 六科关键词 → 风格键，用于从标题/正文猜学科（猜不出返回 ""）
_GUESS = [
    ("物理", "physics"), ("数学", "math"), ("英语", "english"),
    ("化学", "chemistry"), ("生物", "biology"), ("语文", "chinese"),
]


def guess(*texts: str) -> str:
    """从标题/正文里猜学科键；md 忘写 `> 学科：` 时兜底，别再让物理卷戴语文口令。"""
    blob = " ".join(t for t in texts if t)
    for kw, key in _GUESS:
        if kw in blob:
            return key
    return ""


# ------------------------------------------------------------------ 装饰母题
def draw_deco(slide, s: Style) -> None:
    """按学科画角落装饰（很淡，不抢内容）。位置都避开页头文字区。"""
    kind = s.deco
    if kind == "grid":
        # 数学：右上细网格（真实课件的坐标纸感）
        for i in range(5):
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          W - Inches(2.1) + Inches(0.42) * i,
                                          Inches(0.0), Emu(12700), Inches(1.55))
            line.fill.solid()
            line.fill.fore_color.rgb = rgb("EAF0FA")
            line.line.fill.background()
            line.shadow.inherit = False
        for i in range(4):
            line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          W - Inches(2.1), Inches(0.42) * i,
                                          Inches(2.1), Emu(12700))
            line.fill.solid()
            line.fill.fore_color.rgb = rgb("EAF0FA")
            line.line.fill.background()
            line.shadow.inherit = False
    elif kind == "leaf":
        # 语文：右上浅色叶纹水印（真实早读课件的叶片装饰）
        for i, (x, y, w, h, rot) in enumerate([
                (12.05, 0.22, 0.85, 0.34, -28), (12.55, 0.55, 0.7, 0.28, 12),
                (11.85, 0.72, 0.62, 0.24, -8)]):
            leaf = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                          Inches(w), Inches(h))
            leaf.rotation = rot
            leaf.fill.solid()
            leaf.fill.fore_color.rgb = rgb("EDEFE3")
            leaf.line.fill.background()
            leaf.shadow.inherit = False
    elif kind == "duo":
        # 英语：右上青/黄双色小胶囊（呼应青色栏目块 + 黄色高亮）
        for i, (c, w) in enumerate([("BFE8EA", 1.05), ("FBF3C4", 0.72)]):
            cap = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                         W - Inches(2.35 - 1.2 * i), Inches(0.28 + 0.5 * i),
                                         Inches(w), Inches(0.2))
            cap.fill.solid()
            cap.fill.fore_color.rgb = rgb(c)
            cap.line.fill.background()
            cap.shadow.inherit = False
    elif kind == "arc":
        # 物理：右上细弧线（呼应封面轨道弧线）
        for i, (d, x, y) in enumerate([(2.6, 10.45, 0.28), (1.9, 10.85, 0.62)]):
            ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                          Inches(d), Inches(d))
            ring.fill.background()
            ring.line.color.rgb = rgb("E3EAF3")
            ring.line.width = Pt(1.25)
            ring.shadow.inherit = False
    elif kind == "swirl":
        # 化学：右上浅灰旋涡弧（呼应封面旋涡背景）
        for i, (d, x, y) in enumerate([(2.9, 10.15, 0.18), (2.1, 10.55, 0.55),
                                       (1.3, 10.95, 0.95)]):
            ring = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x), Inches(y),
                                          Inches(d), Inches(d))
            ring.fill.background()
            ring.line.color.rgb = rgb("E8EDF2")
            ring.line.width = Pt(1.1)
            ring.shadow.inherit = False
    elif kind == "rings":
        # 生物：右上同心圆（细胞膜感）
        for d, a in [(1.7, 0.55), (1.25, 0.72), (0.8, 0.85)]:
            ring = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                          W - Inches(d + 0.55), Inches(1.35),
                                          Inches(d), Inches(d))
            ring.fill.background()
            ring.line.color.rgb = rgb("DCEBE0")
            ring.line.width = Pt(1.25)
            ring.shadow.inherit = False


# ------------------------------------------------------------------ 字体兜底
_SERIF_PREF = ["宋体", "SimSun", "华文中宋", "STSong", "Noto Serif CJK SC", "Source Han Serif SC"]
_SANS_PREF = ["微软雅黑", "Microsoft YaHei", "黑体", "SimHei", "Noto Sans CJK SC",
              "Source Han Sans SC"]
_KAI_PREF = ["楷体", "KaiTi", "STKaiti", "Noto Serif CJK SC", "Source Han Serif SC"]
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
    """本机没有这个字体就退到同类字形（宋/楷→衬线类，黑/雅黑→黑体类），保证渲染不塌。"""
    if not name:
        return "Noto Sans CJK SC"
    if name in _PICKED:
        return _PICKED[name]
    avail = _available()
    chosen = name
    if avail and name not in avail:
        if "楷" in name:
            chain = _KAI_PREF
        elif any(k in name for k in ("宋", "姚", "Song", "Serif")):
            chain = _SERIF_PREF
        else:
            chain = _SANS_PREF
        chosen = next((f for f in chain if f in avail), name)
    _PICKED[name] = chosen
    return chosen
