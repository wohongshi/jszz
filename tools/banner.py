#!/usr/bin/env python3
"""README 横幅：1280x380 深色，2x 画布后缩。右：六科真实课件封面圆片。"""
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import glob

S = 2
W, H = 1280 * S, 380 * S
BG = (17, 22, 30)
im = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(im)

# 噪点（防大面积纯色出带宽）
noise = Image.effect_noise((W // 2, H // 2), 12).convert("L")
im.paste(Image.merge("RGB", [noise.point(lambda v: 30 + v // 12)] * 3), (0, 0))

# 左侧装饰竖线
d.rectangle([70 * S, 60 * S, 74 * S, 320 * S], fill=(230, 90, 40))

def F(path, size):
    return ImageFont.truetype(path, size * S)

SERIF = "/usr/share/fonts/opentype/noto/NotoSansCJK-Bold.ttc"
if not glob.glob(SERIF):
    import subprocess
    SERIF = subprocess.run(["fc-match", "-f", "%{file}", "Noto Sans CJK SC:bold"],
                           capture_output=True, text=True).stdout.strip()

f_tag = F(SERIF, 13)
f_title = F(SERIF, 62)
f_sub = F(SERIF, 21)
f_small = F(SERIF, 13)

x = 110 * S
d.text((x, 58 * S), "MCP SERVER  ·  SKILL PACK  ·  FOR AI AGENTS",
       font=f_tag, fill=(120, 140, 170))

title = "jszz"
d.text((x, 88 * S), title, font=f_title, fill=(240, 244, 250))
tw = d.textlength(title, font=f_title)

# 标题右的中文副标（自适应缩放，不压右侧圆片区）
sub = "教师备课工具链"
avail = 320 * S
fs = 21
while d.textlength(sub, font=F(SERIF, fs)) > avail and fs > 14:
    fs -= 1
d.text((x + tw + 28 * S, 128 * S), sub, font=F(SERIF, fs), fill=(150, 170, 200))

d.text((x, 205 * S), "B站音视频 · 学科网取材 · 学案/试卷/课件生成",
       font=f_sub, fill=(200, 210, 225))
d.text((x, 242 * S), "九科真实课件视觉 · 公式 OCR · 一键讲评",
       font=f_sub, fill=(200, 210, 225))
d.text((x, 285 * S), "837 份真实教师资料逆向出的排版规范 · 22 个 MCP 工具 · 10 篇写作规范",
       font=f_small, fill=(120, 140, 170))
d.text((x, 318 * S), "MIT License · python 3.10+ · 一条命令起服务",
       font=f_small, fill=(100, 120, 150))

# 右侧：九科真实封面圆片（轻微重叠，右移缩小，不压文字）
covers = sorted(glob.glob("docs/shots/subjects/*.png"))
if covers:
    n = min(6, len(covers))
    r = 46 * S
    cy = 155 * S
    step = int(1.62 * r)
    x0 = W - 70 * S - (2 * r + step * (n - 1))
    for i, p in enumerate(covers[:n]):
        cx = x0 + step * i
        face = Image.open(p).convert("RGB")
        side = min(face.size)
        face = face.crop(((face.width - side) // 2, (face.height - side) // 2,
                          (face.width + side) // 2, (face.height + side) // 2))
        face = face.resize((2 * r, 2 * r), 1)
        mask = Image.new("L", (2 * r, 2 * r), 0)
        ImageDraw.Draw(mask).ellipse([0, 0, 2 * r - 1, 2 * r - 1], fill=255)
        ring = Image.new("RGB", (2 * r + 8 * S, 2 * r + 8 * S), BG)
        ImageDraw.Draw(ring).ellipse([0, 0, 2 * r + 8 * S - 1, 2 * r + 8 * S - 1],
                                     fill=(70, 85, 110))
        im.paste(ring, (cx - 4 * S, cy - 4 * S))
        im.paste(face, (cx, cy), mask)

    cap = "6 科 · 每科一套真实课件视觉"
    cw = d.textlength(cap, font=f_small)
    d.text((W - 70 * S - cw, cy + 2 * r + 22 * S), cap, font=f_small,
           fill=(120, 140, 170))

im = im.resize((1280, 380), 1)
im.save("docs/shots/banner.png")
print("saved docs/shots/banner.png", im.size)
