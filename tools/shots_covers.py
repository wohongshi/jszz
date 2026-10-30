#!/usr/bin/env python3
"""九科真实渲染封面圆片 + 拼图（banner.py 和 README 截图的素材源）。

用法：./tools/py tools/shots_covers.py
产物：docs/shots/subjects/<学科>-01.png（9 张）+ docs/shots/subjects-covers.png（拼图）
"""
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "server"))

from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402
from PIL import Image  # noqa: E402

from jszz.pptx_render import build_xuean_deck  # noqa: E402

OUT = ROOT / "docs/shots/subjects"
TMP = Path("/tmp/shots-covers")
MD = """# {t}
> 学科：{s} | 年级：高二 | 课时：1

## 【学习目标】
- 目标一。

## 【知识梳理】
正文。
"""

TITLES = {
    "语文": "一轮复习　文言文阅读", "数学": "第4节　函数的对称性",
    "英语": "话题1　学校及家庭生活", "物理": "第02讲　匀变速直线运动",
    "化学": "14. 反应热　热化学方程式", "生物": "第4讲　细胞膜与细胞核",
    "政治": "必修四　认识社会与价值选择", "历史": "第9课　古代的商路、贸易与文化交流",
    "地理": "编号11　热力环流和大气水平运动",
}


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    TMP.mkdir(parents=True, exist_ok=True)
    pngs = []
    for subj, title in TITLES.items():
        md = MD.format(s=subj, t=title)
        pptx = TMP / f"{subj}.pptx"
        build_xuean_deck(md, pptx, school="胶州四中")
        subprocess.run(["soffice", "--headless",
                        "-env:UserInstallation=file:///tmp/shots-lo",
                        "--convert-to", "pdf", "--outdir", str(TMP), str(pptx)],
                       capture_output=True, timeout=150)
        pdf = TMP / f"{subj}.pdf"
        subprocess.run(["pdftoppm", "-r", "90", "-png", "-f", "1", "-l", "1",
                        str(pdf), str(TMP / subj)], capture_output=True, timeout=60)
        # pdftoppm 补零位数随总页数变（<10 页出 -1.png，多页出 -01.png）——都认
        cands = sorted(TMP.glob(f"{subj}-*.png"))
        src = cands[0] if cands else None
        if src is None:
            print("MISS", subj)
            continue
        dst = OUT / f"{subj}-01.png"
        dst.write_bytes(src.read_bytes())
        pngs.append(dst)
        print("封面:", dst.name)

    # 拼图 3x3（banner.py 与 README 共用）
    cell_w = 420
    thumbs = []
    for p in pngs:
        im = Image.open(p).convert("RGB")
        r = cell_w / im.width
        thumbs.append(im.resize((cell_w, int(im.height * r)), 1))
    ch = max(t.height for t in thumbs)
    sheet = Image.new("RGB", (cell_w * 3 + 40, ch * 3 + 40), (24, 28, 36))
    for i, t in enumerate(thumbs):
        x = 10 + (i % 3) * (cell_w + 10)
        y = 10 + (i // 3) * (ch + 10)
        sheet.paste(t, (x, y))
    out = ROOT / "docs/shots/subjects-covers.png"
    sheet.save(out)
    print("拼图:", out, sheet.size)


if __name__ == "__main__":
    main()
