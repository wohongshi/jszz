#!/usr/bin/env python3
"""精读真实课件逐页结构：每页标题/字数/形状/图/颜色语义 → analysis/ppt3/。"""
import zipfile
import re
import json
from collections import Counter
from pathlib import Path

BASE = Path("/mnt/usb/11")
OUT = Path(__file__).resolve().parent.parent / "analysis" / "ppt3"
OUT.mkdir(parents=True, exist_ok=True)

JOBS = [
    ("语早读", "语文/10.13早读.pptx"),
    ("语默写", "语文/2026.3.10早自习任务.pptx"),
    ("数函数", "数学/5.3.1 函数的单调性 课件.pptx"),
    ("数计数", "数学/6.1分类加法计数原理与分步乘法计数原理.pptx"),
    ("英写作", "gs1/{2390E21F-A880-4C96-9CE9-A27677B549DA}_10.8-No.9-话题写作案.pptx"),
    ("英检测", "gs1/{1AE6B9FE-4D59-4685-BB77-16D25039FC29}_高三英语一轮-话题1/高三英语一轮-话题1/No.4-B1U1基础检测案.pptx"),
    ("物平抛", "gs1/{75DFFF96-3C34-42B1-8985-8D261AC16B11}_20 实验五：探究平抛运动的特点.pptx"),
    ("物连接", "gs1/{C560018E-A421-4B53-B768-9D4017B02B4E}_第三章　第13课时　专题强化：动力学中的连接体和临界、极值问题.pptx"),
    ("化电源", "gszbjs1/{038775A9-6A34-4923-ACB6-4D0801D2F65C}_18.新型化学电源.pptx"),
    ("化反应", "化学/{B5D450A1-D434-4D9C-BC6C-748794A31F3D}_14.　反应热　热化学方程式.pptx"),
    ("生背诵", "生物/3.2+基因工程的基本操作程序 1背诵.pptx"),
    ("生讲评", "生物/习题讲评4.13.pptx"),
    ("史一轮", "历史/一轮复习第11课.pptx"),
    ("史主观", "历史/2022·山东卷第18题“《工人新村》赏析”解法与拓展.pptx"),
    ("地热力", "地理/编号11  热力环流和大气水平运动.pptx"),
    ("地讲评", "地理/2026年高三期初地理试题讲评.pptx"),
    ("政背诵", "政治/必修一 知识点背诵默写.pptx"),
    ("政学案", "政治/促优8 学案讲评.pptx"),
]

rows = []
for tag, rel in JOBS:
    p = BASE / rel
    if not p.exists():
        print("MISS", tag)
        continue
    z = zipfile.ZipFile(p)
    slides = sorted(
        [n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
        key=lambda s: int(re.search(r"\d+", s.split("/")[-1]).group()),
    )
    pages = []
    for sn in slides:
        x = z.read(sn).decode("utf8", "ignore")
        texts = [t.strip() for t in re.findall(r"<a:t>([^<]*)</a:t>", x) if t.strip()]
        pics = len(re.findall(r"<p:pic>", x))
        tbls = len(re.findall(r"<a:tbl>", x))
        shapes = len(re.findall(r"<p:sp>|<p:pic>", x))
        sizes = Counter(int(m) // 100 for m in re.findall(r'<a:rPr[^>]*sz="(\d+)"', x))
        tcol = Counter()
        for rm in re.finditer(r"<a:rPr[^>]*>.*?</a:rPr>|<a:rPr[^>]*/>", x):
            c = re.search(r'srgbClr val="([0-9A-Fa-f]{6})"', rm.group(0))
            if c:
                tcol[c.group(1).upper()] += 1
        title = ""
        for t in texts:
            if len(t) > 2:
                title = t[:42]
                break
        pages.append({
            "n": int(re.search(r"\d+", sn.split("/")[-1]).group()),
            "title": title,
            "chars": sum(len(t) for t in texts),
            "lines": len(texts),
            "pics": pics, "tbls": tbls, "shapes": shapes,
            "size_top": sizes.most_common(2),
            "col_top": tcol.most_common(3),
            "first5": texts[:5],
        })
    rows.append({"tag": tag, "file": rel.split("/")[-1][:36], "pages": pages})
    (OUT / f"{tag}.json").write_text(json.dumps(pages, ensure_ascii=False, indent=1))
    n = len(pages)
    avg_ch = sum(p["chars"] for p in pages) // max(n, 1)
    avg_sh = sum(p["shapes"] for p in pages) / max(n, 1)
    print(f"{tag}: {n}页 均{avg_ch}字 均{avg_sh:.1f}形 "
          f"图{sum(p['pics'] for p in pages)} 表{sum(p['tbls'] for p in pages)}")

(OUT / "_index.json").write_text(json.dumps(rows, ensure_ascii=False, indent=1))
print("saved ->", OUT)
