#!/usr/bin/env python3
"""第二批样本：更多 pptx + docx（学案/试卷/讲评/史地政）批量转 PNG。"""
import json, subprocess, sys
from pathlib import Path

BASE = Path("/mnt/usb/11")
OUT = Path(__file__).resolve().parent.parent / "analysis/realppt/png2"
OUT.mkdir(parents=True, exist_ok=True)
tmp = Path("/tmp/realppt2"); tmp.mkdir(exist_ok=True)

JOBS = [
    # (tag, path, kind)   kind: ppt=取5页  doc=取前3页
    ("语早2", "语文/10.20早读.pptx", "ppt"),
    ("语默2", "语文/2026.3.10早自习任务.pptx", "ppt"),
    ("数专2", "数学/6.3《二项式定理与杨辉三角》专题复习.pptx", "ppt"),
    ("数导2", "数学/5.2.3简单复合函数的导数课件.pptx", "ppt"),
    ("英评2", "gs1/{1AE6B9FE-4D59-4685-BB77-16D25039FC29}_高三英语一轮-话题1/高三英语一轮-话题1/No.3-话题评测题.pptx", "ppt"),
    ("英晨2", "gs1/{0C4CCC47-1298-4089-A534-9174D2C4E62B}_高二英语下-晨诵暮读晚写11-U6选4 .pptx", "ppt"),
    ("英写2", "gs1/{2390E21F-A880-4C96-9CE9-A27677B549DA}_10.8-No.9-话题写作案.pptx", "ppt"),
    ("物实2", "gs1/{75DFFF96-3C34-42B1-8985-8D261AC16B11}_20 实验五：探究平抛运动的特点.pptx", "ppt"),
    ("物专2", "gs1/{C560018E-A421-4B53-B768-9D4017B02B4E}_第三章　第13课时　专题强化：动力学中的连接体和临界、极值问题.pptx", "ppt"),
    ("化电2", "gszbjs1/{038775A9-6A34-4923-ACB6-4D0801D2F65C}_18.新型化学电源.pptx", "ppt"),
    ("生背2", "生物/3.2+基因工程的基本操作程序 1背诵.pptx", "ppt"),
    ("生评2", "生物/习题讲评4.13.pptx", "ppt"),
    ("史主2", "历史/2022·山东卷第18题“《工人新村》赏析”解法与拓展.pptx", "ppt"),
    ("史轮2", "历史/一轮复习第11课.pptx", "ppt"),
    ("地评2", "地理/2026年高三期初地理试题讲评.pptx", "ppt"),
    ("地课2", "地理/编号11  热力环流和大气水平运动.pptx", "ppt"),
    ("政评2", "政治/促优8 学案讲评.pptx", "ppt"),
    ("政背2", "政治/必修一 知识点背诵默写.pptx", "ppt"),
    ("政主2", "政治/2026山东高考政治主观题.pptx", "ppt"),
    # docx
    ("doc学案1", "历史/第9课学案答案.docx", "doc"),
    ("doc学案2", "历史/第10课学案带答案 - 副本.docx", "doc"),
    ("doc试卷1", "gs1/{D52CC6AA-2F99-4534-869F-C5610B309CC6}_高三物理月考卷10月.docx", "doc"),
    ("doc试卷2", "gs1/{0FCB4F3E-5989-4174-ADA1-7956DB659DDD}_胶州四中高三数学第一次月考试题解析版.docx", "doc"),
    ("doc检测1", "gs1/{1AE6B9FE-4D59-4685-BB77-16D25039FC29}_高三英语一轮-话题1/高三英语一轮-话题1/No.4-B1U1基础检测.docx", "doc"),
    ("doc答案1", "gs1/{DC48E46B-A381-41C7-8E87-10B51B439AF8}_《师说》复习题答案版.docx", "doc"),
    ("doc教师1", "gs1/{D998D589-D8E5-4E3F-8842-EABF35CFD04C}_胶州四中20260918高三级年级下载试题(zxls.com)-教师用卷.docx", "doc"),
    ("doc试卷3", "历史/2026届山东省青岛市高三年级期初调研检测历史试题(zxls.com)-教师用卷.docx", "doc"),
    ("doc评测1", "w物理/高二级部物理学科单元评测试题.DOCX", "doc"),
]

def soffice_pdf(src, tag):
    od = tmp / tag; od.mkdir(exist_ok=True)
    for attempt in range(2):
        r = subprocess.run(["timeout", "280", "soffice", "--headless",
            f"-env:UserInstallation=file:///tmp/lo2_{tag}_{attempt}",
            "--convert-to", "pdf", "--outdir", str(od), str(src)],
            capture_output=True, text=True)
        pdfs = list(od.glob("*.pdf"))
        if pdfs:
            return pdfs[0]
        print("RETRY", tag, r.stdout[-120:], r.stderr[-120:], flush=True)
    print("FAIL", tag, src, flush=True)
    return None

for tag, rel, kind in JOBS:
    if list(OUT.glob(f"{tag}-*.png")):
        print("skip", tag, flush=True)
        continue
    src = BASE / rel
    if not src.exists():
        print("MISSING", tag, rel, flush=True)
        continue
    pdf = soffice_pdf(src, tag)
    if not pdf:
        continue
    n = int(subprocess.run(["pdfinfo", str(pdf)], capture_output=True, text=True)
            .stdout.split("Pages:")[1].split()[0])
    if kind == "ppt":
        pages = sorted(set([1, 2, max(3, n // 2), max(4, n - 1), n]))
    else:
        pages = sorted(set([1, 2, 3])) if n >= 3 else list(range(1, n + 1))
    for pg in pages:
        subprocess.run(["pdftoppm", "-r", "70", "-png", "-f", str(pg), "-l", str(pg),
                        str(pdf), str(OUT / tag)], capture_output=True)
    print(f"{tag} {n}页 -> {pages}", flush=True)
print("DONE")
