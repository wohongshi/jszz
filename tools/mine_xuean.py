#!/usr/bin/env python3
"""把 U 盘真实学案 docx 全文抽到 analysis/xuean3/ 供精读。"""
import zipfile
import re
from pathlib import Path

BASE = Path("/mnt/usb/11")
OUT = Path(__file__).resolve().parent.parent / "analysis" / "xuean3"
OUT.mkdir(parents=True, exist_ok=True)

DOCS = [
    ("历史9答案", "历史/第9课学案答案.docx"),
    ("历史10", "历史/第10课学案带答案 - 副本.docx"),
    ("历史2", "历史/纲要上一轮复习 第2课学案 - 副本(1).docx"),
    ("历史16", "历史/纲要上一轮复习 第16课学案答案.docx"),
    ("历史1轮8", "历史/.一轮第8课学案.docx"),
    ("物理1轮8", "w物理/.一轮第8课学案.docx"),
    ("英自主2", "gs1/{623795FF-42CA-4B22-A980-33B720BA6716}_答案2-高三英语国庆假期自主学习任务.docx"),
    ("化学作1", "化学/胶州四中高三级部化学自主学习手册答案/作业1答案.docx"),
    ("化学作3", "化学/胶州四中高三级部化学自主学习手册答案/作业3答案.docx"),
    ("生寒1", "生物/高二级部生物学科寒假自主学习答案/高二生物寒假自主学习（1）答案.docx"),
    ("生寒10", "生物/高二级部生物学科寒假自主学习答案/高二生物寒假自主学习（10）答案.docx"),
    ("历史背诵", "历史背诵任务.docx"),
]

for tag, rel in DOCS:
    p = BASE / rel
    if not p.exists():
        print("MISS", tag)
        continue
    z = zipfile.ZipFile(p)
    if "word/document.xml" not in z.namelist():
        print("NOT-OOXML", tag)
        continue
    doc = z.read("word/document.xml").decode("utf8", "ignore")
    paras = []
    for pa in re.findall(r"<w:p[ >].*?</w:p>", doc, re.S):
        t = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", pa))
        paras.append(t)
    text = "\n".join(paras)
    (OUT / f"{tag}.txt").write_text(text)
    print(f"{tag}: {len(paras)}段 {len(text)}字")
