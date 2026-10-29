#!/usr/bin/env python3
"""全量抽取 U 盘 PPT/Word/表格（793 个文件）正文与结构 → analysis/full4/。

分派：
  .pptx/.docx/.xlsx  zip+XML 直解（快）
  .xls               xlrd
  .ppt/.doc          旧格式：LibreOffice 批量转成 pptx/docx 后再解（增量，断点续跑）
每文件落一个 .txt + 全局 index.jsonl（逐行追加，崩溃不丢）。
"""
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path

USB = Path("/mnt/usb")
OUT = Path(__file__).resolve().parent.parent / "analysis" / "full4"
TXT = OUT / "txt"
TXT.mkdir(parents=True, exist_ok=True)
IDX = OUT / "index.jsonl"
CONV = Path("/tmp/full4-conv")
CONV.mkdir(exist_ok=True)


def relkey(p: Path) -> str:
    return str(p.relative_to(USB)).replace("/", "__")


def extract_pptx(path: Path) -> dict:
    z = zipfile.ZipFile(path)
    slides = sorted([n for n in z.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
                    key=lambda s: int(re.search(r"\d+", s.split("/")[-1]).group()))
    pages = []
    for sn in slides:
        x = z.read(sn).decode("utf8", "ignore")
        texts = [t.strip() for t in re.findall(r"<a:t>([^<]*)</a:t>", x) if t.strip()]
        pages.append({
            "n": int(re.search(r"\d+", sn.split("/")[-1]).group()),
            "pics": len(re.findall(r"<p:pic>", x)),
            "tbls": len(re.findall(r"<a:tbl>", x)),
            "chars": sum(len(t) for t in texts),
            "text": " / ".join(texts),
        })
    return {"kind": "ppt", "pages": len(pages), "page_texts": pages}


def extract_docx(path: Path) -> dict:
    z = zipfile.ZipFile(path)
    names = z.namelist()
    if "word/document.xml" not in names:
        return {"kind": "doc", "error": "not-ooxml"}
    doc = z.read("word/document.xml").decode("utf8", "ignore")
    paras = []
    for pa in re.findall(r"<w:p[ >].*?</w:p>", doc, re.S):
        t = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", pa))
        if t.strip():
            paras.append(t.strip())
    head = ""
    for h in [n for n in names if re.match(r"word/header\d+\.xml", n)]:
        head += "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", z.read(h).decode("utf8", "ignore")))
    return {"kind": "doc", "pages": 0, "paras": len(paras),
            "chars": sum(len(t) for t in paras), "header": head.strip()[:60],
            "text": "\n".join(paras)}


def extract_xlsx(path: Path) -> dict:
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        ss = z.read("xl/sharedStrings.xml").decode("utf8", "ignore")
        for si in re.findall(r"<si>(.*?)</si>", ss, re.S):
            shared.append("".join(re.findall(r"<t[^>]*>([^<]*)</t>", si)))
    sheets = sorted(n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n))
    out_sheets = []
    for sn in sheets:
        x = z.read(sn).decode("utf8", "ignore")
        vals = []
        for c in re.finditer(r'<c[^>]*t="s"[^>]*>\s*<v>(\d+)</v>', x):
            i = int(c.group(1))
            if i < len(shared) and shared[i].strip():
                vals.append(shared[i].strip())
        for c in re.finditer(r'<c[^>]*><v>([^<]+)</v>', x):
            vals.append(c.group(1).strip())
        out_sheets.append({"sheet": sn.split("/")[-1], "cells": len(vals),
                           "text": " | ".join(vals[:200])})
    return {"kind": "xls", "pages": len(out_sheets), "page_texts": out_sheets}


def extract_xls(path: Path) -> dict:
    import xlrd
    wb = xlrd.open_workbook(str(path))
    out_sheets = []
    for sh in wb.sheets():
        vals = []
        for r in range(min(sh.nrows, 200)):
            for c in range(sh.ncols):
                v = sh.cell_value(r, c)
                if v not in ("", None):
                    vals.append(str(v).strip())
        out_sheets.append({"sheet": sh.name, "cells": len(vals),
                           "text": " | ".join(vals[:200])})
    return {"kind": "xls", "pages": len(out_sheets), "page_texts": out_sheets}


def convert_old(src: Path, target: str) -> Path | None:
    od = CONV / relkey(src)
    od.mkdir(parents=True, exist_ok=True)
    r = subprocess.run(["timeout", "120", "soffice", "--headless",
                        f"-env:UserInstallation=file:///tmp/lo4_{abs(hash(relkey(src)))%99999}",
                        "--convert-to", target, "--outdir", str(od), str(src)],
                       capture_output=True, text=True)
    outs = list(od.glob(f"*.{target}"))
    return outs[0] if outs else None


def main():
    files = sorted(p for p in USB.rglob("*")
                   if p.suffix.lower() in (".pptx", ".ppt", ".docx", ".doc", ".xlsx", ".xls")
                   and not p.name.startswith("~$"))
    done = set()
    if IDX.exists():
        for line in IDX.read_text().splitlines():
            try:
                done.add(json.loads(line)["rel"])
            except Exception:
                pass
    print(f"total={len(files)} done={len(done)}", flush=True)
    for i, p in enumerate(files):
        rel = str(p.relative_to(USB))
        if rel in done:
            continue
        ext = p.suffix.lower()
        rec = {"rel": rel, "ext": ext, "size": p.stat().st_size}
        try:
            if ext == ".pptx":
                rec.update(extract_pptx(p))
            elif ext == ".docx":
                rec.update(extract_docx(p))
            elif ext == ".xlsx":
                rec.update(extract_xlsx(p))
            elif ext == ".xls":
                rec.update(extract_xls(p))
            elif ext in (".ppt", ".doc"):
                tgt = "pptx" if ext == ".ppt" else "docx"
                cv = convert_old(p, tgt)
                if cv is None:
                    rec.update({"kind": "old", "error": "convert-failed"})
                elif ext == ".ppt":
                    rec.update(extract_pptx(cv))
                    rec["converted"] = True
                else:
                    rec.update(extract_docx(cv))
                    rec["converted"] = True
            # 落全文
            if "page_texts" in rec:
                full = "\n".join(f"[{pg.get('sheet') or pg.get('n')}] {pg['text']}"
                                 for pg in rec["page_texts"])
                rec["chars"] = sum(pg.get("chars", len(pg.get("text", ""))) for pg in rec["page_texts"])
            elif "text" in rec:
                full = rec.pop("text")
            else:
                full = ""
            if full.strip():
                (TXT / (relkey(p) + ".txt")).write_text(full)
                rec["txt"] = relkey(p) + ".txt"
            with IDX.open("a") as f:
                f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        except Exception as e:
            rec["error"] = f"{type(e).__name__}: {e}"
            with IDX.open("a") as f:
                f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        if (i + 1) % 25 == 0:
            print(f"progress {i+1}/{len(files)}", flush=True)
    print("DONE", flush=True)


if __name__ == "__main__":
    main()
