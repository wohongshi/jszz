#!/usr/bin/env python3
"""Extract text + layout metadata from every non-video file under /mnt/usb/11.

Writes, per file, a JSON record into analysis/corpus.jsonl and the full text into
analysis/text/<sha1>.txt so later passes can read without re-parsing Office files.

Incremental: re-running skips files already present in corpus.jsonl.
"""
import hashlib
import json
import os
import re
import subprocess
import sys
import time
import traceback
from pathlib import Path

ROOT = Path("/mnt/usb/11")
OUT = Path("/root/.hermes/workspace/jszz/analysis")
TEXT_DIR = OUT / "text"
OUT.mkdir(parents=True, exist_ok=True)
TEXT_DIR.mkdir(parents=True, exist_ok=True)

SKIP_EXT = {".mp4", ".mp3", ".wav", ".m4a", ".flac", ".avi", ".mkv", ".mov", ".wmv"}
OOXML = {".pptx", ".docx", ".xlsx", ".pptm", ".docm", ".xlsm"}
LEGACY = {".doc", ".ppt", ".xls"}


def rel(p: Path) -> str:
    return str(p.relative_to(ROOT))


def key(relpath: str) -> str:
    return hashlib.sha1(relpath.encode("utf-8")).hexdigest()[:16]


def clean(s: str) -> str:
    s = s.replace("\x0b", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def pptx_extract(path: Path):
    from pptx import Presentation
    from pptx.util import Emu

    prs = Presentation(str(path))
    slides = []
    fonts, sizes, colors = {}, {}, {}
    pics = 0
    for idx, slide in enumerate(prs.slides, 1):
        texts = []
        boxes = []
        layout = slide.slide_layout.name if slide.slide_layout is not None else ""
        for sh in slide.shapes:
            try:
                if sh.shape_type == 13:  # picture
                    pics += 1
                box = (round(Emu(sh.left).inches, 2), round(Emu(sh.top).inches, 2),
                       round(Emu(sh.width).inches, 2), round(Emu(sh.height).inches, 2)) if sh.left is not None else None
                if sh.has_text_frame:
                    t = clean(sh.text_frame.text)
                    if t:
                        texts.append(t)
                        boxes.append({"t": t[:80], "box": box})
                    for para in sh.text_frame.paragraphs:
                        for run in para.runs:
                            if run.font.name:
                                fonts[run.font.name] = fonts.get(run.font.name, 0) + len(run.text)
                            if run.font.size is not None:
                                pt = round(run.font.size.pt, 1)
                                sizes[pt] = sizes.get(pt, 0) + len(run.text)
                            try:
                                if run.font.color and run.font.color.rgb is not None:
                                    c = str(run.font.color.rgb)
                                    colors[c] = colors.get(c, 0) + len(run.text)
                            except Exception:
                                pass
                if sh.has_table:
                    rows = []
                    for r in sh.table.rows:
                        rows.append(" | ".join(clean(c.text) for c in r.cells))
                    texts.append("[TABLE]\n" + "\n".join(rows))
            except Exception:
                continue
        notes = ""
        try:
            if slide.has_notes_slide:
                notes = clean(slide.notes_slide.notes_text_frame.text)
        except Exception:
            pass
        slides.append({"no": idx, "layout": layout, "text": "\n".join(texts)})
    meta = {
        "slides": len(prs.slides),
        "size_in": [round(prs.slide_width / 914400, 2), round(prs.slide_height / 914400, 2)],
        "fonts": dict(sorted(fonts.items(), key=lambda kv: -kv[1])[:6]),
        "font_sizes": dict(sorted(sizes.items(), key=lambda kv: -kv[1])[:8]),
        "colors": dict(sorted(colors.items(), key=lambda kv: -kv[1])[:8]),
        "pictures": pics,
    }
    body = []
    for s in slides:
        body.append(f"--- slide {s['no']} [{s['layout']}] ---\n{s['text']}")
    return "\n\n".join(body), meta, slides


def docx_extract(path: Path):
    import docx
    from docx.shared import Pt

    d = docx.Document(str(path))
    out = []
    styles = {}
    fonts, sizes = {}, {}
    for p in d.paragraphs:
        txt = clean(p.text)
        st = p.style.name if p.style is not None else ""
        if txt:
            styles[st] = styles.get(st, 0) + 1
            prefix = ""
            if st.lower().startswith("heading"):
                prefix = "#" * (int(st[-1]) if st[-1].isdigit() else 1) + " "
            out.append(prefix + txt)
            for run in p.runs:
                if run.font.name:
                    fonts[run.font.name] = fonts.get(run.font.name, 0) + len(run.text)
                if run.font.size is not None:
                    pt = round(run.font.size.pt, 1)
                    sizes[pt] = sizes.get(pt, 0) + len(run.text)
    tables = []
    for ti, t in enumerate(d.tables, 1):
        rows = [" | ".join(clean(c.text).replace("\n", " ") for c in r.cells) for r in t.rows]
        tables.append(f"[TABLE {ti}]\n" + "\n".join(rows))
    meta = {
        "paras": len(d.paragraphs),
        "tables": len(d.tables),
        "sections": len(d.sections),
        "page_in": [round(s.page_width.inches, 2) if s.page_width else None for s in d.sections][:3],
        "margins_in": [{k: round(getattr(s, k).inches, 2) for k in ("left_margin", "right_margin", "top_margin", "bottom_margin") if getattr(s, k) is not None} for s in d.sections][:2],
        "styles": dict(sorted(styles.items(), key=lambda kv: -kv[1])[:8]),
        "fonts": dict(sorted(fonts.items(), key=lambda kv: -kv[1])[:6]),
        "font_sizes": dict(sorted(sizes.items(), key=lambda kv: -kv[1])[:8]),
    }
    return "\n".join(out + tables), meta, None


def xlsx_extract(path: Path):
    import openpyxl

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    out, meta = [], {"sheets": []}
    for ws in wb.worksheets:
        rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i >= 40:
                rows.append("...")
                break
            cells = ["" if v is None else str(v) for v in row]
            if any(c.strip() for c in cells):
                rows.append(" | ".join(cells))
        meta["sheets"].append({"name": ws.title, "rows": ws.max_row, "cols": ws.max_column})
        out.append(f"=== sheet: {ws.title} ({ws.max_row}x{ws.max_column}) ===\n" + "\n".join(rows))
    wb.close()
    return "\n\n".join(out), meta, None


def xls_extract(path: Path):
    import xlrd

    wb = xlrd.open_workbook(str(path))
    out, meta = [], {"sheets": []}
    for ws in wb.sheets():
        rows = []
        for i in range(min(ws.nrows, 40)):
            rows.append(" | ".join(str(c.value) for c in ws.row(i)))
        meta["sheets"].append({"name": ws.name, "rows": ws.nrows, "cols": ws.ncols})
        out.append(f"=== sheet: {ws.name} ({ws.nrows}x{ws.ncols}) ===\n" + "\n".join(rows))
    return "\n\n".join(out), meta, None


def pdf_extract(path: Path):
    r = subprocess.run(["pdftotext", "-layout", "-enc", "UTF-8", str(path), "-"],
                       capture_output=True, timeout=120)
    txt = r.stdout.decode("utf-8", "replace")
    info = subprocess.run(["pdfinfo", str(path)], capture_output=True, timeout=60).stdout.decode("utf-8", "replace")
    pages = None
    m = re.search(r"Pages:\s+(\d+)", info)
    if m:
        pages = int(m.group(1))
    return clean(txt), {"pages": pages, "chars": len(txt)}, None


def legacy_extract(path: Path):
    """Convert legacy .doc/.ppt/.xls via libreoffice into txt in a temp dir."""
    tmp = OUT / "_legacy"
    tmp.mkdir(exist_ok=True)
    fmt = {"doc": "txt:Text", "ppt": "txt:Text", "xls": "csv:Text - txt - csv (StarCalc)"}
    ext = path.suffix.lower().lstrip(".")
    target = tmp / (key(rel(path)) + "." + ("csv" if ext == "xls" else "txt"))
    if not target.exists():
        subprocess.run(["soffice", "--headless", "--convert-to", fmt[ext], "--outdir", str(tmp), str(path)],
                       capture_output=True, timeout=180,
                       env={**os.environ, "HOME": str(tmp)})
    if target.exists():
        return clean(target.read_text("utf-8", "replace")), {"legacy": ext}, None
    return "", {"legacy": ext, "failed": True}, None


def kind_of(name: str) -> str:
    n = name.lower()
    rules = [
        (("学案", "导学案", "自主学习", "预学", "导学"), "学案"),
        (("试卷", "试题", "考试", "月考", "期中", "期末", "检测", "评测", "调研", "模拟", "真题", "单元"), "试卷"),
        (("答案", "解析", "讲评", "详解"), "答案讲评"),
        (("课件", "上课", "复习", "一轮", "二轮", "专题", "第", "讲"), "课件"),
        (("背诵", "默写", "晨诵", "早读", "知识点"), "背诵知识"),
        (("成绩", "汇总", "统计", "进展", "填报", "名单"), "成绩数据"),
        (("作业", "任务", "安排", "计划"), "作业任务"),
    ]
    for keys, k in rules:
        if any(x in n for x in keys):
            return k
    return "其他"


def main():
    only = sys.argv[1] if len(sys.argv) > 1 else None
    done = set()
    jsonl = OUT / "corpus.jsonl"
    if jsonl.exists():
        for line in jsonl.read_text("utf-8").splitlines():
            try:
                done.add(json.loads(line)["path"])
            except Exception:
                pass
    files = sorted(p for p in ROOT.rglob("*") if p.is_file())
    todo = [p for p in files if p.suffix.lower() not in SKIP_EXT and not p.name.startswith("~$")]
    print(f"total files={len(files)} todo={len(todo)} done={len(done)}", flush=True)
    t0 = time.time()
    with jsonl.open("a", encoding="utf-8") as fh:
        for i, p in enumerate(todo, 1):
            rp = rel(p)
            if rp in done:
                continue
            if only and only not in rp:
                continue
            ext = p.suffix.lower()
            rec = {"path": rp, "ext": ext, "bytes": p.stat().st_size,
                   "subject": rp.split("/")[0] if "/" in rp else "", "kind": kind_of(p.name),
                   "name": p.name}
            try:
                if ext == ".pptx":
                    text, meta, slides = pptx_extract(p)
                    rec["slides"] = [{"no": s["no"], "layout": s["layout"],
                                      "head": s["text"].split("\n")[0][:60],
                                      "chars": len(s["text"])} for s in slides]
                elif ext == ".docx":
                    text, meta, _ = docx_extract(p)
                elif ext == ".xlsx":
                    text, meta, _ = xlsx_extract(p)
                elif ext == ".xls":
                    text, meta, _ = xls_extract(p)
                elif ext == ".pdf":
                    text, meta, _ = pdf_extract(p)
                elif ext in (".doc", ".ppt"):
                    text, meta, _ = legacy_extract(p)
                elif ext in (".png", ".jpg", ".jpeg"):
                    text, meta = "", {"image": True}
                elif ext in (".html", ".htm"):
                    text, meta = clean(re.sub(r"<[^>]+>", " ", p.read_text("utf-8", "replace"))), {}
                elif ext == ".rar":
                    text, meta = "", {"archive": True}
                else:
                    text, meta = "", {}
                rec["meta"] = meta
                rec["chars"] = len(text)
                tk = key(rp)
                (TEXT_DIR / (tk + ".txt")).write_text(text, encoding="utf-8")
                rec["text_file"] = f"text/{tk}.txt"
                rec["ok"] = True
            except Exception as e:
                rec["ok"] = False
                rec["error"] = f"{type(e).__name__}: {e}"
                rec["trace"] = traceback.format_exc()[-800:]
            fh.write(json.dumps(rec, ensure_ascii=False) + "\n")
            fh.flush()
            if i % 25 == 0:
                print(f"[{i}/{len(todo)}] {time.time()-t0:.0f}s last={rp}", flush=True)
    print(f"done in {time.time()-t0:.0f}s", flush=True)


if __name__ == "__main__":
    main()
