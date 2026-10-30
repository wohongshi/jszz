#!/usr/bin/env python3
"""jszz 素材库：真实课件内嵌图片的零拷贝索引 + 按需抠图。

老师课件里有大量素材图片（照片/示意图/图表）。本工具不复制图片，
只建索引（含「图片被哪一页引用、该页标题」做语义标签），用时从原课件 zip 里抠。

用法：
  python3 tools/asset_lib.py scan                  # 全量建索引（~/.jszz/assets/index.jsonl）
  python3 tools/asset_lib.py search 热力环流 [--subject 地理] [--limit 8]
                                                   # 检索 + 抠出 top-N 到 out/，打印路径
  python3 tools/asset_lib.py get <id> [-o path]    # 按 id 抠一张
"""
import argparse
import hashlib
import io
import json
import re
import zipfile
from pathlib import Path

USB = Path("/mnt/usb/11")
ROOT = Path.home() / ".jszz" / "assets"
INDEX = ROOT / "index.jsonl"
SUBJ = ["语文", "数学", "英语", "物理", "化学", "生物", "政治", "历史", "地理"]
SKIP_EXT = {".mp4", ".mp3", ".wav", ".avi", ".mkv", ".wmv", ".bin", ".wdp"}
REL_RE = re.compile(r'Target="\.\./media/([^"]+)"')
SLIDE_RE = re.compile(r"ppt/slides/(slide\d+)\.xml$")


def _subject_of(p: Path) -> str:
    for s in SUBJ:
        if p.parent.name == s or f"/{s}/" in str(p):
            return s
    return ""


def _slide_titles(z: zipfile.ZipFile) -> dict[str, tuple[str, str]]:
    """{媒体文件名: (slide名, 页标题)}——rels 找引用，slide xml 取页内第一段文字当标题。"""
    out: dict[str, tuple[str, str]] = {}
    for n in z.namelist():
        m = SLIDE_RE.match(n)
        if not m:
            continue
        slide = m.group(1)
        try:
            rels = z.read(f"ppt/slides/_rels/{slide}.xml.rels").decode("utf-8", "ignore")
            xml = z.read(n).decode("utf-8", "ignore")
        except Exception:
            continue
        title = ""
        tm = re.findall(r"<a:t>([^<]{2,60})</a:t>", xml)
        if tm:
            title = "".join(tm)[:50]
        for media in set(REL_RE.findall(rels)):
            out.setdefault(media, (slide, title))
    return out


def cmd_scan(args) -> None:
    ROOT.mkdir(parents=True, exist_ok=True)
    seen: set[str] = set()
    total = 0
    files = list(USB.rglob("*.pptx")) + list(USB.rglob("*.docx"))
    with INDEX.open("w", encoding="utf-8") as fp:
        for p in sorted(files):
            if p.name.startswith("~$"):
                continue
            try:
                z = zipfile.ZipFile(p)
            except Exception:
                continue
            slide_map = _slide_titles(z) if p.suffix == ".pptx" else {}
            for i in z.infolist():
                n = i.filename
                if not n.startswith(("ppt/media/", "word/media/")) or i.is_dir():
                    continue
                e = Path(n).suffix.lower()
                if e in SKIP_EXT:
                    continue
                try:
                    data = z.read(n)
                except Exception:
                    continue
                h = hashlib.sha1(data).hexdigest()
                if h in seen:
                    continue
                seen.add(h)
                # 尺寸（只解码图片头；wmf/emf 跳过解码）
                w = h_ = 0
                if e in (".png", ".jpg", ".jpeg", ".gif", ".tiff", ".tif", ".webp"):
                    try:
                        from PIL import Image
                        with Image.open(io.BytesIO(data)) as im:
                            w, h_ = im.size
                    except Exception:
                        pass
                base = Path(n).name
                slide, title = slide_map.get(base, ("", ""))
                rec = {
                    "id": h[:12], "ext": e, "bytes": i.file_size, "w": w, "h": h_,
                    "kind": ("vector" if e in (".wmf", ".emf")
                             else "photo" if e in (".jpeg", ".jpg") else "raster"),
                    "subject": _subject_of(p), "source": str(p.relative_to(USB)),
                    "zip_path": n, "slide": slide, "title": title,
                }
                fp.write(json.dumps(rec, ensure_ascii=False) + "\n")
                total += 1
    print(f"索引 {total} 张 → {INDEX}")


def _load() -> list[dict]:
    return [json.loads(x) for x in INDEX.read_text("utf-8").splitlines() if x.strip()]


def _score(rec: dict, query: str) -> int:
    """命中来源文件名 / 页标题 / 学科 → 分值（标题命中权重最高）。"""
    q = query.strip()
    if not q:
        return 1
    s = 0
    if q in (rec.get("title") or ""):
        s += 6
    if q in (rec.get("source") or ""):
        s += 3
    for w in re.split(r"\s+", q):
        if len(w) >= 2 and w in (rec.get("title") or ""):
            s += 2
        if w in (rec.get("source") or ""):
            s += 1
    return s


def _extract(rec: dict, out: Path) -> Path:
    src = USB / rec["source"]
    with zipfile.ZipFile(src) as z:
        data = z.read(rec["zip_path"])
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_bytes(data)
    return out


def cmd_search(args) -> None:
    rows = _load()
    if args.subject:
        rows = [r for r in rows if r["subject"] == args.subject]
    # 带语义标签的优先，其次大图；小图标（<4K）不推
    rows = [r for r in rows if r["bytes"] >= 4096]
    scored = sorted(rows, key=lambda r: (-_score(r, args.query), -r["bytes"]))
    top = [r for r in scored if _score(r, args.query) > 0][:args.limit] or scored[:args.limit]
    outdir = ROOT / "out"
    for r in top:
        ext = r["ext"] if r["ext"] != ".jpeg" else ".jpg"
        name = f'{r["id"]}{ext}'
        path = outdir / name
        if not path.exists():
            try:
                _extract(r, path)
            except Exception as e:
                print(f"跳过 {r['id']}（{e}）")
                continue
        tag = r["title"][:32] if r["title"] else Path(r["source"]).name[:32]
        print(f'{r["id"]}  {r["w"]}x{r["h"]}  [{r["subject"] or "通用"}]  {tag}')
        print(f'    {path}')


def cmd_get(args) -> None:
    rows = [r for r in _load() if r["id"].startswith(args.id)]
    if not rows:
        print("无此 id")
        return
    r = rows[0]
    out = Path(args.out) if args.out else ROOT / "out" / (r["id"] + r["ext"])
    print(_extract(r, out))


def main() -> None:
    ap = argparse.ArgumentParser("asset_lib")
    sub = ap.add_subparsers(dest="cmd", required=True)
    sub.add_parser("scan")
    s = sub.add_parser("search")
    s.add_argument("query")
    s.add_argument("--subject", default="")
    s.add_argument("--limit", type=int, default=8)
    g = sub.add_parser("get")
    g.add_argument("id")
    g.add_argument("-o", "--out", default="")
    args = ap.parse_args()
    {"scan": cmd_scan, "search": cmd_search, "get": cmd_get}[args.cmd](args)


if __name__ == "__main__":
    main()
