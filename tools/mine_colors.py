#!/usr/bin/env python3
"""政史地三科真实课件配色取证：从 pptx XML 里量主色（不含白/黑/灰）。

用法：python3 tools/mine_colors.py  （输出 analysis/full4/zsd-colors.json）
"""
import json
import re
import zipfile
from collections import Counter
from pathlib import Path

USB = Path("/mnt/usb/11")
ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "analysis/full4/zsd-colors.json"

FILL = re.compile(r"<a:solidFill>.*?<a:srgbClr val=\"([0-9A-Fa-f]{6})\"", re.S)
FONT = re.compile(r"typeface=\"([^\"]+)\"")


def is_gray(h: str) -> bool:
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return max(r, g, b) - min(r, g, b) < 22 or (r > 240 and g > 240 and b > 240) or (r < 26 and g < 26 and b < 26)


def mine(p: Path) -> dict:
    cols, fonts = Counter(), Counter()
    try:
        with zipfile.ZipFile(p) as z:
            for n in z.namelist():
                if not (n.startswith("ppt/slides/") or n.startswith("ppt/slideMasters/")
                        or n.startswith("ppt/slideLayouts/") or n.startswith("ppt/theme/")):
                    continue
                x = z.read(n).decode("utf-8", "ignore")
                for c in FILL.findall(x):
                    if not is_gray(c.upper()):
                        cols[c.upper()] += 1
                if n.startswith("ppt/slideMasters/"):
                    for f in FONT.findall(x):
                        fonts[f] += 1
    except Exception as e:
        return {"error": str(e)}
    return {"colors": cols.most_common(12), "fonts": fonts.most_common(6)}


def main() -> None:
    targets = []
    for sub in ("政治", "历史", "地理"):
        d = USB / sub
        if not d.exists():
            print("MISS dir", sub)
            continue
        files = sorted([p for p in d.rglob("*.pptx") if not p.name.startswith("~$")],
                       key=lambda p: -p.stat().st_size)[:3]
        for p in files:
            targets.append((sub, p))
    out = {}
    for sub, p in targets:
        r = mine(p)
        r["file"] = str(p.relative_to(USB))
        out.setdefault(sub, []).append(r)
        top = ", ".join(f"#{c}×{n}" for c, n in r.get("colors", [])[:5])
        print(f"{sub} | {p.name[:38]:38} | {top}")
    OUT.write_text(json.dumps(out, ensure_ascii=False, indent=1), "utf-8")
    print("→", OUT)


if __name__ == "__main__":
    main()
