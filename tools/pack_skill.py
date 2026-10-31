#!/usr/bin/env python3
"""打包技能包 dist/jszz.skill = SKILL.md + references/*.md（可复用，替代手工 zip）。"""
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SKILL = ROOT / "skill"
OUT = ROOT / "dist" / "jszz.skill"

def main():
    OUT.parent.mkdir(parents=True, exist_ok=True)
    files = [SKILL / "SKILL.md"] + sorted((SKILL / "references").glob("*.md"))
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            if f.exists():
                z.write(f, str(f.relative_to(SKILL)).replace("\\", "/"))
    with zipfile.ZipFile(OUT) as z:
        names = sorted(z.namelist())
    print(f"→ {OUT}  ({len(names)} 文件)")
    for n in names:
        print("  ", n)

if __name__ == "__main__":
    main()
