#!/usr/bin/env python3
"""debug 轮 QA 驱动：扫 ~/.jszz/out/ 全部 docx/pptx 真实产物，跑 qa.check_*。"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "server"))
from jszz import qa  # noqa: E402

OUT = Path.home() / ".jszz" / "out"
files = sorted(list(OUT.rglob("*.docx")) + list(OUT.rglob("*.pptx")))
# 跳过截图脚本的临时夹具（shot- 前缀），非交付样例
files = [f for f in files if not f.name.startswith("shot-")]
bad = []
soft_only = 0
for f in files:
    try:
        r = qa.check_pptx(f) if f.suffix == ".pptx" else qa.check_docx(f)
        # 用 qa 自带 ok 分级（提示级不算 bad），只揪 hard bug
        if not r.get("ok", True):
            bad.append((f.name, r.get("issues", [])))
        elif r.get("issues"):
            soft_only += 1
    except Exception as e:  # noqa: BLE001
        bad.append((f.name, [{"type": "QA崩溃", "detail": str(e)}]))

print(f"QA 扫描 {len(files)} 件交付产物（已排除 shot- 截图夹具）")
print(f"仅提示级(不计bad): {soft_only}  |  HARD BAD 件数: {len(bad)}")
for name, issues in bad[:30]:
    print(f"  ✗ {name}")
    for it in issues:
        print(f"      - {it.get('type')}: {it.get('detail')}")
print("RESULT:", "ALL CLEAN" if not bad else f"{len(bad)} BAD")
