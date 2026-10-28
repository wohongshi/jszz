#!/usr/bin/env python3
"""端到端验证 v1.3：学科网预览接口（版本清单 / 全文 / 还原 docx / 导图）。

对着真实资料 60192049（广州十校联考数学卷）跑：
  原卷版预览 2/5 页、解析版预览 3/23 页，还原 docx 应含文字与公式图。
"""
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "server"))

from jszz.zxxk import ZxxkClient   # noqa: E402

OUT = ROOT / "out" / "zxxk_test"
OUT.mkdir(parents=True, exist_ok=True)
fails = []


def check(cond, msg):
    print(("  ok   " if cond else "  FAIL ") + msg)
    if not cond:
        fails.append(msg)


c = ZxxkClient()
DOC = 60192049

print("== 版本清单 ==")
m = c.preview_files(DOC)
check(m.get("ok"), "拿得到版本清单")
files = m.get("files") or []
check(len(files) >= 1, f"至少一个版本（{len(files)}）")
for f in files:
    check(f["preview_pages"] <= f["total_pages"], f"预览页数 ≤ 全本：{f['preview_pages']}/{f['total_pages']}")
check(any("原卷" in f["name"] for f in files), "识别出原卷版")

print("== 全文 ==")
t = c.preview_text(DOC, max_pages=0, version=0)
check(t.get("ok"), "读得到正文")
body = "\n".join(p["text"] for p in t.get("pages", []))
check("注意事项" in body, "正文含「注意事项」")
check(len(body) > 500, f"字数合理（{len(body)}）")
check("学科网（北京）" not in body, "已剔除隐藏层水印")

print("== 还原 docx ==")
p = OUT / "restored.docx"
r = c.preview_docx(DOC, p, version=0)
check(r.get("ok") and p.exists(), "生成 docx")
check(r.get("chars", 0) > 800, f"docx 文字量（{r.get('chars')}）")
check(r.get("images", 0) > 0, f"带入公式/图形位图（{r.get('images')} 张）")

if p.exists():
    from docx import Document
    d = Document(str(p))
    text = "\n".join(x.text for x in d.paragraphs)
    check("注意事项" in text, "docx 段落含正文")
    check(len(d.inline_shapes) > 0, f"docx 内嵌图片（{len(d.inline_shapes)}）")

print("== 图片导出 ==")
imgs = OUT / "imgs"
r2 = c.preview_images(DOC, imgs, version=0)
check(r2.get("ok"), "图片导出成功")
check(r2.get("count", 0) > 0, f"导出图片数（{r2.get('count')}）")

print()
if fails:
    print(f"FAILED {len(fails)}")
    for f in fails:
        print(" -", f)
    sys.exit(1)
print("ALL PASS")
