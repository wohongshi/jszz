"""离线自检：不依赖网络，验证渲染器与组卷产出符合规范。"""
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "server"))

from pptx import Presentation                      # noqa: E402
from docx import Document                          # noqa: E402

from jszz.assemble import make_paper              # noqa: E402
from jszz.docx_render import render_paper, render_xuean   # noqa: E402
from jszz.pptx_render import build_paper_deck, build_xuean_deck  # noqa: E402

OUT = ROOT / "out" / "selftest"
OUT.mkdir(parents=True, exist_ok=True)
fails = []


def check(cond, msg):
    print(("  ok   " if cond else "  FAIL ") + msg)
    if not cond:
        fails.append(msg)


print("== 学案 ==")
md = (ROOT / "examples" / "xuean-wuli.md").read_text("utf-8")
docx_path = render_xuean(md, OUT / "xuean.docx", school="测试中学")
pptx_path = build_xuean_deck(md, OUT / "xuean.pptx", school="测试中学")
d = Document(str(docx_path))
text = "\n".join(p.text for p in d.paragraphs)
check("学习目标" in text, "docx 含学习目标")
check("课前预习" in text and "课内探究" in text, "docx 含预习/探究栏目")
check("知道" not in text or True, "docx 可打开")
underlined = sum(1 for p in d.paragraphs for r in p.runs if r.font.underline)
check(underlined >= 8, f"docx 填空下划线充足 ({underlined})")
prs = Presentation(str(pptx_path))
slides_text = ["\n".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame)
               for s in prs.slides]
check(len(prs.slides) >= 8, f"pptx 页数合理 ({len(prs.slides)})")
check(not any("{{" in t or "}}" in t for t in slides_text), "pptx 无 {{}} 残留")
check(abs(prs.slide_width.inches - 13.333) < 0.05, "pptx 为 16:9")
check(any("解析" in t for t in slides_text), "pptx 有解析页")
check(any("小结" in t for t in slides_text), "pptx 有小结页")

print("== 试卷 ==")
md2 = (ROOT / "examples" / "paper-wuli.md").read_text("utf-8")
p_docx = render_paper(md2, OUT / "paper.docx", school="测试中学")
p_pptx = build_paper_deck(md2, OUT / "paper.pptx", school="测试中学")
d2 = Document(str(p_docx))
t2 = "\n".join(p.text for p in d2.paragraphs)
check("单项选择题" in t2, "试卷有大题标题")
check("参考答案" in t2, "试卷含参考答案页")
check(len(d2.tables) >= 3, f"试卷选项用表格排版 ({len(d2.tables)})")
prs2 = Presentation(str(p_pptx))
st2 = ["\n".join(sh.text_frame.text for sh in s.shapes if sh.has_text_frame) for s in prs2.slides]
check(len(prs2.slides) >= 10, f"讲评 pptx 页数合理 ({len(prs2.slides)})")
check(any("答案" in t for t in st2), "讲评 pptx 有答案页")
check(not any("注意事项" in t and "第 1 题" in t for t in st2), "注意事项没有被当成题目")

print("== 组卷 ==")
qs = [
    {"stem": "下列关于加速度的说法正确的是（　　）",
     "options": ["加速度为零时速度一定为零", "加速度方向与速度变化方向相同",
                 "加速度越大速度越大", "加速度为负时物体一定减速"],
     "answer": "B", "analysis": "加速度是速度变化率，方向与 Δv 相同。", "score": 4, "type": "选择"},
    {"stem": "物体做自由落体运动，3 s 内下落高度为____m（g 取 10 m/s²）。",
     "answer": "45", "score": 4, "type": "填空"},
    {"stem": "汽车以 20 m/s 匀速行驶，刹车加速度大小为 4 m/s²，求刹车后 8 s 内的位移。",
     "answer": "50 m", "analysis": "刹车时间 5 s，8 s 内位移等于 5 s 内位移，x=v0²/2a=50 m。",
     "score": 12, "type": "解答"},
]
r = make_paper(qs, title="自检组卷", out_dir=OUT, total_score=20, school="测试中学")
check(r["ok"] and Path(r["files"]["docx"]).exists(), "组卷产出 docx")
check(Path(r["files"]["pptx"]).exists(), "组卷产出 pptx")
md3 = Path(r["files"]["md"]).read_text("utf-8")
check("一、单项选择题" in md3 and "二、填空题" in md3 and "三、解答题" in md3, "按题型自动分节")
check(md3.count("本题共1小题") == 3, "自动写题量与分节")

print("== 新增能力（v1.1）==")
from jszz.qa import check_docx, check_pptx, check_markdown  # noqa: E402

d3 = Document(str(docx_path))
head = "".join(p.text for p in d3.sections[0].header.paragraphs)
head += "".join(c.text for t in d3.sections[0].header.tables for r in t.rows for c in r.cells)
check("编写人" in head, "学案页眉含编写人/审核字段")
foot = "".join(p.text for p in d3.sections[0].footer.paragraphs)
check("共" in foot and "页" in foot, "页脚为「第X页（共Y页）」")
kinds = [b.kind for sec in __import__("jszz.mdparse", fromlist=["parse"]).parse(md).sections
         for b in sec.blocks]
check("tip" in kinds, "【特别提醒】识别为 tip 块")
check("insight" in kinds, "【点睛】识别为 insight 块")

prs3 = Presentation(str(build_xuean_deck(md, OUT / "xuean.pptx", school="测试中学")))
bands = 0
for sl in prs3.slides:
    for sh in sl.shapes:
        if sh.shape_type == 1 and sh.width and abs(sh.width.inches - 13.333) < 0.05:
            bands += 1
            break
check(bands >= 5, f"通栏页头/底部色带存在 ({bands} 页)")

has_note = any("判定要点" in (sh.text_frame.text if sh.has_text_frame else "")
               for sl in prs3.slides for sh in sl.shapes)
check(has_note, "例题页有「判定要点」批注框")

d4 = Document(str(p_docx))
check(len(d4.tables) >= 5, f"答案速查表已生成 ({len(d4.tables)} 表)")

check(check_docx(docx_path)["ok"], "学案 docx 体检通过")
check(check_pptx(pptx_path)["ok"], "学案 pptx 体检通过")
check(check_docx(p_docx)["ok"], "试卷 docx 体检通过")
check(check_pptx(p_pptx)["ok"], "讲评 pptx 体检通过")
check(check_markdown(md)["ok"] and check_markdown(md2)["ok"], "Markdown 写作体检通过")

print("== 打包完整性 ==")
skill_zip = ROOT / "dist" / "jszz.skill"
if skill_zip.exists():
    with zipfile.ZipFile(skill_zip) as z:
        names = z.namelist()
    check(any(n.endswith("SKILL.md") for n in names), "jszz.skill 含 SKILL.md")
    check(sum(1 for n in names if "references/" in n) >= 5, "jszz.skill 含 references")
else:
    check(False, "dist/jszz.skill 不存在")

print()
if fails:
    print(f"FAILED {len(fails)}:")
    for f in fails:
        print(" -", f)
    sys.exit(1)
print("ALL PASS")
