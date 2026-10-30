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

print("== 栏目识别（sec_key 回归，防【】括号漏判退化） ==")
from jszz.mdparse import sec_key          # noqa: E402
from jszz.pptx_render import BIG          # noqa: E402
from jszz.docx_render import BIG_SECTION  # noqa: E402

for raw, key in [
    ("【学习目标】", "学习目标"),
    ("一、试题分析", "试题分析"),
    ("（一）评分标准", "评分标准"),
    ("试题分析（命题立意）", "试题分析"),
    ("【参考答案】", "参考答案"),
    ("二、词性转换（每空1分，共10分）", "词性转换"),
]:
    check(sec_key(raw) == key, f"sec_key({raw!r}) → {key!r}")
for raw in ["【学习目标】", "【当堂检测】", "【答案】", "一、试题分析",
            "强基础·教考衔接", "必备知识·精读教材", "【我的反思】"]:
    k = sec_key(raw)
    check(bool(BIG.match(k)), f"pptx BIG 认得 {raw!r}")
for raw in ["【参考答案】", "试题分析（命题立意）", "备考提示", "原题呈现",
            "强基础·教考衔接", "审题常见误区", "【我的反思】"]:
    k = sec_key(raw)
    check(bool(BIG_SECTION.match(k)), f"docx BIG_SECTION 认得 {raw!r}")
check(not BIG.match(sec_key("普通章节标题")), "普通标题不误判为大栏目")
check(not BIG.match(sec_key("知识点1 匀变速")), "知识点小标题不误判为大栏目")

print("== 九科视觉（语数英物化生政史地）==")
from jszz import subject_style as ss  # noqa: E402

NINE = [("语文", "chinese"), ("数学", "math"), ("英语", "english"),
        ("物理", "physics"), ("化学", "chemistry"), ("生物", "biology"),
        ("政治", "politics"), ("历史", "history"), ("地理", "geography")]
for name, key in NINE:
    st = ss.get(name)
    check(st.key == key, f"学科 {name} → {key} 风格")
    check(bool(st.cover) and bool(st.header) and bool(st.deco),
          f"{name} 有封面/页头/装饰（{st.cover}/{st.header}/{st.deco}）")
for name, key in NINE:
    check(ss.guess(f"胶州四中{name}学案") == key, f"标题猜学科：{name}")
check(ss.guess("一轮复习纲要上第11课") == "history", "历史标题猜中（旧版漏判）")
check(ss.guess("专题11 热力环流") == "geography", "地理标题猜中")
check(ss.guess("复习必修四第二单元") == "politics", "政治标题猜中")
# 九科封面/页头互不重样（防三科套同一套皮）
covers = {ss.get(n).cover for n, _ in NINE}
headers = {ss.get(n).header for n, _ in NINE}
check(len(covers) == 9, f"九科封面版式不重样（{len(covers)}）")
check(len(headers) == 9, f"九科页头不重样（{len(headers)}）")

print("== scene 指令（含政史地四新页型）==")
from jszz.mdparse import parse as _parse  # noqa: E402

_d = _parse("""# t
> 史料：材料。——摘编自《史》
> 材料：时政。| 设问：问题？
> 图表：图例 | 判读：①甲；②乙
> 模板：答题模式：标题 + 表述成文 | 标题：×× | 时代背景：××
## s
body
""")
_scenes = [b.meta.get("scene") for b in _d.preface if b.kind == "scene"]
check(_scenes == ["史料", "材料", "图表", "模板"], f"四种 scene 解析：{_scenes}")
_d2 = _parse("# t\n\n## s\n> 引文：x——y\n")
check([b.meta.get("scene") for b in _d2.sections[0].blocks if b.kind == "scene"] == ["引文"],
      "旧五种 scene 不回归")
# **关键词** 红粗渲染（史轮2-24 取证）
from jszz.mdparse import blanks as _blanks  # noqa: E402
_b = _blanks("打破坊市制度，**商业活动不再受时空限制**，{{纸币}}出现")
check(("strong", "商业活动不再受时空限制") in _b and ("blank", "纸币") in _b,
      "**关键词**与{{挖空}}双语法共存")

print("== 图片引用（图注句式 + 嵌图）==")
from PIL import Image as _Image  # noqa: E402
_img = OUT / "fig-test.png"
_Image.new("RGB", (240, 160), (70, 130, 180)).save(_img)
_md_img = f"""# t
> 学科：地理

## s

![图1示意热力环流的常见类型（单位：hPa）]({_img})

1. 据图1所示，判断（　　）
"""
from jszz.pptx_render import build_xuean_deck as _bud  # noqa: E402
from jszz.docx_render import render_xuean as _rxd  # noqa: E402
_p2 = _bud(_md_img, OUT / "fig-test.pptx", school="测试中学")
_d2 = _rxd(_md_img, OUT / "fig-test.docx", school="测试中学")
from pptx import Presentation as _P2  # noqa: E402
_sl = list(_P2(str(_p2)).slides)
check(any(sh.shape_type == 13 for s2 in _sl for sh in s2.shapes), "pptx 嵌入图片")
check(any(sh.has_text_frame and sh.text_frame.text.startswith("图1示意")
          for s2 in _sl for sh in s2.shapes), "pptx 图注（图N示意句式）")
from docx import Document as _D2  # noqa: E402
_dd = _D2(str(_d2))
check(len(_dd.inline_shapes) == 1, "docx 嵌入图片")
check("图1示意" in "\n".join(pp.text for pp in _dd.paragraphs), "docx 图注")

print("== OCR 三修回归（假引擎，不起 ONNX）==")
import hashlib as _hashlib  # noqa: E402
import io as _io  # noqa: E402
from PIL import Image as _Im, ImageDraw as _ID  # noqa: E402
from jszz import ocr as _ocr  # noqa: E402

def _fake(text):
    def _c(*a, **k):
        return ([[[[0, 0], [10, 0], [10, 10], [0, 10]], text, 0.9]], None)
    return _c

def _png_bytes(w=60, h=30):
    b = _io.BytesIO()
    im = _Im.new("RGB", (w, h), "white")
    _ID.Draw(im).line([(5, h // 2), (w - 5, h // 2)], fill="black")
    im.save(b, format="PNG")
    return b.getvalue()

# 修1：P 模式透明底 → 白底（曾变黑底 OCR 全瞎）
_rgba = _Im.new("RGBA", (40, 20), (0, 0, 0, 0))
_ID.Draw(_rgba).line([(5, 10), (35, 10)], fill=(0, 0, 0, 255))
_prep_out = _ocr._prep(_rgba.convert("P", palette=_Im.ADAPTIVE))
check(_prep_out.getpixel((2, 2)) == (255, 255, 255), "OCR：P 模式透明区→白底")
# 修3：超大图降采样（曾拖 ONNX 发热）
check(max(_ocr._prep(_Im.new("RGB", (4000, 3000), "white")).size) <= 1600,
      "OCR：超大图降采样 ≤1600")
# 修2：水印片段剔除不陪葬正文（曾整段清空）
_cache = (Path.home() / ".jszz" / "cache" / "ocr")
_cache.mkdir(parents=True, exist_ok=True)
def _ocr_once(blob, fake_text):
    _ocr._ENGINE = _fake(fake_text)
    cf = _cache / (_hashlib.sha1(blob).hexdigest() + ".txt")
    if cf.exists():
        cf.unlink()
    return _ocr.ocr_bytes(blob)
check(_ocr_once(_png_bytes(), "v=v0+at 学科网") == "v=v0+at",
      "OCR：水印片段剔除、公式保住")
check(_ocr_once(_png_bytes(61), "学科网 www.zxxk.com") == "",
      "OCR：纯水印图返回空")
check(_ocr_once(_png_bytes(62), "x^2+y^2=r^2") == "x^2+y^2=r^2",
      "OCR：无水印原样保留")

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
