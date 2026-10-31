"""jszz MCP server: B站取素材 / 学科网取题 / 学案与试卷生成（15 个工具）。

传输：stdio（默认，供 Claude Code、WorkBuddy、Cline 等本地 MCP 客户端调用）
      streamable-http（`jszz serve --http` 或 `python -m jszz.server --http`）
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any

from . import config
from .assemble import make_paper as _make_paper
from .bilibili import BilibiliClient
from .docx_render import render_drill, render_grade_report, render_paper, render_xuean
from .pptx_render import build_paper_deck, build_xuean_deck
from .qa import check as _qa_check
from .zxxk import ZxxkClient

try:  # mcp >= 2.0
    from mcp.server.mcpserver import MCPServer as _Server
except Exception:  # mcp 1.x
    from mcp.server.fastmcp import FastMCP as _Server  # type: ignore

INSTRUCTIONS = """jszz 是给中学老师用的备课工具集：
- B站：搜教学视频、取字幕、下载音视频（mp3/mp4）当课堂素材
- 学科网/组卷网：按关键词搜课件/试卷/学案，读预览页文字（免登录），组卷
- 生成：Markdown 学案/试卷 → 可直接打印的 docx，以及 16:9 讲评/新授课 PPT
写内容时用教师常用 Markdown 语法：`#` 标题、`##` 栏目、`###` 知识点、`-` 要点、
`1.` 题目、`A.` 选项、`{{答案}}` 挖空、`【解析】` 解析、`---` 分页。"""

mcp = _Server("jszz", instructions=INSTRUCTIONS)


# ---------------------------------------------------------------- helpers
def _bili() -> BilibiliClient:
    return BilibiliClient(cookie=config.read_cookie("bilibili"))


def _zxxk() -> ZxxkClient:
    ck = config.read_cookie("zxxk")
    client = ZxxkClient(cookie_file=config.cookie_path("zxxk"))
    if ck:
        client.http.set_cookies(ck)
    return client


def _out(sub: str) -> Path:
    return config.out_dir(sub)


def _slug(name: str) -> str:
    return re.sub(r"[\\/:*?\"<>|\s]+", "_", name).strip("_")[:60] or "output"


# ---------------------------------------------------------------- B站
@mcp.tool()
def bili_search(keyword: str, page: int = 1, page_size: int = 20,
                order: str = "totalrank", duration: int = 0) -> dict:
    """搜索 B 站视频。

    keyword: 关键词，如 "匀变速直线运动 一轮复习"
    order:   totalrank(综合) / click(播放) / pubdate(最新) / dm(弹幕) / stow(收藏)
    duration: 0全部 1十分钟内 2十到三十分钟 3三十到六十分钟 4一小时以上
    返回 items[]：bvid / title / author / duration / play / url
    """
    return _bili().search(keyword, page=page, page_size=page_size,
                          order=order, duration=duration)


@mcp.tool()
def bili_video_info(bvid: str) -> dict:
    """取单个视频的详细信息（标题、时长、简介、分P、播放数据）。bvid 形如 BV1xx411c7mD。"""
    return _bili().video_info(bvid)


@mcp.tool()
def bili_subtitle(bvid: str, cid: int | None = None, save: bool = True) -> dict:
    """取视频字幕文本（带 [mm:ss] 时间戳）。多数视频需要登录 cookie（WebUI 里配）。"""
    c = _bili()
    tracks = c.subtitles(bvid, cid)
    if not tracks:
        return {"ok": False, "reason": "该视频没有字幕轨道，或未登录（AI 字幕需登录）"}
    out = []
    for t in tracks[:3]:
        text = c.fetch_subtitle_text(t["url"])
        out.append({"lan": t["lan"], "lan_doc": t["lan_doc"], "text": text})
    result: dict[str, Any] = {"ok": True, "tracks": out}
    if save:
        p = _out("bilibili") / f"{_slug(bvid)}_subtitle.txt"
        p.write_text("\n\n".join(f"# {t['lan_doc']}\n{t['text']}" for t in out), "utf-8")
        result["file"] = str(p)
    return result


@mcp.tool()
def bili_download(bvid: str, audio_only: bool = True, quality: str = "1080",
                  page: int = 1) -> dict:
    """下载 B 站音视频（需要本机装 yt-dlp + ffmpeg）。

    audio_only=True  → mp3（课堂听力/背景音素材）
    audio_only=False → mp4，quality 可为 480/720/1080/2160
    返回文件绝对路径。
    """
    return _bili().download(bvid, _out("bilibili"), audio_only=audio_only,
                            quality=quality, page=page)


@mcp.tool()
def bili_audio_stream(bvid: str, cid: int | None = None) -> dict:
    """取音频直链（DASH），不落盘；适合交给播放器或转写服务。"""
    return _bili().audio_url(bvid, cid)


# ---------------------------------------------------------------- 学科网
@mcp.tool()
def zxxk_search(keyword: str, page: int = 1, page_size: int = 20,
                resource_type: str = "", subject: str = "", grade: str = "") -> dict:
    """学科网资源搜索（免登录）。

    keyword: 如 "匀变速直线运动"
    resource_type: 试卷 / 试题 / 学案 / 课件 / 教案 / 讲义，留空为全部
    subject/grade: 如 "物理" / "高三"
    返回 items[]：document_id / title / type / subject / grade / pages / price / url
    拿到 document_id 后可调 zxxk_preview_text 读正文。
    """
    return _zxxk().search(keyword, page=page, page_size=page_size,
                          resource_type=resource_type, subject=subject, grade=grade)


@mcp.tool()
def zxxk_document(document_id: int) -> dict:
    """取资料详情：标题、简介、学段/年级/版本、页数、作者、是否有预览页。"""
    return _zxxk().detail(document_id)


@mcp.tool()
def zxxk_preview_text(document_id: int, max_pages: int = 0, version: int = 0,
                      save: bool = True) -> dict:
    """读学科网资料的免费预览页正文（预览页是 SVG，文字可直接提取，无需登录/点数）。

    version: 0=原卷版、1=解析版…（先调 zxxk_preview_files 看版本清单）
    max_pages: 0=该版本**全部可预览页**（免费只给前 3–5 页，不是全本）
    返回 pages[]：{page, text}，以及 versions[]（各版本预览页数 / 全本页数）。
    """
    c = _zxxk()
    r = c.preview_text(document_id, max_pages=max_pages, version=version)
    if r.get("ok") and save:
        p = _out("zxxk") / f"{document_id}_preview.txt"
        body = "\n\n".join(f"===== 第{p['page']}页 =====\n{p['text']}" for p in r["pages"])
        p.write_text(f"{r.get('title','')}\n\n{body}", "utf-8")
        r["file"] = str(p)
        r["chars"] = len(body)
    return r


@mcp.tool()
def zxxk_preview_files(document_id: int) -> dict:
    """列出一份学科网资料的可预览**版本清单**（原卷版 / 解析版）。

    返回 files[]: {name, preview_pages, total_pages, pages[]}。
    total_pages 是整份资料真实页数，免费预览只给前 3–5 页 —— 据此判断值不值得下载原件。
    """
    return _zxxk().preview_files(document_id)


@mcp.tool()
def zxxk_preview_docx(document_id: int, version: int = 0, max_pages: int = 0,
                      name: str = "") -> dict:
    """把学科网预览页**还原成可编辑 docx**（文字可编辑 + 公式/图形是位图）。

    这是「拿到一份能改的卷子」的最快路径：不用登录、不花点数。
    version: 0=原卷版 / 1=解析版（见 zxxk_preview_files）
    只覆盖免费预览的那几页（3–5 页）。
    """
    c = _zxxk()
    base = _slug(name or f"zxxk_{document_id}")
    out = _out("zxxk") / f"{base}_v{version}.docx"
    r = c.preview_docx(document_id, out, version=version, max_pages=max_pages,
                       school=config.load().get("school", ""))
    if r.get("ok"):
        r["file"] = str(out)
    return r


@mcp.tool()
def zxxk_preview_images(document_id: int, version: int = 0,
                        max_pages: int = 0) -> dict:
    """导出预览页里的**公式/几何图位图**（做 PPT 素材、拼卷子用）。

    返回保存目录与前 40 个文件名；图片按「页_序号」命名。
    """
    c = _zxxk()
    d = _out("zxxk") / f"imgs_{document_id}_v{version}"
    return c.preview_images(document_id, d, version=version, max_pages=max_pages)


@mcp.tool()
def zxxk_questions(keyword: str, page: int = 1) -> dict:
    """组卷网（zujuan.xkw.com）按关键词搜题。

    注意：组卷网有较强反爬 + 需要登录，且部分机房 IP 会被拦截。
    本机若返回 ok=False，请改用 zxxk_search + zxxk_preview_text 的路径取题，
    或在 WebUI 里导入浏览器 cookie 后再试。
    """
    return _zxxk().zujuan_search(keyword, page=page)


@mcp.tool()
def zxxk_knowledge_questions(knowledge_id: str, subject_prefix: str = "gzsx",
                             page: int = 1, difficulty: str = "") -> dict:
    """按知识点编号抓组卷网题目列表。knowledge_id 形如 zsd28279。

    subject_prefix: gzsx高中数学 gzwl高中物理 gzhx高中化学 gzyw高中语文 gzyy高中英语
                    gzsw高中生物 gzzz高中政治 gzls高中历史 gzdl高中地理（cz* 为初中）
    """
    return _zxxk().zujuan_questions(knowledge_id, subject_prefix=subject_prefix,
                                    page=page, difficulty=difficulty)


# ---------------------------------------------------------------- 生成
@mcp.tool()
def make_xuean(markdown: str, name: str = "", docx: bool = True,
               pptx: bool = True, theme: str = "", answers: bool = False) -> dict:
    """学案 Markdown → 可打印 docx + 新授课 PPT（16:9）。

    answers=True 额外出一份「答案版」docx：{{挖空}} 红粗下划线填回答案
    （真实答案版学案样式，如《师说》复习题答案版、第9课学案答案）。

    推荐的 Markdown 结构：
      # 课题
      > 学科：物理 | 年级：高一 | 课时：1
      ## 学习目标        （- 列表）
      ## 重点难点
      ## 课前预习
      ### 知识点一：xxx   （1. 正文，{{答案}} 表示挖空）
      ## 课内探究
      ### 探究一：xxx     （例1 题干… A. 选项… 【解析】…）
      ## 当堂检测        （1. 题干 + A. 选项）
      ## 课后作业
      ---
      ## 参考答案
    theme: 留空自动按 md 学科选型；也可显式 blue / teal / violet / orange / ink
    """
    md = markdown.strip()
    title = name or (re.search(r"^#\s+(.+)$", md, re.M).group(1) if re.search(r"^#\s+(.+)$", md, re.M) else "学案")
    base = _slug(title)
    d = _out("xuean")
    files: dict[str, str] = {}
    md_path = d / f"{base}.md"
    md_path.write_text(md + "\n", "utf-8")
    files["md"] = str(md_path)
    if docx:
        files["docx"] = str(render_xuean(md, d / f"{base}.docx",
                                         school=config.load().get("school", "")))
        if answers:
            files["docx_answers"] = str(render_xuean(
                md, d / f"{base}-答案版.docx",
                school=config.load().get("school", ""), answers=True))
    if pptx:
        files["pptx"] = str(build_xuean_deck(md, d / f"{base}.pptx", theme=theme,
                                             school=config.load().get("school", "")))
    report = _qa_check(md=md, docx=files.get("docx"), pptx=files.get("pptx"))
    return {"ok": True, "files": files, "title": title,
            "sections": len(re.findall(r"^##\s+", md, re.M)),
            "qa": _qa_summary(report)}


@mcp.tool()
def jszz_ocr(image_path: str) -> dict:
    """图片 OCR → 文字（rapidocr）。

    适用：拍的书页/截图/公式位图。透明底小图会自动合成白底并放大。
    水印（学科网/组卷网 logo）自动过滤。数学公式「尽力而为」，不保证符号全对。
    """
    from . import ocr as _ocr
    if not _ocr.available():
        return {"ok": False, "reason": "缺依赖：pip install rapidocr-onnxruntime"}
    p = Path(image_path)
    if not p.exists():
        return {"ok": False, "reason": f"文件不存在：{image_path}"}
    if p.is_dir():
        out = {}
        for f in sorted(p.glob("*")):
            if f.suffix.lower() in (".png", ".jpg", ".jpeg", ".bmp", ".webp"):
                t = _ocr.ocr_file(f)
                if t:
                    out[f.name] = t
        return {"ok": True, "dir": str(p), "count": len(out), "results": out}
    text = _ocr.ocr_file(p)
    return {"ok": True, "path": str(p), "text": text, "chars": len(text)}


@mcp.tool()
def zxxk_preview_ocr(document_id: int, version: int = 0, max_pages: int = 0) -> dict:
    """读预览页正文 + 公式位图 OCR 回填（比 zxxk_preview_text 多认出公式/符号）。

    返回每页 text（公式文字已插回原行位置）。适合直接给讲评 PPT/组卷用；
    数学公式识别不保证符号全对，正式印刷请人工核对。
    """
    r = _zxxk().preview_text(document_id, max_pages=max_pages, version=version,
                             ocr=True)
    if not r.get("ok"):
        return r
    return {"ok": True, "title": r.get("title"), "version": r.get("version"),
            "version_name": r.get("version_name"), "pages": r.get("pages"),
            "preview_total": r.get("preview_total"),
            "ocr": True, "engine": "rapidocr"}


@mcp.tool()
def zxxk_review_deck(document_id: int, name: str = "", version: int = 0,
                     max_pages: int = 0, docx: bool = True, pptx: bool = True,
                     ocr: bool = True) -> dict:
    """一键链路：学科网资料 → 试卷 Markdown → 打印 docx + 讲评 PPT。

    自动读预览正文（解析版优先：version=1 有答案时并入参考答案节），
    整理成试卷 Markdown 后走 make_paper_md 同款渲染。免登录可用；
    预览只有前 3–5 页（免费限制），要全本先 zxxk_preview_docx。
    ocr=True（默认）把公式位图 OCR 回文字——讲评不再满屏空括号。
    """
    cli = _zxxk()
    r = cli.preview_text(document_id, max_pages=max_pages, version=version,
                         ocr=ocr)
    if not r.get("ok"):
        return {"ok": False, "reason": r.get("reason") or "预览读取失败"}
    from .zxxk import preview_to_paper_md
    title = name or r.get("title") or "试卷"
    detail = r.get("detail") or {}
    md = preview_to_paper_md(r.get("pages") or [], title=title,
                             subject=detail.get("subject") or "",
                             grade=detail.get("grade") or "")
    base = _slug(title)
    d = _out("review")
    files: dict[str, str] = {}
    (d / f"{base}.md").write_text(md, "utf-8")
    files["md"] = str(d / f"{base}.md")
    if docx:
        files["docx"] = str(render_paper(md, d / f"{base}.docx",
                                         school=config.load().get("school", "")))
    if pptx:
        files["pptx"] = str(build_paper_deck(md, d / f"{base}.pptx",
                                             school=config.load().get("school", ""),
                                             subject=detail.get("subject") or ""))
    report = _qa_check(md=md, docx=files.get("docx"), pptx=files.get("pptx"))
    return {"ok": True, "files": files, "title": title,
            "pages_read": len(r.get("pages") or []),
            "preview_total": r.get("preview_total"),
            "questions": len(re.findall(r"^\s*\d{1,2}[.．、]", md, re.M)),
            "has_answers": "参考答案" in md,
            "qa": _qa_summary(report)}


@mcp.tool()
def make_paper_md(markdown: str, name: str = "", docx: bool = True,
                  pptx: bool = True, with_answers: bool = True,
                  theme: str = "") -> dict:
    """试卷 Markdown → 可打印 docx（含参考答案页）+ 讲评 PPT。

    推荐结构：
      # 试卷标题
      > 时间：90分钟 | 满分：100分
      ## 注意事项
      ## 一、单项选择题：本题共8小题，每小题3分，共24分
      1. 题干（　　）
      A. 选项一　B. 选项二　C. 选项三　D. 选项四
      ## 三、非选择题：本题共4小题，共60分
      6. （12分）题干……
      ---
      ## 参考答案
      1. A
      6. 【答案】… 【解析】…
    """
    md = markdown.strip()
    title = name or (re.search(r"^#\s+(.+)$", md, re.M).group(1) if re.search(r"^#\s+(.+)$", md, re.M) else "试卷")
    base = _slug(title)
    d = _out("paper")
    files: dict[str, str] = {}
    (d / f"{base}.md").write_text(md + "\n", "utf-8")
    files["md"] = str(d / f"{base}.md")
    if docx:
        files["docx"] = str(render_paper(md, d / f"{base}.docx",
                                         school=config.load().get("school", ""),
                                         answer_page=with_answers))
    if pptx:
        files["pptx"] = str(build_paper_deck(md, d / f"{base}.pptx", theme=theme,
                                             school=config.load().get("school", "")))
    report = _qa_check(md=md, docx=files.get("docx"), pptx=files.get("pptx"))
    return {"ok": True, "files": files, "title": title,
            "questions": len(re.findall(r"^\s*\d{1,2}[.．、]", md, re.M)),
            "qa": _qa_summary(report)}


@mcp.tool()
def make_drill(markdown: str, name: str = "", answers: bool = False) -> dict:
    """练习/讲义 Markdown → 可打印 docx（每日一练/限时小练/专题培优材料）。

    覆盖教师日常高频产出的两类（照真实资料版式）：
      1. 限时小练/每日一练：纯题组，真题源内联（如「（2026济宁一模）」），紧凑无密封线。
      2. 专题培优/讲义：栏目链「培优知识点及达成目标 → 【典型例题】→ 【解题思路】→
         【变式训练】→ 学生反思总结」，按写作工艺采分/逐项排除体例排。

    推荐结构：
      # 胶州四中高三化学每日一练3
      > 学科：化学 | 年级：高三 | 时间：2026.10
      ## 限时小练
      1.（2026济宁一模）题干（　　）
      A. …　B. …　C. …　D. …
      ## 参考答案
      1.【答案】C

    answers=True 额外出答案版 docx（{{挖空}} 红字填回）。返回 files / questions / qa。
    """
    md = markdown.strip()
    title = name or (re.search(r"^#\s+(.+)$", md, re.M).group(1)
                     if re.search(r"^#\s+(.+)$", md, re.M) else "练习")
    base = _slug(title)
    d = _out("drill")
    files: dict[str, str] = {}
    (d / f"{base}.md").write_text(md + "\n", "utf-8")
    files["md"] = str(d / f"{base}.md")
    files["docx"] = str(render_drill(md, d / f"{base}.docx",
                                     school=config.load().get("school", "")))
    if answers:
        files["docx_answers"] = str(render_drill(
            md, d / f"{base}-答案版.docx",
            school=config.load().get("school", ""), answers=True))
    report = _qa_check(md=md, docx=files.get("docx"))
    return {"ok": True, "files": files, "title": title,
            "questions": len(re.findall(r"^\s*\d{1,2}[.．、]", md, re.M)),
            "qa": _qa_summary(report)}


@mcp.tool()
def make_grade_report(scores: str, name: str = "", subject: str = "",
                      notes: str = "") -> dict:
    """成绩分析报告：成绩表 → 可打印 docx（教师质量分析版式）。

    scores 是成绩表数据，两种给法：
      1. xlsx 文件绝对路径（适配 U 盘真实三种表：多考试长表「考试|班级|姓名|各科」、
         调考成绩单「班级|姓名|学籍号|考号|各科+名次+等级」、单次考试「序号|姓名|各科|总分」）；
      2. JSON 字符串：行数组，字段如 {"姓名":"张三","班级":"1","语文":90,"数学":80,…}。

    自动列检测（学籍号/考号/序号等 ID 列不会误当分数），分析产出：
    各科统计（均分/中位/最高/最低/标准差/及格率/优秀率）+ 分数段分布 + 总分名次榜 +
    单科榜首 + 班级对比 + 多次考试趋势（进步/退步榜）+ 等级分布。
    notes 可追加教学建议（\\n 分段）。返回 files / summary 指标。
    """
    from .grades import analyze, load_scores
    d = load_scores(scores)
    if not d.get("ok"):
        return {"ok": False, "error": d.get("error", "成绩数据解析失败")}
    a = analyze(d["rows"], d["columns"])
    title = name or "成绩分析报告"
    extra = []
    if notes:
        extra.append(("八、教学建议", notes))
    files: dict[str, str] = {}
    out = _out("grade")
    out.mkdir(parents=True, exist_ok=True)
    docx_path = out / f"{_slug(title)}.docx"
    files["docx"] = str(render_grade_report(a, docx_path,
                                            school=config.load().get("school", ""),
                                            title=title, subject=subject,
                                            extra_sections=extra))
    # 摘要指标（供智能体直接读，不必开 docx）
    summary = {
        "n_rows": a.get("n_rows"),
        "subjects": {k: {"mean": v.get("mean"), "pass_rate": v.get("pass_rate"),
                         "excellence_rate": v.get("excellence_rate")}
                     for k, v in (a.get("subjects") or {}).items() if v.get("n")},
        "total_mean": (a.get("total") or {}).get("mean"),
        "top3": (a.get("top") or [])[:3],
        "has_trend": bool(a.get("trend")),
    }
    return {"ok": True, "files": files, "title": title,
            "n_rows": a.get("n_rows"), "summary": summary}


@mcp.tool()
def assemble_paper(questions_json: str, title: str, duration: int = 90,
                   total_score: int | None = None, with_answers: bool = True,
                   docx: bool = True, pptx: bool = True,
                   subject: str = "", grade: str = "") -> dict:
    """组卷：把结构化题目数组拼成试卷（按题型自动分组、自动编大题标题与分值）。

    questions_json 是 JSON 数组（字符串），每项字段：
      stem(题干) options(选项数组,可空) answer(答案) analysis(解析)
      score(分值) type(选择/多选/填空/判断/解答) source(来源,可选)
    例：[{"stem":"关于匀速直线运动…（　　）","options":["A. …"],"answer":"A","score":3,"type":"选择"}]
    """
    try:
        qs = json.loads(questions_json)
        if isinstance(qs, dict) and "questions" in qs:
            qs = qs["questions"]
    except Exception as e:
        return {"ok": False, "error": f"questions_json 解析失败：{e}"}
    if not isinstance(qs, list) or not qs:
        return {"ok": False, "error": "questions_json 必须是非空数组"}
    r = _make_paper(qs, title=title, out_dir=_out("paper"), duration=duration,
                    total_score=total_score, school=config.load().get("school", ""),
                    with_answers=with_answers, subject=subject, grade=grade,
                    formats=("md",) + (("docx",) if docx else ()) + (("pptx",) if pptx else ()))
    r.pop("markdown", None)
    return r


# ---------------------------------------------------------------- 抖音/视频号
@mcp.tool()
def douyin_info(url: str) -> dict:
    """抖音视频解析（元数据 + 无水印直链）。给老师备课取课堂视频素材。

    url: 分享短链 https://v.douyin.com/xxx/ 或视频页 https://www.douyin.com/video/<id> 或图文 /note/<id>
    返回 title(作品描述) / author(作者) / duration / cover / url(无水印播放直链) / source。
    实现：解析分享页 _ROUTER_DATA 抠 play_addr（借鉴 GitHub 开源 Evil0ctal/Douyin_TikTok_Download_API），
    失败降级 yt-dlp。cookie 由无头 chromedriver 过抖音 JS 质询自动收割。
    """
    from .douyin import DouyinClient
    return DouyinClient().info(url)


@mcp.tool()
def douyin_download(url: str, audio_only: bool = False) -> dict:
    """下载抖音视频/音频到本地（课堂素材）。

    audio_only=True → mp3（听力/背景音）；False → mp4（无水印）。
    返回产物绝对路径 file / size / title。主路 play_addr 直链流式下载，失败降级 yt-dlp。
    """
    from .douyin import DouyinClient
    return DouyinClient().download(url, _out("douyin"), audio_only=audio_only)


@mcp.tool()
def douyin_audio_stream(url: str) -> dict:
    """取抖音音频直链（不落盘），交给播放器或转写服务。"""
    from .douyin import DouyinClient
    return DouyinClient().audio_url(url)


@mcp.tool()
def channels_status() -> dict:
    """微信视频号能力边界（如实报告，不假装能自动下载）。

    视频号「链接→视频」无法纯服务端完成：需微信客户端登录态（服务端无客户端 →
    errCode 10012），yt-dlp 无 finder extractor，开源方案靠本地代理拦截微信 PC 客户端
    视频流解密（桌面端机制）。本服务只支持：给定桌面助手抓到的直链+解密key → 下载解密。
    """
    from .channels import ChannelsClient
    return ChannelsClient().status()


@mcp.tool()
def channels_probe(url: str) -> dict:
    """探测一条视频号链接：能否直连、是否需微信登录态，并给出抓取前置条件提示。"""
    from .channels import ChannelsClient
    return ChannelsClient().probe(url)


@mcp.tool()
def channels_download(url: str, key: str = "", filename: str = "") -> dict:
    """下载微信视频号视频（url + 解密 key 由桌面助手 wx_video_download 在微信 PC 端抓取）。

    url: 视频直链（桌面助手抓到的）；key: 解密 key（16/24/32 字节 hex 或原文，可留空存原始字节）；
    filename: 自定义文件名。返回产物路径 + decrypted 标志。
    服务端做下载 + AES 解密；抓取环节必须桌面微信客户端，无法绕过（见 channels_status）。
    """
    from .channels import ChannelsClient
    return ChannelsClient().download(url, _out("channels"), key=key,
                                     filename=filename or f"channels_{int(__import__('time').time())}")


# ---------------------------------------------------------------- 统一取材
@mcp.tool()
def material_fetch(url: str, audio_only: bool = False) -> dict:
    """统一取材入口：贴任意视频/笔记链接，自动识别平台（B站/抖音/小红书/视频号）→ 分发下载。

    老师不用记各平台工具名，一个入口贴链接就行。
    返回 platform / ok / file(s) / 元数据；识别不出平台或平台需登录态时如实报边界。
    audio_only=True 只要音频（B站/抖音支持 mp3）。
    """
    from .bilibili import BilibiliClient, is_bili
    from .channels import ChannelsClient, is_channels
    from .douyin import DouyinClient, is_douyin
    from .xhs import XHSClient, is_xhs

    url = (url or "").strip()
    if not url:
        return {"ok": False, "error": "空链接"}
    # 判平台（顺序：具体域名先判）
    if is_douyin(url):
        r = DouyinClient().download(url, _out("douyin"), audio_only=audio_only)
        return {"platform": "douyin", **r}
    if is_xhs(url):
        r = XHSClient().download(url, _out("xhs"))
        return {"platform": "xhs", **r}
    if is_channels(url):
        return {"platform": "channels", "ok": False, "auto_download": False,
                "error": "视频号需桌面助手抓直链+key（见 channels_status），统一入口只识别不下载",
                "hint": "channels_download(url, key)"}
    if is_bili(url):
        # b23.tv 短链 → bvid 解析走 bilibili 客户端；直接喂 yt-dlp 兜底
        import re as _re
        m = _re.search(r"(BV[\w]+)", url)
        if m:
            r = BilibiliClient(cookie=config.read_cookie("bilibili")).download(
                m.group(1), _out("bilibili"), audio_only=audio_only)
            return {"platform": "bilibili", **r}
        return {"platform": "bilibili", "ok": False,
                "error": "短链未解析出 bvid，用 bili_download 配合 bili_video_info"}
    return {"ok": False, "platform": "unknown",
            "error": f"识别不出平台（支持 B站/抖音/小红书/视频号）: {url[:60]}"}


# ---------------------------------------------------------------- 小红书
@mcp.tool()
def xhs_status() -> dict:
    """小红书能力边界（如实报告，不假装能未登录自动下载）。

    小红书笔记媒体靠登录态 + x-s/x-t 请求签名 API 异步填充，SSR 不含直链：未登录
    noteDetailMap 为空、feed API 返回 code:-1、yt-dlp 报 No video formats found。
    有登录 cookie 时递归抠 masterUrl/imageList 可下载。配 cookie 见 xhs_download。
    """
    from .xhs import XHSClient
    return XHSClient().status()


@mcp.tool()
def xhs_info(url: str) -> dict:
    """小红书笔记解析（元数据 + 视频/图直链）。给老师备课取课堂素材。

    url: 笔记链 https://www.xiaohongshu.com/explore/<id> 或分享短链 xhslink.com/xxx
    返回 title(描述)/author(作者)/type(video|image)/video_urls/image_urls/url。
    无登录 cookie 时如实返回 need_login=True + auto_download=False（不假装）。
    实现：解析 __INITIAL_STATE__ → noteDetailMap 递归抠 masterUrl/imageList（同抖音套路），
    curl_cffi impersonate 过基础反爬。cookie 在 WebUI/CLI 配（登录后浏览器复制）。
    """
    from .xhs import XHSClient
    return XHSClient().info(url)


@mcp.tool()
def xhs_download(url: str, filename: str = "") -> dict:
    """下载小红书视频/图到本地（课堂素材）。需配小红书登录 cookie。

    返回 files[] 产物绝对路径 / count / type。有 cookie 时从 __INITIAL_STATE__ 抠直链下载；
    无 cookie 如实返回 need_login（见 xhs_status 的平台边界说明）。
    """
    from .xhs import XHSClient
    return XHSClient().download(url, _out("xhs"), filename=filename)


# ---------------------------------------------------------------- 写作工艺
@mcp.tool()
def jszz_style_check(markdown: str) -> dict:
    """写作工艺体检（不生成文件）：按全 788 份真实教师资料的语言规律检查一份学案/试卷。

    检查项：
      1. 详解体例——选择题是否逐项排除（X：理由，排除/正确）+「故选」收口
      2. 短句原则——主观题答案是否按采分点分行（阅卷按要点独立给分）
      3. 禁口语——答案/详解里有无「把渣子」「多加点」等阅卷不得分的口语
      4. 归纳/口诀——有没有（真实资料 459/788 有口诀）
      5. 题源标注——真题是否照学科网原文标「（26-27高二上·河北石家庄·阶段检测）」
    返回 issues（必须改）/warnings（建议改）/hints（加分项）+ 计数。
    """
    from .writingcraft import check_writing
    return check_writing(markdown)


@mcp.tool()
def jszz_scoring_suggest(answer: str) -> dict:
    """把一段主观题答案按短句原则拆成采分点草稿（启发式，供人工核对）。

    阅卷按「要点独立给分」，一句含几个得分关键词就得几分。本工具把长句答案
    按分号/句号/换行拆成「要点N：展开」，方便核对是否写足信息点。
    """
    from .writingcraft import suggest_scoring
    pts = suggest_scoring(answer)
    return {"ok": True, "points": pts, "count": len(pts),
            "note": "启发式拆分，分值与关键词请人工核对"}


# ---------------------------------------------------------------- 状态
@mcp.tool()
def lint_markdown(markdown: str) -> dict:
    """写作体检（不生成文件）：挖空是否配对、大题是否写了分值、选项写法、题号是否跳号、有没有答案页。

    在写长文档之前先跑一次，能省掉「生成完才发现结构不对」的返工。
    """
    from .qa import check_markdown
    return check_markdown(markdown)


def _qa_summary(report: dict) -> dict:
    out = {}
    for k, v in report.items():
        if not v:
            continue
        if k == "markdown":
            out[k] = {"ok": v["ok"], "warnings": v.get("warnings", [])[:6]}
        else:
            out[k] = {"ok": v["ok"], "issues": v.get("issues", [])[:6],
                      "pages": v.get("pages")}
    return out


@mcp.tool()
def jszz_status() -> dict:
    """查看本机 jszz 状态：数据目录、已保存的登录 cookie、依赖（yt-dlp/ffmpeg/node）、学校署名。"""
    import shutil

    cfg = config.load()
    return {
        "home": str(config.HOME),
        "out_dir": cfg.get("out_dir"),
        "school": cfg.get("school"),
        "cookies": {
            "bilibili": bool(config.read_cookie("bilibili")),
            "zxxk": bool(config.read_cookie("zxxk")),
        },
        "deps": {k: bool(shutil.which(k)) for k in ("yt-dlp", "ffmpeg", "node")},
        "webui": f"http://127.0.0.1:{cfg.get('webui_port')}/",
    }


@mcp.tool()
def jszz_configure(school: str = "", author: str = "", theme: str = "",
                   webui_port: int = 0) -> dict:
    """设置学校署名 / 作者 / 默认主题 / WebUI 端口（写 ~/.jszz/config.json）。"""
    cfg = config.load()
    if school:
        cfg["school"] = school
    if author:
        cfg["author"] = author
    if theme:
        cfg["theme"] = theme
    if webui_port:
        cfg["webui_port"] = webui_port
    config.save(cfg)
    return {"ok": True, "config": cfg}


def main(argv: list[str] | None = None) -> None:
    argv = argv if argv is not None else sys.argv[1:]
    if "--http" in argv:
        port = 8765
        for i, a in enumerate(argv):
            if a == "--port" and i + 1 < len(argv):
                port = int(argv[i + 1])
        mcp.run(transport="streamable-http", port=port) if _supports_kwargs(mcp) else None
    else:
        mcp.run(transport="stdio")


def _supports_kwargs(server) -> bool:
    return True


if __name__ == "__main__":
    main()
