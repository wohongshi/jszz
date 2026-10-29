---
name: jszz
description: 中学备课：B站找素材、学科网组卷、学案/试卷一键出 docx 和 PPT。用 jszz MCP 工具时先读本技能。
version: 1.0.0
author: jszz
license: MIT
metadata:
  hermes:
    tags: [education, mcp, docx, pptx, bilibili, zxxk, lesson-plan, exam-paper]
    requires:
      mcp: jszz
      binaries: [node]
      optional: [yt-dlp, ffmpeg]
---

# jszz · 备课工作流

一套「找素材 → 出学案/试卷 → 转课件」的固定套路。所有排版参数来自 745 份真实
教师资料（25 份学案、147 份试卷类文档、213 份课件）的统计结果，不是拍脑袋定的。

## 什么时候用

- 老师要一份能直接打印的**学案/导学案**（含学习目标、预习填空、探究例题、当堂检测）
- 要一份**试卷**（选择题/非选择题、分值、答案页），或把题目**组卷**
- 要把学案/试卷**变成上课用的 PPT**（新授课 / 讲评课）
- 要从**学科网**找题、看某份资料的正文预览
- 要从**B站**找教学视频、取字幕、下音频当听力素材

## 六科学案模板（照 100+ 份真实学案逆向）

见 `references/07-六科学案模板.md`，六科各自的栏目顺序、题量配额、答案形态都在里面。
配套 `docs/六科内容地图.md` 是原始素材的内容清单（哪份文件讲了什么）。

## 先决条件

MCP 客户端里配置 `jszz`（stdio 启动）：

```json
{
  "mcpServers": {
    "jszz": {
      "command": "python3",
      "args": ["-m", "jszz.server"],
      "env": { "PYTHONPATH": "/path/to/jszz/server" }
    }
  }
}
```

装依赖（一次性）：

```bash
pip install -r requirements.txt      # mcp requests python-docx python-pptx lxml
# 可选：下 B 站音视频
pip install yt-dlp && apt install ffmpeg
# 过学科网 WAF（必需）
apt install nodejs
```

先调 `jszz_status` 看 cookie 和依赖是否齐。缺 cookie 也能搜、能读预览页，只是
下载原文件和组卷网搜题用不了。

## 核心工作流

### A. 从学科网取材 → 出学案/试卷

1. `zxxk_search(keyword, resource_type="学案")` 找候选，拿 `document_id`
2. `zxxk_preview_files(document_id)` 看**版本清单**（原卷版 / 解析版）与页数：
   `PreviewPage` 是免费能看的（2–5 页），`TotalPage` 才是全本
3. `zxxk_preview_text(document_id, version=0)` 读该版本全部可预览页正文
   （免登录；公式/图形是位图，读不到，但可以 `zxxk_preview_images` 导出）
4. 想要**能改的卷子**：`zxxk_preview_docx(document_id, version=0)` 直接还原成 docx
   （文字可编辑 + 公式图形位图），比手抄快得多
5. **一键讲评**：`zxxk_review_deck(document_id)` 一步到位——预览正文整理成试卷
   Markdown（自动提「时间/满分」元信息、注意事项单列、选项归组）→ 打印 docx +
   讲评 PPT（跟随学科风格）。**默认开 OCR**：公式位图识别回文字，讲评不再满屏空括号
6. **OCR**（rapidocr，结果按图 sha1 落盘缓存）：
   - `zxxk_preview_ocr(document_id)`：预览正文 + 公式 OCR 回填（逐行归位）
   - `jszz_ocr(image_path)`：任意图片/目录 → 文字（拍书页/截图都能用，水印自动过滤）
   - 数学公式识别是「尽力而为」——比空白强，但符号不保证全对，印刷前人工核对
7. 把读到的内容整理成 Markdown（语法见 `references/04-markdown.md`）
8. `make_xuean(markdown, pptx=True)` 或 `make_paper_md(markdown)`
9. 返回的 `files.docx` / `files.pptx` 就是成品路径，直接给老师

> 预览页只覆盖前几页（常见 2–6 页）。整份卷子的后半部分、答案页通常要下载原件，
> 这时需要 `zxxk` cookie（WebUI 里粘贴）。

### B. 组卷（多题拼成一份卷）

`assemble_paper(questions_json, title, total_score=100)`，题目数组字段：

```json
[{"stem":"题干（　　）","options":["A. …","B. …","C. …","D. …"],
  "answer":"A","analysis":"解析","score":3,"type":"选择","source":"2026 山东卷"}]
```

`type` 用 选择/多选/填空/判断/解答，会自动分组、编大题标题（一、二、三）、
按题型给分值，并另起一页放参考答案。

### C. 学案/试卷 → PPT

- 学案：`make_xuean(markdown, docx=True, pptx=True)`，PPT 是新授课结构：
  封面 → 学习目标 → 各知识点（自动分页）→ 探究/例题（题干页 + 解析页）→ 当堂检测 → 小结
- 试卷：`make_paper_md(markdown, pptx=True)`，PPT 是讲评结构：
  封面 → 试卷结构 → 每题（题干+选项 / 答案+解析）→ 易错点回顾

主题：`blue`（默认）、`teal`、`violet`、`orange`、`ink`。一份卷子讲评建议 `ink`。

### D. B站素材

```text
bili_search("匀变速直线运动 一轮复习")          → 挑视频
bili_video_info("BV1xx411c7mD")               → 看分P、时长
bili_subtitle("BV...")                        → 取字幕文本（需登录，AI 字幕也要）
bili_download("BV...", audio_only=True)       → mp3 课堂听力素材
bili_download("BV...", audio_only=False, quality="1080") → mp4
```

下载走 `yt-dlp`，本机没有会在返回里提示。会员清晰度、AI 字幕要配 cookie。

## 写作规范（关键，别凭感觉写）

**学案**（`references/01-xuean.md`）：A4、宋体 12pt 正文、行距 1.5；栏目标签用
【】；主链是「学习目标 → 重点难点 → 课前预习 → 课内探究 → 当堂检测 → 课后作业 → 答案」。
预习部分必须是**填空式**（`{{加速度}}`），探究部分必须是**例题 + 解析**（【解析】）。
一份合格学案 2–4 页，填空密度约每页 6–10 处。
`make_xuean(answers=True)` 额外出**答案版** docx：答案红色压下划线填回
（真实答案版样式，见 `references/10-学案试卷docx取证.md`）。

**试卷**（`references/02-paper.md`）：A4、标题 18pt 居中、`（时间：90分钟　满分：100分）`、
班级/姓名/考号三栏底框线；大题标题写全「一、单项选择题：本题共8小题，每小题3分，共24分」；
选项 4 项一行（长选项自动 2×2）；答案页另起，标 `【答案】` / `【解析】`。
山东卷常见配比：8 单选×3 + 4 多选×4 + 4 非选择×60 ≈ 100 分。

**课件**（`references/03-ppt.md`）：16:9；标题 30–40pt、正文 24–28pt、数字 Times New Roman、
中文微软雅黑；每页 ≤ 6 条、≤ 300 字（素材库中位数 184 字/页）；例题必须「题干页 → 解析页」
两页，不要挤在一页。

## 课件按学科自动换风格

`make_xuean` / `make_paper_md` 生成 PPT 时读 `> 学科：数学`，自动套六套视觉语言之一
（语=提示横条+篇目标签、数=细线框标题牌、英=Learning aims+青色栏目块、
物=深蓝轨道弧线、化=试管卡片蓝胶囊、生=绿块栏目+放射概念图——按 17 份真实课件
取证重做，见 `references/09-六科课件视觉.md`）。另有四行指令可触发专属页型：

```markdown
> 引文：…——出处
> 公式：v = v₀ + at | 条件：… | 结论：…
> 对照：English sentence. | 中文译文。
> 反应：2H₂ + O₂ = 2H₂O | 条件：点燃 | 要点…
> 概念：中心词 > 分支：a;b > 分支：c;d
```

## 体检（写完先跑，别等生成完才返工）

```text
lint_markdown(markdown)      → 挖空配对、大题分值、选项写法、题号跳号、答案页
jszz check --md a.md --docx a.docx --pptx a.pptx   # 命令行同一套
```

`make_xuean` / `make_paper_md` 的返回里也带 `qa` 字段：形状越界、文本框溢出、
占位符残留、空白页、表格超宽、页眉缺失，都会直接列出来。

## 写作语法速查（渲染器只认这套）

| 写法 | 效果 |
|---|---|
| `> 学科：物理 \| 年级：高一 \| 课时：1 \| 编写：… \| 审核：… \| 时间：…` | 学案页眉 + 副标题 |
| `> 时间：90分钟 \| 满分：100分 \| 日期：2026.10` | 试卷标题行 |
| `{{答案}}` | 挖空（docx 下划线 / PPT 下划线） |
| `【解析】…` | 解析块（PPT 独立答案页） |
| `【点睛】…` | 方法要点（PPT 例题右侧红框「判定要点」） |
| `【特别提醒】…` | 易错陷阱（docx 红条 / PPT 单独红页） |
| `> 来源：（2025·山东卷）` | 题源，显示在例题序号块右侧 |

## 交付前自检

- [ ] 学案：有学习目标、有填空、有例题解析、有当堂检测、页脚学校名
- [ ] 试卷：有大题分值说明、有答案页、选项没有断词、题干续行有悬挂缩进
- [ ] PPT：封面有课题与学段、每题不超过 2 页、无 `{{}}` 残留、无空页
- [ ] 文件路径是绝对路径，能直接打开

## 坑（都是从真实素材里踩出来的）

1. **公式是图片**：学科网预览页里的公式、几何图是位图，`zxxk_preview_text` 拿不到。
   涉及数学/物理公式时，让老师补图，或在 Markdown 里用纯文本写（`v＝v0＋at`）。
2. **六科模板不同**：语文一文一案、数学一考点一案、英语一话题四件套、
   物理/化学一讲一案、生物章末填空案。写之前先看 `references/07-六科学案模板.md`，
   别拿物理的骨架去写语文。
3. **预览 JSON 只成功一次**：连续请求会撞阿里云滑块验证，jszz 会缓存首次结果；
   缓存过期（7 天）后才会重新请求。别自己写循环重试，越试越封。
4. **组卷网常被墙**：`zujuan.xkw.com` 有 acw_sc__v2 + 登录双重限制，机房 IP 常被
   拦截（返回 23670 字节的挑战页）。失败就换「搜索 → 预览页取题」这条路，别硬刚。
3. **WAF 需要 node**：zxxk 的挑战页要靠 node 跑页面自己的 JS（`bin/fetch_alicfw.js`），
   没装 node 会一直返回挑战页。
4. **不要用 `{{}}` 之外的挖空写法**：全角下划线、连续 `_` 在换字体/打印时会断线。
5. **别把讲评课件做成「一页一题带解析」**：学生看不清，讲评要题干页留白给学生先做。
6. **署名**：页脚学校名来自 `jszz_configure(school=...)`，默认空；交付前确认。

## 文件结构

```
server/jszz/           MCP 服务与渲染器
skill/                 本技能（SKILL.md + references）
examples/              可直接跑的示例 Markdown
```
