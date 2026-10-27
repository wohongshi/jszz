# 学科网 / 组卷网 取材手册

## 1. 站点与接口（2026-10 实测）

| 用途 | 地址 | 是否需要登录 |
|---|---|---|
| 资源搜索（结构化 JSON） | `https://search.zxxk.com/api-kernel/v1/search-site/document/list?searchKey=…&page=1&pageSize=20` | 否 |
| 资料详情 | `https://www.zxxk.com/soft/<id>.html`（页面里有 JSON-LD） | 否 |
| 预览页正文 | 详情页里 `cdn-preview.xkw.com/.../N.svg` | 否 |
| 下载原文件 | 需账号与点数 | 是 |
| 组卷网搜题 | `https://zujuan.xkw.com/quessearch/?keyword=…` | 是 + 反爬 |

## 2. WAF 怎么过的

两个站点都在阿里云 WAF 后面，会出现两类挑战页：

1. **alicfw**：`<body onload="check()">`，脚本读 `#parm_0/#parm_1` 算 cookie
2. **acw_sc__v2**：`<textarea id="renderData">`，脚本算 `acw_sc__v2` 后 `location.reload()`

处理方式：`server/jszz/bin/fetch_alicfw.js` 用 Node 的 `vm` 模块把页面**自己的 JS**
跑一遍（伪造 document/location/navigator/screen），捕获 `document.cookie` 的赋值，
回填到 `requests.Session` 后重试。最多重试 5 轮，能过第一层。

```python
from jszz.waf import WafSession
s = WafSession()
r = s.get("https://www.zxxk.com/soft/60192049.html")   # 自动解题
```

**已知限制**：acw_sc__v2 的 cookie 在当前机器的出口 IP 被列入黑名单时，反复重试
也过不去（返回 23670 字节挑战页 + 真实浏览器显示「已阻断」）。换网络（家庭宽带）
通常就好；这也是「本工具装在老师自己电脑上」的原因。

## 3. 从一份资料里拿正文

```python
from jszz.zxxk import ZxxkClient
c = ZxxkClient()
print(c.search("匀变速直线运动", resource_type="试卷")["items"][:3])
d = c.detail(60192049)          # preview_pages: 5
t = c.preview_text(60192049, max_pages=5)
for p in t["pages"]:
    print(p["page"], len(p["text"]))
```

预览页是矢量 SVG，`<text>` 节点带绝对坐标。`svg_text()` 按 y 坐标聚行、行内按 x 排序，
还原阅读顺序。实测一份 5 页数学卷，第 1–2 页能还原 500+ 字，题干与选项完整，
**公式与几何图是位图，会缺**。

## 4. 用搜索接口做筛选

`search()` 支持的参数（都是拼在 URL 上）：

```
searchKey=关键词
page / pageSize
resourceTypeName=试卷|试题|学案|课件|教案|讲义
stageName=高中|初中|小学
subjectName=物理|数学|…
gradeName=高三|高一|…
```

返回字段里最好用的：`documentId`、`title`、`resourceTypeName`、`pages`、`displayPrice`、
`downloads`、`pv`、`tags`（含「答案」「解析」「校联考」等），可以据此判断值不值得读。

## 5. 组卷网（zujuan）

题列表 URL 规律（来自开源项目 malyjacob/zujuan 的实测）：

```
https://zujuan.xkw.com/gzsx/zsd28279/o2/?page=1       # 高中数学，知识点 28279
https://zujuan.xkw.com/gzwl/zsd12345/                  # 高中物理
```

列表项选择器：`div.tk-quest-item.quesroot`，题干 `div.exam-item__cnt`，
答案 `div.wrapper.quesdiv`。题型参数 `qt1..qt6`（选择/多选/填空/解答/判断/其他），
难度 `d1..d5`，排序 `o2`（综合）。

**建议**：这条链路对 IP 敏感，做成「能通就用，不能通就换预览页」的降级路径，
不要让老师卡在这里。

## 6. 版权与合规

- 只抓**预览页**（站点本来对未登录用户公开的部分）
- 下载原文件走老师自己的账号，点数由老师的账号承担
- 生成物里保留 `> 来源：…` 注释，方便老师核对出处
- 不要把整站资源批量搬运到公开仓库
