"""学科网 (zxxk.com) / 组卷网 (zujuan.xkw.com) client.

What is verified to work without an account:
  * keyword search            -> https://search.zxxk.com/api-kernel/v1/search-site/document/list
  * document metadata         -> https://www.zxxk.com/soft/<id>.html (JSON-LD)
  * document preview text     -> preview pages are SVG with real <text> nodes, so the
                                 first pages of a 学案/试卷 can be read as text.

What needs the user's own account / a non-blocked IP:
  * downloading the source file (点数 / 会员)
  * 组卷网 question browsing (zujuan) — WAF + login. The client still tries, using the
    same WAF solver, and reports the reason when it cannot.
"""
from __future__ import annotations

import html as htmllib
import json
import re
import urllib.parse
from pathlib import Path
from typing import Any

from . import config
from .svg2docx import svg_text
from .waf import WafSession

CAPTCHA_MARKERS = ("aliyunCaptcha", "aliyun_waf_bb", "nc-container")

SEARCH_API = "https://search.zxxk.com/api-kernel/v1/search-site/document/list"
DOC_URL = "https://www.zxxk.com/soft/{doc_id}.html"
ZUJUAN = "https://zujuan.xkw.com"

RESOURCE_TYPES = {
    "全部": "",
    "试卷": "试卷",
    "试题": "试题",
    "学案": "学案",
    "教案": "教案",
    "课件": "课件",
    "同步": "同步",
    "讲义": "讲义",
}


def _clean(s: str | None) -> str:
    if not s:
        return ""
    s = re.sub(r"<[^>]+>", "", s)
    return htmllib.unescape(s).strip()


class ZxxkClient:
    def __init__(self, cookie_file: Path | None = None, timeout: int = 30):
        self.http = WafSession(cookie_file=cookie_file, timeout=timeout)

    # ---------------- search ----------------
    def search(self, keyword: str, page: int = 1, page_size: int = 20,
               resource_type: str = "", stage: str = "", subject: str = "",
               grade: str = "", sort: str = "") -> dict:
        params = {"searchKey": keyword, "page": page, "pageSize": page_size}
        if resource_type:
            params["resourceTypeName"] = resource_type
        if stage:
            params["stageName"] = stage
        if subject:
            params["subjectName"] = subject
        if grade:
            params["gradeName"] = grade
        if sort:
            params["sort"] = sort
        r = self.http.get(SEARCH_API, params=params,
                          headers={"Referer": "https://search.zxxk.com/doc/"})
        try:
            data = r.json().get("data") or {}
        except Exception:
            return {"ok": False, "error": "搜索接口返回非 JSON（可能被 WAF 拦截）",
                    "status": r.status_code, "raw": r.text[:300]}
        items = []
        for it in data.get("searchList") or []:
            items.append({
                "document_id": it.get("documentId"),
                "title": it.get("title"),
                "type": it.get("resourceTypeName"),
                "subject": it.get("subjectName"),
                "stage": it.get("stageName"),
                "grade": it.get("gradeName"),
                "version": it.get("versionName"),
                "pages": it.get("pages"),
                "file_ext": it.get("fileExt"),
                "downloads": it.get("downloads"),
                "views": it.get("pv"),
                "price": it.get("displayPrice"),
                "free": it.get("isFree"),
                "publish": it.get("publishTime"),
                "tags": [t.get("name") for t in (it.get("tags") or [])],
                "url": DOC_URL.format(doc_id=it.get("documentId")),
            })
        return {
            "ok": True, "total": data.get("bizTotalCount") or data.get("totalCount"),
            "page": data.get("pageIndex"), "page_size": data.get("pageSize"),
            "items": items,
        }

    # ---------------- document ----------------
    def detail(self, doc_id: int | str) -> dict:
        url = DOC_URL.format(doc_id=doc_id)
        r = self.http.get(url)
        text = r.text
        out: dict[str, Any] = {"ok": True, "document_id": int(doc_id), "url": url,
                               "status": r.status_code}
        m = re.search(r'<script type="application/ld\+json">(.*?)</script>', text, re.S)
        if m:
            try:
                ld = json.loads(m.group(1))
                out.update({
                    "title": ld.get("name"),
                    "description": ld.get("description"),
                    "grade": ld.get("educationalLevel"),
                    "type": ld.get("learningResourceType"),
                    "about": ld.get("about"),
                    "author": (ld.get("author") or {}).get("name"),
                    "publish": ld.get("datePublished"),
                    "free": ld.get("isAccessibleForFree"),
                    "size": (ld.get("encoding") or {}).get("contentSize"),
                })
            except Exception:
                pass
        previews = []
        for raw in re.findall(r"https://cdn-preview\.xkw\.com/[^\s\"'<>\\]+", text):
            u = htmllib.unescape(raw).replace("\\u0026", "&").replace("\\/", "/")
            if u not in previews:
                previews.append(u)
        out["preview_pages"] = len(previews)
        out["preview_urls"] = previews[:40]
        out["has_preview"] = bool(previews)
        out["need_login"] = "您好，这份资料暂时无法预览" not in text and not previews
        return out

    # ---------------- preview (JSON, 全部分版本/全部页) ----------------
    def _cache_path(self, doc_id, kind: str = "preview"):
        d = config.HOME / "cache" / "zxxk"
        d.mkdir(parents=True, exist_ok=True)
        return d / f"{kind}_{doc_id}.json"

    def preview_files(self, doc_id: int | str, *, use_cache: bool = True) -> dict:
        """未文档化接口：拿到资料的**版本清单**与每个版本可预览的页。

        返回 files[]: {name, pages:[url...], preview_pages, total_pages}
        价值在于：① 区分原卷版/解析版（详情页 HTML 混在一起）；
        ② 给出 TotalPage（整份资料真实页数），老师据此判断值不值得下载；
        ③ 页面 URL 直接可用，不用再解析详情页。
        注意：免费预览**只给前 3–5 页**（PreviewPage），TotalPage 才是全本。
        """
        cache = self._cache_path(doc_id)
        if use_cache and cache.exists():
            try:
                import json as _json
                import time as _time
                blob = _json.loads(cache.read_text("utf-8"))
                if _time.time() - blob.get("ts", 0) < 7 * 86400:
                    return blob["result"]
            except Exception:
                pass

        url = (f"https://www.zxxk.com/soft/Preview/FirstLoadPreviewJson"
               f"?softID={doc_id}&type=3&product=1&v=2&FullPreview=true")
        r = self.http.get(url, headers={"Referer": f"https://www.zxxk.com/soft/{doc_id}.html"})
        body = r.content.decode("utf-8", "replace")
        if any(m in body[:4000] for m in CAPTCHA_MARKERS):
            # 阿里云滑块验证：脚本解不了，别硬刷（越刷越封）
            return {"ok": False, "status": 200,
                    "reason": "学科网触发了滑块验证（请求过密）。"
                              "稍后重试，或先在浏览器里打开一次学科网，再在 WebUI 导入 cookie。"}
        try:
            import json as _json
            data = (_json.loads(body) or {}).get("data") or {}
        except Exception:
            return {"ok": False, "reason": "预览接口返回非 JSON", "status": r.status_code,
                    "head": body[:120]}
        if not data.get("IsSuccess"):
            return {"ok": False, "reason": "预览接口未成功（可能需要登录/点数）"}

        def page_urls(html: str) -> list[str]:
            out = []
            for raw in re.findall(r'data-original="([^"]+)"', html or ""):
                u = htmllib.unescape(raw).replace("&amp;", "&")
                if u not in out:
                    out.append(u)
            return out

        files = []
        if data.get("IsRar") and data.get("rarPreviewInfo"):
            for f in data["rarPreviewInfo"]:
                urls = page_urls(f.get("Html") or "")
                files.append({"name": f.get("SoftName") or "", "pages": urls,
                              "preview_pages": f.get("PreviewPage") or len(urls),
                              "total_pages": f.get("TotalPage") or len(urls)})
        else:
            urls = page_urls(data.get("Html") or "")
            files.append({"name": data.get("SoftName") or "", "pages": urls,
                          "preview_pages": data.get("PreviewPage") or len(urls),
                          "total_pages": data.get("TotalPage") or len(urls)})
        files = [f for f in files if f["pages"]]
        result = {"ok": bool(files), "document_id": int(doc_id), "files": files,
                  "reason": "" if files else "该资料无可读预览页（可能为扫描件或需登录）"}
        if files:
            try:
                import json as _json
                import time as _time
                cache.write_text(_json.dumps({"ts": _time.time(), "result": result},
                                             ensure_ascii=False), "utf-8")
            except Exception:
                pass
        return result

    # ---------------- preview text ----------------
    def preview_text(self, doc_id: int | str, max_pages: int = 6,
                     version: int = 0) -> dict:
        """读预览页正文。

        version: 0 起，对应 preview_files() 里的 files 顺序（如 0=原卷版, 1=解析版）。
        max_pages: 0 或负数 = 该版本**全部可预览页**（免费只有前 3–5 页，不是全本）。
        """
        meta = self.preview_files(doc_id)
        if not meta.get("ok"):
            # 退化：详情页里能捞到几页算几页
            d = self.detail(doc_id)
            urls = d.get("preview_urls") or []
            if not urls:
                return {"ok": False, "reason": meta.get("reason") or "无可读预览页",
                        "detail": d}
            meta = {"files": [{"name": d.get("title") or "", "pages": urls,
                               "total_pages": len(urls)}]}

        files = meta["files"]
        idx = version if 0 <= version < len(files) else 0
        chosen = files[idx]
        urls = chosen["pages"]
        limit = len(urls) if max_pages <= 0 else min(max_pages, len(urls))
        pages = []
        for i, u in enumerate(urls[:limit], 1):
            r = self.http.s.get(u, timeout=self.http.timeout)
            if r.status_code != 200:
                pages.append({"page": i, "url": u, "text": "", "error": f"HTTP {r.status_code}"})
                continue
            pages.append({"page": i, "url": u, "text": svg_text(r.text)})
        d = self.detail(doc_id)
        return {"ok": True, "document_id": int(doc_id), "title": d.get("title"),
                "version": idx, "version_name": chosen["name"],
                "versions": [{"index": k, "name": f["name"],
                              "preview_pages": f.get("preview_pages"),
                              "total_pages": f["total_pages"]} for k, f in enumerate(files)],
                "pages": pages, "preview_total": chosen["total_pages"],
                "detail": {k: d.get(k) for k in
                           ("title", "type", "subject", "grade", "author", "publish", "size")}}

    def preview_docx(self, doc_id: int | str, out_path: str | Path, *, version: int = 0,
                     max_pages: int = 0, title: str = "", school: str = "") -> dict:
        """把预览页还原成**可编辑 docx**（文字 + 公式/图形的位图）。

        只覆盖免费预览的那几页（3–5 页）；全本要下载原件。
        """
        from .svg2docx import build_docx, parse_svg

        meta = self.preview_files(doc_id)
        if not meta.get("ok"):
            return {"ok": False, "reason": meta.get("reason") or "无可读预览页"}
        files = meta["files"]
        idx = version if 0 <= version < len(files) else 0
        chosen = files[idx]
        urls = chosen["pages"]
        limit = len(urls) if max_pages <= 0 else min(max_pages, len(urls))
        pages = []
        for u in urls[:limit]:
            r = self.http.s.get(u, timeout=self.http.timeout)
            if r.status_code == 200:
                pages.append(parse_svg(r.text))
        if not pages:
            return {"ok": False, "reason": "预览页全部下载失败"}
        d = self.detail(doc_id)
        name = title or (d.get("title") or f"zxxk-{doc_id}")
        footer = "　".join(x for x in [school, name] if x)[:60]
        path = build_docx(pages, out_path, title="", footer=footer)
        return {"ok": True, "file": str(path), "pages": len(pages),
                "chars": sum(len(ln["text"]) for p in pages for ln in p.lines()),
                "images": sum(len(p.pics) for p in pages),
                "version_name": chosen["name"], "title": name}

    def preview_images(self, doc_id: int | str, out_dir: str | Path, *,
                       version: int = 0, max_pages: int = 0) -> dict:
        """导出预览页里的公式/图形位图（做 PPT 素材）。"""
        from .svg2docx import parse_svg, save_images

        meta = self.preview_files(doc_id)
        if not meta.get("ok"):
            return {"ok": False, "reason": meta.get("reason") or "无可读预览页"}
        files = meta["files"]
        idx = version if 0 <= version < len(files) else 0
        urls = files[idx]["pages"]
        limit = len(urls) if max_pages <= 0 else min(max_pages, len(urls))
        pages = []
        for u in urls[:limit]:
            r = self.http.s.get(u, timeout=self.http.timeout)
            if r.status_code == 200:
                pages.append(parse_svg(r.text))
        saved = save_images(pages, out_dir, prefix=f"{doc_id}")
        return {"ok": True, "count": len(saved), "files": saved[:40], "dir": str(out_dir)}

    # ---------------- zujuan (组卷网) ----------------
    def zujuan_questions(self, knowledge_id: str, subject_prefix: str = "gzsx",
                         page: int = 1, qtype: str = "", difficulty: str = "",
                         order: str = "o2") -> dict:
        """List questions of a knowledge point on 组卷网.

        subject_prefix: gzsx 高中数学, gzwl 高中物理, gzhx 高中化学, gzyy 高中英语,
                        gzyw 高中语文, gzsw 高中生物, gzzz 高中政治, gzls 高中历史,
                        gzdl 高中地理; cz* 为初中。
        knowledge_id: e.g. 'zsd28279'. qtype t1..t6 (选择/多选/填空/解答/判断/其他).
        """
        path = f"/{subject_prefix}/{knowledge_id}/"
        params = {}
        if qtype:
            params["qt"] = qtype
        if difficulty:
            params["d"] = difficulty
        url = ZUJUAN + path + (order + "/" if order else "") + (
            f"?page={page}" if page > 1 else "")
        r = self.http.get(url, params=params or None)
        items = parse_zujuan_list(r.text)
        if not items:
            return {"ok": False, "reason": "未解析到题目（可能被 WAF/登录拦截，或知识点编号不对）",
                    "status": r.status_code, "chars": len(r.text), "url": url}
        return {"ok": True, "url": url, "page": page, "items": items}

    def zujuan_search(self, keyword: str, page: int = 1) -> dict:
        url = f"{ZUJUAN}/quessearch/?keyword={urllib.parse.quote(keyword)}"
        if page > 1:
            url += f"&page={page}"
        r = self.http.get(url)
        items = parse_zujuan_list(r.text)
        if not items:
            return {"ok": False, "reason": "未解析到题目（可能被 WAF 拦截，需登录或换网络）",
                    "status": r.status_code, "chars": len(r.text), "url": url}
        return {"ok": True, "url": url, "items": items}


def preview_to_paper_md(pages: list[dict], title: str = "",
                        subject: str = "", grade: str = "") -> str:
    """把预览页正文整理成试卷 Markdown（讲评 PPT/打印 docx 的输入）。

    认三种行：大题「一、…」、题号「1. …」、选项「A. …」；其余按正文归并。
    【答案】/【解析】行保留到参考答案节。
    """
    import re as _re
    lines: list[str] = []
    for pg in pages or []:
        t = (pg.get("text") or "").replace("\u3000", " ")
        lines += [x.strip() for x in t.splitlines() if x.strip()]

    head, body, answers = [], [], []
    in_ans = False
    in_notes = False
    pending_opt: list[str] = []

    def _flush_opts():
        nonlocal pending_opt
        if pending_opt:
            body.append("　".join(pending_opt))
            pending_opt = []

    for ln in lines:
        if _re.match(r"^(参考答案|答案与解析|答案及解析|答案)\s*$", ln) or ln.startswith("【答案】"):
            _flush_opts()
            in_ans = True
        if in_ans:
            answers.append(ln)
            continue
        if _re.match(r"^注意事项\s*[:：]?$", ln):
            _flush_opts()
            in_notes = True
            body.append("")
            body.append("## 注意事项")
            continue
        if _re.match(r"^[一二三四五六七八九十]+[、．.]", ln):
            _flush_opts()
            in_notes = False
            body.append("")
            body.append("## " + ln if not ln.startswith("##") else ln)
        elif in_notes:
            body.append(ln)
        elif _re.match(r"^\d{1,2}[.．、]", ln) and len(ln) > 3:
            _flush_opts()
            body.append("")
            body.append(ln)
        elif _re.match(r"^[A-DＡ-Ｄ]\s*[.．、]", ln):
            pending_opt.append(ln)
            if len(pending_opt) >= 4:
                _flush_opts()
        else:
            _flush_opts()
            body.append(ln)
    _flush_opts()

    title = title or "试卷"
    # 从正文提「满分 150 分 / 用时 120 分钟」进元信息（讲评 PPT 首页要用）
    all_text = "\n".join(lines)
    score_m = _re.search(r"满分\s*(\d+)\s*分", all_text)
    time_m = _re.search(r"(?:考试)?用时\s*(\d+)\s*分钟", all_text)
    meta_bits = []
    if subject:
        meta_bits.append(f"学科：{subject}")
    if grade:
        meta_bits.append(f"年级：{grade}")
    if time_m:
        meta_bits.append(f"时间：{time_m.group(1)}分钟")
    if score_m:
        meta_bits.append(f"满分：{score_m.group(1)}")
    meta = " | ".join(meta_bits)
    out = [f"# {title}"]
    if meta:
        out.append(f"> {meta}")
    out += head + body
    if answers:
        out += ["", "---", "", "## 参考答案"] + answers
    return "\n".join(out) + "\n"


def svg_text_legacy(svg: str) -> str:
    """Rebuild text from an SVG preview page, keeping line order.

    The preview SVG positions every text run absolutely, so we group runs by their
    y coordinate (rounded) and join left-to-right within a line.
    """
    runs = []
    for m in re.finditer(r"<text\b([^>]*)>(.*?)</text>", svg, re.S):
        attrs, inner = m.group(1), m.group(2)
        content = htmllib.unescape(re.sub(r"<[^>]+>", "", inner)).strip()
        if not content:
            continue
        x = _num(re.search(r'\bx="([-\d.]+)"', attrs))
        y = _num(re.search(r'\by="([-\d.]+)"', attrs))
        fs = _num(re.search(r'font-size="([-\d.]+)"', attrs)) or 12.0
        runs.append((round(y / max(fs * 0.5, 1)), x, fs, content))
    if not runs:
        return ""
    runs.sort(key=lambda r: (r[0], r[1]))
    lines: list[list] = []
    for r in runs:
        if lines and r[0] == lines[-1][0]:
            lines[-1][1].append(r[3])
        else:
            lines.append([r[0], [r[3]]])
    out = []
    for _, parts in lines:
        line = "".join(parts)
        line = re.sub(r"\s+", " ", line).strip()
        if line:
            out.append(line)
    return "\n".join(out)


def _num(m: re.Match | None) -> float:
    if not m:
        return 0.0
    try:
        return float(m.group(1))
    except Exception:
        return 0.0


def parse_zujuan_list(html: str) -> list[dict]:
    """Best-effort parse of a 组卷网 question list page."""
    items = []
    if "tk-quest-item" not in html:
        return items
    blocks = re.split(r'<div[^>]*class="[^"]*tk-quest-item[^"]*"', html)[1:]
    for i, b in enumerate(blocks, 1):
        b = b.split("<!--", 1)[0]
        text = re.sub(r"<[^>]+>", " ", b)
        text = htmllib.unescape(re.sub(r"\s+", " ", text)).strip()
        qid = re.search(r'data-question-id="(\d+)"', b) or re.search(r'/q(\d+)\.html', b)
        items.append({
            "index": i,
            "question_id": qid.group(1) if qid else None,
            "text": text[:2000],
            "has_image": "<img" in b,
        })
    return items
