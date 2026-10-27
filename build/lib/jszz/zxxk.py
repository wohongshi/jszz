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

from .waf import WafSession

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

    # ---------------- preview text ----------------
    def preview_text(self, doc_id: int | str, max_pages: int = 6) -> dict:
        """Extract text from the free preview pages (SVG <text> order == reading order)."""
        d = self.detail(doc_id)
        urls = d.get("preview_urls") or []
        if not urls:
            return {"ok": False, "reason": "该资料没有可读预览页（需登录/下载）",
                    "detail": d}
        pages = []
        for i, u in enumerate(urls[:max_pages], 1):
            r = self.http.s.get(u, timeout=self.http.timeout)
            if r.status_code != 200:
                pages.append({"page": i, "text": "", "error": f"HTTP {r.status_code}"})
                continue
            pages.append({"page": i, "text": svg_text(r.text)})
        return {"ok": True, "document_id": int(doc_id), "title": d.get("title"),
                "pages": pages, "preview_total": d.get("preview_pages"),
                "detail": {k: d.get(k) for k in
                           ("title", "type", "subject", "grade", "author", "publish", "size")}}

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


def svg_text(svg: str) -> str:
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
