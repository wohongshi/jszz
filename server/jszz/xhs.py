"""XiaoHongShu (小红书) client — cookie 门控 + 诚实边界。

**平台边界（必须说清）**：小红书笔记媒体（视频/图）是**登录态 + 请求签名 API 异步填充**的：
- SSR HTML 的 `window.__INITIAL_STATE__.note.noteDetailMap` **未登录时为空**（实测三种
  impersonate 指纹、带 cookie、带 xsec_token 均空），视频流不进 HTML。
- 裸调 `/api/sns/web/v1/feed` 返回 `{"code":-1}`——需 `x-s/x-t/x-s-common` 请求签名
  （GitHub 开源 JoeanAmier/XHS-Downloader 整个项目就是维护这套签名 + cookie 池）。
- yt-dlp 有 `XiaoHongShu` extractor 但未登录报 "No video formats found"。

**本模块提供的能力**：
  1. **cookie 门控**：老师在 WebUI/CLI 配小红书 cookie（登录后浏览器复制），有登录态时
     `_INITIAL_STATE__` 的 noteDetailMap 可能带媒体，递归抠 `masterUrl`/`imageList` 直链下载。
  2. **递归提取**（同抖音验证过的套路）：`noteDetailMap.<id>.note.video.media.stream[].masterUrl`
     与 `imageList[].urlDefault`——键名会变，递归找含 `.mp4`/`sns-video`/`xhscdn` 的节点。
  3. **诚实边界**：无 cookie 或 noteDetailMap 仍空时，`status()['auto_download']` 明确 False，
     返回「需登录 cookie 或走 XHS-Downloader 桌面助手」，不假装能自动下载。

即：抓取环节依赖小红书登录态/签名，jszz 负责有 cookie 时的提取 + 下载落盘当课堂素材。
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
_XHS_HOST = "www.xiaohongshu.com"
_COOKIE_JSON = Path.home() / ".jszz" / "cache" / "xhs_cookies.json"

# 笔记 ID：/explore/<hex>、/discovery/item/<hex>、/user/profile/<uid>/<hex>
_ID_RE = re.compile(r"(?:explore|discovery/item|item)/([0-9a-f]{20,32})")
_IS_XHS = re.compile(r"(xiaohongshu\.com|xhslink\.com|小红书)", re.I)


def is_xhs(url: str) -> bool:
    return bool(_IS_XHS.search(url or ""))


def _js_to_json(src: str) -> Any:
    """把 __INITIAL_STATE__ 的 JS 对象转 JSON（用 yt-dlp 的 js_to_json，处理 undefined/尾逗号）。"""
    try:
        from yt_dlp.utils import js_to_json
        return json.loads(js_to_json(src))
    except Exception:  # noqa: BLE001
        # 退化：粗暴替换 undefined/trailing comma
        s = re.sub(r"\bundefined\b", "null", src)
        s = re.sub(r",\s*([}\]])", r"\1", s)
        return json.loads(s)


class XHSClient:
    def __init__(self, cookie: str = "", timeout: int = 30):
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": UA, "Referer": f"https://{_XHS_HOST}/"})
        self.timeout = timeout
        self._cookie = cookie or self._load_cookie()

    # ---------- cookie ----------
    def set_cookie(self, cookie: str) -> None:
        self._cookie = cookie
        try:
            _COOKIE_JSON.parent.mkdir(parents=True, exist_ok=True)
            _COOKIE_JSON.write_text(json.dumps({"cookie": cookie}, ensure_ascii=False))
        except Exception:  # noqa: BLE001
            pass

    @staticmethod
    def _load_cookie() -> str:
        try:
            return json.loads(_COOKIE_JSON.read_text()).get("cookie", "")
        except Exception:  # noqa: BLE001
            return ""

    def has_cookie(self) -> bool:
        return bool(self._cookie and ("a1" in self._cookie or "web_session" in self._cookie
                                     or "xsecappid" in self._cookie))

    # ---------- resolve ----------
    def resolve(self, url: str) -> dict[str, Any]:
        """分享短链/笔记链 → 笔记 ID。支持 explore/discovery/item/xhslink 短链。"""
        url = (url or "").strip()
        if not url:
            return {"ok": False, "error": "空链接"}
        m = _ID_RE.search(url)
        if m:
            return {"ok": True, "id": m.group(1), "share": url,
                    "canonical": f"https://{_XHS_HOST}/explore/{m.group(1)}"}
        # xhslink 短链跟重定向
        try:
            r = self.s.get(url, headers={"User-Agent": UA}, allow_redirects=True,
                           timeout=self.timeout)
            final = r.url
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": f"短链解析失败: {e}"}
        m = _ID_RE.search(final) or _ID_RE.search(r.text or "")
        if not m:
            return {"ok": False, "error": f"落地页未找到笔记 ID: {final[:80]}"}
        return {"ok": True, "id": m.group(1), "share": url, "resolved": final,
                "canonical": f"https://{_XHS_HOST}/explore/{m.group(1)}"}

    # ---------- 提取 ----------
    def _fetch_state(self, url: str) -> dict[str, Any]:
        """抓笔记页 __INITIAL_STATE__（带 cookie 优先；curl_cffi impersonate 兜底）。"""
        headers = {"User-Agent": UA, "Referer": f"https://{_XHS_HOST}/"}
        if self._cookie:
            headers["Cookie"] = self._cookie
        text = ""
        # 优先 curl_cffi impersonate（TLS 指纹，过基础反爬）
        try:
            from curl_cffi import requests as cr
            r = cr.get(url, impersonate="chrome124", headers=headers,
                       timeout=self.timeout)
            text = r.text
        except Exception:  # noqa: BLE001
            try:
                text = self.s.get(url, headers=headers, timeout=self.timeout).text
            except Exception as e:  # noqa: BLE001
                return {"ok": False, "error": f"抓取失败: {e}"}
        m = re.search(r"window\.__INITIAL_STATE__\s*=\s*(\{.*?\})\s*</script>", text, re.S)
        if not m:
            return {"ok": False, "error": "页面无 __INITIAL_STATE__（结构变化）"}
        try:
            d = _js_to_json(m.group(1))
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": f"__INITIAL_STATE__ 解析失败: {e}"}
        return {"ok": True, "state": d, "login": bool((d.get("login") or {}))}

    def _extract_media(self, state: dict, nid: str) -> dict[str, Any]:
        """从 noteDetailMap 递归抠媒体直链（视频 masterUrl / 图 imageList）——同抖音套路。"""
        ndm = ((state.get("note") or {}).get("noteDetailMap")) or {}
        note = (ndm.get(nid) or {}).get("note") or {}
        # 递归兜底：键名会变
        video_urls, image_urls, desc, author = [], [], "", ""

        def walk(o):
            nonlocal desc, author
            if isinstance(o, dict):
                for k, v in o.items():
                    kl = k.lower()
                    if isinstance(v, str):
                        if ".mp4" in v or "sns-video" in v:
                            video_urls.append(v)
                        elif ("xhscdn" in v and ("jpg" in v or "jpeg" in v or "webp" in v
                                                 or "/spectrum/" in v)):
                            image_urls.append(v)
                    if kl in ("masterurl",) and isinstance(v, str):
                        video_urls.append(v)
                    if kl in ("urdefault", "url_default", "urldefault") and isinstance(v, str):
                        image_urls.append(v)
                    if kl in ("desc", "title") and isinstance(v, str) and not desc:
                        desc = v
                    if kl == "nickname" and isinstance(v, str) and not author:
                        author = v
                    walk(v)
            elif isinstance(o, list):
                for v in o[:20]:
                    walk(v)

        walk(note if note else state)
        video_urls = list(dict.fromkeys(video_urls))
        image_urls = list(dict.fromkeys(image_urls))
        return {"ok": bool(video_urls or image_urls),
                "video_urls": video_urls, "image_urls": image_urls,
                "desc": desc, "author": author, "type":
                ("video" if video_urls else "image" if image_urls else "unknown")}

    def info(self, url: str) -> dict[str, Any]:
        """元数据 + 媒体直链。无登录态/noteDetailMap 空时如实报边界（不假装）。"""
        res = self.resolve(url)
        if not res.get("ok"):
            return res
        target = res.get("resolved") or res["canonical"]
        st = self._fetch_state(target)
        if not st.get("ok"):
            return {**res, **st, "auto_download": False}
        media = self._extract_media(st["state"], res["id"])
        if media["ok"]:
            return {"ok": True, "id": res["id"], "type": media["type"],
                    "title": media["desc"], "author": media["author"],
                    "video_urls": media["video_urls"], "image_urls": media["image_urls"],
                    "url": (media["video_urls"] or media["image_urls"])[0],
                    "auto_download": True, "source": "initial_state"}
        return {"ok": False, "id": res["id"], "auto_download": False,
                "need_login": not self.has_cookie(),
                "error": ("未登录/noteDetailMap 为空：小红书媒体靠登录态+签名 API 异步填充，"
                          "SSR 不含直链。请在 WebUI 配小红书登录 cookie（浏览器登录后复制），"
                          "或用 GitHub 开源 XHS-Downloader 桌面助手抓取"),
                "hint": "cookies 配好后重试 xhs_info"}

    # ---------- 下载 ----------
    def download(self, url: str, out_dir: str | Path, filename: str = "") -> dict[str, Any]:
        """下载小红书视频/图到 out_dir（有 cookie 抠出直链才成功）。"""
        meta = self.info(url)
        if not meta.get("ok"):
            return meta
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        base = filename or (meta.get("title") or f"xhs_{meta['id']}")[:80]
        base = re.sub(r"[\\/:*?\"<>|\s]+", "_", base)
        headers = {"User-Agent": UA, "Referer": f"https://{_XHS_HOST}/"}
        if self._cookie:
            headers["Cookie"] = self._cookie
        files = []
        try:
            if meta["type"] == "video":
                dest = out_dir / f"{base}.mp4"
                self._stream(meta["video_urls"][0], headers, dest)
                files.append(dest)
            else:
                for i, u in enumerate(meta["image_urls"][:12], 1):
                    dest = out_dir / f"{base}_{i}.jpg"
                    self._stream(u, headers, dest)
                    files.append(dest)
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": f"下载失败: {e}", "id": meta["id"]}
        files = [f for f in files if f.exists() and f.stat().st_size > 0]
        if not files:
            return {"ok": False, "error": "下载得到空文件", "id": meta["id"]}
        return {"ok": True, "id": meta["id"], "type": meta["type"],
                "files": [str(f) for f in files], "count": len(files),
                "title": meta.get("title", ""), "author": meta.get("author", "")}

    @staticmethod
    def _stream(url: str, headers: dict, dest: Path, chunk: int = 1 << 16) -> None:
        with requests.get(url, headers=headers, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(dest, "wb") as f:
                for part in r.iter_content(chunk):
                    f.write(part)

    def status(self) -> dict[str, Any]:
        """如实报告小红书能力边界（不假装能未登录自动下载）。"""
        return {
            "ok": True,
            "auto_download": self.has_cookie(),
            "has_cookie": self.has_cookie(),
            "reason": ("小红书笔记媒体靠登录态 + x-s/x-t 请求签名 API 异步填充，SSR 不含直链；"
                       "未登录 noteDetailMap 为空、feed API 返回 code:-1、yt-dlp extractor 报 "
                       "No video formats found。有登录 cookie 时递归抠 masterUrl/imageList 可下载。"),
            "supported": "配小红书登录 cookie（WebUI/CLI）→ xhs_info/xhs_download 提取+下载",
            "desktop_helper": "JoeanAmier/XHS-Downloader（GitHub 开源，维护签名 + cookie 池）",
            "usage": "xhs_cookies 设 cookie → xhs_download(url)；无 cookie 时如实返回 need_login",
        }
