"""Bilibili client: WBI-signed search, video info, subtitles, audio/video download.

No third-party SDK: the WBI signature is ~40 lines and yt-dlp (optional) handles
DASH muxing. Cookies (SESSDATA etc.) are optional but unlock higher quality and
member-only content; they are stored by the WebUI.
"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import subprocess
import time
import urllib.parse
from pathlib import Path
from typing import Any, Iterable

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

API = "https://api.bilibili.com"

# 64-position mixin table used by WBI signing
_MIXIN_KEY_ENC_TAB = [
    46, 47, 18, 2, 53, 8, 23, 32, 15, 50, 10, 31, 58, 3, 45, 35, 27, 43, 5, 49,
    33, 9, 42, 19, 29, 28, 14, 39, 12, 38, 41, 13, 37, 48, 7, 16, 24, 55, 40,
    61, 26, 17, 0, 1, 60, 51, 30, 4, 22, 25, 54, 21, 56, 59, 6, 63, 57, 62, 11,
    36, 20, 34, 44, 52,
]


def _mixin_key(raw: str) -> str:
    return "".join(raw[i] for i in _MIXIN_KEY_ENC_TAB)[:32]


def _strip_bad(s: str) -> str:
    return re.sub(r"[!'()*]", "", s)


class BilibiliClient:
    def __init__(self, cookie: str = "", timeout: int = 20):
        self.s = requests.Session()
        self.s.headers.update({
            "User-Agent": UA,
            "Referer": "https://www.bilibili.com/",
            "Accept": "application/json, text/plain, */*",
        })
        if cookie:
            self.set_cookie(cookie)
        self.timeout = timeout
        self._wbi: tuple[str, str] | None = None
        self._wbi_ts = 0.0

    # ---------- auth ----------
    def set_cookie(self, cookie: str) -> None:
        for part in cookie.replace("\n", ";").split(";"):
            if "=" in part:
                k, v = part.split("=", 1)
                self.s.cookies.set(k.strip(), v.strip(), domain=".bilibili.com")

    def cookie_string(self) -> str:
        return "; ".join(f"{c.name}={c.value}" for c in self.s.cookies)

    def nav(self) -> dict[str, Any]:
        r = self.s.get(f"{API}/x/web-interface/nav", timeout=self.timeout)
        return r.json()

    def is_login(self) -> bool:
        try:
            return bool(self.nav().get("data", {}).get("isLogin"))
        except Exception:
            return False

    # ---------- WBI ----------
    def _wbi_keys(self) -> tuple[str, str]:
        if self._wbi and time.time() - self._wbi_ts < 3600:
            return self._wbi
        try:
            data = self.nav()["data"]["wbi_img"]
            img = data["img_url"].rsplit("/", 1)[-1].split(".")[0]
            sub = data["sub_url"].rsplit("/", 1)[-1].split(".")[0]
            self._wbi = (_mixin_key(img + sub), "")
            self._wbi_ts = time.time()
        except Exception:
            # fall back to the long-lived public pair (works for anonymous search)
            self._wbi = (_mixin_key(
                "7cd084941338484aae1ad9425b84077c" "4932caff0ff746eab6f01bf08b70ac45"), "")
            self._wbi_ts = time.time() - 3000
        return self._wbi

    def signed_params(self, params: dict[str, Any]) -> dict[str, Any]:
        mixin, _ = self._wbi_keys()
        p = dict(params)
        p["wts"] = int(time.time())
        items = sorted(p.items())
        query = urllib.parse.urlencode(
            [(k, _strip_bad(str(v))) for k, v in items if v is not None]
        )
        p["w_rid"] = hashlib.md5((query + mixin).encode()).hexdigest()
        return p

    def get(self, path: str, params: dict[str, Any] | None = None, wbi: bool = True) -> dict:
        params = dict(params or {})
        if wbi:
            params = self.signed_params(params)
        r = self.s.get(f"{API}{path}", params=params, timeout=self.timeout)
        r.raise_for_status()
        return r.json()

    # ---------- content ----------
    def search(self, keyword: str, page: int = 1, page_size: int = 20,
               order: str = "totalrank", duration: int = 0, tids: int = 0) -> dict:
        """Search videos. order: totalrank|click|pubdate|dm|stow; duration: 0全部 1<10min 2 10-30 3 30-60 4>60."""
        data = self.get("/x/web-interface/wbi/search/type", {
            "search_type": "video", "keyword": keyword, "page": page,
            "page_size": page_size, "order": order, "duration": duration, "tids": tids,
        })
        res = data.get("data") or {}
        items = []
        for it in res.get("result") or []:
            items.append({
                "bvid": it.get("bvid"),
                "aid": it.get("aid"),
                "title": re.sub(r"</?em[^>]*>", "", it.get("title", "")),
                "author": it.get("author"),
                "mid": it.get("mid"),
                "duration": it.get("duration"),
                "play": it.get("play"),
                "danmaku": it.get("video_review"),
                "pubdate": it.get("pubdate"),
                "description": re.sub(r"</?em[^>]*>", "", it.get("description", "")),
                "pic": it.get("pic"),
                "url": f"https://www.bilibili.com/video/{it.get('bvid')}",
            })
        return {"total": res.get("numResults"), "page": page, "items": items}

    def video_info(self, bvid: str) -> dict:
        data = self.get("/x/web-interface/view", {"bvid": bvid}, wbi=False)
        d = data.get("data") or {}
        pages = [{
            "cid": p.get("cid"), "page": p.get("page"), "part": p.get("part"),
            "duration": p.get("duration"),
        } for p in d.get("pages") or []]
        return {
            "bvid": d.get("bvid"), "aid": d.get("aid"), "title": d.get("title"),
            "desc": d.get("desc"), "duration": d.get("duration"), "pubdate": d.get("pubdate"),
            "owner": (d.get("owner") or {}).get("name"), "owner_mid": (d.get("owner") or {}).get("mid"),
            "stat": d.get("stat"), "pages": pages,
            "pic": d.get("pic"),
        }

    def subtitles(self, bvid: str, cid: int | None = None) -> list[dict]:
        """Return subtitle tracks (requires login cookies for most videos)."""
        if cid is None:
            info = self.video_info(bvid)
            cid = (info["pages"] or [{}])[0].get("cid")
        data = self.get("/x/player/v2", {"bvid": bvid, "cid": cid}, wbi=False)
        subs = ((data.get("data") or {}).get("subtitle") or {}).get("subtitles") or []
        out = []
        for s in subs:
            url = s.get("subtitle_url") or ""
            if url.startswith("//"):
                url = "https:" + url
            out.append({"lan": s.get("lan"), "lan_doc": s.get("lan_doc"), "url": url})
        return out

    def fetch_subtitle_text(self, sub_url: str) -> str:
        r = self.s.get(sub_url, timeout=self.timeout)
        data = r.json()
        body = data.get("body") or []
        lines = []
        for seg in body:
            t = int(float(seg.get("from", 0)))
            lines.append(f"[{t//60:02d}:{t%60:02d}] {seg.get('content','').strip()}")
        return "\n".join(lines)

    def audio_url(self, bvid: str, cid: int | None = None) -> dict:
        """Direct audio stream URL (DASH) — best-effort, no download."""
        if cid is None:
            info = self.video_info(bvid)
            cid = (info["pages"] or [{}])[0].get("cid")
        data = self.get("/x/player/playurl", {
            "bvid": bvid, "cid": cid, "fnval": 16, "fourk": 1, "qn": 80,
        }, wbi=False)
        dash = (data.get("data") or {}).get("dash") or {}
        audios = dash.get("audio") or []
        if not audios:
            return {"ok": False, "reason": data.get("message", "no dash audio")}
        audios.sort(key=lambda a: a.get("bandwidth", 0), reverse=True)
        a = audios[0]
        base = a.get("baseUrl") or a.get("base_url")
        return {"ok": True, "url": base, "bandwidth": a.get("bandwidth"),
                "codec": a.get("codecs"), "backup": a.get("backupUrl") or []}

    # ---------- download ----------
    def download(self, bvid: str, out_dir: str | Path, audio_only: bool = False,
                 quality: str = "1080", page: int = 1) -> dict:
        """Download via yt-dlp (installed separately). Returns the produced file path."""
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        if not shutil.which("yt-dlp"):
            return {"ok": False, "error": "yt-dlp not installed: pip install yt-dlp (and ffmpeg)"}
        url = f"https://www.bilibili.com/video/{bvid}"
        if page and page > 1:
            url += f"?p={page}"
        if audio_only:
            candidates = ["bestaudio/best"]
        else:
            # 竖屏视频 height 恒为 1920，不能只按 height 过滤，逐个降级试
            candidates = [
                f"bv*[height<={quality}][ext=mp4]+ba[ext=m4a]/b[height<={quality}]",
                f"bv*[width<={quality}]+ba/bv*[height<={quality}]+ba",
                "bv*+ba/b",
                "b",
            ]
        base = [
            "yt-dlp", url, "-o", str(out_dir / "%(title).120B.%(ext)s"),
            "--no-playlist", "-f", candidates[0], "--print", "after_move:filepath",
        ]
        if audio_only:
            base += ["-x", "--audio-format", "mp3", "--audio-quality", "0"]
        if self.cookie_string():
            base += ["--add-header", f"Cookie:{self.cookie_string()}"]
        base += ["--add-header", f"User-Agent:{UA}", "--add-header", "Referer:https://www.bilibili.com/"]
        last_err = ""
        for fmt in candidates:
            attempt = list(base)
            attempt[attempt.index("-f") + 1] = fmt
            proc = subprocess.run(attempt, capture_output=True, text=True, timeout=1800)
            if proc.returncode == 0:
                path = (proc.stdout or "").strip().splitlines()[-1] if proc.stdout.strip() else ""
                return {"ok": True, "path": path, "audio_only": audio_only, "format": fmt}
            last_err = (proc.stderr or "")[-600:]
        return {"ok": False, "error": last_err}

    def hot(self, pn: int = 1, ps: int = 20) -> list[dict]:
        data = self.get("/x/web-interface/popular", {"pn": pn, "ps": ps}, wbi=False)
        return [{"bvid": v.get("bvid"), "title": v.get("title"), "owner": (v.get("owner") or {}).get("name"),
                 "duration": v.get("duration"), "stat": v.get("stat")}
                for v in (data.get("data") or {}).get("list") or []]

    def comments(self, bvid: str, pn: int = 1, ps: int = 20) -> list[dict]:
        info = self.video_info(bvid)
        data = self.get("/x/v2/reply", {"type": 1, "oid": info["aid"], "pn": pn, "ps": ps}, wbi=False)
        out = []
        for c in ((data.get("data") or {}).get("replies") or []):
            out.append({"user": (c.get("member") or {}).get("uname"), "like": c.get("like"),
                        "content": (c.get("content") or {}).get("message"), "ctime": c.get("ctime")})
        return out
