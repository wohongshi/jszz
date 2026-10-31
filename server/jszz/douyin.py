"""Douyin (抖音) client: 分享链解析 + yt-dlp 出元数据/下载。

借鉴 GitHub 开源项目：
- yt-dlp（自带 Douyin extractor，签名/cookie/反爬它处理，我们只做薄封装）
- Evil0ctal/Douyin_TikTok_Download_API（分享短链 v.douyin.com → canonical 视频 URL 的解析思路）

零第三方 SDK：requests 解析分享短链拿 canonical URL/ID → yt-dlp 出元数据与文件。
用途定位：给老师备课取**课堂教学素材**（实验演示、科普短视频），下载文件落盘当素材。
"""
from __future__ import annotations

import re
import shutil
import json
import subprocess
from pathlib import Path
from typing import Any

import requests

UA = ("Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 "
      "(KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1")
DESKTOP_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
CHROME_BIN = "/usr/bin/chromium"
CHROMEDRIVER = "/usr/bin/chromedriver"
_COOKIE_JSON = Path.home() / ".jszz" / "cache" / "douyin_cookies.json"
_COOKIE_FILE = Path.home() / ".jszz" / "cache" / "douyin_cookies.txt"

# 视频 ID 提取：/video/<id>、/note/<id>、/share/video/<id>
_ID_RE = re.compile(r"(?:video|note|share/video)/(\d{15,25})")
_SHORT_RE = re.compile(r"https?://v\.douyin\.com/[\w-]+/?")
_IS_DOUYIN = re.compile(r"(douyin\.com|iesdouyin\.com)", re.I)


def _harvest_with_chrome(target: str) -> dict[str, Any]:
    """无头 chromium 打开抖音过 JS 质询 → CDP Network.getAllCookies 收割 → 写 Netscape cookie 文件。"""
    import json as _json
    import subprocess as _sp
    import time as _time
    import urllib.request as _u
    import http.cookiejar as _cj
    # 起 chromedriver
    s = __import__("socket").socket(); s.bind(("127.0.0.1", 0)); port = s.getsockname()[1]; s.close()
    proc = _sp.Popen([CHROMEDRIVER, f"--port={port}"],
                     stdout=_sp.DEVNULL, stderr=_sp.DEVNULL)
    base = f"http://127.0.0.1:{port}"
    try:
        for _ in range(40):
            try:
                _u.urlopen(base + "/status", timeout=1); break
            except Exception:
                _time.sleep(0.25)
        args = ["--headless=new", "--no-sandbox", "--disable-gpu",
                "--disable-dev-shm-usage", "--window-size=1200,900"]
        try:
            _u.urlopen("http://127.0.0.1:7890", timeout=1)
            args.append("--proxy-server=http://127.0.0.1:7890")
        except Exception:
            pass
        body = _json.dumps({"capabilities": {"alwaysMatch": {
            "browserName": "chrome",
            "goog:chromeOptions": {"binary": CHROME_BIN, "args": args}}}})
        req = _u.Request(base + "/session", data=body.encode(),
                         headers={"Content-Type": "application/json"})
        sid = _json.loads(_u.urlopen(req, timeout=30).read())["value"]["sessionId"]
        sess = f"{base}/session/{sid}"
        try:
            # 打开目标页，等 JS 质询跑完（acw_sc__v2 由脚本回填 cookie）
            _u.urlopen(_u.Request(sess + "/url",
                       data=_json.dumps({"url": target}).encode(),
                       headers={"Content-Type": "application/json"}), timeout=45)
            _time.sleep(3.5)
            # CDP 取所有 cookie（含 httpOnly）
            cdp = _json.dumps({"cmd": "Network.getAllCookies", "params": {}})
            rq = _u.Request(sess + "/goog/cdp/execute", data=cdp.encode(),
                            headers={"Content-Type": "application/json"})
            allc = _json.loads(_u.urlopen(rq, timeout=20).read())["value"]["cookies"]
        finally:
            try:
                _u.urlopen(_u.Request(sess, method="DELETE"), timeout=5)
            except Exception:
                pass
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
    names = [c.get("name") for c in allc]
    # 空名 cookie 会写坏文件/污染头，必须剔除
    allc = [c for c in allc if (c.get("name") or "").strip()]
    if not allc:
        return {"ok": False, "error": "未收割到有效 cookie（页面可能是验证墙）"}
    # 用 MozillaCookieJar.save() 写 Netscape 文件——格式保证 _really_load 能回读。
    # 手搓常栽在断言 domain_specified == initial_dot：第 2 列必须为 TRUE 当且仅当域名以 '.' 开头。
    from http.cookiejar import MozillaCookieJar, Cookie as _Cookie
    _COOKIE_FILE.parent.mkdir(parents=True, exist_ok=True)
    jar = MozillaCookieJar(str(_COOKIE_FILE))
    for c in allc:
        domain = c.get("domain") or ".douyin.com"
        # 归一到父域（带点）：host-only（www.douyin.com）cookie 发不到 iesdouyin.com 等
        # 子域，yt-dlp 的 Douyin extractor 跨子域取数据时会丢 cookie → Fresh cookies needed
        if "douyin.com" in domain:
            domain = ".douyin.com"
        elif "iesdouyin.com" in domain:
            domain = ".iesdouyin.com"
        initial_dot = domain.startswith(".")
        try:
            exp = int(float(c.get("expires") or c.get("expirationDate") or 0))
        except (TypeError, ValueError):
            exp = 0
        if exp <= 0:
            # session cookie（CDP 给 -1/0）会被 http.cookiejar 当 invalid expires 跳过；
            # 给一个近未来时间戳保证不被丢（不影响会话有效性）
            exp = int(_time.time()) + 3600
        jar.set_cookie(_Cookie(
            version=0, name=c["name"], value=c.get("value", ""),
            port=None, port_specified=False,
            domain=domain, domain_specified=initial_dot, domain_initial_dot=initial_dot,
            path=c.get("path") or "/", path_specified=True,
            secure=bool(c.get("secure")), expires=exp,
            discard=False, comment=None, comment_url=None, rest={}, rfc2109=False))
    jar.save(ignore_discard=True, ignore_expires=True)
    cookie_str = "; ".join(f"{c['name']}={c.get('value','')}" for c in allc)
    _COOKIE_JSON.parent.mkdir(parents=True, exist_ok=True)
    _COOKIE_JSON.write_text(json.dumps(
        {"cookie_str": cookie_str, "names": names, "count": len(allc)},
        ensure_ascii=False))
    return {"ok": True, "cookie_str": cookie_str, "cookie_file": str(_COOKIE_FILE),
            "names": names, "count": len(allc)}


def is_douyin(url: str) -> bool:
    return bool(_IS_DOUYIN.search(url or ""))


class DouyinClient:
    def __init__(self, cookie: str = "", timeout: int = 25):
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": DESKTOP_UA,
                               "Referer": "https://www.douyin.com/",
                               "Accept-Language": "zh-CN,zh;q=0.9"})
        if cookie:
            self.set_cookie(cookie)
        self.timeout = timeout

    # ---------- auth ----------
    def set_cookie(self, cookie: str) -> None:
        for part in cookie.replace("\n", ";").split(";"):
            if "=" in part:
                k, v = part.split("=", 1)
                self.s.cookies.set(k.strip(), v.strip(), domain=".douyin.com")

    def cookie_string(self) -> str:
        return "; ".join(f"{c.name}={c.value}" for c in self.s.cookies)

    # ---------- 新鲜 cookie 收割（过 JS 反爬质询）----------
    def harvest_cookies(self, video_url: str = "") -> dict[str, Any]:
        """用无头 chromedriver 过抖音 JS 质询，收割会话 cookie（含 httpOnly）。

        yt-dlp 的 Douyin extractor 要「新鲜 cookie」（ttwid/msToken/acw_sc__v2 等），
        curl 拿不到（只下发 __ac_nonce 前置 cookie）——必须真浏览器跑完质询脚本。
        CDP Network.getAllCookies 连 httpOnly 一起取。返回 {ok, cookie_file, names}。
        """
        if not shutil.which("chromedriver") or not Path(CHROME_BIN).exists():
            return {"ok": False, "error": "缺 chromedriver/chromium，无法过抖音 JS 质询"}
        target = video_url or "https://www.douyin.com/"
        return _harvest_with_chrome(target)

    # ---------- resolve ----------
    def resolve(self, url: str) -> dict[str, Any]:
        """分享短链/分享页 → canonical 视频 URL + 视频 ID。

        支持：v.douyin.com/<code>/ 短链、iesdouyin.com/share/video/<id>、
        douyin.com/video/<id>、douyin.com/share/video/<id>、/note/<id>（图文）。
        短链用移动端 UA 跟重定向拿落地页（桌面 UA 常被反爬拦到验证页）。
        """
        url = (url or "").strip()
        if not url:
            return {"ok": False, "error": "空链接"}
        # 先试原地抠 ID
        m = _ID_RE.search(url)
        if m:
            vid = m.group(1)
            kind = "note" if "/note/" in url else "video"
            return {"ok": True, "id": vid, "kind": kind,
                    "canonical": f"https://www.douyin.com/video/{vid}",
                    "share": url}
        # 短链：跟重定向
        try:
            r = self.s.get(url, headers={"User-Agent": UA}, allow_redirects=True,
                           timeout=self.timeout)
            final = r.url
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": f"短链解析失败: {e}"}
        m = _ID_RE.search(final) or _ID_RE.search(r.text or "")
        if not m:
            return {"ok": False, "error": f"落地页未找到视频 ID（可能触发验证页）: {final[:80]}"}
        vid = m.group(1)
        kind = "note" if "/note/" in final else "video"
        return {"ok": True, "id": vid, "kind": kind,
                "canonical": f"https://www.douyin.com/video/{vid}",
                "share": url, "resolved": final}

    # ---------- yt-dlp ----------
    @staticmethod
    def _ytdlp() -> str | None:
        return shutil.which("yt-dlp")

    def _run_ytdlp(self, args: list[str], cookie_file: str = "") -> subprocess.CompletedProcess:
        full = [self._ytdlp() or "yt-dlp"]
        if cookie_file and Path(cookie_file).exists():
            # yt-dlp 的 Douyin extractor 只从 --cookies 文件/cookiejar 读会话 cookie，
            # --add-header Cookie: 它不认（会仍报 Fresh cookies needed）。文件由
            # MozillaCookieJar.save() 写出，格式保证能回读（见 _harvest_with_chrome）。
            full += ["--cookies", cookie_file]
        full += args
        return subprocess.run(full, capture_output=True, text=True, timeout=180)

    def _fresh_cookie_file(self, url: str = "") -> str:
        """收割新鲜 cookie → Netscape 文件路径；15 分钟内缓存复用（起浏览器慢）。"""
        cache = _COOKIE_FILE
        try:
            if cache.exists():
                import time
                if time.time() - cache.stat().st_mtime < 900:
                    return str(cache)
        except Exception:  # noqa: BLE001
            pass
        r = self.harvest_cookies(url)
        return r.get("cookie_file", "") if r.get("ok") else ""

    # ---------- 提取（主路：分享页 _ROUTER_DATA → play_addr）----------
    def _fetch_play_addr(self, vid: str, cookie_file: str = "") -> dict[str, Any]:
        """分享页 _ROUTER_DATA 里抠 aweme item（play_addr/desc/author）——Evil0ctal 核心思路。

        抖音把视频数据 URL-编码塞进 `window._ROUTER_DATA`，递归找含 play_addr 的节点。
        返回 {ok, urls, desc, author, cover, duration}。绕开 yt-dlp 那个挑剔的 Douyin extractor。
        """
        cookies = {}
        if cookie_file and Path(cookie_file).exists():
            for line in Path(cookie_file).read_text().splitlines():
                if line.startswith("#") or not line.strip():
                    continue
                p = line.split("\t")
                if len(p) >= 7 and p[6]:
                    cookies[p[5]] = p[6]
        headers = {"User-Agent": UA,
                   "Cookie": "; ".join(f"{k}={v}" for k, v in cookies.items()),
                   "Referer": "https://www.douyin.com/"}
        target = f"https://www.iesdouyin.com/share/video/{vid}/"
        try:
            r = self.s.get(target, headers=headers, timeout=self.timeout)
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": f"分享页请求失败: {e}"}
        m = re.search(r"window\._ROUTER_DATA\s*=\s*(\{.*?\})\s*</script>", r.text, re.S)
        if not m:
            if "验证" in r.text:
                return {"ok": False, "error": "抖音验证墙（反爬质询），稍后重试或更新 cookie"}
            return {"ok": False, "error": "分享页无 _ROUTER_DATA（页面结构变化或视频不存在）"}
        try:
            d = json.loads(m.group(1))
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": f"_ROUTER_DATA JSON 解析失败: {e}"}
        video, aweme = None, None

        def walk(o):
            nonlocal video, aweme
            if isinstance(o, dict):
                if "play_addr" in o and isinstance(o.get("play_addr"), dict):
                    video = o
                if "author" in o and "video" in o and "desc" in o:
                    aweme = o
                for v in o.values():
                    walk(v)
            elif isinstance(o, list):
                for v in o:
                    walk(v)

        walk(d)
        urls = ((video or {}).get("play_addr") or {}).get("url_list") or []
        if not urls:
            return {"ok": False, "error": "未找到 play_addr（视频可能已删除/私密）"}
        aweme = aweme or {}
        return {"ok": True, "urls": urls,
                "desc": aweme.get("desc", ""),
                "author": (aweme.get("author") or {}).get("nickname", ""),
                "cover": ((video or {}).get("cover") or {}).get("url_list", [""])[0],
                "duration": (video or {}).get("duration")}

    def _best_play_url(self, urls: list[str]) -> str:
        # playwm（带水印）→ play（无水印）；优先无水印
        u = urls[0] if urls else ""
        return u.replace("playwm", "play") if "playwm" in u else u

    def info(self, url: str) -> dict[str, Any]:
        """元数据（标题/作者/时长/封面）——分享页 play_addr 提取，失败降级 yt-dlp。"""
        res = self.resolve(url)
        if not res.get("ok"):
            return res
        ck = self._fresh_cookie_file(res.get("resolved") or res["canonical"])
        pa = self._fetch_play_addr(res["id"], cookie_file=ck)
        if pa.get("ok"):
            return {"ok": True, "id": res["id"], "kind": res["kind"],
                    "title": pa.get("desc", ""), "author": pa.get("author", ""),
                    "duration": (pa.get("duration") or 0) / 1000 if pa.get("duration") else None,
                    "cover": pa.get("cover"), "url": self._best_play_url(pa["urls"]),
                    "source": "play_addr"}
        # 降级 yt-dlp
        if not self._ytdlp():
            return {"ok": False, "error": pa.get("error", ""),
                    "id": res["id"], "hint": "装 yt-dlp 可作兜底解析"}
        target = res.get("resolved") or res["canonical"]
        proc = self._run_ytdlp([target, "--no-playlist", "--dump-json",
                                "--no-warnings", "--no-check-certificates"],
                               cookie_file=ck)
        if proc.returncode != 0:
            return {"ok": False, "error": (proc.stderr or "yt-dlp 解析失败")[-300:],
                    "id": res["id"]}
        try:
            d = json.loads(proc.stdout)
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": f"元数据 JSON 解析失败: {e}"}
        return {
            "ok": True, "id": res["id"], "kind": res["kind"],
            "title": d.get("title", ""), "author": d.get("uploader", "")
                    or d.get("creator", ""),
            "duration": d.get("duration"), "view_count": d.get("view_count"),
            "like_count": d.get("like_count"), "comment_count": d.get("comment_count"),
            "cover": d.get("thumbnail"), "url": target, "source": "yt-dlp",
            "description": (d.get("description") or "")[:200],
        }

    def download(self, url: str, out_dir: str | Path, audio_only: bool = False,
                 quality: str = "1080") -> dict[str, Any]:
        """下载视频/音频到 out_dir。主路走 play_addr 直链（requests 流式），失败降级 yt-dlp。"""
        res = self.resolve(url)
        if not res.get("ok"):
            return res
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        ck = self._fresh_cookie_file(res.get("resolved") or res["canonical"])
        pa = self._fetch_play_addr(res["id"], cookie_file=ck)
        if pa.get("ok"):
            dl = self._best_play_url(pa["urls"])
            title = re.sub(r"[\\/:*?\"<>|\s]+", "_", (pa.get("desc") or f"douyin_{res['id']}"))[:80]
            try:
                cookies = {}
                if ck and Path(ck).exists():
                    for line in Path(ck).read_text().splitlines():
                        if line.startswith("#") or not line.strip():
                            continue
                        p = line.split("\t")
                        if len(p) >= 7 and p[6]:
                            cookies[p[5]] = p[6]
                h = {"User-Agent": UA, "Referer": "https://www.douyin.com/",
                     "Cookie": "; ".join(f"{k}={v}" for k, v in cookies.items())}
                if audio_only:
                    # 先下 mp4 再 ffmpeg 抽音轨
                    raw = out_dir / f"{title}.mp4"
                    self._stream_to(dl, h, raw)
                    mp3 = out_dir / f"{title}.mp3"
                    subprocess.run(["ffmpeg", "-y", "-i", str(raw), "-vn",
                                    "-acodec", "libmp3lame", str(mp3)],
                                   capture_output=True, timeout=120)
                    raw.unlink(missing_ok=True)
                    final = mp3 if mp3.exists() else raw
                else:
                    final = out_dir / f"{title}.mp4"
                    self._stream_to(dl, h, final)
                if not final.exists() or final.stat().st_size == 0:
                    return {"ok": False, "error": "play_addr 下载失败（空文件）", "id": res["id"]}
                return {"ok": True, "id": res["id"], "file": str(final),
                        "size": final.stat().st_size, "title": final.stem,
                        "author": pa.get("author", ""), "source": "play_addr",
                        "audio_only": audio_only}
            except Exception as e:  # noqa: BLE001
                return {"ok": False, "error": f"play_addr 下载异常: {e}", "id": res["id"]}
        # 降级 yt-dlp
        if not self._ytdlp():
            return {"ok": False, "error": pa.get("error", ""), "id": res["id"],
                    "hint": "装 yt-dlp 可作兜底下载"}
        target = res.get("resolved") or res["canonical"]
        out_tpl = str(out_dir / "%(title).80B.%(ext)s")
        base = [target, "-o", out_tpl, "--no-playlist",
                "--no-warnings", "--no-check-certificates",
                "--print", "after_move:filepath"]
        if audio_only:
            base += ["-f", "bestaudio", "-x", "--audio-format", "mp3"]
        else:
            base += ["-f", f"bv*[height<={quality}]+ba/b[height<={quality}]/b",
                     "--merge-output-format", "mp4"]
        proc = self._run_ytdlp(base, cookie_file=ck)
        path = (proc.stdout or "").strip().splitlines()
        path = path[-1] if path else ""
        if proc.returncode != 0 or not path or not Path(path).exists():
            return {"ok": False, "error": (proc.stderr or "下载失败")[-300:], "id": res["id"]}
        p = Path(path)
        return {"ok": True, "id": res["id"], "file": str(p),
                "size": p.stat().st_size, "title": p.stem, "source": "yt-dlp",
                "audio_only": audio_only}

    @staticmethod
    def _stream_to(url: str, headers: dict, dest: Path, chunk: int = 1 << 16) -> None:
        with requests.get(url, headers=headers, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(dest, "wb") as f:
                for part in r.iter_content(chunk):
                    f.write(part)

    def audio_url(self, url: str) -> dict[str, Any]:
        """音频直链（bestaudio）——不落盘，交给播放器/转写。"""
        if not self._ytdlp():
            return {"ok": False, "error": "yt-dlp 未安装"}
        res = self.resolve(url)
        if not res.get("ok"):
            return res
        target = res.get("resolved") or res["canonical"]
        ck = self._fresh_cookie_file(target)
        proc = self._run_ytdlp([target, "-f", "bestaudio", "-g",
                                "--no-playlist", "--no-warnings"], cookie_file=ck)
        if proc.returncode != 0:
            return {"ok": False, "error": (proc.stderr or "取直链失败")[-200:], "id": res["id"]}
        urls = (proc.stdout or "").strip().splitlines()
        return {"ok": True, "id": res["id"], "audio_url": urls[0] if urls else ""}
