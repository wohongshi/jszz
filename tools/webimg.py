#!/usr/bin/env python3
"""网络找图：按关键词搜教学示意图/素材图，下载入库并记来源署名。

图源（免 key）：
  1. 必应图搜（cn.bing.com/images/async，直连可用，中文教学图最全）——主力
  2. Wikimedia Commons（需代理 proxy-run，CC 授权、带作者信息）——兜底

用法：
  python3 tools/webimg.py search "热力环流示意图" [--count 6] [--source bing|commons]
  python3 tools/webimg.py download "<图URL>" [--title "..."]
"""
import argparse
import hashlib
import json
import re
import subprocess
import sys
import urllib.parse
import urllib.request
from pathlib import Path

ROOT = Path.home() / ".jszz" / "assets" / "web"
INDEX = ROOT / "index.jsonl"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/120.0 Safari/537.36")
CHROME_BIN = "/usr/bin/chromium"
CHROMEDRIVER = "/usr/bin/chromedriver"


class Browser:
    """chromedriver 原生 REST 驱动（零 pip 依赖）——webimg 的完整浏览器兜底。

    HTTP 抓取被反爬/超时挡住时用：搜索走真浏览器 DOM，下载走页面上下文 fetch
    （自带 Referer/Cookie，防盗链图也能取）。
    """

    def __init__(self):
        self.proc = None
        self.sid = ""
        self.port = 0

    def __enter__(self):
        import socket
        s = socket.socket()
        s.bind(("127.0.0.1", 0))
        self.port = s.getsockname()[1]
        s.close()
        # 境外图源（Wikimedia 等）直连不通；clash 按需代理在 7890 时自动带
        args = [CHROMEDRIVER, f"--port={self.port}"]
        self.proc = subprocess.Popen(
            args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        base = f"http://127.0.0.1:{self.port}"
        for _ in range(40):
            try:
                urllib.request.urlopen(base + "/status", timeout=1)
                break
            except Exception:
                import time
                time.sleep(0.25)
        chrome_args = [
            "--headless=new", "--no-sandbox", "--disable-gpu",
            "--disable-dev-shm-usage", "--window-size=1400,1000"]
        try:
            urllib.request.urlopen("http://127.0.0.1:7890", timeout=1)
            chrome_args.append("--proxy-server=http://127.0.0.1:7890")
        except Exception:
            pass
        body = json.dumps({"capabilities": {"alwaysMatch": {
            "browserName": "chrome",
            "goog:chromeOptions": {"binary": CHROME_BIN, "args": chrome_args}}}})
        req = urllib.request.Request(base + "/session", data=body.encode(),
                                     headers={"Content-Type": "application/json"})
        self.sid = json.loads(urllib.request.urlopen(req, timeout=30).read())["value"]["sessionId"]
        self.base = base
        return self

    def __exit__(self, *a):
        try:
            urllib.request.urlopen(
                urllib.request.Request(f"{self.base}/session/{self.sid}", method="DELETE"),
                timeout=5)
        except Exception:
            pass
        if self.proc:
            self.proc.terminate()
            try:
                self.proc.wait(timeout=5)
            except Exception:
                self.proc.kill()

    def get(self, url: str):
        body = json.dumps({"url": url})
        urllib.request.urlopen(urllib.request.Request(
            f"{self.base}/session/{self.sid}/url", data=body.encode(),
            headers={"Content-Type": "application/json"}), timeout=45)

    def js(self, script: str, timeout: int = 45, async_: bool = False):
        # async_=True 走 execute/async：回调即 arguments[-1]，脚本里直接用
        endpoint = "execute/async" if async_ else "execute/sync"
        body = json.dumps({"script": script, "args": []})
        req = urllib.request.Request(
            f"{self.base}/session/{self.sid}/{endpoint}", data=body.encode(),
            headers={"Content-Type": "application/json"})
        return json.loads(urllib.request.urlopen(req, timeout=timeout).read())["value"]


def search_bing_browser(query: str, count: int = 6) -> list[dict]:
    """兜底：真浏览器打开必应图搜页，从 DOM 里取 murl（HTTP 路径被反爬时用）。"""
    try:
        with Browser() as b:
            b.get("https://cn.bing.com/images/search?q=" + urllib.parse.quote(query))
            raw = b.js("""
                return Array.from(document.querySelectorAll('a.iusc'))
                  .slice(0, 30).map(a => a.getAttribute('m')).filter(Boolean);
            """)
    except Exception:
        return []
    out = []
    for item in raw or []:
        try:
            d = json.loads(item.replace("&quot;", '"'))
        except Exception:
            continue
        if d.get("murl"):
            out.append({"title": (d.get("t") or "").split("|")[0].strip()[:80],
                        "url": d["murl"], "page": d.get("purl", ""), "source": "bing"})
        if len(out) >= count:
            break
    return out


def download_browser(url: str) -> bytes:
    """兜底：页面上下文 fetch 取图（带 Referer/Cookie，防盗链可破）；失败返回 b""。

    坑：目标站常把 http 重定向到 https，之后页面内再 fetch `http://…` 会被
    Chrome 按「混合内容」拦掉（XHR ERR 0）——所以用**同源相对路径** fetch，
    协议自动跟随最终落地页。
    """
    try:
        with Browser() as b:
            parsed = urllib.parse.urlparse(url)
            b.get(f"{parsed.scheme}://{parsed.netloc}/")
            rel = parsed.path + ("?" + parsed.query if parsed.query else "")
            b64 = b.js(f"""
                const done = arguments[arguments.length - 1];
                const grab = (u) => fetch(u, {{credentials: 'include'}})
                  .then(r => r.blob())
                  .then(bl => {{const fr = new FileReader();
                     fr.onloadend = () => done(fr.result.split(',')[1] || '');
                     fr.readAsDataURL(bl);}});
                grab({json.dumps(rel)}).catch(() => grab({json.dumps(url)}))
                  .catch(() => done(''));
            """, timeout=60, async_=True)
            import base64
            return base64.b64decode(b64) if b64 else b""
    except Exception:
        return b""


def _get(url: str, timeout: int = 15, binary: bool = False):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read() if binary else r.read().decode("utf-8", "ignore")


def search_bing(query: str, count: int = 6) -> list[dict]:
    """必应图搜：解析 iusc 元素的 m=JSON（含 murl 真图地址）。HTTP 被挡 → 浏览器兜底。"""
    url = ("https://cn.bing.com/images/async?q=" + urllib.parse.quote(query) +
           f"&first=1&count={count * 2}")
    try:
        html = _get(url)
    except Exception:
        return search_bing_browser(query, count)
    out = []
    for raw in re.findall(r'class="iusc"[^>]*m="([^"]+)"', html):
        try:
            d = json.loads(raw.replace("&quot;", '"'))
        except Exception:
            continue
        if d.get("murl"):
            out.append({"title": (d.get("t") or "").split("|")[0].strip()[:80],
                        "url": d["murl"], "page": d.get("purl", ""), "source": "bing"})
        if len(out) >= count:
            break
    return out or search_bing_browser(query, count)


def search_commons(query: str, count: int = 6) -> list[dict]:
    """Wikimedia Commons API（免 key；境外站，直连不通时走代理）。"""
    api = ("https://commons.wikimedia.org/w/api.php?action=query&format=json"
           "&generator=search&gsrsearch=" + urllib.parse.quote(query) +
           f"&gsrnamespace=6&gsrlimit={count}&prop=imageinfo"
           "&iiprop=url|extmetadata&iiurlwidth=1280")
    try:
        raw = _get(api)
    except Exception:
        # 直连不通 → 经本机代理重试（clash 按需拉起）
        try:
            raw = subprocess.run(
                ["proxy-run", "--", "curl", "-s", "-m", "20", api],
                capture_output=True, text=True, timeout=60).stdout
        except Exception:
            return []
    try:
        pages = list(json.loads(raw).get("query", {}).get("pages", {}).values())
    except Exception:
        return []
    out = []
    for p in pages:
        ii = (p.get("imageinfo") or [{}])[0]
        if not ii.get("url"):
            continue
        meta = ii.get("extmetadata") or {}
        out.append({
            "title": (p.get("title") or "")[5:85],
            "url": ii.get("thumburl") or ii["url"],
            "page": ii.get("descriptionurl", ""),
            "author": re.sub(r"<[^>]+>", "",
                             (meta.get("Artist", {}) or {}).get("value", ""))[:60],
            "license": (meta.get("LicenseShortName", {}) or {}).get("value", ""),
            "source": "commons",
        })
    return out[:count]


def download(url: str, title: str = "", query: str = "") -> dict:
    """下载图片入库；同 URL 不重复下载。返回索引记录。"""
    ROOT.mkdir(parents=True, exist_ok=True)
    h = hashlib.sha1(url.encode()).hexdigest()[:16]
    for line in (INDEX.read_text("utf-8").splitlines()
                 if INDEX.exists() else []):
        rec = json.loads(line)
        if rec["url"] == url:
            return rec
    try:
        data = _get(url, binary=True)
    except Exception:
        data = b""
    if not data:
        # 境外站直连不通 → 经本机代理取（比浏览器兜底快）
        try:
            r = subprocess.run(["proxy-run", "--", "curl", "-sL", "-m", "25",
                                "-A", UA, url], capture_output=True, timeout=60)
            data = r.stdout
        except Exception:
            data = b""
    if not data:
        data = download_browser(url)   # 再不行 → 浏览器页面上下文 fetch
    if not data:
        raise RuntimeError(f"下载失败（HTTP/代理/浏览器均未取到）: {url}")
    ext = ".jpg"
    m = re.search(r"\.(png|jpe?g|gif|webp|svg)", urllib.parse.urlparse(url).path.lower())
    if m:
        ext = "." + m.group(1).replace("jpeg", "jpg")
    path = ROOT / f"{h}{ext}"
    path.write_bytes(data)
    rec = {"id": h, "path": str(path), "title": title, "query": query,
           "url": url, "bytes": len(data)}
    with INDEX.open("a", encoding="utf-8") as fp:
        fp.write(json.dumps(rec, ensure_ascii=False) + "\n")
    return rec


def main() -> None:
    ap = argparse.ArgumentParser("webimg")
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("search")
    s.add_argument("query")
    s.add_argument("--count", type=int, default=6)
    s.add_argument("--source", default="bing", choices=["bing", "commons", "all"])
    d = sub.add_parser("download")
    d.add_argument("url")
    d.add_argument("--title", default="")
    d.add_argument("--query", default="")
    args = ap.parse_args()

    if args.cmd == "search":
        fn = {"bing": search_bing, "commons": search_commons}.get(args.source)
        res = (fn(args.query, args.count) if fn
               else search_bing(args.query, args.count) +
                    search_commons(args.query, args.count))
        for r in res:
            print(json.dumps(r, ensure_ascii=False))
    else:
        print(json.dumps(download(args.url, args.title, args.query),
                         ensure_ascii=False))


if __name__ == "__main__":
    main()
