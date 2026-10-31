"""HTTP fetcher that survives the Aliyun WAF used by zxxk.com / zujuan.xkw.com.

Two challenge families appear, sometimes chained:
  * `alicfw`  — a `check()` script driven by #parm_0/#parm_1 inputs
  * `acw_sc__v2` — a `reload()` script reading JSON out of #renderData

Both are plain JS that only needs document/location/navigator, so we run the page's
own script under Node (tools/fetch_alicfw.js) instead of shipping a browser.
"""
from __future__ import annotations

import json
import subprocess
import threading
from http.cookies import SimpleCookie
from pathlib import Path
from typing import Optional

import requests

HERE = Path(__file__).resolve().parent
SOLVER = HERE / "bin" / "fetch_alicfw.js"

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

CHALLENGE_MARKERS = ("id=\"renderData\"", "onload=\"check()\"", "aliyun_waf_aa", "parm_0")


def is_challenge(html: str) -> bool:
    if len(html) > 60000:
        return False
    return any(m in html for m in CHALLENGE_MARKERS)


def solve_challenge(html: str, origin: str, timeout: int = 20) -> list[str]:
    """Run the challenge page's own JS in Node; return Set-Cookie fragments."""
    try:
        proc = subprocess.run(
            ["node", str(SOLVER), origin],
            input=html.encode("utf-8"), capture_output=True, timeout=timeout,
        )
        data = json.loads(proc.stdout.decode("utf-8", "replace") or "{}")
        return data.get("cookies") or []
    except Exception:
        return []


class WafSession:
    """requests.Session + on-demand WAF challenge solving."""

    def __init__(self, cookie_file: Optional[Path] = None, timeout: int = 30):
        self.s = requests.Session()
        self.s.headers.update({
            "User-Agent": UA,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        })
        self.timeout = timeout
        self.cookie_file = cookie_file
        self._lock = threading.Lock()
        if cookie_file and cookie_file.exists():
            self.load(cookie_file)

    # ---- cookies -------------------------------------------------------
    def load(self, path: Path) -> None:
        try:
            data = json.loads(path.read_text("utf-8"))
        except Exception:
            return
        for host, items in data.items():
            for name, value in items.items():
                self.s.cookies.set(name, value, domain=host)

    def save(self, path: Optional[Path] = None) -> None:
        path = path or self.cookie_file
        if not path:
            return
        out: dict[str, dict[str, str]] = {}
        for c in self.s.cookies:
            out.setdefault(c.domain or "", {})[c.name] = c.value
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(out, ensure_ascii=False, indent=1), "utf-8")

    def set_cookies(self, cookie_string: str, domain: str = "") -> None:
        """Accept a raw `name=value; name2=value2` string pasted by the user.

        坑：domain=None 会让 requests 的 create_cookie 炸（NoneType.startswith）——
        空域名传空串即可，requests 自己按 host-only cookie 处理。
        """
        c = SimpleCookie()
        c.load(cookie_string)
        for k, morsel in c.items():
            self.s.cookies.set(k, morsel.value, domain=domain or "")

    def cookie_header(self, url: str) -> str:
        return "; ".join(f"{k}={v}" for k, v in self.s.cookies.get_dict(url).items())

    # ---- requests ------------------------------------------------------
    def get(self, url: str, *, params=None, headers=None, allow_redirects=True,
            tries: int = 5) -> requests.Response:
        last: Optional[requests.Response] = None
        for _ in range(tries):
            r = self.s.get(url, params=params, headers=headers,
                           timeout=self.timeout, allow_redirects=allow_redirects)
            last = r
            r.encoding = r.apparent_encoding or "utf-8"
            if not is_challenge(r.text):
                return r
            with self._lock:
                cookies = solve_challenge(r.text, url)
                if not cookies:
                    return r
                for frag in cookies:
                    head = frag.split(";")[0]
                    if "=" in head:
                        name, value = head.split("=", 1)
                        host = requests.utils.urlparse(url).hostname or ""
                        self.s.cookies.set(name.strip(), value.strip(), domain=host)
        return last  # type: ignore[return-value]

    def post(self, url: str, *, json_body=None, data=None, headers=None, tries: int = 4):
        last = None
        for _ in range(tries):
            r = self.s.post(url, json=json_body, data=data, headers=headers,
                            timeout=self.timeout)
            last = r
            r.encoding = r.apparent_encoding or "utf-8"
            if not is_challenge(r.text):
                return r
            with self._lock:
                cookies = solve_challenge(r.text, url)
                if not cookies:
                    return r
                for frag in cookies:
                    head = frag.split(";")[0]
                    if "=" in head:
                        name, value = head.split("=", 1)
                        host = requests.utils.urlparse(url).hostname or ""
                        self.s.cookies.set(name.strip(), value.strip(), domain=host)
        return last
