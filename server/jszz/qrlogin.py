"""扫码登录：B 站二维码（官方 API，无需浏览器）。

流程（与网页端一致）：
  1. GET  /x/passport-login/web/qrcode/generate  -> {url, qrcode_key}
  2. 把 url 画成二维码，用户用 B 站 App 扫
  3. 轮询 /x/passport-login/web/qrcode/poll?qrcode_key=…
       code 86101 未扫码 / 86090 已扫码未确认 / 86038 二维码失效 / 0 成功
  4. 成功后返回的 data.url 里带 SESSDATA / bili_jct / DedeUserID，存成 cookie 串

学科网没有公开的扫码接口，WebUI 里仍用粘贴 cookie。
"""
from __future__ import annotations

import base64
import io
import re
import time
import urllib.parse
from typing import Any

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

GENERATE = "https://passport.bilibili.com/x/passport-login/web/qrcode/generate"
POLL = "https://passport.bilibili.com/x/passport-login/web/qrcode/poll"

STATUS_TEXT = {
    0: "登录成功",
    86101: "等待扫码",
    86090: "已扫码，请在手机上确认",
    86038: "二维码已失效",
}

# 进程内保存 key -> 创建时间，用于判断过期（2 分钟）
_PENDING: dict[str, float] = {}


def _qr_png_data_url(data: str, scale: int = 6) -> str:
    """把二维码画成 PNG 的 data URL（前端直接 <img src>）。"""
    import segno

    qr = segno.make(data, error="m")
    buf = io.BytesIO()
    qr.save(buf, kind="png", scale=scale, border=2, dark="#0F2B55")
    b64 = base64.b64encode(buf.getvalue()).decode()
    return "data:image/png;base64," + b64


def start(site: str = "bilibili") -> dict[str, Any]:
    if site != "bilibili":
        return {"ok": False, "error": "该站点没有扫码接口，请粘贴 cookie"}
    try:
        r = requests.get(GENERATE, headers={"User-Agent": UA, "Referer": "https://www.bilibili.com/"},
                         timeout=15)
        data = r.json()
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}
    if data.get("code") != 0:
        return {"ok": False, "error": data.get("message") or "生成二维码失败"}
    key = data["data"]["qrcode_key"]
    url = data["data"]["url"]
    _PENDING[key] = time.time()
    try:
        img = _qr_png_data_url(url)
    except Exception as e:
        return {"ok": False, "error": f"二维码渲染失败（需要 segno）：{e}"}
    return {"ok": True, "key": key, "url": url, "qr": img,
            "expires_in": 180, "hint": "用 B 站手机客户端扫码"}


def _cookie_from_url(url: str) -> str:
    """从回调 url 里取出 SESSDATA / bili_jct / DedeUserID / sid。"""
    q = urllib.parse.urlparse(url).query
    params = urllib.parse.parse_qs(q)
    keep = ("SESSDATA", "bili_jct", "DedeUserID", "DedeUserID__ckMd5", "sid")
    parts = []
    for k in keep:
        if params.get(k):
            parts.append(f"{k}={params[k][0]}")
    return "; ".join(parts)


def poll(key: str, site: str = "bilibili") -> dict[str, Any]:
    if site != "bilibili":
        return {"ok": False, "error": "unsupported"}
    if not re.fullmatch(r"[0-9a-zA-Z]{16,64}", key or ""):
        return {"ok": False, "error": "无效的 qrcode_key"}
    try:
        r = requests.get(POLL, params={"qrcode_key": key},
                         headers={"User-Agent": UA, "Referer": "https://www.bilibili.com/"},
                         timeout=15)
        data = r.json()
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}
    if data.get("code") != 0:
        return {"ok": False, "error": data.get("message") or "轮询失败"}
    d = data.get("data") or {}
    code = d.get("code")
    state = {0: "ok", 86101: "waiting", 86090: "scanned", 86038: "expired"}.get(code, "waiting")
    out: dict[str, Any] = {"ok": True, "state": state,
                           "message": STATUS_TEXT.get(code, d.get("message") or "")}
    if state == "ok":
        cookie = _cookie_from_url(d.get("url") or "")
        out["cookie"] = cookie
        out["has_sessdata"] = "SESSDATA=" in cookie
        _PENDING.pop(key, None)
    elif state == "expired":
        _PENDING.pop(key, None)
    return out


def verify(cookie: str) -> dict[str, Any]:
    """用 cookie 调 nav 接口确认真的登录了（扫码后自检）。"""
    if not cookie:
        return {"ok": False, "reason": "empty cookie"}
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Referer": "https://www.bilibili.com/"})
    for part in cookie.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            s.cookies.set(k.strip(), v.strip(), domain=".bilibili.com")
    try:
        j = s.get("https://api.bilibili.com/x/web-interface/nav", timeout=15).json()
    except Exception as e:
        return {"ok": False, "reason": f"{type(e).__name__}: {e}"}
    d = j.get("data") or {}
    return {"ok": bool(d.get("isLogin")), "uname": d.get("uname"),
            "vip": (d.get("vipStatus") == 1), "level": d.get("level_info", {}).get("current_level")}
