"""学科网 SSO 扫码登录（微信扫码，流程复刻 sso.zxxk.com 页面 JS）。

逆向自 sso.zxxk.com 的 scan-code-login-v2.js / scan-code-processor-v1.js：
  1. GET /user/qrcode/qrcode-login/v2/content?token=&service=…
     -> data.content（二维码内容 URL，含 key=）、data.queryToken
  2. 把 content 画成二维码，微信扫码打开 -> 手机上点「同意学科网登录」
  3. 状态走 WebSocket：wss://sso.zxxk.com/user/qrcode/socketServer
       ?qrcodeKey=<key>&type=qrcode-login
     推送 {"data":"<状态>"}：LOGINED 成功 / LOGINING 已扫待确认 /
     QUIT_LOGIN 用户取消 / EXPIRED 过期；另有 valid/invalid 心跳应答
  4. LOGINED 后 GET /user/wechat-applet/dologin?service=…（session 绑定，
     无 key 参数）——响应 Set-Cookie（ssoid 等），再访问 service 域名
     落 www.zxxk.com 侧 cookie（zxxk_member_token 等）

注：HTTP 轮询 /user/wechat-applet/logincheck 参数语义不明（实测 1013/1019
恒报错），真实页面的主通道就是 WebSocket——本模块照此实现。
"""
from __future__ import annotations

import io
import json
import threading
import time
from typing import Any

from .waf import WafSession

SERVICE = "https://www.zxxk.com/"
SSO = "https://sso.zxxk.com"

STATUS_TEXT = {
    "LOGINING": "已扫码，请在手机上确认",
    "LOGINED": "登录成功",
    "QUIT_LOGIN": "手机端取消了登录",
    "EXPIRED": "二维码已失效",
    "ACCOUNT_BIND": "需要绑定账号，请改用粘贴 cookie",
}

# key -> 会话状态（WafSession + WS 连接），2 分钟过期
_PENDING: dict[str, dict[str, Any]] = {}
_LOCK = threading.Lock()


def _qr_png_data_url(data: str, scale: int = 6) -> str:
    import base64
    import segno

    qr = segno.make(data, error="m")
    buf = io.BytesIO()
    qr.save(buf, kind="png", scale=scale, border=2, dark="#005BAC")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def start(service: str = SERVICE) -> dict[str, Any]:
    """取二维码 + 挂上状态通道（WebSocket）。"""
    import websockets.sync.client as wsc

    sess = WafSession()
    try:
        sess.get(SSO + "/login", params={"dialogTheme": "true", "service": service})
        r = sess.get(SSO + "/user/qrcode/qrcode-login/v2/content",
                     params={"token": "", "service": service},
                     headers={"Referer": SSO + "/login",
                              "X-Requested-With": "XMLHttpRequest"})
        data = (r.json() or {}).get("data") or {}
    except Exception as e:
        return {"ok": False, "error": f"{type(e).__name__}: {e}"}
    content = data.get("content") or ""
    qtok = (data.get("queryToken") or {}).get("token") or ""
    if not content or "key=" not in content:
        return {"ok": False, "error": "未取到二维码内容（可重试一次）"}
    key = content.split("key=")[1].split("&")[0]

    conn = None
    try:
        cookie_str = "; ".join(f"{k}={v}" for k, v in sess.s.cookies.get_dict().items())
        conn = wsc.connect(
            f"wss://sso.zxxk.com/user/qrcode/socketServer?qrcodeKey={key}&type=qrcode-login",
            additional_headers={"Cookie": cookie_str, "Origin": SSO,
                                "User-Agent": "Mozilla/5.0"},
            open_timeout=15)
        try:  # 首条是会话注册回执（32 位 hex），收掉
            conn.recv(timeout=5)
        except Exception:
            pass
    except Exception:
        conn = None  # WS 挂了就退化为只靠 dologin 试探（仍可用，状态提示少）

    with _LOCK:
        _PENDING[key] = {"sess": sess, "conn": conn, "qtoken": qtok,
                         "service": service, "created": time.time(),
                         "state": "waiting"}
    try:
        img = _qr_png_data_url(content)
    except Exception as e:
        return {"ok": False, "error": f"二维码渲染失败（需要 segno）：{e}"}
    return {"ok": True, "key": key, "url": content, "qr": img,
            "expires_in": 180,
            "hint": "用微信扫码，在手机上点「同意学科网登录」"}


def _drain(key: str) -> tuple[str, str]:
    """收 WS 推送，返回 (状态, 文案)。状态：waiting/scanned/ok/expired。"""
    with _LOCK:
        ent = _PENDING.get(key)
    if not ent:
        return "expired", "二维码已失效"
    if time.time() - ent["created"] > 180:
        with _LOCK:
            _PENDING.pop(key, None)
        _close(ent)
        return "expired", "二维码已失效"

    conn = ent.get("conn")
    if conn is not None:
        while True:
            try:
                msg = conn.recv(timeout=0.15)
            except Exception:
                break  # 超时 = 没有新推送
            raw = msg if isinstance(msg, str) else ""
            if raw in ("valid", "invalid"):
                if raw == "invalid":
                    ent["state"] = "expired"
                continue
            try:
                st = (json.loads(raw) or {}).get("data", "")
            except Exception:
                continue
            if st == "LOGINED":
                ent["state"] = "scanned"  # _finish 会再判 ok
                return "scanned", STATUS_TEXT["LOGINED"]
            if st in STATUS_TEXT:
                ent["state"] = "scanned" if st == "LOGINING" else "waiting"
                if st in ("QUIT_LOGIN", "EXPIRED"):
                    ent["state"] = "expired"
                return {"LOGINING": "scanned", "LOGINED": "scanned",
                        "QUIT_LOGIN": "expired", "EXPIRED": "expired",
                        "ACCOUNT_BIND": "waiting"}[st], STATUS_TEXT[st]
    return ent["state"], "等待扫码（3 分钟内有效）"


def _close(ent: dict[str, Any]) -> None:
    conn = ent.get("conn")
    if conn is not None:
        try:
            conn.close()
        except Exception:
            pass


def _finish(key: str) -> dict[str, Any]:
    """扫码确认后：dologin -> service 落 cookie -> 导出 cookie 串。"""
    with _LOCK:
        ent = _PENDING.get(key)
    if not ent:
        return {"ok": False, "state": "expired", "message": "二维码已失效"}
    sess: WafSession = ent["sess"]
    service = ent["service"]
    H = {"Referer": SSO + "/login", "X-Requested-With": "XMLHttpRequest"}
    try:
        r = sess.get(SSO + "/user/wechat-applet/dologin",
                     params={"service": service}, headers=H)
        r.encoding = "utf-8"
        j = r.json()
    except Exception as e:
        return {"ok": False, "state": "waiting", "message": f"{type(e).__name__}: {e}"}
    if j.get("status") != 200:
        msg = j.get("message") or "登录未完成"
        # 「同意」弹窗类消息：手机端确认后重试一次即可，这里如实上报
        return {"ok": False, "state": "waiting", "message": msg}
    try:  # 访问 service 域，把 www.zxxk.com 侧 cookie 落进 jar
        sess.get(service, headers={"Referer": SSO + "/login"})
    except Exception:
        pass
    parts = [f"{c.name}={c.value}" for c in sess.s.cookies
             if "zxxk" in (c.domain or "") or "xkw" in (c.domain or "")]
    cookie = "; ".join(parts)
    with _LOCK:
        _PENDING.pop(key, None)
    _close(ent)
    return {"ok": True, "state": "ok", "cookie": cookie,
            "has_token": "zxxk_member_token" in cookie,
            "message": "登录成功"}


def poll(key: str) -> dict[str, Any]:
    state, text = _drain(key)
    if state == "scanned" and text == STATUS_TEXT["LOGINED"]:
        fin = _finish(key)
        if fin.get("ok"):
            return fin
        return {"ok": True, "state": "scanned", "message": fin.get("message") or text}
    return {"ok": True, "state": state, "message": text}


def verify(cookie: str) -> dict[str, Any]:
    """检查 cookie 是否含登录凭证并探一次首页。"""
    if not cookie:
        return {"ok": False, "reason": "empty cookie"}
    import requests
    has = "zxxk_member_token" in cookie
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0", "Referer": SERVICE})
    for part in cookie.split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            s.cookies.set(k.strip(), v.strip(), domain=".zxxk.com")
    try:
        r = s.get(SERVICE, timeout=15)
        ok = r.status_code == 200 and ("退出" in r.text or "个人中心" in r.text or has)
    except Exception:
        ok = has
    return {"ok": bool(ok), "has_token": has}
