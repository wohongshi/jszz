#!/usr/bin/env python3
"""活码验证：生成学科网扫码登录二维码并后台轮询，扫成后落 cookie。

用法：python3 live_qr.py  （二维码 PNG 落 /tmp/zxxk-live-qr.png，日志 /tmp/zxxk-live.log）
"""
import sys, time, pathlib

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parent.parent / "server"))

from jszz import qrlogin, zxxk_qr  # noqa: E402

LOG = pathlib.Path("/tmp/zxxk-live.log")
PNG = pathlib.Path("/tmp/zxxk-live-qr.png")


def log(msg):
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    with LOG.open("a") as f:
        f.write(line + "\n")


r = qrlogin.start("zxxk")
if not r.get("ok"):
    log(f"start FAIL: {r.get('error')}")
    sys.exit(1)
key = r["key"]
log(f"start ok key={key[:10]} url={r['url'][:70]}")

# 把 data URL 转成 PNG 文件（微信里可长按识别）
import base64  # noqa: E402
png = base64.b64decode(r["qr"].split(",", 1)[1])
PNG.write_bytes(png)
log(f"QR saved {PNG} ({len(png)} bytes)")

deadline = time.time() + 175
while time.time() < deadline:
    p = zxxk_qr.poll(key)
    log(f"poll -> {p.get('state')} | {p.get('message')}")
    if p.get("state") == "ok":
        cookie = p.get("cookie") or ""
        log(f"LOGIN OK cookie_len={len(cookie)} has_token={p.get('has_token')}")
        try:
            from jszz import config
            config.write_cookie("zxxk", cookie)
            log(f"cookie saved -> {config.cookie_path('zxxk') if hasattr(config,'cookie_path') else 'config'}")
        except Exception as e:
            log(f"cookie save err: {e}")
        break
    if p.get("state") == "expired":
        log("EXPIRED")
        break
    time.sleep(2)
else:
    log("TIMEOUT 3min")
