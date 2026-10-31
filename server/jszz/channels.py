"""WeChat Channels (微信视频号) client — 有限能力 + 诚实边界。

**平台硬边界（必须说清）**：微信视频号「链接 → 视频」无法纯服务端完成。
- `channels.weixin.qq.com` 无微信客户端登录态直接返回 `errCode:10012 非法请求`。
- yt-dlp 无 finder/wechat extractor（0 个）。
- 所有 GitHub 开源方案（wx_video_download、wx_channels_download）原理都是
  **本地代理 + 根证书拦截微信 PC 客户端视频流 → 拿真实 URL + 解密 key → 下载解密**——
  那是桌面端机制，无头服务器上没有微信客户端可拦。

**本模块提供的能力**（借鉴 wx_video_download 的 `download --url --key` 模式）：
  1. `download(url, key=...)`：给定桌面助手抓到的**视频直链 + 解密 key**，服务端下载 + AES 解密。
  2. `decrypt_file(path, key)`：对已下载的加密视频就地解密。
  3. `status()`：如实报告视频号自动下载不可用、给出桌面助手路径。

即：URL + key 由老师在装了 wx_video_download 的 Windows/Mac 微信 PC 端抓到，
jszz 负责服务端批量下载 + 解密落盘当课堂素材。抓取环节必须桌面端，无法绕过。
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
_CHANNELS_HOST = "channels.weixin.qq.com"


def is_channels(url: str) -> bool:
    return bool(re.search(r"channels\.weixin\.qq\.com|finder|视频号", url or ""))


class ChannelsClient:
    """视频号：下载+解密（URL+key 由桌面助手提供），不假装能自动解析链接。"""

    def __init__(self, timeout: int = 30):
        self.s = requests.Session()
        self.s.headers.update({"User-Agent": UA, "Referer": f"https://{_CHANNELS_HOST}/"})
        self.timeout = timeout

    # ---------- 探测 / 诚实状态 ----------
    def status(self) -> dict[str, Any]:
        """如实报告视频号能力边界（不假装能自动下载）。"""
        return {
            "ok": True,
            "auto_download": False,
            "reason": ("微信视频号需微信客户端登录态（服务端无客户端 → errCode 10012 非法请求），"
                       "yt-dlp 无 finder extractor；开源方案靠本地代理拦截微信 PC 客户端视频流解密，"
                       "属桌面端机制，无头服务器无法独立完成。"),
            "supported": "给定视频直链 + 解密 key（桌面助手 wx_video_download 抓取）→ 服务端下载 + 解密",
            "desktop_helper": "wx_video_download / wx_channels_download（GitHub 开源，装在微信 PC 端同机）",
            "usage": "channels_download(url=<视频直链>, key=<解密key>) 或 decrypt_file(path, key)",
        }

    def probe(self, url: str) -> dict[str, Any]:
        """探一条视频号链接：是否可直连、是否需登录态。用于告诉用户抓取前置条件。"""
        if not is_channels(url):
            return {"ok": False, "error": "不是视频号链接"}
        try:
            r = self.s.get(url, timeout=self.timeout)
            text = r.text[:400]
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": f"请求失败: {e}"}
        need_login = ("非法请求" in text) or ("10012" in text) or ("登录" in text and "finder" not in text)
        return {"ok": True, "status_code": r.status_code, "need_login": need_login,
                "snippet": text[:120],
                "hint": ("需微信客户端登录态——请在装了 wx_video_download 的微信 PC 端抓 "
                         "视频直链 + 解密 key，再用 channels_download 下载") if need_login
                        else "页面可直连，但仍需视频直链 + 解密 key（客户端抓取）"}

    # ---------- 下载 + 解密（URL+key 由桌面助手提供）----------
    def download(self, url: str, out_dir: str | Path, key: str = "",
                 filename: str = "") -> dict[str, Any]:
        """下载视频直链到 out_dir；给了 key 就 AES 解密（视频号视频流是 AES 加密的）。

        url/key 由桌面助手（wx_video_download）在微信 PC 端抓取——本方法只做服务端下载+解密。
        """
        if not url:
            return {"ok": False, "error": "缺视频直链 url（桌面助手抓取）"}
        out_dir = Path(out_dir)
        out_dir.mkdir(parents=True, exist_ok=True)
        name = filename or "channels_video"
        name = re.sub(r"[\\/:*?\"<>|\s]+", "_", name)[:80]
        raw = out_dir / f"{name}.bin"
        try:
            with self.s.get(url, stream=True, timeout=60) as r:
                r.raise_for_status()
                with open(raw, "wb") as f:
                    for part in r.iter_content(1 << 16):
                        f.write(part)
        except Exception as e:  # noqa: BLE001
            return {"ok": False, "error": f"下载失败: {e}"}
        if raw.stat().st_size == 0:
            return {"ok": False, "error": "下载得到空文件"}
        if not key:
            # 无 key 就留原始字节（可能已明文，也可能未解密）
            out = raw.with_suffix(".mp4")
            raw.rename(out)
            return {"ok": True, "file": str(out), "size": out.stat().st_size,
                    "decrypted": False,
                    "note": "未提供解密 key，已存原始字节；若无法播放请用 decrypt_file 补解密"}
        dec = self.decrypt_file(raw, key, dest=raw.with_suffix(".mp4"))
        if dec.get("ok"):
            raw.unlink(missing_ok=True)
            return {"ok": True, "file": dec["file"], "size": dec["size"], "decrypted": True}
        # 解密失败也保留原始下载，别丢
        out = raw.with_suffix(".mp4")
        raw.rename(out)
        return {"ok": True, "file": str(out), "size": out.stat().st_size,
                "decrypted": False, "note": f"解密失败已存原始字节: {dec.get('error')}"}

    def decrypt_file(self, path: str | Path, key: str,
                     dest: str | Path | None = None) -> dict[str, Any]:
        """对已下载的加密视频 AES-128 解密（视频号视频流加密）。

        key 是 16/24/32 字节（hex 或原文）。视频号常见 AES-128-ECB/CBC 分块加密，
        此处按 ECB 优先、CBC 兜底两种都试，产出能播的即为准（诚实：不同版本加密略有差异，
        解不出会明确报错而不是出坏文件）。"""
        path = Path(path)
        if not path.exists():
            return {"ok": False, "error": f"文件不存在: {path}"}
        dest = Path(dest) if dest else path.with_suffix(".mp4")
        raw = path.read_bytes()
        k = self._parse_key(key)
        if not k:
            return {"ok": False, "error": "解密 key 无法解析（需 16/24/32 字节 hex 或原文）"}
        out = self._aes_decrypt(raw, k)
        if out is None:
            return {"ok": False, "error": "AES 解密失败（key 或加密模式不匹配）"}
        dest.write_bytes(out)
        return {"ok": True, "file": str(dest), "size": dest.stat().st_size}

    @staticmethod
    def _parse_key(key: str) -> bytes:
        key = (key or "").strip()
        if not key:
            return b""
        # hex（32/48/64 hex chars）
        if re.fullmatch(r"[0-9a-fA-F]{32}|[0-9a-fA-F]{48}|[0-9a-fA-F]{64}", key):
            return bytes.fromhex(key)
        kb = key.encode()
        if len(kb) in (16, 24, 32):
            return kb
        return b""

    @staticmethod
    def _aes_decrypt(raw: bytes, key: bytes) -> bytes | None:
        try:
            from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
        except Exception:  # noqa: BLE001
            return None
        # 优先 ECB（视频号常见），再 CBC（iv 取前 16 字节或全零）
        for mode in ("ecb", "cbc_zero", "cbc_head"):
            try:
                if mode == "ecb":
                    m = modes.ECB()
                elif mode == "cbc_zero":
                    m = modes.CBC(b"\x00" * 16)
                else:
                    m = modes.CBC(raw[:16])
                cipher = Cipher(algorithms.AES(key), m)
                d = cipher.decryptor()
                body = raw if mode != "cbc_head" else raw[16:]
                out = d.update(body) + d.finalize()
                # 简单判据：去掉 PKCS7 填充后若以 mp4/moov/ftyp 常见头开头即认为成功
                unpadded = out[:-out[-1]] if out and 0 < out[-1] <= 16 else out
                if unpadded[:4] in (b"\x00\x00\x00\x18", b"\x00\x00\x00\x1c",
                                    b"\x00\x00\x00 ", b"ftyp") or b"ftyp" in unpadded[:32]:
                    return unpadded
                if mode == "ecb":
                    ecb_out = out  # 至少 ECB 结果先记下，最后兜底返回
            except Exception:  # noqa: BLE001
                continue
        # 都没命中 mp4 头：返回 ECB 原始解密结果（可能 key 对但头判断没中）
        try:
            cipher = Cipher(algorithms.AES(key), modes.ECB())
            d = cipher.decryptor()
            return d.update(raw) + d.finalize()
        except Exception:  # noqa: BLE001
            return None
