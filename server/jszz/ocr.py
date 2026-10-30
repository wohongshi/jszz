"""OCR：公式位图 → 文字（rapidocr-onnxruntime）。

学科网预览 SVG 里公式/图形是位图，文字层提不到——本模块补这一环：
  * 透明底合成白底 + 小图 2x 放大（rapidocr 对透明底/小图识别差，实测必做）
  * 水印图（学科网/组卷网 logo）OCR 后识别为水印即丢弃
  * 按 SVG 版面位置（parse_svg + _flow 行归并）把公式文字插回正确行内

依赖 rapidocr-onnxruntime（纯 pip，含 onnxruntime）；缺依赖时降级返回空，
不阻塞主链路。
"""
from __future__ import annotations

import io
import re
from pathlib import Path

_ENGINE = None
_MISSING = False

WATERMARK_RE = re.compile(
    r"学科网|组卷网|zxxk|zujuan|www\.|\.com|\.cn|中考|高考资源网", re.I)


def _engine():
    global _ENGINE, _MISSING
    if _ENGINE is None and not _MISSING:
        try:
            from rapidocr_onnxruntime import RapidOCR
            _ENGINE = RapidOCR()
        except Exception:
            _MISSING = True
    return _ENGINE


def available() -> bool:
    return _engine() is not None


def _prep(im, scale: int = 2):
    """RGBA 合成白底；小图放大 scale 倍（公式位图多为透明底细线小图）；超大图缩到 1600px 内。

    坑 1：P 模式（PNG-8/GIF）的透明区直接 convert('RGB') 会变黑底——OCR 全瞎。
    坑 2：超大图不缩会把 ONNX 推理拖到几十秒（发热大户），必须先降采样。
    """
    from PIL import Image
    if im.mode == "P":
        im = im.convert("RGBA")        # 带透明的调色板图先展开，保住 alpha
    if im.mode in ("RGBA", "LA"):
        bg = Image.new("RGB", im.size, "white")
        bg.paste(im, mask=im.split()[-1])
        im = bg
    elif im.mode != "RGB":
        im = im.convert("RGB")
    if max(im.size) > 1600:
        r = 1600 / max(im.size)
        im = im.resize((max(1, int(im.width * r)), max(1, int(im.height * r))), 1)
    if max(im.size) < 260:
        im = im.resize((im.width * scale, im.height * scale), 1)  # LANCZOS
    return im


def ocr_bytes(data: bytes, *, use_cache: bool = True) -> str:
    """图片字节 → 文字（水印/空图返回空串）。

    结果按 sha1 落盘缓存（~/.jszz/cache/ocr/）——ONNX 推理发热大，
    同一张图只推一次（thermal-gate 友好）。
    """
    import hashlib
    key = hashlib.sha1(data).hexdigest()
    cache_dir = Path.home() / ".jszz" / "cache" / "ocr"
    cache_file = cache_dir / f"{key}.txt"
    if use_cache and cache_file.exists():
        return cache_file.read_text("utf-8")
    eng = _engine()
    if eng is None:
        return ""
    try:
        from PIL import Image
        im = _prep(Image.open(io.BytesIO(data)))
    except Exception:
        return ""
    if im.width < 12 or im.height < 12:
        return ""
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    try:
        res, _ = eng(buf.getvalue())
    except Exception:
        try:
            res, _ = eng(im)
        except Exception:
            return ""
    text = " ".join(t for _, t, _ in (res or [])).strip()
    if text:
        # 水印过滤：图里常是「公式/图形 + 水印 logo」混排——剔掉水印片段保住正文，
        # 只有整段都是水印才丢（此前 search 命中就整段清空，公式全陪葬——误杀）
        kept = [seg for seg in re.split(r"\s+", text) if seg and not WATERMARK_RE.search(seg)]
        text = " ".join(kept).strip()
    try:
        cache_dir.mkdir(parents=True, exist_ok=True)
        cache_file.write_text(text, "utf-8")
    except Exception:
        pass
    return text


def ocr_file(path: str | Path) -> str:
    try:
        return ocr_bytes(Path(path).read_bytes())
    except Exception:
        return ""


def ocr_page_with_formulas(page) -> list[dict]:
    """一页（svg2docx.parse_svg 的 Page）→ 排好序的行列表。

    每行 {y, text}；行内小图 OCR 后按 x 插进文字流（公式位图补回文字层）。
    大图（插图/几何图，行高超限的）不 OCR 排版，返回 {y, text, image: True} 标记。
    """
    from .svg2docx import _flow, INLINE_MAX_H
    bands = _flow(page)
    out = []
    for b in bands:
        parts = []
        for kind, obj in b["items"]:
            if kind == "text":
                parts.append((obj.x, obj.text, False))
            else:
                is_inline = obj.h <= INLINE_MAX_H
                txt = ocr_bytes(obj.data) if is_inline else ""
                parts.append((obj.x, txt or "", not is_inline))
        parts.sort(key=lambda p: p[0])
        line = ""
        has_big_img = False
        for _, txt, big in parts:
            if big:
                has_big_img = True
                if not line:
                    line = "[图]"
                continue
            if txt:
                line += txt
        if line or has_big_img:
            out.append({"y": b["y"], "text": line.strip(), "image": has_big_img})
    return out


def page_text_with_formulas(page) -> str:
    """一页 → 纯文本（公式位图的 OCR 文字插回原行位置）。"""
    lines = ocr_page_with_formulas(page)
    return "\n".join(x["text"] for x in lines if x["text"])
