"""配置与本地状态：~/.jszz/"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

HOME = Path(os.environ.get("JSZZ_HOME", Path.home() / ".jszz"))
CONFIG = HOME / "config.json"
COOKIE_DIR = HOME / "cookies"
OUT_DIR = HOME / "out"
LOG_DIR = HOME / "logs"

DEFAULTS: dict[str, Any] = {
    "school": "",
    "author": "",
    "theme": "blue",
    "paper_theme": "ink",
    "out_dir": str(OUT_DIR),
    "webui_port": 8770,
    "bilibili_quality": "1080",
}


def ensure_dirs() -> None:
    for d in (HOME, COOKIE_DIR, OUT_DIR, LOG_DIR, Path(DEFAULTS["out_dir"])):
        d.mkdir(parents=True, exist_ok=True)


def load() -> dict[str, Any]:
    ensure_dirs()
    cfg = dict(DEFAULTS)
    if CONFIG.exists():
        try:
            cfg.update(json.loads(CONFIG.read_text("utf-8")))
        except Exception:
            pass
    return cfg


def save(cfg: dict[str, Any]) -> None:
    ensure_dirs()
    CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), "utf-8")


def set_value(key: str, value: Any) -> dict[str, Any]:
    cfg = load()
    cfg[key] = value
    save(cfg)
    return cfg


def cookie_path(site: str) -> Path:
    ensure_dirs()
    return COOKIE_DIR / f"{site}.txt"


def read_cookie(site: str) -> str:
    p = cookie_path(site)
    return p.read_text("utf-8").strip() if p.exists() else ""


def write_cookie(site: str, value: str) -> Path:
    p = cookie_path(site)
    p.write_text(value.strip() + "\n", "utf-8")
    os.chmod(p, 0o600)
    return p


def out_dir(sub: str = "") -> Path:
    cfg = load()
    d = Path(cfg.get("out_dir") or OUT_DIR)
    if sub:
        d = d / sub
    d.mkdir(parents=True, exist_ok=True)
    return d
