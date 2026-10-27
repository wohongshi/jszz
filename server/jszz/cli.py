"""命令行入口：jszz serve / web / login / search / xuean / paper / status。"""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from pathlib import Path

from . import config
from .bilibili import BilibiliClient
from .docx_render import render_paper, render_xuean
from .pptx_render import build_paper_deck, build_xuean_deck
from .zxxk import ZxxkClient


def cmd_serve(args):
    from . import server
    if args.http:
        server.mcp.run(transport="streamable-http", host="127.0.0.1", port=args.port)
    else:
        server.mcp.run(transport="stdio")


def cmd_web(args):
    from .webui import serve
    serve(port=args.port, host=args.host)


def cmd_login(args):
    if not args.site:
        print("用法：jszz login bilibili|zxxk '<cookie 字符串>'")
        print("  bilibili: 浏览器登录 bilibili.com → F12 → Application → Cookies，复制 SESSDATA")
        print("  zxxk:     浏览器登录 zxxk.com → F12 → Network → 任一请求的 Cookie 头整段复制")
        return
    if not args.cookie:
        p = config.cookie_path(args.site)
        print(f"{args.site} 当前 cookie：{'已保存' if p.exists() else '未保存'} ({p})")
        return
    p = config.write_cookie(args.site, args.cookie)
    print(f"已保存 {args.site} cookie → {p}")


def cmd_status(_args):
    print(json.dumps(_status_raw(), ensure_ascii=False, indent=1))


def _status_raw():
    import shutil
    cfg = config.load()
    return {
        "home": str(config.HOME), "out_dir": cfg.get("out_dir"), "school": cfg.get("school"),
        "cookies": {s: bool(config.read_cookie(s)) for s in ("bilibili", "zxxk")},
        "deps": {k: bool(shutil.which(k)) for k in ("yt-dlp", "ffmpeg", "node")},
    }


def cmd_search(args):
    if args.source in ("bili", "all"):
        c = BilibiliClient(cookie=config.read_cookie("bilibili"))
        r = c.search(args.keyword, page_size=args.limit)
        print(f"\n[B站] 共 {r.get('total')} 条")
        for it in r["items"]:
            print(f"  {it['bvid']}  {it['duration']:>8}  {it['play']:>9}  {it['title'][:60]}")
    if args.source in ("zxxk", "all"):
        c = ZxxkClient(cookie_file=config.cookie_path("zxxk"))
        ck = config.read_cookie("zxxk")
        if ck:
            c.http.set_cookies(ck)
        r = c.search(args.keyword, page_size=args.limit)
        if not r.get("ok"):
            print(f"\n[学科网] 失败：{r.get('error')}")
            return
        print(f"\n[学科网] 共 {r.get('total')} 条")
        for it in r["items"]:
            print(f"  {it['document_id']:>10}  {it['type'] or '-':<4} {it['subject'] or '-':<4}"
                  f" {it['title'][:52]}  {it['price']}")


def cmd_xuean(args):
    md = Path(args.file).read_text("utf-8")
    base = args.name or Path(args.file).stem
    out = config.out_dir("xuean")
    files = []
    if not args.no_docx:
        files.append(render_xuean(md, out / f"{base}.docx", school=config.load().get("school", "")))
    if not args.no_pptx:
        files.append(build_xuean_deck(md, out / f"{base}.pptx", theme=args.theme,
                                      school=config.load().get("school", "")))
    for f in files:
        print(f)


def cmd_paper(args):
    md = Path(args.file).read_text("utf-8")
    base = args.name or Path(args.file).stem
    out = config.out_dir("paper")
    files = []
    if not args.no_docx:
        files.append(render_paper(md, out / f"{base}.docx", school=config.load().get("school", ""),
                                  answer_page=not args.no_answers))
    if not args.no_pptx:
        files.append(build_paper_deck(md, out / f"{base}.pptx", theme=args.theme,
                                      school=config.load().get("school", "")))
    for f in files:
        print(f)


def cmd_check(args):
    from .qa import check
    r = check(md=Path(args.md).read_text("utf-8") if args.md else None,
              docx=args.docx, pptx=args.pptx)
    bad = 0
    for kind, rep in r.items():
        if not rep:
            continue
        items = rep.get("warnings") or rep.get("issues") or []
        print(f"[{kind}] {'OK' if rep.get('ok') else 'FOUND ' + str(len(items))}")
        for it in items:
            print("   ", it if isinstance(it, str) else
                  f"{it.get('page', '')} {it.get('type')}: {it.get('detail')}")
        bad += 0 if rep.get("ok") else 1
    raise SystemExit(1 if bad else 0)


def cmd_set(args):
    cfg = config.load()
    cfg[args.key] = args.value
    config.save(cfg)
    print(json.dumps(cfg, ensure_ascii=False, indent=1))


def main(argv=None):
    p = argparse.ArgumentParser("jszz", description="备课工具：B站素材 / 学科网题目 / 学案与试卷生成")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("serve", help="启动 MCP 服务（stdio，或 --http）")
    s.add_argument("--http", action="store_true")
    s.add_argument("--port", type=int, default=8765)
    s.set_defaults(func=cmd_serve)

    s = sub.add_parser("web", help="启动本地控制台（登录/依赖/产物）")
    s.add_argument("--port", type=int, default=None)
    s.add_argument("--host", default="127.0.0.1")
    s.set_defaults(func=cmd_web)

    s = sub.add_parser("login", help="保存 B站 / 学科网 cookie")
    s.add_argument("site", nargs="?", choices=["bilibili", "zxxk"])
    s.add_argument("cookie", nargs="?")
    s.set_defaults(func=cmd_login)

    s = sub.add_parser("status", help="查看状态")
    s.set_defaults(func=cmd_status)

    s = sub.add_parser("search", help="命令行搜索")
    s.add_argument("keyword")
    s.add_argument("--source", choices=["bili", "zxxk", "all"], default="all")
    s.add_argument("--limit", type=int, default=8)
    s.set_defaults(func=cmd_search)

    s = sub.add_parser("xuean", help="学案 Markdown → docx + pptx")
    s.add_argument("file")
    s.add_argument("--name")
    s.add_argument("--theme", default="blue")
    s.add_argument("--no-docx", action="store_true")
    s.add_argument("--no-pptx", action="store_true")
    s.set_defaults(func=cmd_xuean)

    s = sub.add_parser("paper", help="试卷 Markdown → docx + 讲评 pptx")
    s.add_argument("file")
    s.add_argument("--name")
    s.add_argument("--theme", default="ink")
    s.add_argument("--no-docx", action="store_true")
    s.add_argument("--no-pptx", action="store_true")
    s.add_argument("--no-answers", action="store_true")
    s.set_defaults(func=cmd_paper)

    s = sub.add_parser("check", help="体检：Markdown / docx / pptx")
    s.add_argument("--md")
    s.add_argument("--docx")
    s.add_argument("--pptx")
    s.set_defaults(func=cmd_check)

    s = sub.add_parser("set", help="改配置，如 jszz set school 胶州四中")
    s.add_argument("key")
    s.add_argument("value")
    s.set_defaults(func=cmd_set)

    args = p.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main()
