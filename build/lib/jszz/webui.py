"""本地 Web 控制台：登录态、依赖自检、产物下载。

只用标准库（http.server），不引额外依赖，方便装到任何一台教师机上。
默认只监听 127.0.0.1，凭本机访问即可，不要暴露到公网。
"""
from __future__ import annotations

import html
import json
import mimetypes
import shutil
import threading
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from . import config
from .bilibili import BilibiliClient
from .zxxk import ZxxkClient

PAGE = """<!doctype html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>jszz · 备课工具控制台</title>
<style>
:root{--ink:#15202b;--muted:#6b7785;--line:#e5e9f0;--bg:#f6f8fb;--card:#fff;
--accent:#2563eb;--ok:#16a34a;--warn:#b45309;--bad:#dc2626}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);
font:15px/1.6 -apple-system,BlinkMacSystemFont,"Segoe UI","Microsoft YaHei",sans-serif}
.wrap{max-width:900px;margin:0 auto;padding:28px 18px 60px}
header{display:flex;align-items:baseline;gap:12px;margin-bottom:6px}
h1{font-size:26px;margin:0;letter-spacing:.5px}
.sub{color:var(--muted);font-size:13px}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:18px}
@media(max-width:720px){.grid{grid-template-columns:1fr}}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px}
.card h2{font-size:15px;margin:0 0 10px;display:flex;justify-content:space-between;align-items:center}
.badge{font-size:12px;padding:2px 8px;border-radius:999px;border:1px solid var(--line)}
.badge.ok{color:var(--ok);border-color:#bbf7d0;background:#f0fdf4}
.badge.bad{color:var(--bad);border-color:#fecaca;background:#fef2f2}
.badge.warn{color:var(--warn);border-color:#fde68a;background:#fffbeb}
.badge.idle{color:var(--muted);border-color:var(--line);background:#f8fafc}
code{overflow-wrap:anywhere}
label{display:block;font-size:13px;color:var(--muted);margin:10px 0 4px}
input,textarea,select{width:100%;padding:9px 10px;border:1px solid var(--line);border-radius:8px;
font:inherit;background:#fbfcfe}
textarea{min-height:88px;resize:vertical;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:13px}
button{margin-top:10px;background:var(--accent);color:#fff;border:0;border-radius:8px;
padding:9px 16px;font:inherit;cursor:pointer}
button.ghost{background:#eef2ff;color:var(--accent)}
button:disabled{opacity:.5;cursor:default}
.msg{margin-top:10px;font-size:13px;white-space:pre-wrap}
.msg.ok{color:var(--ok)} .msg.bad{color:var(--bad)}
ul{margin:8px 0 0;padding-left:18px}
li{font-size:13px;color:var(--muted)}
code{background:#eef2f7;padding:1px 5px;border-radius:4px;font-size:12px}
a{color:var(--accent);text-decoration:none}
.hint{font-size:12px;color:var(--muted);margin-top:6px}
</style></head><body><div class="wrap">
<header><h1>jszz</h1><span class="sub">备课工具 · 本地控制台</span></header>
<div class="sub">B站素材 / 学科网题目 / 学案与试卷生成 —— 登录态与依赖都在本机，不上传任何账号信息。</div>

<div class="grid">
  <div class="card">
    <h2>B站账号 <span id="b-badge" class="badge">检查中</span></h2>
    <div class="hint">未登录也能搜索和下载 1080P；登录后可取 AI 字幕、会员清晰度。</div>
    <label>Cookie（浏览器 F12 → Application → Cookies → bilibili.com，复制 SESSDATA 即可）</label>
    <textarea id="b-cookie" placeholder="SESSDATA=xxxxx; bili_jct=xxxxx"></textarea>
    <button onclick="save('bilibili')">保存 Cookie</button>
    <button class="ghost" onclick="test('bilibili')">测试连通</button>
    <div class="msg" id="b-msg"></div>
  </div>

  <div class="card">
    <h2>学科网 / 组卷网 <span id="z-badge" class="badge">检查中</span></h2>
    <div class="hint">搜索与预览页读取免登录；下载原文件、组卷网搜题需要账号 cookie。</div>
    <label>Cookie（zxxk.com 登录后复制整段 cookie）</label>
    <textarea id="z-cookie" placeholder="zxxk_member_token=...; ..."></textarea>
    <button onclick="save('zxxk')">保存 Cookie</button>
    <button class="ghost" onclick="test('zxxk')">测试连通</button>
    <div class="msg" id="z-msg"></div>
  </div>

  <div class="card">
    <h2>出卷署名</h2>
    <label>学校（打印在页脚与课件首页）</label>
    <input id="school" placeholder="如：胶州四中">
    <label>默认主题</label>
    <select id="theme">
      <option value="blue">蓝（默认）</option><option value="teal">青绿</option>
      <option value="violet">紫</option><option value="orange">橙</option><option value="ink">墨黑</option>
    </select>
    <button onclick="saveConfig()">保存</button>
    <div class="msg" id="c-msg"></div>
  </div>

  <div class="card">
    <h2>运行环境</h2>
    <ul id="deps"></ul>
    <div class="hint">下载 B 站视频需要 <code>yt-dlp</code> 与 <code>ffmpeg</code>；过学科网 WAF 需要 <code>node</code>。</div>
  </div>
</div>

<div class="card" style="margin-top:14px">
  <h2>最近产物</h2>
  <ul id="files"><li>加载中…</li></ul>
</div>

<div class="card" style="margin-top:14px">
  <h2>接入 MCP</h2>
  <div class="hint">复制到 Claude Code / WorkBuddy / Cline 的 MCP 配置里（stdio）：</div>
  <pre id="mcpcfg" style="background:#0f172a;color:#e2e8f0;padding:12px;border-radius:10px;overflow:auto;font-size:12px"></pre>
</div>

<script>
const $ = (id) => document.getElementById(id);
async function api(path, body) {
  const r = await fetch(path, body ? {method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify(body)} : {});
  return r.json();
}
function badge(el, tone, text){ el.className = 'badge ' + tone; el.textContent = text; }
async function refresh(){
  const s = await api('/api/status');
  badge($('b-badge'), s.bilibili.logged_in?'ok':'idle', s.bilibili.logged_in?'已登录':'未登录');
  badge($('z-badge'), s.zxxk.has_cookie?'ok':'warn', s.zxxk.has_cookie?'已导入 Cookie':'免登录可用');
  $('school').value = s.config.school || '';
  $('theme').value = s.config.theme || 'blue';
  const d = s.deps;
  $('deps').innerHTML = [
    ['yt-dlp', d['yt-dlp']], ['ffmpeg', d.ffmpeg], ['node', d.node]
  ].map(([k,v]) => `<li>${k}: <span style="color:${v?'#16a34a':'#dc2626'}">${v?'可用':'缺失'}</span></li>`).join('')
   + `<li>数据目录: <code>${s.home}</code></li>`;
  const f = s.recent_files || [];
  $('files').innerHTML = f.length ? f.map(x =>
    `<li><a href="/files/${encodeURIComponent(x.rel)}" target="_blank">${x.name}</a>
     <span style="color:#94a3b8"> · ${x.size_kb} KB</span></li>`).join('') : '<li>还没有产物，去 MCP 里调一次 <code>make_xuean</code></li>';
  $('mcpcfg').textContent = JSON.stringify({
    mcpServers: { jszz: {
      command: s.python, args: ['-m', 'jszz.server'],
      env: { PYTHONPATH: s.server_path, JSZZ_HOME: s.home }
    } }
  }, null, 2);
}
async function save(site){
  const val = site === 'bilibili' ? $('b-cookie').value : $('z-cookie').value;
  const el = site === 'bilibili' ? $('b-msg') : $('z-msg');
  const r = await api('/api/cookie', {site, cookie: val});
  el.className = 'msg ' + (r.ok?'ok':'bad');
  el.textContent = r.ok ? '已保存到 ' + r.path : ('保存失败：' + r.error);
  refresh();
}
async function test(site){
  const el = site === 'bilibili' ? $('b-msg') : $('z-msg');
  el.className = 'msg'; el.textContent = '测试中…';
  const r = await api('/api/test', {site});
  el.className = 'msg ' + (r.ok?'ok':'bad');
  el.textContent = r.ok ? r.detail : ('失败：' + r.detail);
}
async function saveConfig(){
  const r = await api('/api/config', {school:$('school').value, theme:$('theme').value});
  $('c-msg').className = 'msg ok'; $('c-msg').textContent = r.ok ? '已保存' : '失败';
}
refresh();
</script></div></body></html>
"""


def _status() -> dict:
    cfg = config.load()
    bili_ok = False
    try:
        bili_ok = BilibiliClient(cookie=config.read_cookie("bilibili"), timeout=8).is_login()
    except Exception:
        bili_ok = False
    files = []
    out = Path(cfg.get("out_dir") or config.OUT_DIR)
    if out.exists():
        for p in sorted(out.rglob("*"), key=lambda x: x.stat().st_mtime, reverse=True)[:12]:
            if p.is_file() and p.suffix.lower() in (".docx", ".pptx", ".pdf", ".md", ".mp3", ".mp4", ".txt"):
                files.append({"name": p.name, "rel": str(p.relative_to(out)),
                              "size_kb": max(1, p.stat().st_size // 1024)})
    return {
        "home": str(config.HOME),
        "config": cfg,
        "bilibili": {"logged_in": bili_ok, "has_cookie": bool(config.read_cookie("bilibili"))},
        "zxxk": {"has_cookie": bool(config.read_cookie("zxxk"))},
        "deps": {k: bool(shutil.which(k)) for k in ("yt-dlp", "ffmpeg", "node")},
        "recent_files": files,
        "python": shutil.which("python3") or "python3",
        "server_path": str(Path(__file__).resolve().parent.parent),
    }


class Handler(BaseHTTPRequestHandler):
    server_version = "jszz"

    def log_message(self, fmt, *args):  # 静音访问日志
        return

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _text(self, text, code=200, ctype="text/html; charset=utf-8"):
        body = text.encode()
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urllib.parse.urlparse(self.path).path
        if path in ("/", "/index.html"):
            return self._text(PAGE)
        if path == "/api/status":
            return self._json(_status())
        if path.startswith("/files/"):
            rel = urllib.parse.unquote(path[len("/files/"):])
            cfg = config.load()
            root = Path(cfg.get("out_dir") or config.OUT_DIR).resolve()
            target = (root / rel).resolve()
            if not str(target).startswith(str(root)) or not target.is_file():
                return self._text("not found", 404)
            ctype = mimetypes.guess_type(str(target))[0] or "application/octet-stream"
            body = target.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Content-Disposition",
                             f"attachment; filename*=UTF-8''{urllib.parse.quote(target.name)}")
            self.end_headers()
            self.wfile.write(body)
            return
        return self._text("not found", 404)

    def do_POST(self):
        path = urllib.parse.urlparse(self.path).path
        length = int(self.headers.get("Content-Length") or 0)
        try:
            body = json.loads(self.rfile.read(length) or b"{}")
        except Exception:
            return self._json({"ok": False, "error": "bad json"}, 400)

        if path == "/api/cookie":
            site = str(body.get("site") or "")
            if site not in ("bilibili", "zxxk"):
                return self._json({"ok": False, "error": "unknown site"}, 400)
            cookie = str(body.get("cookie") or "").strip()
            if not cookie:
                return self._json({"ok": False, "error": "cookie 为空"})
            p = config.write_cookie(site, cookie)
            return self._json({"ok": True, "path": str(p)})

        if path == "/api/config":
            cfg = config.load()
            for k in ("school", "author", "theme"):
                if k in body:
                    cfg[k] = body[k]
            config.save(cfg)
            return self._json({"ok": True, "config": cfg})

        if path == "/api/test":
            site = body.get("site")
            if site == "bilibili":
                c = BilibiliClient(cookie=config.read_cookie("bilibili"), timeout=10)
                try:
                    r = c.search("高考", page_size=2)
                    login = c.is_login()
                    return self._json({"ok": True,
                                       "detail": f"搜索正常，返回 {r.get('total')} 条；"
                                                 f"{'已登录' if login else '未登录（只能 1080P）'}"})
                except Exception as e:
                    return self._json({"ok": False, "detail": f"{type(e).__name__}: {e}"})
            if site == "zxxk":
                try:
                    c = ZxxkClient(cookie_file=config.cookie_path("zxxk"))
                    ck = config.read_cookie("zxxk")
                    if ck:
                        c.http.set_cookies(ck)
                    r = c.search("高考", page_size=2)
                    if r.get("ok"):
                        return self._json({"ok": True,
                                           "detail": f"学科网接口正常，命中 {r.get('total')} 条"})
                    return self._json({"ok": False, "detail": r.get("error") or "被拦截"})
                except Exception as e:
                    return self._json({"ok": False, "detail": f"{type(e).__name__}: {e}"})
            return self._json({"ok": False, "detail": "unknown site"})
        return self._json({"ok": False, "error": "not found"}, 404)


def serve(port: int | None = None, host: str = "127.0.0.1") -> None:
    config.ensure_dirs()
    port = port or int(config.load().get("webui_port") or 8770)
    srv = ThreadingHTTPServer((host, port), Handler)
    print(f"jszz 控制台: http://{host}:{port}/  （Ctrl+C 退出）", flush=True)
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        srv.server_close()


if __name__ == "__main__":
    serve()
