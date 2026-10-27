// alicfw JS-challenge solver: stdin = challenge HTML, argv[2] = site origin.
// Prints JSON {cookies:[...], reloaded:bool}. Cookie strings are full Set-Cookie
// fragments ("name=value;expires=...;path=/").
const fs = require('fs');
const vm = require('vm');

const html = fs.readFileSync(0, 'utf8');
const origin = process.argv[2] || 'https://www.zxxk.com/';
let u;
try { u = new URL(origin); } catch (e) { u = new URL('https://www.zxxk.com/'); }

const cookies = [];
let reloaded = false;
const parms = {};
for (const m of html.matchAll(/id="?(parm_\d+)"?[^>]*>\s*([^<\s]+)/g)) parms[m[1]] = m[2];
for (const m of html.matchAll(/name="?(parm_\d+)"?\s+value=\s*([^\s<]+)/g)) parms[m[1]] = m[2];

// textarea/other elements addressed by id (aliyun acw_sc__v2 puts JSON in #renderData)
const elems = {};
for (const m of html.matchAll(/<textarea[^>]*id="([^"]+)"[^>]*>([\s\S]*?)<\/textarea>/g)) elems[m[1]] = m[2];
for (const m of html.matchAll(/<input[^>]*id="([^"]+)"[^>]*value="([^"]*)"/g)) elems[m[1]] = m[2];
for (const m of html.matchAll(/<input[^>]*value="([^"]*)"[^>]*id="([^"]+)"/g)) elems[m[2]] = m[1];

const scripts = [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)]
  .map((m) => m[1]).filter((s) => s.trim());

const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36';

const location = {
  href: u.href, host: u.host, hostname: u.hostname, protocol: u.protocol,
  pathname: u.pathname, search: u.search, hash: '', origin: u.origin,
  reload: () => { reloaded = true; }, replace: () => { reloaded = true; }, assign: () => { reloaded = true; },
};

const doc = {
  location,
  getElementById: (id) => {
    if (parms[id] !== undefined) return { value: parms[id], innerHTML: parms[id], textContent: parms[id] };
    const v = elems[id] !== undefined ? elems[id] : '';
    return { value: v, innerHTML: v, textContent: v, innerText: v };
  },
  getElementsByTagName: () => [],
  querySelector: () => null,
  addEventListener: () => {},
  createElement: () => ({ style: {}, setAttribute() {}, appendChild() {} }),
  body: { appendChild() {}, removeChild() {} },
  documentElement: {},
  referrer: '',
  readyState: 'complete',
  cookie: '',
};
Object.defineProperty(doc, 'cookie', {
  get() { return cookies.map((c) => c.split(';')[0]).join('; '); },
  set(v) { cookies.push(String(v)); },
  configurable: true,
});

const sandbox = {
  document: doc, navigator: { userAgent: UA, platform: 'Win32', language: 'zh-CN', languages: ['zh-CN'] },
  location, screen: { width: 1920, height: 1080, availWidth: 1920, availHeight: 1040, colorDepth: 24 },
  innerWidth: 1920, innerHeight: 900, outerWidth: 1920, outerHeight: 1040, devicePixelRatio: 1,
  setTimeout: (fn) => { try { fn(); } catch (e) {} return 0; }, clearTimeout: () => {},
  setInterval: () => 0, clearInterval: () => {},
  console: { log: () => {}, warn: () => {}, error: () => {}, info: () => {} },
  Date, Math, String, Number, Boolean, Array, Object, JSON, RegExp, Error, TypeError,
  parseInt, parseFloat, isNaN, isFinite, encodeURIComponent, decodeURIComponent, unescape, escape,
};
sandbox.window = sandbox; sandbox.self = sandbox; sandbox.top = sandbox;
sandbox.parent = sandbox; sandbox.frames = sandbox; sandbox.globalThis = sandbox;
vm.createContext(sandbox);

const results = [];
for (const s of scripts) {
  try { vm.runInContext(s, sandbox, { timeout: 5000 }); results.push('ok'); }
  catch (e) { results.push('err:' + e.message); }
}
// The page calls check() from <body onload="check()"> (or similar).
const names = new Set(['check', 'onload']);
for (const m of html.matchAll(/onload="([A-Za-z0-9_$.]+)\s*\(/g)) names.add(m[1]);
for (const fn of names) {
  if (typeof sandbox[fn] === 'function') {
    try { sandbox[fn](); results.push('called:' + fn); }
    catch (e) { results.push('callerr:' + fn + ':' + e.message); }
  }
}
process.stdout.write(JSON.stringify({ cookies, reloaded, parms, scripts: scripts.length, results }));
