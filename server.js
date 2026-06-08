// server.js — дашборд с СЕРВЕРНОЙ авторизацией (cookie-сессия).
// Пароль берётся из переменной окружения DASH_PASS (Railway). Если не задана — дефолт 'uzum2026'.
// Отдаётся ТОЛЬКО index.html и только после входа; все прочие файлы (Эталон.xlsx, *.py, *.md, копии) — 404.
const http = require('http');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

const ROOT   = __dirname;
const PORT   = process.env.PORT || 3000;
const PASS   = process.env.DASH_PASS || 'uzum2026';
const SECRET = process.env.DASH_SECRET || crypto.createHash('sha256').update('uzum-dash|' + PASS).digest('hex');
const COOKIE = 'dash_sess';
const MAXAGE = 30 * 24 * 3600;            // 30 дней (сек)

const sign = exp => exp + '.' + crypto.createHmac('sha256', SECRET).update(String(exp)).digest('base64url');
function valid(tok) {
  if (!tok) return false;
  const i = tok.indexOf('.');
  if (i <= 0) return false;
  const exp = tok.slice(0, i), sig = tok.slice(i + 1);
  if (!/^\d+$/.test(exp) || +exp < Date.now()) return false;
  const expect = crypto.createHmac('sha256', SECRET).update(exp).digest('base64url');
  const a = Buffer.from(sig), b = Buffer.from(expect);
  return a.length === b.length && crypto.timingSafeEqual(a, b);
}
function cookieVal(req, name) {
  for (const part of (req.headers.cookie || '').split(';')) {
    const k = part.trim().slice(0, name.length), eq = part.trim()[name.length];
    if (k === name && eq === '=') return decodeURIComponent(part.trim().slice(name.length + 1));
  }
  return null;
}
const authed = req => valid(cookieVal(req, COOKIE));

const LOGIN = err => `<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0"><title>Вход · ПВЗ</title>
<meta property="og:title" content="ПВЗ — интерактивный дашборд"><meta property="og:description" content="Факт, прогноз и точка безубыточности по 18 ПВЗ."><meta property="og:type" content="website"><meta name="twitter:card" content="summary"><meta name="twitter:title" content="ПВЗ — интерактивный дашборд"><meta name="twitter:description" content="Факт, прогноз и точка безубыточности по 18 ПВЗ."><meta name="theme-color" content="#0a0c11">
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<style>
@import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,400..800&family=Hanken+Grotesk:wght@400;500;600;700&display=swap');
*{margin:0;padding:0;box-sizing:border-box}
body{min-height:100vh;background:#0a0c11;color:#e8ecf2;font-family:'Hanken Grotesk',system-ui,sans-serif;display:grid;place-items:center;background-image:radial-gradient(700px 400px at 50% 32%,rgba(139,109,255,.18),transparent 60%)}
.gbox{text-align:center;width:min(340px,88vw);padding:34px 28px;background:#13171f;border:1px solid #2c3543;border-radius:20px;box-shadow:0 30px 80px rgba(0,0,0,.6)}
.glogo{width:54px;height:54px;border-radius:15px;margin:0 auto 16px;background:conic-gradient(from 140deg,#8b6dff,#49c8ff,#ffb84d,#8b6dff);display:grid;place-items:center;font-family:'Bricolage Grotesque',sans-serif;font-weight:800;font-size:28px;color:#0a0c11}
h2{font-family:'Bricolage Grotesque',sans-serif;font-size:26px;font-weight:800;margin-bottom:6px}
p{color:#8893a5;font-size:13px;margin-bottom:18px;line-height:1.4}
input{width:100%;padding:12px 14px;border-radius:11px;border:1px solid #2c3543;background:#0a0c11;color:#e8ecf2;font-family:inherit;font-size:15px;outline:none;text-align:center}
input:focus{border-color:#8b6dff}
button{width:100%;margin-top:11px;padding:12px;border:0;border-radius:11px;background:#8b6dff;color:#0a0c11;font-family:'Bricolage Grotesque',sans-serif;font-weight:700;font-size:15px;cursor:pointer;transition:transform .12s}
button:hover{transform:translateY(-1px)}
.err{color:#ff5c72;font-size:12.5px;margin-top:10px;min-height:16px}
</style></head><body>
<form class="gbox" method="POST" action="/__auth">
  <div class="glogo">U</div>
  <h2>ПВЗ</h2>
  <p>Введите пароль для доступа к дашборду</p>
  <input type="password" name="pw" placeholder="Пароль" autocomplete="current-password" autofocus>
  <button type="submit">Войти</button>
  <div class="err">${err ? 'Неверный пароль' : ''}</div>
</form></body></html>`;

function send(res, code, type, body, extra) {
  res.writeHead(code, Object.assign({ 'Content-Type': type, 'Cache-Control': 'no-store' }, extra || {}));
  res.end(body);
}

http.createServer((req, res) => {
  let p;
  try { p = decodeURIComponent(new URL(req.url, 'http://x').pathname); } catch (e) { p = req.url; }

  // вход
  if (req.method === 'POST' && p === '/__auth') {
    let body = '';
    req.on('data', c => { body += c; if (body.length > 8192) req.destroy(); });
    req.on('end', () => {
      const m = /(?:^|&)pw=([^&]*)/.exec(body);
      let pw = ''; try { pw = decodeURIComponent((m ? m[1] : '').replace(/\+/g, ' ')); } catch (e) {}
      const ok = pw.length === PASS.length && crypto.timingSafeEqual(Buffer.from(pw), Buffer.from(PASS));
      if (ok) {
        const tok = sign(Date.now() + MAXAGE * 1000);
        send(res, 302, 'text/plain', '', { 'Set-Cookie': `${COOKIE}=${encodeURIComponent(tok)}; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=${MAXAGE}`, 'Location': '/' });
      } else {
        send(res, 401, 'text/html; charset=utf-8', LOGIN(true));
      }
    });
    return;
  }
  // выход
  if (p === '/__logout') {
    send(res, 302, 'text/plain', '', { 'Set-Cookie': `${COOKIE}=; HttpOnly; Secure; SameSite=Lax; Path=/; Max-Age=0`, 'Location': '/' });
    return;
  }
  // без сессии → страница входа на любой GET (URL сохраняется), 200 чтобы healthcheck Railway был зелёным
  if (!authed(req)) { send(res, 200, 'text/html; charset=utf-8', LOGIN(false)); return; }

  // с сессией:
  if (p === '/v2' || p === '/v2.html') { send(res, 302, 'text/plain', '', { 'Location': '/' }); return; }
  if (p === '/' || p === '/index.html') {
    fs.readFile(path.join(ROOT, 'index.html'), (e, buf) => e
      ? send(res, 500, 'text/plain; charset=utf-8', 'index.html missing')
      : send(res, 200, 'text/html; charset=utf-8', buf));
    return;
  }
  // данные дашборда — только авторизованным; data.json (источник правды) отдаём как JS-обёртку с теми же const
  if (p === '/data.js') {
    fs.readFile(path.join(ROOT, 'data.json'), 'utf8', (e, txt) => e
      ? send(res, 500, 'application/javascript; charset=utf-8', 'console.error("data.json missing")')
      : send(res, 200, 'application/javascript; charset=utf-8',
          'const __D=' + txt + ';const DATA=__D.DATA,WK=__D.WK,FACT=__D.FACT,SAL=__D.SAL;'));
    return;
  }
  // ВСЁ остальное (эталон, *.py, *.md, копии, выгрузки) НЕ отдаём
  send(res, 404, 'text/plain; charset=utf-8', 'Not found');
}).listen(PORT, () => console.log('dashboard on :' + PORT + (process.env.DASH_PASS ? ' [DASH_PASS env]' : ' [default pass]')));
