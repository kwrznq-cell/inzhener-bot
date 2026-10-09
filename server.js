// Локальный запуск без Vercel: node server.js  (переменные берутся из .env)
const http = require('http');
const fs = require('fs');
const path = require('path');

if (fs.existsSync('.env')) {
  for (const line of fs.readFileSync('.env', 'utf8').split('\n')) {
    const m = line.match(/^\s*([A-Z_]+)\s*=\s*(.*)\s*$/);
    if (m && !process.env[m[1]]) process.env[m[1]] = m[2].replace(/^["']|["']$/g, '');
  }
}

const routes = { '/api/chat': './api/chat', '/api/login': './api/login', '/api/logout': './api/logout', '/api/me': './api/me', '/api/files': './api/files' };
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png' };

http.createServer(async (req, res) => {
  const url = new URL(req.url, 'http://x');
  if (routes[url.pathname]) {
    try { await require(routes[url.pathname])(req, res); }
    catch (e) { console.error(e); if (!res.headersSent) res.statusCode = 500; res.end(); }
    return;
  }
  const file = path.join(__dirname, 'public', url.pathname === '/' ? 'index.html' : path.normalize(url.pathname));
  if (!file.startsWith(path.join(__dirname, 'public')) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) {
    res.statusCode = 404; return res.end('Not found');
  }
  res.setHeader('Content-Type', types[path.extname(file)] || 'application/octet-stream');
  fs.createReadStream(file).pipe(res);
}).listen(process.env.PORT || 3000, () => console.log(`Инженер запущен: http://localhost:${process.env.PORT || 3000}`));
