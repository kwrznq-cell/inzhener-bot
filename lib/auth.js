const crypto = require('crypto');

const COOKIE = 'eb_session';
const MAX_AGE = 60 * 60 * 24 * 30; // 30 дней

function secret() {
  return process.env.SESSION_SECRET || process.env.STAFF_PASSWORD || 'change-me';
}

function sign(payload) {
  const data = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const sig = crypto.createHmac('sha256', secret()).update(data).digest('base64url');
  return `${data}.${sig}`;
}

function verify(token) {
  if (!token || !token.includes('.')) return null;
  const [data, sig] = token.split('.');
  const expected = crypto.createHmac('sha256', secret()).update(data).digest('base64url');
  const a = Buffer.from(sig);
  const b = Buffer.from(expected);
  if (a.length !== b.length || !crypto.timingSafeEqual(a, b)) return null;
  try {
    const payload = JSON.parse(Buffer.from(data, 'base64url').toString());
    if (payload.exp < Date.now()) return null;
    return payload;
  } catch {
    return null;
  }
}

function parseCookies(req) {
  const out = {};
  (req.headers.cookie || '').split(';').forEach((p) => {
    const i = p.indexOf('=');
    if (i > 0) out[p.slice(0, i).trim()] = decodeURIComponent(p.slice(i + 1).trim());
  });
  return out;
}

function getSession(req) {
  return verify(parseCookies(req)[COOKIE]);
}

function setSessionCookie(res, name) {
  const token = sign({ name: name || 'сотрудник', exp: Date.now() + MAX_AGE * 1000 });
  const secure = process.env.NODE_ENV === 'production' || process.env.VERCEL ? '; Secure' : '';
  res.setHeader('Set-Cookie', `${COOKIE}=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${MAX_AGE}${secure}`);
}

function clearSessionCookie(res) {
  res.setHeader('Set-Cookie', `${COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0`);
}

function checkPassword(input) {
  const list = (process.env.STAFF_PASSWORD || '')
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
  if (!list.length) return false;
  const a = crypto.createHash('sha256').update(String(input || '')).digest();
  return list.some((p) => crypto.timingSafeEqual(a, crypto.createHash('sha256').update(p).digest()));
}

async function readJson(req) {
  if (req.body && typeof req.body === 'object') return req.body;
  if (typeof req.body === 'string') return JSON.parse(req.body || '{}');
  const chunks = [];
  for await (const c of req) chunks.push(c);
  const raw = Buffer.concat(chunks).toString();
  return raw ? JSON.parse(raw) : {};
}

module.exports = { getSession, setSessionCookie, clearSessionCookie, checkPassword, readJson };
