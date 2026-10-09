const { checkPassword, setSessionCookie, readJson } = require('../lib/auth');

module.exports = async (req, res) => {
  if (req.method !== 'POST') {
    res.statusCode = 405;
    return res.end();
  }
  const body = await readJson(req).catch(() => ({}));
  res.setHeader('Content-Type', 'application/json');
  if (!process.env.STAFF_PASSWORD) {
    res.statusCode = 500;
    return res.end(JSON.stringify({ error: 'На сервере не задан STAFF_PASSWORD' }));
  }
  if (!checkPassword(body.password)) {
    await new Promise((r) => setTimeout(r, 600));
    res.statusCode = 401;
    return res.end(JSON.stringify({ error: 'Неверный пароль' }));
  }
  const name = String(body.name || '').slice(0, 60).trim();
  setSessionCookie(res, name);
  res.end(JSON.stringify({ ok: true, name }));
};
