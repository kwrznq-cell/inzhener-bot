const { clearSessionCookie } = require('../lib/auth');

module.exports = async (req, res) => {
  clearSessionCookie(res);
  res.setHeader('Content-Type', 'application/json');
  res.end(JSON.stringify({ ok: true }));
};
