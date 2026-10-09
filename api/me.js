const { getSession } = require('../lib/auth');
const { listKnowledge } = require('../lib/knowledge');
const { DEFAULTS } = require('../lib/settings');

module.exports = async (req, res) => {
  const s = getSession(req);
  res.setHeader('Content-Type', 'application/json');
  if (!s) {
    res.statusCode = 401;
    return res.end(JSON.stringify({ auth: false }));
  }
  res.end(JSON.stringify({ auth: true, name: s.name, knowledge: listKnowledge(), defaults: DEFAULTS }));
};
