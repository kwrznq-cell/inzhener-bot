const { getSession, readJson } = require('../lib/auth');
const { listFiles, uploadFile, deleteFile } = require('../lib/files');

module.exports = async (req, res) => {
  res.setHeader('Content-Type', 'application/json');
  const session = getSession(req);
  if (!session) { res.statusCode = 401; return res.end(JSON.stringify({ error: 'Требуется вход' })); }
  if (!process.env.ANTHROPIC_API_KEY) { res.statusCode = 500; return res.end(JSON.stringify({ error: 'На сервере не задан ANTHROPIC_API_KEY' })); }

  try {
    if (req.method === 'GET') {
      return res.end(JSON.stringify({ files: await listFiles() }));
    }
    if (req.method === 'POST') {
      const body = await readJson(req);
      const file = await uploadFile({ name: body.name, mime: body.mime, buffer: Buffer.from(String(body.data || ''), 'base64') });
      return res.end(JSON.stringify({ file }));
    }
    if (req.method === 'DELETE') {
      const id = new URL(req.url, 'http://x').searchParams.get('id');
      await deleteFile(id);
      return res.end(JSON.stringify({ ok: true }));
    }
    res.statusCode = 405;
    res.end(JSON.stringify({ error: 'Метод не поддерживается' }));
  } catch (e) {
    res.statusCode = e.status === 401 ? 500 : 400;
    res.end(JSON.stringify({ error: e.message }));
  }
};
