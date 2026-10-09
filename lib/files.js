// Хранилище базы знаний на стороне Anthropic (Files API).
// Файлы живут не в коде: их загружают и удаляют через интерфейс сайта.
const { docxToText } = require('./knowledge');

const BASE = process.env.ANTHROPIC_BASE_URL || 'https://api.anthropic.com';
const DOC_MIME = ['application/pdf', 'text/plain'];
const IMG_MIME = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];

function headers(extra = {}) {
  return { 'x-api-key': process.env.ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01', ...extra };
}

async function call(path, opts = {}) {
  const r = await fetch(BASE + path, opts);
  if (!r.ok) {
    const text = await r.text();
    let msg = text;
    try { msg = JSON.parse(text).error?.message || text; } catch {}
    const err = new Error(msg);
    err.status = r.status;
    throw err;
  }
  return r.status === 204 ? null : r.json();
}

// Список файлов базы знаний
async function listFiles() {
  const out = [];
  let path = '/v1/files?limit=100';
  for (let i = 0; i < 20 && path; i++) {
    const page = await call(path, { headers: headers() });
    out.push(...(page.data || []));
    path = page.next_page || null;
    if (path && !path.startsWith('/')) path = '/v1/files?page=' + encodeURIComponent(page.next_page);
  }
  return out
    .map((f) => ({
      id: f.id,
      name: f.filename,
      mime: f.mime_type,
      size: f.size_bytes,
      created: f.created_at,
      kind: IMG_MIME.includes(f.mime_type) ? 'image' : DOC_MIME.includes(f.mime_type) ? 'document' : 'other',
    }))
    .sort((a, b) => String(a.name).localeCompare(String(b.name), 'ru'));
}

// Загрузка нового файла. .docx превращаем в текст, .doc не поддерживается Anthropic.
async function uploadFile({ name, mime, buffer }) {
  let filename = String(name || 'файл').slice(0, 200);
  let data = buffer;
  let type = mime || '';

  if (/\.docx$/i.test(filename)) {
    data = Buffer.from(docxToText(buffer), 'utf8');
    filename = filename.replace(/\.docx$/i, '.txt');
    type = 'text/plain';
  } else if (/\.pdf$/i.test(filename)) {
    type = 'application/pdf';
  } else if (/\.(txt|md|csv|json|rtf)$/i.test(filename)) {
    if (/\.(md|csv|json|rtf)$/i.test(filename)) filename = filename.replace(/\.[^.]+$/, '') + '.txt';
    type = 'text/plain';
  } else if (IMG_MIME.includes(type)) {
    // изображение — оставляем как есть
  } else if (/\.docx?$/i.test(filename)) {
    throw new Error('Старый формат .doc не поддерживается. Пересохраните файл как .docx или PDF.');
  } else {
    throw new Error('Неподдерживаемый формат. Загрузите PDF, DOCX, TXT, MD, CSV или изображение.');
  }

  const form = new FormData();
  form.append('file', new Blob([data], { type }), filename);
  const res = await call('/v1/files', { method: 'POST', headers: headers(), body: form });
  return { id: res.id, name: res.filename, mime: res.mime_type, size: res.size_bytes, created: res.created_at };
}

async function deleteFile(id) {
  if (!/^file_[A-Za-z0-9_-]+$/.test(String(id))) throw new Error('Некорректный идентификатор файла');
  await call('/v1/files/' + id, { method: 'DELETE', headers: headers() });
}

// Блоки для запроса к модели. Идентификаторы всегда сверяем со списком на сервере.
async function buildKnowledgeBlocks(requestedIds) {
  const files = await listFiles();
  const use = Array.isArray(requestedIds) && requestedIds.length
    ? files.filter((f) => requestedIds.includes(f.id))
    : files;
  return use
    .filter((f) => f.kind !== 'other')
    .map((f) => f.kind === 'image'
      ? { type: 'image', source: { type: 'file', file_id: f.id } }
      : { type: 'document', source: { type: 'file', file_id: f.id }, title: f.name, citations: { enabled: false } });
}

module.exports = { listFiles, uploadFile, deleteFile, buildKnowledgeBlocks };
