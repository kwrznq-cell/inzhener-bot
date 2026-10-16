// Хранилище базы знаний на стороне Anthropic (Files API).
// Файлы живут не в коде: их загружают и удаляют через интерфейс сайта.
//
// РАЗДЕЛЕНИЕ БАЗ МЕЖДУ БОТАМИ.
// Хранилище Anthropic общее для всех ботов, работающих на одном ключе API.
// Чтобы боты не видели документы друг друга, имя каждого файла помечается меткой бота:
// «[yurist] Протокол разногласий.txt». Бот видит, отдаёт модели и может удалять только
// свои файлы — проверка на сервере, подделать список идентификаторов с клиента нельзя.
//
// BOT_ID      — метка этого бота: yurist, inzhener и т. д. Не задан — разделения нет.
// BOT_LEGACY  — «1» у того бота, которому принадлежат старые файлы без метки
//               (загруженные до разделения). Ставится ровно одному боту.
const { docxToText } = require('./knowledge');

const BASE = process.env.ANTHROPIC_BASE_URL || 'https://api.anthropic.com';
const DOC_MIME = ['application/pdf', 'text/plain'];
const IMG_MIME = ['image/jpeg', 'image/png', 'image/gif', 'image/webp'];

const BOT_ID = String(process.env.BOT_ID || '').trim().toLowerCase().replace(/[^a-z0-9_-]/g, '');
const BOT_LEGACY = ['1', 'true', 'yes'].includes(String(process.env.BOT_LEGACY || '').trim().toLowerCase());
const TAG = BOT_ID ? `[${BOT_ID}] ` : '';

// Чья метка стоит на файле: «[yurist] Договор.txt» → 'yurist'; без метки → null
function ownerOf(filename) {
  const m = String(filename || '').match(/^\[([a-z0-9_-]{1,40})\]\s/i);
  return m ? m[1].toLowerCase() : null;
}

// Виден ли файл этому боту
function belongs(filename) {
  if (!BOT_ID) return true;                       // разделение выключено
  const owner = ownerOf(filename);
  if (owner === null) return BOT_LEGACY;          // старый файл без метки
  return owner === BOT_ID;
}

// Своя метка в показе не нужна — сотрудник видит обычное имя файла
function display(filename) {
  const name = String(filename || '');
  return BOT_ID && ownerOf(name) === BOT_ID ? name.slice(TAG.length) : name;
}

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

// Список файлов базы знаний — только те, что принадлежат этому боту
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
    .filter((f) => belongs(f.filename))
    .map((f) => ({
      id: f.id,
      name: display(f.filename),
      mime: f.mime_type,
      size: f.size_bytes,
      created: f.created_at,
      kind: IMG_MIME.includes(f.mime_type) ? 'image' : DOC_MIME.includes(f.mime_type) ? 'document' : 'other',
    }))
    .sort((a, b) => String(a.name).localeCompare(String(b.name), 'ru'));
}

// Загрузка нового файла. .docx превращаем в текст, .doc не поддерживается Anthropic.
async function uploadFile({ name, mime, buffer }) {
  let filename = String(name || 'файл').replace(/^\[[a-z0-9_-]{1,40}\]\s*/i, '').slice(0, 180);
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
  form.append('file', new Blob([data], { type }), TAG + filename);
  const res = await call('/v1/files', { method: 'POST', headers: headers(), body: form });
  return { id: res.id, name: display(res.filename), mime: res.mime_type, size: res.size_bytes, created: res.created_at };
}

// Удалять можно только свой файл: идентификатор с клиента сверяем со списком бота
async function deleteFile(id) {
  if (!/^file_[A-Za-z0-9_-]+$/.test(String(id))) throw new Error('Некорректный идентификатор файла');
  if (BOT_ID) {
    const mine = await listFiles();
    if (!mine.some((f) => f.id === id)) throw new Error('Файл не принадлежит базе знаний этого ассистента');
  }
  await call('/v1/files/' + id, { method: 'DELETE', headers: headers() });
}

// Блоки для запроса к модели. listFiles уже отфильтрован по боту, поэтому чужой
// идентификатор, присланный с клиента, просто не найдётся и в запрос не попадёт.
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

module.exports = { listFiles, uploadFile, deleteFile, buildKnowledgeBlocks, ownerOf, belongs, display };
