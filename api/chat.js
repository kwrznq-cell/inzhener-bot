const { getSession, readJson } = require('../lib/auth');
const { loadKnowledge } = require('../lib/knowledge');
const { buildKnowledgeBlocks } = require('../lib/files');
const { buildSystemPrompt } = require('../lib/prompt');
const { clean } = require('../lib/settings');
const { buildDocx } = require('../lib/docxgen');
const { buildHtml } = require('../lib/dochtml');
const { applyEdits } = require('../lib/docxedit');

const API_URL = (process.env.ANTHROPIC_BASE_URL || 'https://api.anthropic.com') + '/v1/messages';
const CANDIDATE_MODELS = [process.env.ANTHROPIC_MODEL, 'claude-opus-5', 'claude-opus-4-6', 'claude-opus-4-5', 'claude-sonnet-4-5'].filter(Boolean);
let workingModel = null;

// Инструмент памяти: модель сама записывает правила, о которых просит инженер.
// Сервер ничего не хранит — запись уходит в браузер событием {t:'memory'} и живёт в настройках.
const MEMORY_TOOL = {
  name: 'remember',
  description: 'Запомнить указание инженера на будущее: правило подбора, предпочтительный материал или производитель, типовое решение. Запись сохраняется в памяти ассистента и подставляется во все следующие ответы.',
  input_schema: {
    type: 'object',
    properties: {
      rule: { type: 'string', description: 'Правило одной-двумя фразами: что делаем, в каком случае, точная формулировка в кавычках, если она есть.' },
    },
    required: ['rule'],
  },
};


const BLOCK_SCHEMA = {
  type: 'object',
  description: 'Блок документа.',
  properties: {
    t: { type: 'string', enum: ['p', 'h1', 'h2', 'center', 'right', 'small', 'li', 'num', 'clause', 'table', 'sign', 'spacer', 'pagebreak'], description: 'Тип блока: p — абзац, h1/h2 — заголовки, center/right — выключка, li — маркированный пункт, num — нумерованный, clause — пункт договора с номером вида 5.6., table — таблица, sign — подписи сторон, spacer — отступ, pagebreak — новая страница.' },
    text: { type: 'string', description: 'Текст блока. Разметка: **жирный**, *курсив*, ==жёлтая заливка==, __подчёркнутый__.' },
    n: { type: 'string', description: 'Номер пункта для num/clause, например «5.6.».' },
    header: { type: 'array', items: { type: 'string' }, description: 'Шапка таблицы.' },
    rows: { type: 'array', items: { type: 'array', items: { type: 'string' } }, description: 'Строки таблицы.' },
    widths: { type: 'array', items: { type: 'number' }, description: 'Относительные ширины колонок, например [1, 2.4, 2.4].' },
    left: { type: 'object', description: 'Левая подпись: { role, company, title, person }.' },
    right: { type: 'object', description: 'Правая подпись: { role, company, title, person }.' },
  },
  required: ['t'],
};

const DOC_TOOL = {
  name: 'make_doc',
  description: 'Собрать готовый документ Word (.docx) на бланке нашей компании: опросный лист, техническое предложение, спецификацию подобранной арматуры, технические требования, обоснование выбора материалов, письмо. Файл сразу появится у пользователя для скачивания.',
  input_schema: {
    type: 'object',
    properties: {
      filename: { type: 'string', description: 'Имя файла без расширения, например «Опросный лист на подбор арматуры — ТЭЦ-3».' },
      company: { type: 'string', enum: ['belven', 'armeta', 'none'], description: 'От чьего имени документ: belven — ООО «Белвен», armeta — ООО «Армета», none — без шапки.' },
      title: { type: 'string', description: 'Заголовок документа, например «ОПРОСНЫЙ ЛИСТ НА ПОДБОР ТРУБОПРОВОДНОЙ АРМАТУРЫ».' },
      subtitle: { type: 'string', description: 'Подзаголовок: объект, заказчик, номер запроса, дата, город.' },
      footerNote: { type: 'string', description: 'Сноска внизу каждой страницы (необязательно).' },
      blocks: { type: 'array', items: BLOCK_SCHEMA, description: 'Содержимое документа по порядку.' },
    },
    required: ['filename', 'company', 'blocks'],
  },
};

const EDIT_TOOL = {
  name: 'edit_docx',
  description: 'Внести правки прямо в присланный файл .docx (опросный лист, техническое задание, спецификация): найти точный фрагмент текста и заменить его нашей редакцией. Новая редакция выделяется жёлтым, оформление документа сохраняется.',
  input_schema: {
    type: 'object',
    properties: {
      source: { type: 'string', description: 'Имя присланного файла, в который вносим правки (если прислан один — можно не указывать).' },
      filename: { type: 'string', description: 'Имя итогового файла без расширения.' },
      edits: {
        type: 'array',
        description: 'Список правок по порядку.',
        items: {
          type: 'object',
          properties: {
            find: { type: 'string', description: 'Точный фрагмент текста из документа (копируй дословно, вместе с началом и концом фразы).' },
            replace: { type: 'string', description: 'Новая редакция этого фрагмента.' },
            keepOld: { type: 'boolean', description: 'Оставить прежнюю редакцию зачёркнутой рядом (по умолчанию нет).' },
          },
          required: ['find', 'replace'],
        },
      },
    },
    required: ['filename', 'edits'],
  },
};

const fs = require('fs');
const path = require('path');
const logoCache = new Map();
function fileLogo(name) {
  if (!name) return null;
  if (logoCache.has(name)) return logoCache.get(name);
  let out = null;
  for (const dir of [path.join(__dirname, '..', 'public'), path.join(process.cwd(), 'public')]) {
    try {
      const buf = fs.readFileSync(path.join(dir, name));
      out = { data: buf.toString('base64'), ext: /jpe?g$/i.test(name) ? 'jpg' : 'png' };
      break;
    } catch {}
  }
  logoCache.set(name, out);
  return out;
}

// размеры PNG/JPEG, чтобы логотип в шапке не был растянут
function imageSize(buf) {
  try {
    if (buf.length > 24 && buf.readUInt32BE(0) === 0x89504e47) return { w: buf.readUInt32BE(16), h: buf.readUInt32BE(20) };
    let i = 2;
    while (i < buf.length) {
      if (buf[i] !== 0xff) { i++; continue; }
      const m = buf[i + 1];
      if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc) return { h: buf.readUInt16BE(i + 5), w: buf.readUInt16BE(i + 7) };
      i += 2 + buf.readUInt16BE(i + 2);
    }
  } catch {}
  return null;
}

function blankFor(settings, key) {
  const b = (settings.blanks || {})[key === 'armeta' ? 'armeta' : 'belven'];
  if (!b || key === 'none') return {};
  const logo = (b.logo && b.logo.data) ? b.logo : fileLogo(b.logoFile);
  const out = {
    company: { name: b.name, requisites: b.requisites, contacts: b.contacts },
    logo: logo || null,
  };
  if (logo) {
    const size = imageSize(Buffer.from(logo.data, 'base64'));
    const cy = 580000;                                  // ~1,6 см по высоте
    const cx = size && size.h ? Math.round(cy * size.w / size.h) : 1200000;
    out.logoCx = Math.min(cx, 2200000);
    out.logoCy = out.logoCx === cx ? cy : Math.round(2200000 * (size ? size.h / size.w : 0.4));
  }
  return out;
}

function safeName(s, fallback) {
  const n = String(s || '').replace(/[\\/:*?"<>|\u0000-\u001f]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 120);
  return n || fallback;
}

// ---- проверка входящих сообщений ----
function sanitizeMessages(raw) {
  if (!Array.isArray(raw) || !raw.length) throw new Error('Пустой диалог');
  const out = [];
  for (const m of raw.slice(-40)) {
    const role = m.role === 'assistant' ? 'assistant' : 'user';
    const blocks = (Array.isArray(m.content) ? m.content : [{ type: 'text', text: String(m.content || '') }])
      .map((b) => {
        if (b.type === 'text' && b.text) return { type: 'text', text: String(b.text) };
        if (role !== 'user') return null;
        if (b.type === 'image' && b.source?.type === 'base64')
          return { type: 'image', source: { type: 'base64', media_type: b.source.media_type, data: b.source.data } };
        if (b.type === 'document' && b.source?.type === 'base64' && b.source.media_type === 'application/pdf')
          return { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: b.source.data }, title: String(b.title || 'документ.pdf').slice(0, 200) };
        if (b.type === 'document' && b.source?.type === 'text')
          return { type: 'document', source: { type: 'text', media_type: 'text/plain', data: String(b.source.data || '(пусто)') }, title: String(b.title || 'документ').slice(0, 200) };
        return null;
      })
      .filter(Boolean);
    if (!blocks.length) continue;
    // склеиваем подряд идущие реплики одной роли
    if (out.length && out[out.length - 1].role === role) out[out.length - 1].content.push(...blocks);
    else out.push({ role, content: blocks });
  }
  while (out.length && out[0].role !== 'user') out.shift();
  if (!out.length || out[out.length - 1].role !== 'user') throw new Error('Последнее сообщение должно быть от пользователя');
  return out;
}

// ---- один потоковый запрос к Claude ----
async function streamOnce({ model, system, messages, tools, emit, settings }) {
  const body = { model, max_tokens: settings.maxTokens, system, messages, stream: true };
  if (tools.length) body.tools = tools;
  if (settings.thinking) body.thinking = { type: 'enabled', budget_tokens: Math.max(1024, Math.min(10000, settings.maxTokens - 2000)) };

  const resp = await fetch(API_URL, {
    method: 'POST',
    headers: {
      'content-type': 'application/json',
      'x-api-key': process.env.ANTHROPIC_API_KEY,
      'anthropic-version': '2023-06-01',
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    const text = await resp.text();
    const err = new Error(text);
    err.status = resp.status;
    err.isModelError = (resp.status === 404 || resp.status === 400) && /model/i.test(text);
    throw err;
  }

  const blocks = [];
  let stopReason = null;
  const decoder = new TextDecoder();
  let buffer = '';

  const handle = (evt) => {
    switch (evt.type) {
      case 'content_block_start': {
        const b = { ...evt.content_block };
        if (b.type === 'text') { b.text = b.text || ''; b.citations = b.citations || undefined; }
        if (b.type === 'server_tool_use' || b.type === 'tool_use') { b._json = ''; }
        if (b.type === 'web_search_tool_result' && Array.isArray(b.content)) {
          emit({ t: 'sources', v: b.content.filter((r) => r.url).map((r) => ({ url: r.url, title: r.title })) });
        }
        blocks[evt.index] = b;
        break;
      }
      case 'content_block_delta': {
        const b = blocks[evt.index];
        const d = evt.delta;
        if (d.type === 'text_delta') { b.text += d.text; emit({ t: 'text', v: d.text }); }
        else if (d.type === 'thinking_delta') { b.thinking = (b.thinking || '') + d.thinking; emit({ t: 'think', v: d.thinking }); }
        else if (d.type === 'signature_delta') { b.signature = d.signature; }
        else if (d.type === 'input_json_delta') { b._json += d.partial_json; }
        else if (d.type === 'citations_delta') {
          (b.citations = b.citations || []).push(d.citation);
          if (d.citation.url) emit({ t: 'cite', v: { url: d.citation.url, title: d.citation.title } });
        }
        break;
      }
      case 'content_block_stop': {
        const b = blocks[evt.index];
        if (b && (b.type === 'server_tool_use' || b.type === 'tool_use')) {
          try { b.input = JSON.parse(b._json || '{}'); } catch { b.input = {}; }
          delete b._json;
          if (b.type === 'server_tool_use' && b.input.query) emit({ t: 'search', v: b.input.query });
        }
        if (b && b.type === 'text' && !b.citations) delete b.citations;
        break;
      }
      case 'message_delta':
        if (evt.delta?.stop_reason) stopReason = evt.delta.stop_reason;
        break;
      case 'error':
        throw new Error(evt.error?.message || 'Ошибка API');
    }
  };

  for await (const chunk of resp.body) {
    buffer += decoder.decode(chunk, { stream: true });
    let idx;
    while ((idx = buffer.indexOf('\n\n')) >= 0) {
      const raw = buffer.slice(0, idx);
      buffer = buffer.slice(idx + 2);
      const dataLine = raw.split('\n').filter((l) => l.startsWith('data:')).map((l) => l.slice(5).trim()).join('');
      if (!dataLine) continue;
      let evt;
      try { evt = JSON.parse(dataLine); } catch { continue; }
      handle(evt);
    }
  }
  return { blocks: blocks.filter(Boolean), stopReason };
}

module.exports = async (req, res) => {
  if (req.method !== 'POST') { res.statusCode = 405; return res.end(); }
  const session = getSession(req);
  if (!session) { res.statusCode = 401; res.setHeader('Content-Type', 'application/json'); return res.end(JSON.stringify({ error: 'Требуется вход' })); }
  if (!process.env.ANTHROPIC_API_KEY) { res.statusCode = 500; res.setHeader('Content-Type', 'application/json'); return res.end(JSON.stringify({ error: 'На сервере не задан ANTHROPIC_API_KEY' })); }

  let messages, requestedFileIds = null, settings = clean(null), sourceDocs = [];
  try {
    const body = await readJson(req);
    messages = sanitizeMessages(body.messages);
    if (Array.isArray(body.fileIds)) requestedFileIds = body.fileIds.filter((x) => typeof x === 'string').slice(0, 100);
    if (Array.isArray(body.sourceDocs)) {
      sourceDocs = body.sourceDocs
        .filter((d) => d && typeof d.name === 'string' && typeof d.data === 'string')
        .slice(-5)
        .map((d) => ({ name: d.name.slice(0, 200), data: d.data }));
    }
    settings = clean(body.settings);
  } catch (e) {
    res.statusCode = 400; res.setHeader('Content-Type', 'application/json');
    return res.end(JSON.stringify({ error: e.message }));
  }

  res.statusCode = 200;
  res.setHeader('Content-Type', 'application/x-ndjson; charset=utf-8');
  res.setHeader('Cache-Control', 'no-cache, no-transform');
  res.setHeader('X-Accel-Buffering', 'no');
  const emit = (obj) => res.write(JSON.stringify(obj) + '\n');

  // база знаний: файлы, загруженные через интерфейс (Files API) + папка knowledge/ в репозитории
  let kbBlocks = [];
  try {
    kbBlocks = await buildKnowledgeBlocks(requestedFileIds);
  } catch (e) {
    console.error('База знаний недоступна:', e.message);
    emit({ t: 'notice', v: 'База знаний временно недоступна, отвечаю без неё: ' + e.message });
  }
  kbBlocks = [...loadKnowledge().blocks.map((b) => ({ ...b })), ...kbBlocks];

  const system = [{ type: 'text', text: buildSystemPrompt({ userName: session.name, hasKnowledge: kbBlocks.length > 0, settings }), cache_control: { type: 'ephemeral' } }];
  if (kbBlocks.length) {
    kbBlocks[kbBlocks.length - 1] = { ...kbBlocks[kbBlocks.length - 1], cache_control: { type: 'ephemeral' } };
    messages[0].content = [...kbBlocks, { type: 'text', text: '(Выше — документы внутренней базы знаний компании: шаблоны, критерии, примеры. Опирайся на них в первую очередь.)' }, ...messages[0].content];
  }
  const last = messages[messages.length - 1];
  last.content[last.content.length - 1] = { ...last.content[last.content.length - 1], cache_control: { type: 'ephemeral' } };

  const tools = [
    ...(settings.webSearch ? [{
      type: 'web_search_20250305',
      name: 'web_search',
      max_uses: settings.webSearchMax,
      user_location: { type: 'approximate', country: 'RU', timezone: 'Europe/Moscow' },
      ...(settings.domains.length ? { allowed_domains: settings.domains } : {}),
    }] : []),
    MEMORY_TOOL,
    DOC_TOOL,
    EDIT_TOOL,
  ];

  try {
    let convo = messages;
    for (let turn = 0; turn < 6; turn++) {
      const models = [...new Set([settings.model, ...(workingModel ? [workingModel] : []), ...CANDIDATE_MODELS])];
      let result;
      for (let i = 0; i < models.length; i++) {
        try {
          result = await streamOnce({ model: models[i], system, messages: convo, tools, emit, settings });
          if (workingModel !== models[i]) { workingModel = models[i]; emit({ t: 'model', v: models[i] }); }
          break;
        } catch (e) {
          if (e.isModelError && i < models.length - 1) continue;
          throw e;
        }
      }
      if (result.stopReason === 'tool_use') {
        const calls = result.blocks.filter((b) => b.type === 'tool_use');
        const answers = [];
        for (const c of calls) {
          if (c.name === 'make_doc') {
            try {
              const inp = c.input || {};
              const blank = blankFor(settings, inp.company);
              const doc = {
                title: inp.title || '',
                subtitle: inp.subtitle || '',
                footerNote: inp.footerNote || '',
                blocks: Array.isArray(inp.blocks) ? inp.blocks : [],
                ...blank,
              };
              const name = safeName(inp.filename, 'Документ') + '.docx';
              const buf = buildDocx(doc);
              emit({ t: 'doc', v: { name, size: buf.length, data: buf.toString('base64'), html: buildHtml(doc) } });
              answers.push({ type: 'tool_result', tool_use_id: c.id, content: `Файл «${name}» собран и передан пользователю (${Math.round(buf.length / 1024)} КБ).` });
            } catch (err) {
              answers.push({ type: 'tool_result', tool_use_id: c.id, content: 'Не удалось собрать документ: ' + err.message, is_error: true });
            }
            continue;
          }
          if (c.name === 'edit_docx') {
            try {
              const inp = c.input || {};
              const want = String(inp.source || '').toLowerCase();
              const src = sourceDocs.find((d) => want && d.name.toLowerCase().includes(want.replace(/\.docx$/, ''))) || sourceDocs[sourceDocs.length - 1];
              if (!src) throw new Error('исходный .docx не приложен к диалогу — попроси прислать файл опросного листа или спецификации');
              const edits = Array.isArray(inp.edits) ? inp.edits.slice(0, 200) : [];
              const { buffer, report } = applyEdits(Buffer.from(src.data, 'base64'), edits);
              const name = safeName(inp.filename, 'Документ с правками') + '.docx';
              emit({ t: 'doc', v: { name, size: buffer.length, data: buffer.toString('base64') } });
              const bad = report.filter((r) => !r.ok);
              answers.push({
                type: 'tool_result',
                tool_use_id: c.id,
                content: `Файл «${name}» готов: внесено правок ${report.length - bad.length} из ${report.length}.`
                  + (bad.length ? ' Не найдены фрагменты: ' + bad.map((b) => JSON.stringify(b.find.slice(0, 60))).join('; ') + '. Уточни формулировки дословно по тексту документа и вызови инструмент ещё раз только для них.' : ''),
              });
            } catch (err) {
              answers.push({ type: 'tool_result', tool_use_id: c.id, content: 'Не удалось внести правки: ' + err.message, is_error: true });
            }
            continue;
          }
          if (c.name !== 'remember') { answers.push({ type: 'tool_result', tool_use_id: c.id, content: 'Неизвестный инструмент.', is_error: true }); continue; }
          const text = String((c.input && c.input.rule) || '').trim().slice(0, 4000);
          if (!text) { answers.push({ type: 'tool_result', tool_use_id: c.id, content: 'Пустая формулировка — ничего не записано.', is_error: true }); continue; }
          const item = { id: 'm' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), text, ts: Date.now(), src: 'bot' };
          settings.memory = [...(settings.memory || []), item];
          emit({ t: 'memory', v: item });
          answers.push({ type: 'tool_result', tool_use_id: c.id, content: 'Записано в память.' });
        }
        convo = [...convo, { role: 'assistant', content: result.blocks }, { role: 'user', content: answers }];
        continue;
      }
      if (result.stopReason === 'pause_turn') {
        convo = [...convo, { role: 'assistant', content: result.blocks }];
        continue;
      }
      if (result.stopReason === 'max_tokens') emit({ t: 'notice', v: 'Ответ обрезан по длине. Напишите «продолжи».' });
      break;
    }
    emit({ t: 'done' });
  } catch (e) {
    console.error(e);
    let msg = e.message;
    try { msg = JSON.parse(e.message).error?.message || msg; } catch {}
    if (e.status === 401) msg = 'Неверный ANTHROPIC_API_KEY';
    if (e.status === 429) msg = 'Превышен лимит запросов к API, попробуйте через минуту';
    if (e.status === 529) msg = 'Сервис Anthropic перегружен, попробуйте ещё раз';
    emit({ t: 'error', v: msg });
  }
  res.end();
};
