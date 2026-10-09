// Правки в чужой .docx: находим фрагмент текста и заменяем его,
// подсвечивая новую редакцию жёлтым (и, если нужно, зачёркивая старую).
const { zip, unzip } = require('./zip');

const escXml = (s) => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const unescXml = (s) => String(s)
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&amp;/g, '&');

// текст абзаца (склеиваем все w:t, учитываем разрывы строк и табы)
function paraText(xml) {
  let out = '';
  const re = /<w:(t|tab|br)(\s[^>]*)?(\/>|>)/g;
  let m;
  while ((m = re.exec(xml))) {
    if (m[1] === 'tab') { out += '\t'; continue; }
    if (m[1] === 'br') { out += '\n'; continue; }
    if (m[3] === '/>') continue;
    const end = xml.indexOf('</w:t>', re.lastIndex);
    if (end < 0) break;
    out += unescXml(xml.slice(re.lastIndex, end));
    re.lastIndex = end + 6;
  }
  return out;
}

const norm = (s) => String(s || '').replace(/ /g, ' ').replace(/[‑–—]/g, '-').replace(/\s+/g, ' ').trim();

function firstRpr(xml) {
  const m = xml.match(/<w:r\b[^>]*>\s*<w:rPr>([\s\S]*?)<\/w:rPr>/);
  return m ? m[1] : '';
}
function paraPr(xml) {
  const m = xml.match(/<w:pPr>[\s\S]*?<\/w:pPr>/);
  return m ? m[0] : '';
}

function run(text, rprInner, opts = {}) {
  if (text === '') return '';
  const extra = (opts.highlight ? '<w:highlight w:val="yellow"/>' : '') + (opts.strike ? '<w:strike/>' : '');
  const rpr = rprInner || extra ? `<w:rPr>${(rprInner || '').replace(/<w:highlight[^>]*\/>/g, '')}${extra}</w:rPr>` : '';
  return text.split('\n').map((line, i) =>
    (i ? '<w:r><w:br/></w:r>' : '') + (line ? `<w:r>${rpr}<w:t xml:space="preserve">${escXml(line)}</w:t></w:r>` : '')
  ).join('');
}

// edits: [{ find, replace, keepOld?: true — оставить прежнюю редакцию зачёркнутой }]
function applyEdits(buf, edits) {
  const files = unzip(buf);
  const key = 'word/document.xml';
  if (!files[key]) throw new Error('В файле нет word/document.xml — это не .docx');
  let xml = files[key].toString('utf8');

  const paras = xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  const report = [];

  for (const ed of edits) {
    const find = String(ed.find || '');
    const replace = String(ed.replace == null ? '' : ed.replace);
    if (!find.trim()) { report.push({ find, ok: false, why: 'пустой поиск' }); continue; }
    const nf = norm(find);
    let done = false;

    for (const p of paras) {
      if (!xml.includes(p)) continue;              // абзац уже заменён
      const txt = paraText(p);
      const nt = norm(txt);
      const at = nt.indexOf(nf);
      if (at < 0) continue;

      // позиции в исходном тексте ищем по нормализованному соответствию
      const before = sliceByNorm(txt, 0, at);
      const after = sliceByNorm(txt, at + nf.length, Infinity);
      const rpr = firstRpr(p);
      const pPr = paraPr(p);
      const body = run(before, rpr)
        + (ed.keepOld ? run(sliceByNorm(txt, at, at + nf.length), rpr, { strike: true }) + run(' ', rpr) : '')
        + run(replace, rpr, { highlight: true })
        + run(after, rpr);
      xml = xml.replace(p, `<w:p>${pPr}${body}</w:p>`);
      report.push({ find, ok: true });
      done = true;
      break;
    }
    if (!done) report.push({ find, ok: false, why: 'фрагмент не найден в документе' });
  }

  files[key] = Buffer.from(xml, 'utf8');
  const entries = Object.keys(files).map((n) => ({ name: n, data: files[n] }));
  return { buffer: zip(entries), report };
}

// вырезаем кусок исходного текста, считая позиции по нормализованной версии
function sliceByNorm(text, from, to) {
  let n = 0, start = -1, end = text.length;
  let prevSpace = true;
  for (let i = 0; i <= text.length; i++) {
    const ch = text[i];
    const isSpace = ch === undefined ? true : /\s| /.test(ch);
    if (n === from && start < 0) start = i;
    if (n >= to && end === text.length) { end = i; break; }
    if (ch === undefined) break;
    if (isSpace) { if (!prevSpace) n++; prevSpace = true; }
    else { n++; prevSpace = false; }
  }
  if (start < 0) start = text.length;
  return text.slice(start, Math.max(start, end));
}

// простая выжимка текста документа (для проверки и показа)
function docxText(buf) {
  const files = unzip(buf);
  const xml = (files['word/document.xml'] || Buffer.from('')).toString('utf8');
  return (xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || []).map(paraText).join('\n');
}

module.exports = { applyEdits, docxText, paraText };
