// База знаний: все файлы из папки /knowledge подмешиваются в каждый диалог.
// Поддерживаются .md .txt .csv .pdf .docx
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');

const DIR = path.join(process.cwd(), 'knowledge');
const TEXT_EXT = ['.md', '.txt', '.csv'];
const ALL_EXT = [...TEXT_EXT, '.pdf', '.docx'];

let cache = null;

function walk(dir) {
  if (!fs.existsSync(dir)) return [];
  return fs.readdirSync(dir, { withFileTypes: true }).flatMap((e) => {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) return walk(p);
    if (e.name.startsWith('.') || e.name.toLowerCase() === 'readme.md') return [];
    return ALL_EXT.includes(path.extname(e.name).toLowerCase()) ? [p] : [];
  });
}

// Минимальный распаковщик .docx (zip) без внешних зависимостей
function docxToText(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('bad docx');
  const count = buf.readUInt16LE(eocd + 10);
  let off = buf.readUInt32LE(eocd + 16);
  for (let n = 0; n < count; n++) {
    const method = buf.readUInt16LE(off + 10);
    const csize = buf.readUInt32LE(off + 20);
    const nameLen = buf.readUInt16LE(off + 28);
    const extraLen = buf.readUInt16LE(off + 30);
    const commentLen = buf.readUInt16LE(off + 32);
    const local = buf.readUInt32LE(off + 42);
    const name = buf.slice(off + 46, off + 46 + nameLen).toString();
    if (name === 'word/document.xml') {
      const lNameLen = buf.readUInt16LE(local + 26);
      const lExtraLen = buf.readUInt16LE(local + 28);
      const start = local + 30 + lNameLen + lExtraLen;
      const data = buf.slice(start, start + csize);
      const xml = (method === 8 ? zlib.inflateRawSync(data) : data).toString('utf8');
      return xmlToText(xml);
    }
    off += 46 + nameLen + extraLen + commentLen;
  }
  throw new Error('document.xml not found');
}

function xmlToText(xml) {
  return xml
    .replace(/<w:tab\/>/g, '\t')
    .replace(/<w:br[^>]*\/>/g, '\n')
    .replace(/<\/w:p>/g, '\n')
    .replace(/<\/w:tc>/g, ' | ')
    .replace(/<[^>]+>/g, '')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function loadKnowledge() {
  if (cache) return cache;
  const files = walk(DIR);
  const blocks = [];
  const list = [];
  for (const file of files) {
    const rel = path.relative(DIR, file);
    const ext = path.extname(file).toLowerCase();
    try {
      const buf = fs.readFileSync(file);
      if (ext === '.pdf') {
        blocks.push({
          type: 'document',
          source: { type: 'base64', media_type: 'application/pdf', data: buf.toString('base64') },
          title: rel,
          citations: { enabled: false },
        });
      } else {
        const text = ext === '.docx' ? docxToText(buf) : buf.toString('utf8');
        blocks.push({
          type: 'document',
          source: { type: 'text', media_type: 'text/plain', data: text || '(пусто)' },
          title: rel,
          citations: { enabled: false },
        });
      }
      list.push(rel);
    } catch (e) {
      console.error('Не удалось прочитать', rel, e.message);
    }
  }
  cache = { blocks, list };
  return cache;
}

function listKnowledge() {
  return loadKnowledge().list;
}

module.exports = { loadKnowledge, listKnowledge, docxToText };
