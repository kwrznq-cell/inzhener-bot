// Сборка .docx без зависимостей: строгий деловой бланк, таблицы, подписи, колонтитулы.
const { zip } = require('./zip');

const esc = (s) => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;');

const TW = { cm: (v) => Math.round(v * 566.93) };

// ---- инлайн-разметка: **жирный**, *курсив*, ==жёлтая заливка==, __подчёркнутый__
function runs(text, base = {}) {
  const out = [];
  const src = String(text == null ? '' : text);
  const re = /(\*\*[^*]+\*\*|==[^=]+==|__[^_]+__|\*[^*\n]+\*)/g;
  let last = 0, m;
  const push = (t, o) => {
    if (t === '') return;
    const parts = String(t).split('\n');
    parts.forEach((p, i) => {
      if (i) out.push({ br: true });
      if (p) out.push({ text: p, ...base, ...o });
    });
  };
  while ((m = re.exec(src))) {
    push(src.slice(last, m.index), {});
    const tok = m[0];
    if (tok.startsWith('**')) push(tok.slice(2, -2), { b: true });
    else if (tok.startsWith('==')) push(tok.slice(2, -2), { hl: 'yellow' });
    else if (tok.startsWith('__')) push(tok.slice(2, -2), { u: true });
    else push(tok.slice(1, -1), { i: true });
    last = m.index + tok.length;
  }
  push(src.slice(last), {});
  return out;
}

function runXml(r, font, size) {
  if (r.br) return '<w:r><w:br/></w:r>';
  const rpr = [`<w:rFonts w:ascii="${esc(font)}" w:hAnsi="${esc(font)}" w:cs="${esc(font)}"/>`];
  if (r.b) rpr.push('<w:b/>');
  if (r.i) rpr.push('<w:i/>');
  if (r.u) rpr.push('<w:u w:val="single"/>');
  if (r.hl) rpr.push(`<w:highlight w:val="${esc(r.hl)}"/>`);
  if (r.caps) rpr.push('<w:caps/>');
  rpr.push(`<w:sz w:val="${r.size || size}"/><w:szCs w:val="${r.size || size}"/>`);
  if (r.color) rpr.push(`<w:color w:val="${esc(r.color)}"/>`);
  return `<w:r><w:rPr>${rpr.join('')}</w:rPr><w:t xml:space="preserve">${esc(r.text)}</w:t></w:r>`;
}

function para(text, opt = {}, S) {
  const pr = [];
  const jc = opt.align || 'both';
  pr.push(`<w:jc w:val="${jc}"/>`);
  const ind = [];
  if (opt.firstLine) ind.push(`w:firstLine="${opt.firstLine}"`);
  if (opt.left) ind.push(`w:left="${opt.left}"`);
  if (opt.hanging) ind.push(`w:hanging="${opt.hanging}"`);
  if (ind.length) pr.push(`<w:ind ${ind.join(' ')}/>`);
  pr.push(`<w:spacing w:before="${opt.before || 0}" w:after="${opt.after == null ? 120 : opt.after}" w:line="${opt.line || 276}" w:lineRule="auto"/>`);
  if (opt.keepNext) pr.push('<w:keepNext/>');
  if (opt.border) pr.push('<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="2" w:color="8A8F98"/></w:pBdr>');
  const rr = runs(text, opt.run || {}).map((r) => runXml(r, S.font, opt.size || S.size)).join('');
  return `<w:p><w:pPr>${pr.join('')}</w:pPr>${rr}</w:p>`;
}

function cell(content, width, opt, S) {
  const shade = opt.shade ? `<w:shd w:val="clear" w:color="auto" w:fill="${opt.shade}"/>` : '';
  const body = String(content == null ? '' : content).split('\n\n').map((t) =>
    para(t, { align: opt.align || 'both', after: 40, line: 240, size: opt.size || S.tableSize, run: opt.run }, S)
  ).join('') || para('', { after: 40 }, S);
  return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/>${shade}<w:vAlign w:val="top"/></w:tcPr>${body}</w:tc>`;
}

function table(b, S) {
  const cols = (b.header && b.header.length) || ((b.rows && b.rows[0]) || []).length || 1;
  const total = 9354; // ширина полосы набора при полях 3/1.5 см
  const ws = (b.widths && b.widths.length === cols) ? b.widths : Array(cols).fill(1);
  const sum = ws.reduce((a, c) => a + c, 0);
  const widths = ws.map((w) => Math.round(total * w / sum));
  const grid = widths.map((w) => `<w:gridCol w:w="${w}"/>`).join('');
  const borders = b.noBorders
    ? '<w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders>'
    : '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map((s) => `<w:${s} w:val="single" w:sz="6" w:space="0" w:color="333333"/>`).join('') + '</w:tblBorders>';
  const head = b.header && b.header.length
    ? `<w:tr><w:trPr><w:tblHeader/></w:trPr>${b.header.map((h, i) => cell(h, widths[i], { shade: 'EFEFEF', align: 'center', run: { b: true } }, S)).join('')}</w:tr>`
    : '';
  const rows = (b.rows || []).map((r) => `<w:tr>${widths.map((w, i) => cell(r[i], w, { align: b.align || 'left' }, S)).join('')}</w:tr>`).join('');
  return `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/>${borders}<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="60" w:type="dxa"/><w:left w:w="90" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${head}${rows}</w:tbl><w:p><w:pPr><w:spacing w:after="80"/></w:pPr></w:p>`;
}

function signatures(b, S) {
  const left = b.left || {}; const right = b.right || {};
  const line = (s) => [
    '**' + (s.role || '') + '**',
    s.company || '',
    s.title || '',
    '',
    '_____________________ / ' + (s.person || '________________') + ' /',
    s.stamp === false ? '' : 'М.П.',
  ].filter((x) => x !== null && x !== undefined).join('\n\n');
  return table({ noBorders: true, rows: [[line(left), line(right)]], widths: [1, 1], align: 'left' }, S);
}

function block(b, S) {
  if (typeof b === 'string') return para(b, { firstLine: S.indent }, S);
  switch (b.t) {
    case 'h1': return para(b.text, { align: 'center', run: { b: true, caps: !!b.caps }, size: S.size + 4, after: 160, before: 120, keepNext: true }, S);
    case 'h2': return para(b.text, { align: b.align || 'left', run: { b: true }, after: 120, before: 160, keepNext: true }, S);
    case 'center': return para(b.text, { align: 'center', after: b.after == null ? 120 : b.after }, S);
    case 'right': return para(b.text, { align: 'right' }, S);
    case 'small': return para(b.text, { size: S.size - 4, align: b.align || 'left' }, S);
    case 'li': return para('— ' + b.text, { left: 397, hanging: 227 }, S);
    case 'num': return para(`${b.n}. ${b.text}`, { left: 567, hanging: 567 }, S);
    case 'clause': return para(`${b.n} ${b.text}`, { left: 567, hanging: 567 }, S);
    case 'table': return table(b, S);
    case 'sign': return signatures(b, S);
    case 'pagebreak': return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>';
    case 'spacer': return `<w:p><w:pPr><w:spacing w:after="${b.h || 200}"/></w:pPr></w:p>`;
    case 'p':
    default: return para(b.text, { firstLine: b.noIndent ? 0 : S.indent, align: b.align || 'both' }, S);
  }
}

function headerXml(doc, S, hasLogo) {
  const c = doc.company || {};
  if (!c.name && !hasLogo) return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>';
  const cx = doc.logoCx || 990000, cy = doc.logoCy || 330000;
  const logo = hasLogo
    ? `<w:p><w:pPr><w:jc w:val="left"/><w:spacing w:after="40"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="1" name="logo"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="logo"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
    : '';
  const lines = [c.name, c.requisites, c.contacts].filter(Boolean);
  const body = [
    logo,
    lines.length ? para(lines[0], { align: 'left', run: { b: true }, size: S.size, after: 0 }, S) : '',
    ...lines.slice(1).map((l) => para(l, { align: 'left', size: S.size - 4, after: 0 }, S)),
    lines.length ? para('', { border: true, after: 120 }, S) : '',
  ].join('');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">${body || '<w:p/>'}</w:hdr>`;
}

function footerXml(doc, S) {
  const note = doc.footerNote ? para(doc.footerNote, { align: 'left', size: S.size - 6, after: 0 }, S) : '';
  const f = esc(S.font);
  const page = `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:after="0"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}"/><w:sz w:val="${S.size - 4}"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}"/><w:sz w:val="${S.size - 4}"/></w:rPr><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="${f}" w:hAnsi="${f}"/><w:sz w:val="${S.size - 4}"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${note}${page}</w:ftr>`;
}

// doc = { title, subtitle, company:{name,requisites,contacts}, logo:{data(base64),ext}, blocks:[], footerNote }
function buildDocx(doc = {}) {
  const S = {
    font: doc.font || 'Times New Roman',
    size: Math.round((doc.fontSize || 12) * 2),
    tableSize: Math.round((doc.tableFontSize || 11) * 2),
    indent: doc.indent === 0 ? 0 : TW.cm(1.25),
  };
  const hasLogo = !!(doc.logo && doc.logo.data);
  const hasHead = hasLogo || !!(doc.company && doc.company.name);
  const head = [];
  if (doc.title) head.push(block({ t: 'h1', text: doc.title, caps: doc.titleCaps !== false }, S));
  if (doc.subtitle) head.push(block({ t: 'center', text: doc.subtitle, after: 200 }, S));

  const body = (doc.blocks || []).map((b) => block(b, S)).join('');
  const sect = `<w:sectPr>${hasHead ? '<w:headerReference w:type="default" r:id="rId6"/>' : ''}<w:footerReference w:type="default" r:id="rId7"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="${TW.cm(2)}" w:right="${TW.cm(1.5)}" w:bottom="${TW.cm(2)}" w:left="${TW.cm(3)}" w:header="${TW.cm(1)}" w:footer="${TW.cm(1)}" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr>`;

  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${head.join('')}${body}${sect}</w:body></w:document>`;

  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="${esc(S.font)}" w:hAnsi="${esc(S.font)}" w:cs="${esc(S.font)}"/><w:sz w:val="${S.size}"/><w:szCs w:val="${S.size}"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="276" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style></w:styles>`;

  const rels = [
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    '<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
  ].join('');

  const types = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="jpg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`;

  const core = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${esc(doc.title || 'Документ')}</dc:title><dc:creator>${esc((doc.company && doc.company.name) || 'Инженер')}</dc:creator><cp:lastModifiedBy>${esc((doc.company && doc.company.name) || 'Инженер')}</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString().slice(0, 19)}Z</dcterms:created></cp:coreProperties>`;

  const entries = [
    { name: '[Content_Types].xml', data: types },
    { name: '_rels/.rels', data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>' },
    { name: 'docProps/core.xml', data: core },
    { name: 'word/document.xml', data: document },
    { name: 'word/styles.xml', data: styles },
    { name: 'word/header1.xml', data: headerXml(doc, S, hasLogo) },
    { name: 'word/footer1.xml', data: footerXml(doc, S) },
    { name: 'word/_rels/document.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rels}</Relationships>` },
  ];
  if (hasLogo) {
    const ext = (doc.logo.ext || 'png').toLowerCase().replace(/[^a-z]/g, '') || 'png';
    entries.push({ name: 'word/_rels/header1.xml.rels', data: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/logo.${ext}"/></Relationships>` });
    entries.push({ name: `word/media/logo.${ext}`, data: Buffer.from(doc.logo.data, 'base64'), store: true });
  }
  return zip(entries);
}

module.exports = { buildDocx };
