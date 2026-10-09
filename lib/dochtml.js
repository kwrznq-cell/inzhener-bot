// Тот же документ, что уходит в .docx, но в виде HTML — для предпросмотра,
// печати в PDF и сохранения картинкой.
const esc = (s) => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function inline(text) {
  return esc(text)
    .replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>')
    .replace(/==([^=]+)==/g, '<mark>$1</mark>')
    .replace(/__([^_]+)__/g, '<u>$1</u>')
    .replace(/\*([^*\n]+)\*/g, '<i>$1</i>')
    .replace(/\n/g, '<br>');
}

function cellHtml(v) {
  return String(v == null ? '' : v).split('\n\n').map((p) => `<p>${inline(p)}</p>`).join('');
}

function block(b) {
  if (typeof b === 'string') return `<p class="ind">${inline(b)}</p>`;
  switch (b.t) {
    case 'h1': return `<h1>${inline(b.text)}</h1>`;
    case 'h2': return `<h2>${inline(b.text)}</h2>`;
    case 'center': return `<p class="c">${inline(b.text)}</p>`;
    case 'right': return `<p class="r">${inline(b.text)}</p>`;
    case 'small': return `<p class="s">${inline(b.text)}</p>`;
    case 'li': return `<p class="li">— ${inline(b.text)}</p>`;
    case 'num': return `<p class="cl"><span class="n">${esc(b.n)}.</span> ${inline(b.text)}</p>`;
    case 'clause': return `<p class="cl"><span class="n">${esc(b.n)}</span> ${inline(b.text)}</p>`;
    case 'table': {
      const head = b.header && b.header.length ? `<thead><tr>${b.header.map((h) => `<th>${inline(h)}</th>`).join('')}</tr></thead>` : '';
      const ws = b.widths && b.widths.length ? b.widths : null;
      const cols = ws ? `<colgroup>${ws.map((w) => `<col style="width:${(100 * w / ws.reduce((a, c) => a + c, 0)).toFixed(2)}%">`).join('')}</colgroup>` : '';
      const rows = (b.rows || []).map((r) => `<tr>${r.map((c) => `<td>${cellHtml(c)}</td>`).join('')}</tr>`).join('');
      return `<table>${cols}${head}<tbody>${rows}</tbody></table>`;
    }
    case 'sign': {
      const one = (s = {}) => `<div class="sg"><p><b>${esc(s.role || '')}</b></p><p>${esc(s.company || '')}</p><p>${esc(s.title || '')}</p><p class="ln">_____________________ / ${esc(s.person || '________________')} /</p>${s.stamp === false ? '' : '<p>М.П.</p>'}</div>`;
      return `<div class="signs">${one(b.left)}${one(b.right)}</div>`;
    }
    case 'pagebreak': return '<div class="pb"></div>';
    case 'spacer': return '<div class="sp"></div>';
    case 'p':
    default: return `<p class="${b.noIndent ? '' : 'ind'}${b.align === 'center' ? ' c' : ''}">${inline(b.text)}</p>`;
  }
}

function buildHtml(doc = {}) {
  const c = doc.company || {};
  const logo = doc.logo && doc.logo.data ? `<img class="logo" src="data:image/${doc.logo.ext === 'jpg' ? 'jpeg' : (doc.logo.ext || 'png')};base64,${doc.logo.data}" alt="">` : '';
  const head = (c.name || logo)
    ? `<header>${logo}<div class="co"><b>${esc(c.name || '')}</b>${c.requisites ? `<span>${esc(c.requisites)}</span>` : ''}${c.contacts ? `<span>${esc(c.contacts)}</span>` : ''}</div></header>`
    : '';
  const title = doc.title ? `<h1 class="title">${esc(doc.title)}</h1>` : '';
  const sub = doc.subtitle ? `<p class="c sub">${inline(doc.subtitle)}</p>` : '';
  const body = (doc.blocks || []).map(block).join('');
  const foot = doc.footerNote ? `<footer>${esc(doc.footerNote)}</footer>` : '';
  return `<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>${esc(doc.title || 'Документ')}</title>
<style>
  @page { size: A4; margin: 20mm 15mm 20mm 30mm; }
  html,body{margin:0;background:#fff;color:#111}
  body{font-family:"Times New Roman",Georgia,serif;font-size:12pt;line-height:1.45}
  .sheet{width:210mm;min-height:297mm;box-sizing:border-box;padding:20mm 15mm 20mm 30mm;margin:0 auto;background:#fff}
  header{display:flex;gap:12px;align-items:flex-start;border-bottom:1px solid #999;padding-bottom:6px;margin-bottom:18px}
  header .logo{max-height:16mm;max-width:45mm;object-fit:contain}
  header .co{display:flex;flex-direction:column;font-size:10pt;line-height:1.35}
  header .co b{font-size:12pt}
  h1.title{font-size:14pt;text-align:center;margin:10px 0 6px;text-transform:uppercase;letter-spacing:.02em}
  h1{font-size:13pt;text-align:center;margin:14px 0 8px}
  h2{font-size:12pt;margin:14px 0 6px}
  p{margin:0 0 6px;text-align:justify}
  p.ind{text-indent:1.25cm}
  p.c{text-align:center}p.r{text-align:right}p.s{font-size:10pt}
  p.li{margin-left:.7cm;text-indent:-.4cm}
  p.cl{margin-left:1cm;text-indent:-1cm}
  p.cl .n{display:inline-block;min-width:1cm}
  mark{background:#ffef60;padding:0 1px}
  table{width:100%;border-collapse:collapse;margin:8px 0 12px;font-size:11pt}
  th,td{border:1px solid #333;padding:4px 6px;vertical-align:top;text-align:left}
  th{background:#efefef;text-align:center;font-weight:bold}
  td p,th p{margin:0 0 4px;text-align:left}
  .signs{display:flex;gap:24px;margin-top:26px}
  .signs .sg{flex:1}
  .signs p{margin:0 0 4px;text-align:left}
  .signs .ln{margin-top:16px}
  .sp{height:14px}
  .pb{page-break-after:always;height:0}
  footer{margin-top:24px;border-top:1px solid #bbb;padding-top:6px;font-size:9pt;color:#444}
  @media print{ .sheet{width:auto;min-height:0;padding:0;margin:0} body{font-size:12pt} }
</style></head><body><div class="sheet">${head}${title}${sub}${body}${foot}</div></body></html>`;
}

module.exports = { buildHtml };
