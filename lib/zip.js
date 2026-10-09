// Минимальный ZIP-упаковщик (без зависимостей) — нужен для сборки .docx
const zlib = require('zlib');

const CRC = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();
function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

// entries: [{ name, data: Buffer|string, store?: true }]
function zip(entries) {
  const files = [];
  const chunks = [];
  let offset = 0;

  for (const e of entries) {
    const name = Buffer.from(e.name, 'utf8');
    const raw = Buffer.isBuffer(e.data) ? e.data : Buffer.from(e.data, 'utf8');
    const store = !!e.store;
    const body = store ? raw : zlib.deflateRawSync(raw, { level: 9 });
    const crc = crc32(raw);

    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);            // version
    local.writeUInt16LE(0x0800, 6);        // utf-8 names
    local.writeUInt16LE(store ? 0 : 8, 8); // method
    local.writeUInt16LE(0, 10);            // time
    local.writeUInt16LE(0x21, 12);         // date (1996-01-01)
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(body.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(name.length, 26);
    local.writeUInt16LE(0, 28);

    chunks.push(local, name, body);
    files.push({ name, crc, csize: body.length, size: raw.length, offset, store });
    offset += local.length + name.length + body.length;
  }

  const central = [];
  let clen = 0;
  for (const f of files) {
    const h = Buffer.alloc(46);
    h.writeUInt32LE(0x02014b50, 0);
    h.writeUInt16LE(20, 4);
    h.writeUInt16LE(20, 6);
    h.writeUInt16LE(0x0800, 8);
    h.writeUInt16LE(f.store ? 0 : 8, 10);
    h.writeUInt16LE(0, 12);
    h.writeUInt16LE(0x21, 14);
    h.writeUInt32LE(f.crc, 16);
    h.writeUInt32LE(f.csize, 20);
    h.writeUInt32LE(f.size, 24);
    h.writeUInt16LE(f.name.length, 28);
    h.writeUInt32LE(0, 38);                // external attrs
    h.writeUInt32LE(f.offset, 42);
    central.push(h, f.name);
    clen += h.length + f.name.length;
  }

  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(clen, 12);
  end.writeUInt32LE(offset, 16);

  return Buffer.concat([...chunks, ...central, end]);
}

// читаем .docx: имя -> Buffer
function unzip(buf) {
  const out = {};
  let p = buf.length - 22;
  while (p >= 0 && buf.readUInt32LE(p) !== 0x06054b50) p--;
  if (p < 0) throw new Error('Не похоже на .docx (нет конца архива)');
  const count = buf.readUInt16LE(p + 10);
  let cp = buf.readUInt32LE(p + 16);
  for (let i = 0; i < count; i++) {
    if (buf.readUInt32LE(cp) !== 0x02014b50) break;
    const method = buf.readUInt16LE(cp + 10);
    const csize = buf.readUInt32LE(cp + 20);
    const nlen = buf.readUInt16LE(cp + 28);
    const elen = buf.readUInt16LE(cp + 30);
    const clen = buf.readUInt16LE(cp + 32);
    const lo = buf.readUInt32LE(cp + 42);
    const name = buf.slice(cp + 46, cp + 46 + nlen).toString('utf8');
    const ln = buf.readUInt16LE(lo + 26), le = buf.readUInt16LE(lo + 28);
    const start = lo + 30 + ln + le;
    const body = buf.slice(start, start + csize);
    out[name] = method === 0 ? body : zlib.inflateRawSync(body);
    cp += 46 + nlen + elen + clen;
  }
  return out;
}

module.exports = { zip, unzip, crc32 };
