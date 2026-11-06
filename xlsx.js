/* xlsx.js: a small reader and writer for Excel .xlsx files, written for this app.
   No outside code and no network: it unpacks the file in memory with the browser's own
   DecompressionStream, reads the sheets, and writes new files as an uncompressed zip. */
const XLSXLite = (() => {
  'use strict';
  const td = new TextDecoder('utf-8');
  const te = new TextEncoder();
  const u16 = (b, o) => b[o] | (b[o + 1] << 8);
  const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
  const fail = code => { const e = new Error(code); e.code = code; return e; };

  /* ---------- reading ---------- */
  async function inflateRaw(bytes) {
    if (typeof DecompressionStream !== 'function') throw fail('nodecompress');
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
    return new Uint8Array(await new Response(stream).arrayBuffer());
  }
  function unzip(input) {
    const b = input instanceof Uint8Array ? input : new Uint8Array(input);
    if (b.length < 22 || u32(b, 0) !== 0x04034b50) throw fail('notzip');
    let eocd = -1;
    for (let i = b.length - 22; i >= 0 && i >= b.length - 65558; i--) if (u32(b, i) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw fail('notzip');
    const count = u16(b, eocd + 10);
    let off = u32(b, eocd + 16);
    const entries = new Map();
    for (let k = 0; k < count; k++) {
      if (off + 46 > b.length || u32(b, off) !== 0x02014b50) throw fail('badzip');
      const nlen = u16(b, off + 28), xlen = u16(b, off + 30), clen = u16(b, off + 32);
      const name = td.decode(b.subarray(off + 46, off + 46 + nlen)).replace(/\\/g, '/');
      entries.set(name, { flags: u16(b, off + 8), method: u16(b, off + 10), csize: u32(b, off + 20), lho: u32(b, off + 42) });
      off += 46 + nlen + xlen + clen;
    }
    return { b, entries };
  }
  async function entryText(z, name) {
    let e = z.entries.get(name);
    if (!e) { const low = name.toLowerCase(); for (const [k, v] of z.entries) if (k.toLowerCase() === low) { e = v; break; } }
    if (!e) return null;
    if (e.flags & 1) throw fail('encrypted');
    if (e.csize === 0xFFFFFFFF || e.lho === 0xFFFFFFFF) throw fail('toolarge');
    const b = z.b, o = e.lho;
    if (u32(b, o) !== 0x04034b50) throw fail('badzip');
    const start = o + 30 + u16(b, o + 26) + u16(b, o + 28);
    const data = b.subarray(start, start + e.csize);
    if (e.method === 0) return td.decode(data);
    if (e.method === 8) return td.decode(await inflateRaw(data));
    throw fail('method');
  }
  const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  function unesc(s) {
    if (s.indexOf('&') < 0) return s;
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (m, e) => {
      if (e[0] === '#') {
        const cp = e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
        return cp > 0 && cp < 0x110000 ? String.fromCodePoint(cp) : m;
      }
      return ENT[e] != null ? ENT[e] : m;
    });
  }
  function attrs(s) {
    const o = {};
    if (s) s.replace(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g, (m, k, a, b2) => { o[k] = unesc(a != null ? a : b2); return m; });
    return o;
  }
  const unx = s => s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, hex) => String.fromCharCode(parseInt(hex, 16)));
  function textOf(xml) {
    const clean = xml.replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g, '');
    let out = '';
    clean.replace(/<(?:\w+:)?t\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?t>)/g, (m, inner) => { if (inner) out += inner; return m; });
    return unx(unesc(out));
  }
  function parseRels(xml) {
    const out = [];
    if (xml) xml.replace(/<(?:\w+:)?Relationship\b([^>]*?)\/?>/g, (m, a) => { out.push(attrs(a)); return m; });
    return out;
  }
  function normPath(p) {
    const parts = [];
    for (const seg of p.split('/')) { if (seg === '..') parts.pop(); else if (seg && seg !== '.') parts.push(seg); }
    return parts.join('/');
  }
  const DATE_IDS = new Set([14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 57, 58]);
  function isDateFormat(id, code) {
    if (DATE_IDS.has(id)) return true;
    if (!code) return false;
    const c = code.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/[_*]./g, '').replace(/\[[^\]]*\]/g, '').replace(/general/gi, '');
    return /[dy]/i.test(c) || /m{3,}/i.test(c);
  }
  function colIndex(ref) {
    let n = 0;
    for (let i = 0; i < ref.length; i++) {
      const c = ref.charCodeAt(i) | 32;
      if (c >= 97 && c <= 122) n = n * 26 + (c - 96); else break;
    }
    return n - 1;
  }
  function serialToIso(n, date1904) {
    const days = Math.floor(n + 1e-7) + (date1904 ? 1462 : 0) - 25569;
    const d = new Date(days * 86400000);
    return isNaN(d.getTime()) ? '' : d.toISOString().slice(0, 10);
  }
  function parseSheet(xml, sst, xfDate, date1904) {
    const rows = [];
    const sd = /<(?:\w+:)?sheetData\b[^>]*>([\s\S]*?)<\/(?:\w+:)?sheetData>/.exec(xml);
    if (!sd) return rows;
    const rowRe = /<(?:\w+:)?row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?row>)/g;
    let rm;
    while ((rm = rowRe.exec(sd[1]))) {
      if (!rm[2]) continue;
      const row = [];
      const cellRe = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
      let cm, next = 0, any = false;
      while ((cm = cellRe.exec(rm[2]))) {
        const ca = attrs(cm[1]);
        const ci = ca.r ? colIndex(ca.r) : next;
        next = ci + 1;
        const body = cm[2] || '';
        const vm = /<(?:\w+:)?v\b[^>]*>([\s\S]*?)<\/(?:\w+:)?v>/.exec(body);
        const v = vm ? unesc(vm[1]) : null;
        const t = ca.t || 'n';
        let val = '';
        if (t === 's') val = v != null && sst[+v] != null ? sst[+v] : '';
        else if (t === 'inlineStr') { const im = /<(?:\w+:)?is\b[^>]*>([\s\S]*?)<\/(?:\w+:)?is>/.exec(body); val = im ? textOf(im[1]) : ''; }
        else if (t === 'str') val = v != null ? unx(v) : '';
        else if (t === 'b') val = v === '1' || v === 'true';
        else if (t === 'e') val = '';
        else if (t === 'd') val = v ? String(v).slice(0, 10) : '';
        else if (v != null && v !== '') {
          const n = Number(v);
          if (!isFinite(n)) val = v;
          else if (xfDate[+(ca.s || 0)]) val = serialToIso(n, date1904);
          else val = n;
        }
        if (ci >= 0 && ci < 16384) { row[ci] = val; if (val !== '') any = true; }
      }
      if (!any) continue;
      for (let j = 0; j < row.length; j++) if (row[j] === undefined) row[j] = '';
      rows.push(row);
    }
    return rows;
  }
  async function read(input) {
    const z = unzip(input);
    let wbPath = 'xl/workbook.xml';
    const od = parseRels(await entryText(z, '_rels/.rels')).find(r => /\/officeDocument$/.test(r.Type || ''));
    if (od && od.Target) wbPath = normPath(od.Target.replace(/^\//, ''));
    const wbDir = wbPath.includes('/') ? wbPath.slice(0, wbPath.lastIndexOf('/') + 1) : '';
    const wbXml = await entryText(z, wbPath);
    if (!wbXml) throw fail('noworkbook');
    const rels = parseRels(await entryText(z, wbDir + '_rels/' + wbPath.slice(wbDir.length) + '.rels'));
    const resolve = t => (t.startsWith('/') ? normPath(t.slice(1)) : normPath(wbDir + t));
    const date1904 = /<(?:\w+:)?workbookPr\b[^>]*\bdate1904\s*=\s*["'](?:1|true)["']/i.test(wbXml);
    const sstRel = rels.find(r => /\/sharedStrings$/.test(r.Type || ''));
    const stRel = rels.find(r => /\/styles$/.test(r.Type || ''));
    const sst = [];
    const sstXml = await entryText(z, sstRel ? resolve(sstRel.Target) : wbDir + 'sharedStrings.xml');
    if (sstXml) sstXml.replace(/<(?:\w+:)?si\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:\w+:)?si>)/g, (m, inner) => { sst.push(inner ? textOf(inner) : ''); return m; });
    const stXml = (await entryText(z, stRel ? resolve(stRel.Target) : wbDir + 'styles.xml')) || '';
    const fmts = {};
    stXml.replace(/<(?:\w+:)?numFmt\b([^>]*?)\/?>/g, (m, a) => { const at = attrs(a); if (at.numFmtId != null) fmts[+at.numFmtId] = at.formatCode || ''; return m; });
    const xfDate = [];
    const cx = /<(?:\w+:)?cellXfs\b[^>]*>([\s\S]*?)<\/(?:\w+:)?cellXfs>/.exec(stXml);
    if (cx) cx[1].replace(/<(?:\w+:)?xf\b([^>]*?)\/?>/g, (m, a) => { const id = +(attrs(a).numFmtId || 0); xfDate.push(isDateFormat(id, fmts[id])); return m; });
    const tags = [];
    wbXml.replace(/<(?:\w+:)?sheet\b([^>]*?)\/?>/g, (m, a) => { tags.push(attrs(a)); return m; });
    const sheets = [];
    for (const at of tags) {
      const ridKey = Object.keys(at).find(k => /(^|:)id$/i.test(k) && k !== 'sheetId');
      const rel = rels.find(r => r.Id === at[ridKey]);
      if (!rel || !/\/worksheet$/.test(rel.Type || '')) continue;
      const xml = await entryText(z, resolve(rel.Target));
      if (xml == null) continue;
      sheets.push({ name: at.name || 'Sheet' + (sheets.length + 1), hidden: at.state === 'hidden' || at.state === 'veryHidden', rows: parseSheet(xml, sst, xfDate, date1904) });
    }
    return { sheets, date1904 };
  }

  /* ---------- writing ---------- */
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
    return t;
  })();
  function crc32(buf) { let c = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) c = CRC[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
  function zipStore(files) {
    const now = new Date();
    const dosTime = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const dosDate = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    const parts = [], central = [];
    let offset = 0;
    for (const f of files) {
      const name = te.encode(f.name), data = f.data, crc = crc32(data);
      const lh = new Uint8Array(30 + name.length), lv = new DataView(lh.buffer);
      lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, 0x0800, true); lv.setUint16(8, 0, true);
      lv.setUint16(10, dosTime, true); lv.setUint16(12, dosDate, true); lv.setUint32(14, crc, true);
      lv.setUint32(18, data.length, true); lv.setUint32(22, data.length, true); lv.setUint16(26, name.length, true); lv.setUint16(28, 0, true);
      lh.set(name, 30);
      parts.push(lh, data);
      const ch = new Uint8Array(46 + name.length), cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0x0800, true); cv.setUint16(10, 0, true);
      cv.setUint16(12, dosTime, true); cv.setUint16(14, dosDate, true); cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true); cv.setUint32(24, data.length, true); cv.setUint16(28, name.length, true);
      cv.setUint32(42, offset, true);
      ch.set(name, 46);
      central.push(ch);
      offset += lh.length + data.length;
    }
    const cdSize = central.reduce((s, c) => s + c.length, 0);
    const end = new Uint8Array(22), ev = new DataView(end.buffer);
    ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, files.length, true); ev.setUint16(10, files.length, true);
    ev.setUint32(12, cdSize, true); ev.setUint32(16, offset, true);
    const out = new Uint8Array(offset + cdSize + 22);
    let p = 0;
    for (const c of parts.concat(central, [end])) { out.set(c, p); p += c.length; }
    return out;
  }
  const esc = s => String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, '');
  function colName(i) { let s = ''; for (let n = i + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; }
  const STYLE = { text: 0, date: 1, money: 2, header: 3, percent: 4, note: 5, decimal: 6, title: 7 };
  function cellXml(ref, c) {
    if (c == null || c === '') return '';
    if (typeof c !== 'object') c = { v: c };
    const s = c.style && STYLE[c.style] ? ' s="' + STYLE[c.style] + '"' : '';
    if (c.f) {
      const v = c.v;
      if (typeof v === 'number' && isFinite(v)) return '<c r="' + ref + '"' + s + '><f>' + esc(c.f) + '</f><v>' + v + '</v></c>';
      if (typeof v === 'string') return '<c r="' + ref + '"' + s + ' t="str"><f>' + esc(c.f) + '</f><v>' + esc(v) + '</v></c>';
      return '<c r="' + ref + '"' + s + '><f>' + esc(c.f) + '</f></c>';
    }
    if (typeof c.v === 'number') return isFinite(c.v) ? '<c r="' + ref + '"' + s + '><v>' + c.v + '</v></c>' : '';
    if (typeof c.v === 'boolean') return '<c r="' + ref + '"' + s + ' t="b"><v>' + (c.v ? 1 : 0) + '</v></c>';
    if (c.v == null || c.v === '') return '';
    return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(c.v) + '</t></is></c>';
  }
  const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
  function sheetXml(sh, active) {
    const rows = sh.rows.map((r, i) => {
      const cells = (r || []).map((c, j) => cellXml(colName(j) + (i + 1), c)).join('');
      return cells ? '<row r="' + (i + 1) + '">' + cells + '</row>' : '';
    }).join('');
    const sel = active ? ' tabSelected="1"' : '';
    const view = sh.freeze
      ? '<sheetView workbookViewId="0"' + sel + '><pane ySplit="' + sh.freeze + '" topLeftCell="A' + (sh.freeze + 1) + '" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A' + (sh.freeze + 1) + '" sqref="A' + (sh.freeze + 1) + '"/></sheetView>'
      : '<sheetView workbookViewId="0"' + sel + '/>';
    const cols = sh.cols && sh.cols.length ? '<cols>' + sh.cols.map((w, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>').join('') + '</cols>' : '';
    return HEAD + '<worksheet ' + NS + '><sheetViews>' + view + '</sheetViews><sheetFormatPr defaultRowHeight="15"/>' + cols + '<sheetData>' + rows + '</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
  }
  const STYLES_XML = HEAD + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="3"><numFmt numFmtId="164" formatCode="d mmm yyyy"/><numFmt numFmtId="165" formatCode="#,##0.00"/><numFmt numFmtId="166" formatCode="0.0%"/></numFmts>' +
    '<fonts count="4"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font>' +
    '<font><i/><sz val="10"/><color rgb="FF5A6A70"/><name val="Calibri"/><family val="2"/></font><font><b/><sz val="14"/><name val="Calibri"/><family val="2"/></font></fonts>' +
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="8">' +
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '<xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
    '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  /* book: { sheets: [{ name, rows: [[cell]], cols?: [widths], freeze?: rows }], active?: index }
     cell: number | string | { v, style?: 'date'|'money'|'header'|'percent'|'note'|'decimal'|'title', f?: formula } */
  function write(book) {
    const n = book.sheets.length, active = book.active || 0;
    const enc = s => te.encode(s);
    const files = [];
    let types = HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
    for (let i = 1; i <= n; i++) types += '<Override PartName="/xl/worksheets/sheet' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
    types += '</Types>';
    files.push({ name: '[Content_Types].xml', data: enc(types) });
    files.push({ name: '_rels/.rels', data: enc(HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>') });
    let wb = HEAD + '<workbook ' + NS + '><workbookPr/><bookViews><workbookView activeTab="' + active + '"/></bookViews><sheets>';
    let rels = HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    book.sheets.forEach((sh, i) => {
      wb += '<sheet name="' + esc(String(sh.name).slice(0, 31)) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
      rels += '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>';
    });
    wb += '</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';
    rels += '<Relationship Id="rId' + (n + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    files.push({ name: 'xl/workbook.xml', data: enc(wb) });
    files.push({ name: 'xl/_rels/workbook.xml.rels', data: enc(rels) });
    files.push({ name: 'xl/styles.xml', data: enc(STYLES_XML) });
    book.sheets.forEach((sh, i) => files.push({ name: 'xl/worksheets/sheet' + (i + 1) + '.xml', data: enc(sheetXml(sh, i === active)) }));
    return zipStore(files);
  }

  return { read, write, serialToIso };
})();
if (typeof module !== 'undefined') module.exports = { XLSXLite };
