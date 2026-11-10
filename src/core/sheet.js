'use strict';
// Workbook reading and writing on top of SheetJS (passed in as XLSX, so tests and the worker bundle share this code).
// readWorkbook returns a plain model the extension host can turn into any table text:
//   { sheets: [{ name, ref, rows: string[][], fmt: {"r,c": text}, links: {"r,c": url}, merges: [[r1,c1,r2,c2]], truncated }] }
// rows hold values as stored (numbers in plain form, dates as ISO 8601). fmt, links and merges are what the Pro tool
// "Keep Excel formats, links and merged cells" adds. Every read is capped (file size, zip bomb guard, rows, columns, cells).
const T = require('./table');
const { checkZip, isZip } = require('./zipcheck');

const SHEET_LIMITS = { maxFileBytes: 100 * 1024 * 1024, maxCsvBytes: 60 * 1024 * 1024, maxRows: 100000, maxCols: 2000, maxCells: 2000000, maxSheets: 255 };
const WORKBOOK_EXT = /\.(xlsx|xlsm|xlsb|xls|ods|fods)$/i;
const TEXT_EXT = /\.(csv|tsv|tab|txt)$/i;

const PASSWORD = 'This workbook is password-protected. Open it in Excel or LibreOffice, remove the password, save it, and open it here again.';

class SheetError extends Error { constructor(m, code = 'SHEET') { super(m); this.code = code; } }

const pad2 = (n) => String(n).padStart(2, '0');

function plainNumber(v) {
  if (!Number.isFinite(v)) return String(v);
  if (Number.isInteger(v)) return Math.abs(v) < 1e21 ? String(v) : v.toExponential();
  const s = String(Number(v.toPrecision(15)));
  return s;
}

function isoDate(XLSX, v) {
  const d = XLSX.SSF.parse_date_code(v);
  if (!d) return plainNumber(v);
  const date = `${String(d.y).padStart(4, '0')}-${pad2(d.m)}-${pad2(d.d)}`;
  const hasTime = d.H || d.M || d.S;
  const time = `${pad2(d.H)}:${pad2(d.M)}${d.S ? ':' + pad2(d.S) : ''}`;
  if (v < 1 && v >= 0) return time; // a time without a date
  return hasTime ? `${date} ${time}` : date;
}

/** A cell as stored: numbers plain, dates ISO, booleans TRUE/FALSE, errors as Excel shows them. */
function rawText(XLSX, cell) {
  if (!cell) return '';
  switch (cell.t) {
    case 'n': return cell.z && cell.z !== 'General' && XLSX.SSF.is_date(cell.z) ? isoDate(XLSX, cell.v) : plainNumber(cell.v);
    case 'd': return cell.v instanceof Date ? cell.v.toISOString().replace('T', ' ').replace(/(\.000)?Z$/, '').replace(/ 00:00:00$/, '') : String(cell.v);
    case 'b': return cell.v ? 'TRUE' : 'FALSE';
    case 'e': return cell.w || '#ERROR';
    case 'z': return '';
    default: return cell.v == null ? '' : String(cell.v);
  }
}

function cellAt(ws, r, c, XLSX) {
  if (ws['!data']) { const row = ws['!data'][r]; return row ? row[c] : undefined; }
  return ws[XLSX.utils.encode_cell({ r, c })];
}

/** Reads a workbook buffer (or CSV/TSV text) into the plain model. name picks the reader by extension. */
function readWorkbook(XLSX, buf, name, limits = {}) {
  const L = { ...SHEET_LIMITS, ...limits };
  const ext = (/\.[^./\\]+$/.exec(String(name || '')) || [''])[0].toLowerCase();
  if (TEXT_EXT.test(ext)) {
    if (buf.length > L.maxCsvBytes) throw new SheetError(`The file is ${Math.round(buf.length / 1048576)} MB; CSV files up to ${Math.round(L.maxCsvBytes / 1048576)} MB are read.`, 'TOO_BIG');
    let text;
    if (buf[0] === 0xff && buf[1] === 0xfe) text = Buffer.from(buf).toString('utf16le').slice(1);
    else text = Buffer.from(buf).toString('utf8');
    const p = T.parseDelimited(text, { delimiter: ext === '.tsv' || ext === '.tab' ? '\t' : undefined, limits: { maxChars: L.maxCsvBytes, maxRows: L.maxRows, maxCols: L.maxCols, maxCells: L.maxCells } });
    const rows = p.rows;
    const ref = rows.length ? `A1:${T.colName(Math.max(0, p.columns - 1))}${rows.length}` : 'A1:A1';
    return { kind: 'text', sheets: [{ name: String(name || 'Sheet1').split(/[\\/]/).pop(), ref, rows, fmt: {}, links: {}, merges: [], truncated: '' }] };
  }
  if (buf.length > L.maxFileBytes) throw new SheetError(`The file is ${Math.round(buf.length / 1048576)} MB; workbooks up to ${Math.round(L.maxFileBytes / 1048576)} MB are read.`, 'TOO_BIG');
  if (buf.length === 0) throw new SheetError('The file is empty.', 'EMPTY');
  const zip = isZip(buf);
  if (zip) {
    const z = checkZip(Buffer.from(buf.buffer, buf.byteOffset, buf.byteLength));
    if (/encryption-data/.test(z.manifest || '')) throw new SheetError(PASSWORD, 'PASSWORD');
  }
  else if (/^\.(xlsx|xlsm|ods)$/.test(ext) && !(buf[0] === 0xd0 && buf[1] === 0xcf)) throw new SheetError(`This ${ext} file is not a zip archive, so it is not an Excel or OpenDocument workbook. It may be damaged or renamed.`, 'NOT_WORKBOOK');
  else if (!(buf[0] === 0xd0 && buf[1] === 0xcf) && !(buf[0] === 0x09 && [0, 2, 4, 8].includes(buf[1])) && !((buf[0] === 0xff && buf[1] === 0xfe) || (buf[0] === 0xfe && buf[1] === 0xff))) {
    // not zip, not OLE (.xls), not an old BIFF stream: only text formats are left (SpreadsheetML XML, HTML or CSV saved
    // with an .xls name). Binary data here is not a workbook, so it is refused instead of shown as a garbage sheet.
    const head = buf.subarray(0, 4096);
    let bad = 0;
    for (const b of head) if (b < 9 || (b > 13 && b < 32)) bad++;
    if (bad > head.length * 0.02) throw new SheetError(`This ${ext || 'file'} file holds binary data that is not a workbook Excel or LibreOffice writes. It may be damaged or renamed.`, 'NOT_WORKBOOK');
  }
  let wb;
  try {
    wb = XLSX.read(buf, { type: 'buffer', dense: true, cellNF: true, cellText: true, cellDates: false, cellFormula: false, cellHTML: false, cellStyles: false, sheetRows: L.maxRows + 1, WTF: false });
  } catch (e) {
    if (/password|encrypt/i.test(String(e && e.message))) throw new SheetError(PASSWORD, 'PASSWORD');
    throw new SheetError(`The workbook could not be read: ${String(e && e.message || e).slice(0, 200)}`, 'READ');
  }
  const names = (wb.SheetNames || []).slice(0, L.maxSheets);
  const sheets = [];
  let cells = 0;
  for (const sn of names) {
    const ws = wb.Sheets[sn];
    const s = { name: String(sn), ref: '', rows: [], fmt: {}, links: {}, merges: [], truncated: '' };
    sheets.push(s);
    if (!ws || !ws['!ref']) continue;
    const full = XLSX.utils.decode_range(ws['!ref']);
    const r0 = Math.max(0, full.s.r), c0 = Math.max(0, full.s.c);
    let r1 = full.e.r, c1 = full.e.c;
    // Excel ranges start at A1 for the table; leading empty rows/columns are kept so A1 refs stay right
    const notes = [];
    if (r1 - 0 + 1 > L.maxRows) { r1 = L.maxRows - 1; notes.push(`first ${T.fmtInt(L.maxRows)} rows`); }
    if (c1 - 0 + 1 > L.maxCols) { c1 = L.maxCols - 1; notes.push(`first ${T.fmtInt(L.maxCols)} columns`); }
    const left = L.maxCells - cells;
    const width = c1 + 1;
    if (left <= 0) { s.truncated = 'not read: the workbook cell limit was reached'; continue; }
    if ((r1 + 1) * width > left) { r1 = Math.max(0, Math.floor(left / width) - 1); notes.push(`first ${T.fmtInt(r1 + 1)} rows (cell limit)`); }
    // trim trailing empty rows and columns of the used range
    for (let r = 0; r <= r1; r++) {
      const row = new Array(width).fill('');
      if (r >= r0) {
        for (let c = c0; c <= c1; c++) {
          const cell = cellAt(ws, r, c, XLSX);
          if (!cell) continue;
          const raw = rawText(XLSX, cell);
          row[c] = raw;
          if (cell.w != null && cell.t !== 's' && cell.w !== raw) s.fmt[r + ',' + c] = String(cell.w);
          if (cell.l && cell.l.Target) s.links[r + ',' + c] = String(cell.l.Target).slice(0, 2000);
        }
      }
      s.rows.push(row);
    }
    cells += s.rows.length * width;
    while (s.rows.length && s.rows[s.rows.length - 1].every((v) => v === '')) s.rows.pop();
    let maxC = -1;
    for (const row of s.rows) for (let c = row.length - 1; c > maxC; c--) if (row[c] !== '') { maxC = c; break; }
    for (const row of s.rows) row.length = maxC + 1;
    for (const m of (ws['!merges'] || []).slice(0, 100000)) s.merges.push([m.s.r, m.s.c, m.e.r, m.e.c]);
    s.ref = s.rows.length && maxC >= 0 ? `A1:${T.colName(maxC)}${s.rows.length}` : '';
    if (notes.length || (ws['!fullref'] && ws['!fullref'] !== ws['!ref'])) s.truncated = notes.length ? notes.join(', ') : `first ${T.fmtInt(L.maxRows)} rows`;
  }
  if (!sheets.length) throw new SheetError('The workbook has no sheets.', 'EMPTY');
  return { kind: zip ? 'zip' : 'binary', sheets };
}

/**
 * The cells of one sheet and range as table rows.
 * o: { range: 'A1:D9' | '', formats: false | true (Pro: Excel display text, links, merged cells filled), trim: true }
 */
function sheetRows(sheet, o = {}) {
  const rowsAll = sheet.rows || [];
  const width = rowsAll.reduce((m, r) => Math.max(m, r.length), 0);
  if (!rowsAll.length || !width) return { rows: [], range: '' };
  const bounds = { s: { r: 0, c: 0 }, e: { r: rowsAll.length - 1, c: width - 1 } };
  const rg = T.parseRange(o.range || '', bounds);
  const e = { r: Math.min(rg.e.r, bounds.e.r), c: Math.min(rg.e.c, bounds.e.c) };
  if (rg.s.r > e.r || rg.s.c > e.c) return { rows: [], range: T.rangeName({ s: rg.s, e: rg.e }) };
  let fill = null;
  if (o.formats && sheet.merges && sheet.merges.length) {
    fill = new Map();
    for (const [r1, c1, r2, c2] of sheet.merges) {
      if ((r2 - r1 + 1) * (c2 - c1 + 1) > 1000000) continue;
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) if (r !== r1 || c !== c1) fill.set(r + ',' + c, r1 + ',' + c1);
    }
  }
  const value = (r, c) => {
    const k = r + ',' + c;
    const src = fill && fill.get(k);
    const key = src || k;
    const [rr, cc] = src ? src.split(',').map(Number) : [r, c];
    let t = (rowsAll[rr] && rowsAll[rr][cc]) || '';
    if (o.formats) {
      if (sheet.fmt && sheet.fmt[key] != null) t = sheet.fmt[key];
      if (sheet.links && sheet.links[key]) return { t, href: sheet.links[key] };
    }
    return t;
  };
  const out = [];
  for (let r = rg.s.r; r <= e.r; r++) {
    const row = [];
    for (let c = rg.s.c; c <= e.c; c++) row.push(value(r, c));
    out.push(row);
  }
  if (o.trim !== false) {
    while (out.length && out[out.length - 1].every((v) => T.cellText(v) === '')) out.pop();
  }
  return { rows: out, range: T.rangeName({ s: rg.s, e }) };
}

const numberCell = /^-?\d+(\.\d+)?([eE][-+]?\d+)?$/;
const percentCell = /^-?\d+(\.\d+)?%$/;

/** Writes tables to an .xlsx buffer: [{ name, rows: string[][] }]. Numbers and percentages become numbers. */
function writeXlsx(XLSX, tables) {
  const wb = XLSX.utils.book_new();
  const used = new Set();
  for (let i = 0; i < tables.length; i++) {
    const t = tables[i];
    let base = String(t.name || `Table ${i + 1}`).replace(/[[\]:*?/\\]/g, ' ').replace(/^'+|'+$/g, '').replace(/\s+/g, ' ').trim().slice(0, 31) || `Table ${i + 1}`;
    let name = base, k = 2;
    while (used.has(name.toLowerCase())) { const suf = ` (${k++})`; name = base.slice(0, 31 - suf.length) + suf; }
    used.add(name.toLowerCase());
    const aoa = t.rows.map((r) => r.map((v) => {
      const s = T.cellText(v);
      if (numberCell.test(s) && s.length < 16) return Number(s);
      if (percentCell.test(s)) return { t: 'n', v: Number(s.slice(0, -1)) / 100, z: '0%' };
      return s;
    }));
    const ws = XLSX.utils.aoa_to_sheet(aoa.map((r) => r.map((v) => (v && typeof v === 'object' ? v.v : v))));
    aoa.forEach((r, ri) => r.forEach((v, ci) => { if (v && typeof v === 'object') { const a = XLSX.utils.encode_cell({ r: ri, c: ci }); if (ws[a]) ws[a].z = v.z; } }));
    const widths = [];
    for (const r of t.rows) r.forEach((v, c) => { widths[c] = Math.min(60, Math.max(widths[c] || 6, T.strWidth(T.cellText(v)) + 2)); });
    ws['!cols'] = widths.map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, name);
  }
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', compression: true });
}

module.exports = { readWorkbook, sheetRows, writeXlsx, rawText, isoDate, plainNumber, SheetError, SHEET_LIMITS, WORKBOOK_EXT, TEXT_EXT };
