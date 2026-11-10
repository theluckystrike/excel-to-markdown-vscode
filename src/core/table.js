'use strict';
// Plain table functions with no vscode dependency: read TSV/CSV text (what Excel, Google Sheets, Numbers and
// LibreOffice put on the clipboard), write Markdown, HTML, AsciiDoc, Jira and CSV tables, read Markdown tables back,
// and measure text width for alignment (CJK and emoji take two columns). Every reader is bounded by LIMITS.

const LIMITS = { maxChars: 10 * 1024 * 1024, maxRows: 200000, maxCols: 2000, maxCells: 2000000 };

class TableError extends Error {
  constructor(message, code) { super(message); this.code = code; }
}

const fmtInt = (n) => String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
const mb = (n) => (n / 1048576).toFixed(1);

// ---------------- delimited text ----------------

/** Picks the delimiter of CSV/TSV text: a tab wins; else the one of , ; | that splits the first lines evenly. */
function detectDelimiter(text) {
  const sample = text.slice(0, 65536);
  if (sample.indexOf('\t') >= 0) return '\t';
  const lines = [];
  let cur = '', q = false;
  for (let i = 0; i < sample.length && lines.length < 20; i++) {
    const ch = sample[i];
    if (ch === '"') q = !q;
    if (!q && (ch === '\n' || ch === '\r')) { if (cur) lines.push(cur); cur = ''; continue; }
    if (!q) cur += ch; else if (ch !== '"') cur += ' ';
  }
  if (cur && lines.length < 20) lines.push(cur);
  let best = '\t', bestScore = 0;
  for (const d of [',', ';', '|']) {
    const counts = lines.map((l) => l.split(d).length - 1);
    const nonzero = counts.filter((c) => c > 0);
    if (!nonzero.length) continue;
    const first = counts[0];
    const same = counts.filter((c) => c === first).length;
    const score = nonzero.length * 2 + same * 3 + (first > 0 ? 1 : 0);
    if (score > bestScore) { best = d; bestScore = score; }
  }
  return best;
}

/**
 * Reads delimited text into rows of strings. Quoted cells may hold the delimiter, line breaks and "" for a quote
 * (Excel and Google Sheets write cells with a tab or a line break that way). Ragged rows are padded.
 * opts: { delimiter?, limits? } -> { rows, delimiter, columns }
 */
function parseDelimited(text, opts = {}) {
  const L = { ...LIMITS, ...(opts.limits || {}) };
  text = String(text == null ? '' : text);
  if (text.length > L.maxChars) throw new TableError(`The text is ${mb(text.length)} MB. Up to ${mb(L.maxChars)} MB is converted at a time.`, 'TOO_BIG');
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  const d = opts.delimiter || detectDelimiter(text);
  const rows = [];
  let row = [], cell = '', i = 0, cells = 0, maxCols = 0;
  const n = text.length;
  const endCell = () => {
    row.push(cell); cell = '';
    if (row.length > L.maxCols) throw new TableError(`A row has more than ${fmtInt(L.maxCols)} columns. Markdown tables are made up to ${fmtInt(L.maxCols)} columns.`, 'TOO_WIDE');
  };
  const endRow = () => {
    endCell();
    if (row.length > maxCols) maxCols = row.length;
    cells += row.length;
    rows.push(row); row = [];
    if (rows.length > L.maxRows) throw new TableError(`The text has more than ${fmtInt(L.maxRows)} rows. Up to ${fmtInt(L.maxRows)} rows are converted at a time.`, 'TOO_MANY_ROWS');
    if (cells > L.maxCells) throw new TableError(`The text has more than ${fmtInt(L.maxCells)} cells. Up to ${fmtInt(L.maxCells)} cells are converted at a time.`, 'TOO_MANY_CELLS');
  };
  while (i < n) {
    const ch = text[i];
    if (ch === '"' && cell === '') {
      // quoted cell
      let j = i + 1, buf = '';
      for (;;) {
        const k = text.indexOf('"', j);
        if (k < 0) { buf += text.slice(j); j = n; break; }
        buf += text.slice(j, k);
        if (text[k + 1] === '"') { buf += '"'; j = k + 2; continue; }
        j = k + 1; break;
      }
      // text after the closing quote up to the delimiter stays in the cell (lenient, like Excel)
      let k = j;
      while (k < n && text[k] !== d && text[k] !== '\n' && text[k] !== '\r') k++;
      cell = buf + text.slice(j, k);
      if (cell === '') cell = '\u0000'; // a quoted empty cell is still a cell; replaced below
      i = k;
      continue;
    }
    if (ch === d) { endCell(); i++; continue; }
    if (ch === '\r' || ch === '\n') {
      endRow();
      i += ch === '\r' && text[i + 1] === '\n' ? 2 : 1;
      continue;
    }
    // plain run up to the next special character
    let k = i + 1;
    while (k < n) { const c = text[k]; if (c === d || c === '\n' || c === '\r') break; k++; }
    cell += text.slice(i, k);
    i = k;
  }
  if (cell !== '' || row.length) endRow();
  for (const r of rows) {
    for (let c = 0; c < r.length; c++) if (r[c] === '\u0000') r[c] = '';
    while (r.length < maxCols) r.push('');
  }
  return { rows, delimiter: d, columns: maxCols };
}

/** Writes rows as CSV or TSV: cells with the delimiter, a quote, a line break or edge spaces are quoted. */
function toDelimited(rows, delimiter = ',') {
  const needs = (s) => s.indexOf(delimiter) >= 0 || /["\r\n]/.test(s) || /^\s|\s$/.test(s);
  return rows.map((r) => r.map((c) => { const s = cellText(c); return needs(s) ? `"${s.replace(/"/g, '""')}"` : s; }).join(delimiter)).join('\n') + '\n';
}

// ---------------- width ----------------

let segmenter = null;
try { segmenter = new Intl.Segmenter(undefined, { granularity: 'grapheme' }); } catch { segmenter = null; }
const ZERO = /^[\u0300-\u036f\u0483-\u0489\u0591-\u05bd\u0610-\u061a\u064b-\u065f\u0e31\u0e34-\u0e3a\u0e47-\u0e4e\u1ab0-\u1aff\u1dc0-\u1dff\u200b-\u200f\u2028-\u202e\u2060-\u2064\u20d0-\u20ff\ufe00-\ufe0f\ufe20-\ufe2f\ufeff]+$/;
let EMOJI = null;
try { EMOJI = /\p{Extended_Pictographic}/u; } catch { EMOJI = null; }

function wideCode(cp) {
  return (cp >= 0x1100 && cp <= 0x115f) || (cp >= 0x2e80 && cp <= 0x303e) || (cp >= 0x3041 && cp <= 0x33ff) ||
    (cp >= 0x3400 && cp <= 0x4dbf) || (cp >= 0x4e00 && cp <= 0x9fff) || (cp >= 0xa000 && cp <= 0xa4cf) ||
    (cp >= 0xa960 && cp <= 0xa97f) || (cp >= 0xac00 && cp <= 0xd7a3) || (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xfe10 && cp <= 0xfe19) || (cp >= 0xfe30 && cp <= 0xfe6f) || (cp >= 0xff00 && cp <= 0xff60) ||
    (cp >= 0xffe0 && cp <= 0xffe6) || (cp >= 0x1f300 && cp <= 0x1f64f) || (cp >= 0x1f900 && cp <= 0x1f9ff) ||
    (cp >= 0x1f680 && cp <= 0x1f6ff) || (cp >= 0x1fa70 && cp <= 0x1faff) || (cp >= 0x20000 && cp <= 0x3fffd);
}

/** Display width in a monospace editor: ASCII 1, CJK and emoji 2, combining marks 0. */
function strWidth(s) {
  s = String(s);
  if (/^[\x20-\x7e]*$/.test(s)) return s.length;
  let w = 0;
  const graphemes = segmenter ? Array.from(segmenter.segment(s), (g) => g.segment) : Array.from(s);
  for (const g of graphemes) {
    if (ZERO.test(g)) continue;
    const cp = g.codePointAt(0);
    if (cp < 0x20 || (cp >= 0x7f && cp < 0xa0)) continue;
    if (wideCode(cp) || (g.length > 1 && EMOJI && EMOJI.test(g) && /\ufe0f|\u200d|[\u{1f3fb}-\u{1f3ff}]/u.test(g)) || (EMOJI && EMOJI.test(g) && cp > 0xffff)) w += 2;
    else w += 1;
  }
  return w;
}

// ---------------- cells ----------------

/** A cell is a string, or { t: text, href } for a hyperlink (Pro "Keep Excel formats"). */
function cellText(c) { return c && typeof c === 'object' ? String(c.t == null ? '' : c.t) : String(c == null ? '' : c); }

const NUMERIC = /^[-+\u2212]?\(?[-+]?(?:[$\u20ac\u00a3\u00a5\u20b9\u20a9]\s?)?(?:\d{1,3}(?:[,\u00a0\u202f ]\d{3})+|\d+)?(?:[.,]\d+)?(?:[eE][-+]?\d+)?\)?\s?(?:%|\u2030|[$\u20ac\u00a3\u00a5\u20b9\u20a9]|[A-Z]{3})?$/;
function isNumeric(s) {
  s = cellText(s).trim();
  if (!s || !/\d/.test(s)) return false;
  return NUMERIC.test(s);
}

/** Escape rules: pipes and backslashes always; line breaks as <br> (or a space); HTML and Markdown on request. */
function escapeCell(c, o = {}) {
  let s = cellText(c).replace(/\r\n?/g, '\n').replace(/\t/g, ' ');
  s = s.replace(/\\/g, '\\\\').replace(/\|/g, '\\|');
  if (o.escapeHtml !== false) s = s.replace(/&(?=#?\w+;)/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  if (o.escapeMarkdown) s = s.replace(/([*_`[\]#~])/g, '\\$1');
  s = s.replace(/\n/g, o.lineBreak === 'space' ? ' ' : '<br>').replace(/\u0000/g, '');
  s = s.trim();
  if (c && typeof c === 'object' && c.href) {
    const href = String(c.href).replace(/[\s()<>|]/g, (m) => '%' + m.charCodeAt(0).toString(16).toUpperCase().padStart(2, '0'));
    if (/^(https?:|mailto:)/i.test(href)) s = `[${s.replace(/[[\]]/g, '\\$&')}](${href})`;
  }
  return s;
}

function normalize(rows) {
  const w = rows.reduce((m, r) => Math.max(m, r.length), 0);
  return rows.map((r) => { const o = r.slice(); while (o.length < w) o.push(''); return o; });
}

/** Column alignment: 'auto' (numbers right, others none), 'left', 'center', 'right', 'none'. */
function alignments(rows, mode, headerRow) {
  const cols = rows.length ? rows[0].length : 0;
  const out = [];
  for (let c = 0; c < cols; c++) {
    if (mode === 'left' || mode === 'center' || mode === 'right') { out.push(mode); continue; }
    if (mode !== 'auto') { out.push('none'); continue; }
    let nums = 0, other = 0;
    for (let r = headerRow ? 1 : 0; r < rows.length; r++) {
      const v = cellText(rows[r][c]).trim();
      if (!v) continue;
      if (isNumeric(v)) nums++; else { other++; break; }
    }
    out.push(nums > 0 && other === 0 ? 'right' : 'none');
  }
  return out;
}

const DEFAULTS = { header: 'first', align: 'auto', maxWidth: 0, compact: false, escapeHtml: true, escapeMarkdown: false, lineBreak: '<br>' };

/**
 * Writes rows as a GitHub Markdown table.
 * o: { header: 'first'|'empty', align: 'auto'|'left'|'center'|'right'|'none' or an array per column,
 *      maxWidth: 0 (pad every cell) or n (cells wider than n are not padded and do not widen the column),
 *      compact, escapeHtml, escapeMarkdown, lineBreak: '<br>'|'space', raw: cells are Markdown already }
 */
function toMarkdown(rows, opts = {}) {
  const o = { ...DEFAULTS, ...opts };
  rows = normalize(rows || []);
  if (!rows.length || !rows[0].length) return '';
  const headerRow = o.header !== 'empty';
  const al = Array.isArray(o.align) ? rows[0].map((_, i) => o.align[i] || 'none') : alignments(rows, o.align, headerRow);
  const cells = rows.map((r) => r.map((c) => (o.raw ? cellText(c).trim() : escapeCell(c, o))));
  const head = headerRow ? cells[0] : cells[0].map(() => '');
  const body = headerRow ? cells.slice(1) : cells;
  const cols = head.length;
  const max = Number(o.maxWidth) > 0 ? Number(o.maxWidth) : Infinity;
  const widths = new Array(cols).fill(3);
  const wcache = o.compact ? null : [head, ...body].map((r) => r.map((c) => strWidth(c)));
  if (!o.compact) {
    for (const r of wcache) for (let c = 0; c < cols; c++) { const w = r[c]; if (w <= max && w > widths[c]) widths[c] = w; }
    for (let c = 0; c < cols; c++) if (widths[c] < 3) widths[c] = 3;
  }
  const pad = (s, w, a, sw) => {
    if (o.compact || sw >= w) return s;
    const gap = w - sw;
    if (a === 'right') return ' '.repeat(gap) + s;
    if (a === 'center') { const l = Math.floor(gap / 2); return ' '.repeat(l) + s + ' '.repeat(gap - l); }
    return s + ' '.repeat(gap);
  };
  const line = (r, ri) => '| ' + r.map((c, i) => pad(c, widths[i], al[i], wcache ? wcache[ri][i] : 0)).join(' | ') + ' |';
  const dash = (i) => {
    const w = o.compact ? 3 : widths[i];
    if (al[i] === 'right') return '-'.repeat(w - 1) + ':';
    if (al[i] === 'left') return ':' + '-'.repeat(w - 1);
    if (al[i] === 'center') return ':' + '-'.repeat(Math.max(1, w - 2)) + ':';
    return '-'.repeat(w);
  };
  const out = [line(head, 0), '| ' + head.map((_, i) => dash(i)).join(' | ') + ' |'];
  for (let r = 0; r < body.length; r++) out.push(line(body[r], r + 1));
  return out.join('\n') + '\n';
}

// ---------------- other outputs (free extras) ----------------

const escHtml = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function toHtml(rows, opts = {}) {
  const o = { ...DEFAULTS, ...opts };
  rows = normalize(rows || []);
  if (!rows.length) return '';
  const headerRow = o.header !== 'empty';
  const al = alignments(rows, o.align, headerRow);
  const cell = (c, tag, i) => {
    let s = escHtml(cellText(c)).replace(/\r\n?|\n/g, '<br>');
    if (c && typeof c === 'object' && c.href && /^(https?:|mailto:)/i.test(c.href)) s = `<a href="${escHtml(String(c.href))}">${s}</a>`;
    const a = al[i] === 'right' || al[i] === 'center' || al[i] === 'left' ? ` style="text-align: ${al[i]}"` : '';
    return `<${tag}${a}>${s}</${tag}>`;
  };
  const out = ['<table>'];
  if (headerRow) out.push('  <thead>', '    <tr>' + rows[0].map((c, i) => cell(c, 'th', i)).join('') + '</tr>', '  </thead>');
  out.push('  <tbody>');
  for (const r of rows.slice(headerRow ? 1 : 0)) out.push('    <tr>' + r.map((c, i) => cell(c, 'td', i)).join('') + '</tr>');
  out.push('  </tbody>', '</table>');
  return out.join('\n') + '\n';
}

function toAsciiDoc(rows, opts = {}) {
  const o = { ...DEFAULTS, ...opts };
  rows = normalize(rows || []);
  if (!rows.length) return '';
  const headerRow = o.header !== 'empty';
  const al = alignments(rows, o.align, headerRow);
  const spec = al.map((a) => (a === 'right' ? '>' : a === 'center' ? '^' : '<')).join(',');
  const cell = (c) => {
    let s = cellText(c).replace(/\r\n?/g, '\n').replace(/\|/g, '\\|').replace(/\n/g, ' +\n');
    if (c && typeof c === 'object' && c.href && /^(https?:|mailto:)/i.test(c.href)) s = `${c.href}[${s.replace(/]/g, '\\]')}]`;
    return '| ' + s;
  };
  const out = [`[cols="${spec}"${headerRow ? ', options="header"' : ''}]`, '|==='];
  rows.forEach((r, i) => { out.push(r.map(cell).join(' ')); if (i === 0 && headerRow) out.push(''); });
  out.push('|===');
  return out.join('\n') + '\n';
}

function toJira(rows, opts = {}) {
  const o = { ...DEFAULTS, ...opts };
  rows = normalize(rows || []);
  if (!rows.length) return '';
  const headerRow = o.header !== 'empty';
  const cell = (c) => {
    let s = cellText(c).replace(/\r\n?/g, '\n').replace(/\|/g, '\\|').replace(/\n/g, ' \\\\ ');
    if (!s.trim()) s = ' ';
    if (c && typeof c === 'object' && c.href && /^(https?:|mailto:)/i.test(c.href)) s = `[${s}|${c.href}]`;
    return s;
  };
  return rows.map((r, i) => (i === 0 && headerRow ? '||' + r.map(cell).join('||') + '||' : '|' + r.map(cell).join('|') + '|')).join('\n') + '\n';
}

function render(format, rows, opts) {
  if (format === 'html') return toHtml(rows, opts);
  if (format === 'asciidoc') return toAsciiDoc(rows, opts);
  if (format === 'jira') return toJira(rows, opts);
  if (format === 'csv') return toDelimited(rows, ',');
  if (format === 'tsv') return toDelimited(rows, '\t');
  return toMarkdown(rows, opts);
}

// ---------------- Markdown tables back ----------------

const DELIM_ROW = /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/;

/** Splits one Markdown table line into raw cells (escaped pipes stay in the cell). */
function splitRow(line) {
  let s = line.trim();
  if (s.startsWith('|')) s = s.slice(1);
  const cells = [];
  let cur = '';
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    if (ch === '\\' && i + 1 < s.length) { cur += ch + s[i + 1]; i++; continue; }
    if (ch === '|') { cells.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim() !== '' || !s.endsWith('|')) cells.push(cur.trim());
  return cells;
}

/** Markdown cell text -> plain text: <br> to a line break, backslash escapes and the HTML entities undone. */
function unescapeCell(s) {
  return String(s)
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/\\([!-/:-@[-`{-~])/g, '$1')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#124;/g, '|').replace(/&amp;/g, '&');
}

/** Parses the Markdown table in `text` (header, delimiter row, body). Returns null when there is none. */
function parseMarkdownTable(text, opts = {}) {
  const lines = String(text).replace(/\r\n?/g, '\n').split('\n').filter((l) => l.trim() !== '');
  if (lines.length < 2 || lines.length > LIMITS.maxRows) return null;
  const di = lines.findIndex((l) => DELIM_ROW.test(l) && l.includes('-'));
  if (di !== 1) return null;
  const header = splitRow(lines[0]);
  const delim = splitRow(lines[1]);
  const align = delim.map((d) => { const l = d.startsWith(':'), r = d.endsWith(':'); return l && r ? 'center' : r ? 'right' : l ? 'left' : 'none'; });
  const raw = [header, ...lines.slice(2).map(splitRow)];
  const cols = Math.min(LIMITS.maxCols, Math.max(header.length, ...raw.map((r) => r.length)));
  for (const r of raw) { r.length = Math.min(r.length, cols); while (r.length < cols) r.push(''); }
  while (align.length < cols) align.push('none');
  align.length = cols;
  return { raw, rows: opts.keepRaw ? raw : raw.map((r) => r.map(unescapeCell)), align };
}

const looksLikeRow = (l) => /\|/.test(l) && l.trim() !== '';

/** Finds the Markdown table around line `at` of `lines`: { start, end } (inclusive) or null. */
function findTableAt(lines, at) {
  if (at < 0 || at >= lines.length || !looksLikeRow(lines[at])) return null;
  let s = at, e = at;
  while (s > 0 && looksLikeRow(lines[s - 1]) && at - s < LIMITS.maxRows) s--;
  while (e < lines.length - 1 && looksLikeRow(lines[e + 1]) && e - at < LIMITS.maxRows) e++;
  // the table starts at the line above the first delimiter row of the block
  for (let k = s; k < e; k++) {
    if (DELIM_ROW.test(lines[k + 1]) && lines[k + 1].includes('-')) return at >= k ? { start: k, end: e } : null;
  }
  return null;
}

/** Every Markdown table of a document: [{ start, end, title }] with the nearest heading above as the title. */
function findAllTables(text, limit = 500) {
  const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
  const out = [];
  let heading = '', inFence = false;
  for (let i = 0; i < lines.length && out.length < limit; i++) {
    const l = lines[i];
    if (/^\s*(```|~~~)/.test(l)) { inFence = !inFence; continue; }
    if (inFence) continue;
    const h = /^\s{0,3}#{1,6}\s+(.*?)\s*#*\s*$/.exec(l);
    if (h) { heading = h[1]; continue; }
    if (looksLikeRow(l) && i + 1 < lines.length && DELIM_ROW.test(lines[i + 1]) && lines[i + 1].includes('-')) {
      let e = i + 1;
      while (e + 1 < lines.length && looksLikeRow(lines[e + 1])) e++;
      out.push({ start: i, end: e, title: heading, text: lines.slice(i, e + 1).join('\n') });
      i = e;
    }
  }
  return out;
}

/** Re-aligns a Markdown table; keeps its cells and its alignment row. */
function formatMarkdownTable(text, opts = {}) {
  const t = parseMarkdownTable(text, { keepRaw: true });
  if (!t) return null;
  const align = t.align.some((a) => a !== 'none') ? t.align : (opts.align === 'auto' || !opts.align ? alignments(t.raw.map((r) => r.map(unescapeCell)), 'auto', true) : alignments(t.raw, opts.align, true));
  return toMarkdown(t.raw, { ...opts, raw: true, header: 'first', align });
}

// ---------------- A1 ranges ----------------

function colIndex(letters) { let n = 0; for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; }
function colName(i) { let s = ''; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }

/** 'B2:D10', 'B:D', '2:10', 'C5' -> { s: {r, c}, e: {r, c} } inside `bounds`; '' -> bounds. Throws on bad input. */
function parseRange(text, bounds) {
  const t = String(text || '').trim().replace(/\$/g, '');
  if (!t) return bounds;
  const m = /^([A-Za-z]{1,3})?(\d{1,7})?(?::([A-Za-z]{1,3})?(\d{1,7})?)?$/.exec(t);
  if (!m || (!m[1] && !m[2])) throw new TableError(`"${t.slice(0, 40)}" is not a range. Use A1 notation, for example A1:D20, B:D or 2:10.`, 'RANGE');
  const c1 = m[1] ? colIndex(m[1]) : bounds.s.c, r1 = m[2] ? Number(m[2]) - 1 : bounds.s.r;
  const hasEnd = t.includes(':');
  const c2 = hasEnd ? (m[3] ? colIndex(m[3]) : (m[1] && !m[2] ? c1 : bounds.e.c)) : (m[1] ? c1 : bounds.e.c);
  const r2 = hasEnd ? (m[4] ? Number(m[4]) - 1 : (m[2] && !m[1] ? r1 : bounds.e.r)) : (m[2] ? r1 : bounds.e.r);
  const s = { r: Math.max(Math.min(r1, r2), 0), c: Math.max(Math.min(c1, c2), 0) };
  const e = { r: Math.max(r1, r2), c: Math.max(c1, c2) };
  if (r1 < 0 || r2 < 0) throw new TableError('Row numbers start at 1.', 'RANGE');
  return { s, e };
}

function rangeName(r) { return `${colName(r.s.c)}${r.s.r + 1}:${colName(r.e.c)}${r.e.r + 1}`; }

module.exports = {
  LIMITS, DEFAULTS, TableError, detectDelimiter, parseDelimited, toDelimited, strWidth, isNumeric, escapeCell, unescapeCell, alignments,
  toMarkdown, toHtml, toAsciiDoc, toJira, render, splitRow, parseMarkdownTable, findTableAt, findAllTables, formatMarkdownTable,
  parseRange, rangeName, colName, colIndex, cellText, fmtInt,
};
