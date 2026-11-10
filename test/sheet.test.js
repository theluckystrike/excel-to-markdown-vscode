'use strict';
// Workbooks through the real worker bundle (dist/worker.js): values as stored, ISO dates, Excel display text, links,
// merged cells, several sheets, .ods/.xls/.csv, writing .xlsx, and the caps (zip bombs, cut files, a hung worker).
// Fixtures are built in memory with SheetJS from the sibling table-viewer-vscode tree (a dev-only dependency).
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const zlib = require('zlib');
const XLSX = require(path.join(__dirname, '..', '..', 'table-viewer-vscode', 'node_modules', 'xlsx'));
const { runJob } = require('../src/runner');
const S = require('../src/core/sheet');
const T = require('../src/core/table');
const { checkZip } = require('../src/core/zipcheck');

const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'e2m-sheet-'));
test.after(() => fs.rmSync(dir, { recursive: true, force: true }));
const put = (name, buf) => { const p = path.join(dir, name); fs.writeFileSync(p, buf); return p; };

function priceBook() {
  const ws = XLSX.utils.aoa_to_sheet([
    ['Item', 'Price', 'Sold on', 'Share', 'Page'],
    ['Green tea', 1234.5, 46027, 0.25, 'tea page'],
    ['Merged label', null, null, null, null],
    ['Cake | slice', -3, 46027.5, 1, 'none'],
  ]);
  ws.B2.z = '#,##0.00'; ws.C2.z = 'yyyy-mm-dd'; ws.D2.z = '0%'; ws.C4.z = 'yyyy-mm-dd hh:mm'; ws.D4.z = '0%';
  ws.E2.l = { Target: 'https://example.com/tea' };
  ws['!merges'] = [{ s: { r: 2, c: 0 }, e: { r: 2, c: 2 } }];
  const team = XLSX.utils.aoa_to_sheet([['\u540d\u524d', 'Role'], ['\u5c71\u7530', 'Lead \ud83d\ude00'], ['Ann', '<b>Dev</b>']]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Prices');
  XLSX.utils.book_append_sheet(wb, team, 'Team');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([]), 'Empty');
  return wb;
}

test('xlsx: values as stored, ISO dates, several sheets; Pro formats give Excel text, links and merged cells', async () => {
  const f = put('prices.xlsx', XLSX.write(priceBook(), { type: 'buffer', bookType: 'xlsx' }));
  const m = await runJob({ op: 'read', file: f, name: f });
  assert.deepStrictEqual(m.sheets.map((s) => s.name), ['Prices', 'Team', 'Empty']);
  const p = m.sheets[0];
  assert.deepStrictEqual(p.rows[1], ['Green tea', '1234.5', '2026-01-05', '0.25', 'tea page']);
  assert.deepStrictEqual(p.rows[3], ['Cake | slice', '-3', '2026-01-05 12:00', '1', 'none']);
  assert.strictEqual(p.ref, 'A1:E4');
  const plain = T.toMarkdown(S.sheetRows(p).rows);
  assert.match(plain, /\| Green tea +\| +1234\.5 \| 2026-01-05 +\| +0\.25 \| tea page \|/);
  assert.match(plain, /Cake \\\| slice/);
  const rich = T.toMarkdown(S.sheetRows(p, { formats: true }).rows);
  assert.match(rich, /1,234\.50/);
  assert.match(rich, /25%/);
  assert.match(rich, /\[tea page\]\(https:\/\/example\.com\/tea\)/);
  assert.match(rich, /\| Merged label +\| Merged label +\| Merged label +\|/);
  const team = T.toMarkdown(S.sheetRows(m.sheets[1]).rows);
  assert.match(team, /&lt;b&gt;Dev&lt;\/b&gt;/);
  assert.strictEqual(m.sheets[2].rows.length, 0);
});

test('ranges of a sheet: A1:B2, a column range, a range past the data', async () => {
  const f = put('range.xlsx', XLSX.write(priceBook(), { type: 'buffer', bookType: 'xlsx' }));
  const m = await runJob({ op: 'read', file: f, name: f });
  assert.deepStrictEqual(S.sheetRows(m.sheets[0], { range: 'A1:B2' }).rows, [['Item', 'Price'], ['Green tea', '1234.5']]);
  assert.deepStrictEqual(S.sheetRows(m.sheets[0], { range: 'E:E' }).rows.map((r) => r[0]), ['Page', 'tea page', '', 'none']);
  assert.deepStrictEqual(S.sheetRows(m.sheets[0], { range: 'Z90:Z99' }).rows, []);
  assert.throws(() => S.sheetRows(m.sheets[0], { range: 'nope!' }), /not a range/);
});

test('.ods, .xls (BIFF8) and .csv / .tsv files read the same cells', async () => {
  const wb = priceBook();
  for (const [ext, type] of [['ods', 'ods'], ['xls', 'biff8']]) {
    const f = put('book.' + ext, XLSX.write(wb, { type: 'buffer', bookType: type }));
    const m = await runJob({ op: 'read', file: f, name: f });
    assert.strictEqual(m.sheets[0].rows[1][0], 'Green tea', ext);
    assert.strictEqual(m.sheets[0].rows[1][1], '1234.5', ext);
    assert.strictEqual(m.sheets[1].rows[1][0], '\u5c71\u7530', ext);
  }
  const csv = put('list.csv', Buffer.from('\ufeffid,name\n1,"Smith, Ann"\n2,"two\nlines"\n'));
  const c = await runJob({ op: 'read', file: csv, name: csv });
  assert.deepStrictEqual(c.sheets[0].rows, [['id', 'name'], ['1', 'Smith, Ann'], ['2', 'two\nlines']]);
  const tsv = put('list.tsv', Buffer.from('a,b\tc\n1\t2\n'));
  assert.deepStrictEqual((await runJob({ op: 'read', file: tsv, name: tsv })).sheets[0].rows, [['a,b', 'c'], ['1', '2']]);
  const u16 = put('u16.csv', Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('x,y\n1,2\n', 'utf16le')]));
  assert.deepStrictEqual((await runJob({ op: 'read', file: u16, name: u16 })).sheets[0].rows, [['x', 'y'], ['1', '2']]);
});

test('write .xlsx from Markdown tables: one sheet per table, safe unique names, numbers stored as numbers', async () => {
  const tables = [
    { name: 'Sales: Q1/Q2', rows: [['Item', 'Qty', 'Share'], ['Tea', '3', '25%'], ['Cake', '12.5', '']] },
    { name: 'Sales: Q1/Q2', rows: [['a'], ['<b>x</b>']] },
    { name: '', rows: [['only'], ['row']] },
  ];
  const buf = Buffer.from(await runJob({ op: 'writeXlsx', tables }));
  const wb = XLSX.read(buf, { type: 'buffer', cellNF: true });
  assert.deepStrictEqual(wb.SheetNames, ['Sales Q1 Q2', 'Sales Q1 Q2 (2)', 'Table 3']);
  const ws = wb.Sheets[wb.SheetNames[0]];
  assert.strictEqual(ws.B2.t, 'n'); assert.strictEqual(ws.B2.v, 3);
  assert.strictEqual(ws.C2.v, 0.25); assert.strictEqual(ws.C2.z, '0%');
  assert.strictEqual(ws.A2.v, 'Tea');
  assert.strictEqual(wb.Sheets[wb.SheetNames[1]].A2.v, '<b>x</b>');
  // and it reads back through our reader
  const f = put('round.xlsx', buf);
  const m = await runJob({ op: 'read', file: f, name: f });
  assert.deepStrictEqual(m.sheets[0].rows[1], ['Tea', '3', '0.25']);
});

/** A zip with one stored or deflated entry, sizes as given (to fake a lying directory). */
function zipWith(name, data, { method = 8, claim } = {}) {
  const comp = method === 8 ? zlib.deflateRawSync(data) : data;
  const usize = claim == null ? data.length : claim;
  const nb = Buffer.from(name);
  const lh = Buffer.alloc(30); lh.writeUInt32LE(0x04034b50, 0); lh.writeUInt16LE(20, 4); lh.writeUInt16LE(method, 8); lh.writeUInt32LE(comp.length, 18); lh.writeUInt32LE(usize, 22); lh.writeUInt16LE(nb.length, 26);
  const cd = Buffer.alloc(46); cd.writeUInt32LE(0x02014b50, 0); cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6); cd.writeUInt16LE(method, 10); cd.writeUInt32LE(comp.length, 20); cd.writeUInt32LE(usize, 24); cd.writeUInt16LE(nb.length, 28); cd.writeUInt32LE(0, 42);
  const body = Buffer.concat([lh, nb, comp]);
  const cdb = Buffer.concat([cd, nb]);
  const eo = Buffer.alloc(22); eo.writeUInt32LE(0x06054b50, 0); eo.writeUInt16LE(1, 8); eo.writeUInt16LE(1, 10); eo.writeUInt32LE(cdb.length, 12); eo.writeUInt32LE(body.length, 16);
  return Buffer.concat([body, cdb, eo]);
}

test('zip bomb caps: a part that unpacks past its claim, a 1000x ratio, too many parts, a huge total, a cut zip', async () => {
  // 20 MB of zeros packs to about 20 KB; the directory says 1,000 bytes
  const liar = zipWith('xl/worksheets/sheet1.xml', Buffer.alloc(20 * 1024 * 1024), { claim: 1000 });
  assert.throws(() => checkZip(liar), /unpacks to more than the 1000 bytes it claims/);
  // honest but extreme ratio: refused from the directory alone
  const ratio = zipWith('xl/sharedStrings.xml', Buffer.alloc(2 * 1024 * 1024));
  assert.throws(() => checkZip(ratio), /zip bomb/);
  const many = Buffer.from(liar); many.writeUInt16LE(20000, many.length - 22 + 10);
  assert.throws(() => checkZip(many), /20000 parts/);
  const huge = zipWith('a.xml', Buffer.from('x'.repeat(100)), { method: 0 });
  assert.throws(() => checkZip(huge, { maxTotal: 50 }), /unpacks to more than/);
  assert.throws(() => checkZip(liar.subarray(0, liar.length - 30)), /not a complete zip|damaged|past the end/);
  // through the worker: a message, quickly, and the next good file still reads
  const t0 = Date.now();
  const bomb = put('bomb.xlsx', liar);
  await assert.rejects(runJob({ op: 'read', file: bomb, name: bomb }), /claims/);
  assert.ok(Date.now() - t0 < 5000);
  const good = put('good.xlsx', XLSX.write(priceBook(), { type: 'buffer', bookType: 'xlsx' }));
  assert.strictEqual((await runJob({ op: 'read', file: good, name: good })).sheets.length, 3);
});

test('hostile files: empty, renamed text, random bytes, cut workbook, 3,000 columns; each gives a message', async () => {
  const cases = {
    'empty.xlsx': Buffer.alloc(0),
    'text.xlsx': Buffer.from('hello, this is not a workbook\n'),
    'random.xls': Buffer.from(Array.from({ length: 4096 }, (_, i) => (i * 7919 + 13) & 255)),
    'cut.xlsx': XLSX.write(priceBook(), { type: 'buffer', bookType: 'xlsx' }).subarray(0, 900),
  };
  for (const [n, b] of Object.entries(cases)) {
    const f = put(n, b);
    const t0 = Date.now();
    let err = null;
    try { await runJob({ op: 'read', file: f, name: f }, { ms: 15000 }); } catch (e) { err = e; }
    if (n === 'random.xls' && !err) continue; // SheetJS may read random bytes as a text sheet; any result is fine as long as it ends
    assert.ok(err && err.message.length > 5, n);
    assert.ok(Date.now() - t0 < 15000, n);
  }
  // wide sheet: columns past the cap are cut with a note
  const row = Array.from({ length: 3000 }, (_, i) => i);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([row, row]), 'Wide');
  const f = put('wide.xlsx', XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
  const m = await runJob({ op: 'read', file: f, name: f });
  assert.strictEqual(m.sheets[0].rows[0].length, 2000);
  assert.match(m.sheets[0].truncated, /first 2,000 columns/);
});

test('rows cap: a sheet with more rows than the limit is cut with a note (limit lowered for the test)', () => {
  const aoa = Array.from({ length: 50 }, (_, i) => [i, 'r' + i]);
  const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), 'Long');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  const m = S.readWorkbook(XLSX, buf, 'long.xlsx', { maxRows: 10 });
  assert.strictEqual(m.sheets[0].rows.length, 10);
  assert.match(m.sheets[0].truncated, /first 10 rows/);
  assert.throws(() => S.readWorkbook(XLSX, buf, 'long.xlsx', { maxFileBytes: 100 }), /workbooks up to/);
});

test('watchdog: a worker that hangs is stopped and the host goes on', async () => {
  const t0 = Date.now();
  await assert.rejects(runJob({ op: 'hang' }, { ms: 800 }), (e) => e.code === 'TIMEOUT');
  assert.ok(Date.now() - t0 < 4000);
});

test('number and date text: plain numbers, big integers, time-only cells', () => {
  assert.strictEqual(S.plainNumber(0.1 + 0.2), '0.3');
  assert.strictEqual(S.plainNumber(12345678901234), '12345678901234');
  assert.strictEqual(S.plainNumber(1e22), '1e+22');
  assert.strictEqual(S.isoDate(XLSX, 0.5), '12:00');
  assert.strictEqual(S.isoDate(XLSX, 45658.75), '2025-01-01 18:00');
  assert.strictEqual(S.rawText(XLSX, { t: 'b', v: true }), 'TRUE');
  assert.strictEqual(S.rawText(XLSX, { t: 'e', v: 42, w: '#N/A' }), '#N/A');
});
