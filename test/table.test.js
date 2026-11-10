'use strict';
// Clipboard and text tables: Excel, Google Sheets, Numbers and LibreOffice TSV, CSV, Markdown output and alignment,
// escaping, the round trip back from Markdown, formatting, ranges, and hostile input (5 MB, 10,000 columns, HTML).
const test = require('node:test');
const assert = require('node:assert');
const T = require('../src/core/table');

test('Excel TSV: CRLF rows, quoted cells with a tab, a line break and "" quotes', () => {
  // what Excel for Windows and Mac puts on the clipboard for 3 x 3 cells
  const clip = 'Name\tNote\tQty\r\n"Multi\nline"\t"He said ""hi"""\t3\r\n"Tab\there"\t\t1,250\r\n';
  const p = T.parseDelimited(clip);
  assert.strictEqual(p.delimiter, '\t');
  assert.deepStrictEqual(p.rows, [['Name', 'Note', 'Qty'], ['Multi\nline', 'He said "hi"', '3'], ['Tab\there', '', '1,250']]);
  const md = T.toMarkdown(p.rows);
  assert.strictEqual(md, [
    '| Name          | Note         |   Qty |',
    '| ------------- | ------------ | ----: |',
    '| Multi<br>line | He said "hi" |     3 |',
    '| Tab here      |              | 1,250 |',
    ''].join('\n'));
});

test('Google Sheets TSV (LF, quoted line breaks), Numbers and LibreOffice copies', () => {
  const sheets = 'City\tPop\nTokyo\t37400068\n"Rio de\nJaneiro"\t13634274';
  assert.deepStrictEqual(T.parseDelimited(sheets).rows, [['City', 'Pop'], ['Tokyo', '37400068'], ['Rio de\nJaneiro', '13634274']]);
  // Numbers: trailing tab for an empty last cell, no final newline
  const numbers = 'A\tB\tC\n1\t2\t\n3\t\t';
  assert.deepStrictEqual(T.parseDelimited(numbers).rows, [['A', 'B', 'C'], ['1', '2', ''], ['3', '', '']]);
  // LibreOffice: a quote inside a cell that is not at its start stays as it is
  const lo = 'Size\tLabel\n12"\tscreen';
  assert.deepStrictEqual(T.parseDelimited(lo).rows, [['Size', 'Label'], ['12"', 'screen']]);
});

test('CSV and semicolon CSV are detected; a BOM is dropped', () => {
  assert.strictEqual(T.detectDelimiter('a,b,c\n1,2,3'), ',');
  assert.strictEqual(T.detectDelimiter('a;b;c\n1,5;2,5;3'), ';');
  const p = T.parseDelimited('\ufeffid,name\n1,"Smith, Ann"\n');
  assert.deepStrictEqual(p.rows, [['id', 'name'], ['1', 'Smith, Ann']]);
});

test('empty cells and ragged rows are padded to the widest row', () => {
  const p = T.parseDelimited('a\tb\tc\n1\n2\t3\n\t\t\n');
  assert.deepStrictEqual(p.rows, [['a', 'b', 'c'], ['1', '', ''], ['2', '3', ''], ['', '', '']]);
  const md = T.toMarkdown(p.rows);
  assert.strictEqual(md.split('\n').length, 6);
  assert.ok(md.split('\n').slice(0, 5).every((l) => (l.match(/(?<!\\)\|/g) || []).length === 4), md);
});

test('pipes and backslashes are escaped; HTML in cells stays text', () => {
  const md = T.toMarkdown([['col'], ['a|b'], ['C:\\temp\\x'], ['<script>alert(1)</script>'], ['<img src=x onerror=alert(1)>'], ['&copy; & co']]);
  assert.match(md, /a\\\|b/);
  assert.match(md, /C:\\\\temp\\\\x/);
  assert.ok(!/<script|<img/i.test(md), md);
  assert.match(md, /&lt;script&gt;alert\(1\)&lt;\/script&gt;/);
  assert.match(md, /&amp;copy; & co/);
  // escapeHtml off keeps the text as typed (pipes still escaped)
  assert.match(T.toMarkdown([['h'], ['<b>x|y</b>']], { escapeHtml: false }), /<b>x\\\|y<\/b>/);
  // escapeMarkdown escapes the Markdown characters
  assert.match(T.toMarkdown([['h'], ['*bold* _i_ `c`']], { escapeMarkdown: true }), /\\\*bold\\\* \\_i\\_ \\`c\\`/);
});

test('alignment: numbers right, text default; currencies, percent, negatives and thousands count as numbers', () => {
  for (const n of ['1', '-4', '1,234.56', '$1,200.50', '12%', '(35)', '1.5e3', '\u20ac 12', '7 USD', '0.25']) assert.ok(T.isNumeric(n), n);
  for (const s of ['2024-01-05', 'abc', '1.2.3', '', '12 apples', '10:30']) assert.ok(!T.isNumeric(s), s);
  const md = T.toMarkdown([['Item', 'Price', 'Code'], ['Tea', '3.50', 'A1'], ['Cake', '12', 'B2']]);
  assert.match(md.split('\n')[1], /^\| ---- \| ----: \| ---- \|$/);
  assert.match(T.toMarkdown([['a', 'b'], ['1', 'x']], { align: 'left' }).split('\n')[1], /^\| :-- \| :-- \|$/);
  assert.match(T.toMarkdown([['a', 'b'], ['1', 'x']], { align: 'center' }).split('\n')[1], /^\| :-: \| :-: \|$/);
  assert.match(T.toMarkdown([['a', 'b'], ['1', 'x']], { align: 'none' }).split('\n')[1], /^\| --- \| --- \|$/);
});

test('unicode width: CJK and emoji take two columns, so the pipes line up in a monospace editor', () => {
  assert.strictEqual(T.strWidth('abc'), 3);
  assert.strictEqual(T.strWidth('\u6771\u4eac'), 4); // Tokyo in kanji
  assert.strictEqual(T.strWidth('\ud55c\uad6d\uc5b4'), 6); // Korean
  assert.strictEqual(T.strWidth('\ud83d\ude00'), 2); // grinning face
  assert.strictEqual(T.strWidth('\ud83d\udc68\u200d\ud83d\udc69\u200d\ud83d\udc67'), 2); // family (ZWJ sequence)
  assert.strictEqual(T.strWidth('e\u0301'), 1); // e + combining acute
  const md = T.toMarkdown([['Name', 'City'], ['\u5c71\u7530', '\u6771\u4eac\u90fd'], ['Ann', 'Rome \ud83d\ude00']]);
  const widths = md.trim().split('\n').map((l) => T.strWidth(l));
  assert.ok(widths.every((w) => w === widths[0]), JSON.stringify(widths) + '\n' + md);
});

test('maxWidth: a long cell is not padded and does not widen its column; compact writes no padding', () => {
  const long = 'x'.repeat(80);
  const md = T.toMarkdown([['a', 'b'], [long, '1'], ['short', '2']], { maxWidth: 10 });
  assert.match(md, /\| short \| +2 \|/);
  assert.ok(md.split('\n')[0].length < 30);
  assert.strictEqual(T.toMarkdown([['a', 'bb'], ['1', '22']], { compact: true }), '| a | bb |\n| --: | --: |\n| 1 | 22 |\n');
});

test('header row: first row, or an empty header with every row in the body; line breaks as a space', () => {
  const md = T.toMarkdown([['1', '2'], ['3', '4']], { header: 'empty' });
  assert.deepStrictEqual(md.split('\n').slice(0, 3), ['|     |     |', '| --: | --: |', '|   1 |   2 |']);
  assert.match(T.toMarkdown([['h'], ['a\nb']], { lineBreak: 'space' }), /\| a b \|/);
});

test('Markdown round trip: rows -> Markdown -> rows keeps every cell, including pipes, backslashes, <br> and HTML', () => {
  const rows = [['Name', 'Path', 'Note'], ['a|b', 'C:\\x\\y', 'two\nlines'], ['<tag>', '', '50% & more'], ['\u6771\u4eac', '\ud83d\ude00', '*x*']];
  for (const o of [{}, { escapeMarkdown: true }, { compact: true }]) {
    const t = T.parseMarkdownTable(T.toMarkdown(rows, o));
    assert.deepStrictEqual(t.rows, rows, JSON.stringify(o));
  }
  // Markdown -> CSV -> rows
  const csv = T.toDelimited(rows, ',');
  assert.deepStrictEqual(T.parseDelimited(csv.replace(/\n$/, ''), { delimiter: ',' }).rows, rows);
  const tsv = T.toDelimited(rows, '\t');
  assert.deepStrictEqual(T.parseDelimited(tsv.replace(/\n$/, '')).rows, rows);
});

test('reading Markdown tables: no outer pipes, alignment row, short rows, not a table', () => {
  const t = T.parseMarkdownTable('a | b | c\n:-- | :-: | --:\n1 | 2\n');
  assert.deepStrictEqual(t.align, ['left', 'center', 'right']);
  assert.deepStrictEqual(t.rows, [['a', 'b', 'c'], ['1', '2', '']]);
  assert.strictEqual(T.parseMarkdownTable('just text\nmore text'), null);
  assert.strictEqual(T.parseMarkdownTable('| a |\n| b |'), null);
});

test('format a messy table: columns lined up, its own alignment row kept', () => {
  const out = T.formatMarkdownTable('|Item|Qty|\n|:-|-:|\n|Green tea|3|\n|Cake|12|');
  assert.strictEqual(out, '| Item      | Qty |\n| :-------- | --: |\n| Green tea |   3 |\n| Cake      |  12 |\n');
  // a table with no alignment marks gets numbers right
  assert.match(T.formatMarkdownTable('|a|b|\n|-|-|\n|x|1|'), /\| --- \| --: \|/);
  // escaped pipes stay escaped
  assert.match(T.formatMarkdownTable('|a|\n|-|\n|x\\|y|'), /x\\\|y/);
});

test('find the table under the cursor and every table of a document (code fences skipped, heading as title)', () => {
  const lines = ['Intro | not a table', '', '| a | b |', '| - | - |', '| 1 | 2 |', '', 'end'];
  assert.deepStrictEqual(T.findTableAt(lines, 4), { start: 2, end: 4 });
  assert.deepStrictEqual(T.findTableAt(lines, 2), { start: 2, end: 4 });
  assert.strictEqual(T.findTableAt(lines, 0), null);
  assert.strictEqual(T.findTableAt(lines, 6), null);
  const doc = '# Sales\n\n| q | n |\n|---|---|\n| 1 | 2 |\n\n```\n| x | y |\n|---|---|\n```\n\n## Costs 2026\n\n| c |\n|---|\n| 9 |\n';
  const all = T.findAllTables(doc);
  assert.deepStrictEqual(all.map((t) => [t.title, t.start, t.end]), [['Sales', 2, 4], ['Costs 2026', 13, 15]]);
});

test('HTML, AsciiDoc, Jira and CSV outputs escape their own special characters', () => {
  const rows = [['A', 'N'], ['<b>x</b> | y', '5']];
  const html = T.toHtml(rows);
  assert.match(html, /<th>A<\/th><th style="text-align: right">N<\/th>/);
  assert.match(html, /&lt;b&gt;x&lt;\/b&gt; \| y/);
  assert.match(T.toAsciiDoc(rows), /^\[cols="<,>", options="header"\]\n\|===\n\| A \| N\n\n\| <b>x<\/b> \\\| y \| 5\n\|===\n$/);
  assert.strictEqual(T.toJira(rows), '||A||N||\n|<b>x</b> \\| y|5|\n');
  assert.strictEqual(T.render('csv', [['a,b', 'c']]), '"a,b",c\n');
  // a hyperlink cell
  const link = [['Site'], [{ t: 'Zovo docs', href: 'https://example.com/a b' }]];
  assert.match(T.toMarkdown(link), /\[Zovo docs\]\(https:\/\/example\.com\/a%20b\)/);
  assert.match(T.toHtml(link), /<a href="https:\/\/example\.com\/a b">Zovo docs<\/a>/);
  assert.ok(!/javascript/.test(T.toMarkdown([['x'], [{ t: 'bad', href: 'javascript:alert(1)' }]])));
});

test('A1 ranges: cells, columns, rows, and a clear message for bad input', () => {
  const b = { s: { r: 0, c: 0 }, e: { r: 99, c: 9 } };
  assert.deepStrictEqual(T.parseRange('', b), b);
  assert.deepStrictEqual(T.parseRange('b2:d10', b), { s: { r: 1, c: 1 }, e: { r: 9, c: 3 } });
  assert.deepStrictEqual(T.parseRange('$D$10:$B$2', b), { s: { r: 1, c: 1 }, e: { r: 9, c: 3 } });
  assert.deepStrictEqual(T.parseRange('C:E', b), { s: { r: 0, c: 2 }, e: { r: 99, c: 4 } });
  assert.deepStrictEqual(T.parseRange('3:5', b), { s: { r: 2, c: 0 }, e: { r: 4, c: 9 } });
  assert.deepStrictEqual(T.parseRange('C5', b), { s: { r: 4, c: 2 }, e: { r: 4, c: 2 } });
  assert.throws(() => T.parseRange('hello world', b), /not a range/);
  assert.throws(() => T.parseRange('A0', b), /start at 1/);
  assert.strictEqual(T.colName(0) + T.colName(25) + T.colName(26) + T.colName(701) + T.colName(702), 'AZAAZZAAA');
  assert.strictEqual(T.colIndex('AAA'), 702);
});

test('timing: a 10,000-row paste (8 columns) becomes a Markdown table in under 1.5 s', () => {
  const lines = ['id\tname\tcity\tqty\tprice\tdate\tnote\tflag'];
  for (let i = 0; i < 10000; i++) lines.push(`${i}\tName ${i}\tCity ${i % 50}\t${i % 97}\t${(i * 1.25).toFixed(2)}\t2026-01-${String(1 + (i % 28)).padStart(2, '0')}\t"note ${i}\nsecond line"\t${i % 2 ? 'TRUE' : 'FALSE'}`);
  const text = lines.join('\r\n');
  const t0 = Date.now();
  const p = T.parseDelimited(text);
  const md = T.toMarkdown(p.rows);
  const ms = Date.now() - t0;
  assert.strictEqual(p.rows.length, 10001);
  assert.strictEqual(md.split('\n').length, 10003);
  assert.ok(ms < 1500, `${ms} ms`);
});

test('hostile: a 5 MB clipboard converts in under 3 s; over 10 MB gives a message, not a hang', () => {
  const row = 'alpha\tbeta\tgamma\t12345\t"quoted\tcell"\n';
  const text = row.repeat(Math.ceil((5 * 1024 * 1024) / row.length));
  const t0 = Date.now();
  const p = T.parseDelimited(text);
  const md = T.toMarkdown(p.rows);
  assert.ok(Date.now() - t0 < 3000, `${Date.now() - t0} ms`);
  assert.ok(md.length > 5 * 1024 * 1024);
  assert.throws(() => T.parseDelimited('x'.repeat(10 * 1024 * 1024 + 1)), (e) => e.code === 'TOO_BIG' && /Up to 10\.0 MB/.test(e.message));
});

test('hostile: 10,000 columns, 300,000 rows, an unclosed quote and control characters give messages or text, never a hang', () => {
  const wide = Array.from({ length: 10000 }, (_, i) => 'c' + i).join('\t');
  assert.throws(() => T.parseDelimited(wide), (e) => e.code === 'TOO_WIDE' && /2,000 columns/.test(e.message));
  assert.throws(() => T.parseDelimited('a\n'.repeat(300000)), (e) => e.code === 'TOO_MANY_ROWS');
  assert.throws(() => T.parseDelimited(('a\t'.repeat(1999) + 'a\n').repeat(1100)), (e) => e.code === 'TOO_MANY_CELLS');
  const open = T.parseDelimited('a\tb\n"never closed\tx\ny');
  assert.strictEqual(open.rows.length, 2);
  const ctl = T.toMarkdown([['h'], ['a\u0000b\u0007c']]);
  assert.ok(!ctl.includes('\u0000'));
  // a Markdown table with 100,000 rows of one pipe each parses without stalling
  const big = '| a |\n| - |\n' + '| x |\n'.repeat(100000);
  const t0 = Date.now();
  assert.strictEqual(T.parseMarkdownTable(big).rows.length, 100001);
  assert.ok(Date.now() - t0 < 3000);
});
