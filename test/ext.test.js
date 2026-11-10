'use strict';
// The extension host with a stub vscode: activation with no network and no message, the paste and convert commands in
// a text editor, the workbook preview, the 3 Pro tools (tries counted only after output, Escape keeps a try, the panel
// on the 4th use with the module's buttons), the showProTools setting, the CSP and the test clipboard file.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const http = require('http');
const https = require('https');
const net = require('net');
const XLSX = require(path.join(__dirname, '..', '..', 'table-viewer-vscode', 'node_modules', 'xlsx'));
const { loadExtension, openPreview, Selection } = require('./support/vscode-stub');
const { buildHtml } = require('../src/html');
const { makeClipboard } = require('../src/clip');

const CUR = '‸'; // cursor marker for openEditor
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'e2m-ext-'));
test.after(() => fs.rmSync(dir, { recursive: true, force: true }));
const put = (name, buf) => { const p = path.join(dir, name); fs.mkdirSync(path.dirname(p), { recursive: true }); fs.writeFileSync(p, buf); return p; };
function book(rows, name = 'Sheet1') { const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name); return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }); }
const SAMPLE = path.join(__dirname, '..', 'samples', 'sample.xlsx');
const BAD = put('bad.xlsx', Buffer.from('not a zip at all'));

function trapNetwork() {
  const calls = [];
  const saved = { fetch: global.fetch, hr: http.request, hg: http.get, sr: https.request, sg: https.get, nc: net.connect, ncc: net.createConnection };
  global.fetch = async (u) => { calls.push(String(u)); throw new Error('network blocked in test'); };
  http.request = http.get = https.request = https.get = (...a) => { calls.push(String(a[0] && (a[0].href || a[0].host || a[0]))); throw new Error('network blocked in test'); };
  net.connect = net.createConnection = (...a) => { calls.push('net ' + JSON.stringify(a[0])); throw new Error('network blocked in test'); };
  return { calls, restore: () => { global.fetch = saved.fetch; http.request = saved.hr; http.get = saved.hg; https.request = saved.sr; https.get = saved.sg; net.connect = saved.nc; net.createConnection = saved.ncc; } };
}

function boot() { const s = loadExtension(); s.ext.activate(s.context); return s; }
const run = (s, id, ...a) => s.vscode.commands.executeCommand(`excelToMarkdown.${id}`, ...a);

test('activation registers the commands and makes no network call, no message, no Pro text', async () => {
  const nw = trapNetwork();
  const s = loadExtension();
  try {
    s.ext.activate(s.context);
    for (const c of ['paste', 'csvToTable', 'tableToCsv', 'tableToTsv', 'formatTable', 'openSpreadsheet', 'openSample', 'addKeybinding', 'workbookToMarkdown', 'tablesToXlsx', 'getPro', 'payWithPayPal', 'enterLicenseKey', 'payByCard']) assert.ok(s.commands.has(`excelToMarkdown.${c}`), c);
    assert.ok(s.providers.has('excelToMarkdown.preview'));
    const v = await openPreview(s, SAMPLE);
    assert.strictEqual(v.last('init').sheets.length, 3);
    assert.deepStrictEqual(v.last('init').pro.tries, { batch: 3, xlsx: 3, formats: 3 });
    assert.deepStrictEqual(nw.calls, []);
    assert.deepStrictEqual([s.log.info.length, s.log.warn.length, s.log.error.length], [0, 0, 0]);
    assert.deepStrictEqual(s.log.external, []);
    assert.ok(!s.log.status.some((t) => /Pro/.test(t)));
  } finally { nw.restore(); s.restore(); }
});

test('paste: Excel cells at the cursor become a table on lines of their own; the clipboard is only read', async () => {
  const s = boot();
  try {
    s.clipboard.text = 'Item\tQty\r\n"Tea\nleaves"\t3\r\nCake|slice\t12\r\n';
    const ed = s.openEditor(`# Notes\nSee ${CUR}here\n`);
    await run(s, 'paste');
    assert.strictEqual(ed.document.text, '# Notes\nSee \n\n| Item          | Qty |\n| ------------- | --: |\n| Tea<br>leaves |   3 |\n| Cake\\|slice   |  12 |\n\nhere\n');
    assert.deepStrictEqual(s.log.clipboardWrites, []);
    assert.match(s.log.status.at(-1), /2 columns, 3 rows/);
    // no editor: a new Markdown file with the table
    s.vscode.window.activeTextEditor = undefined; s.vscode.window.visibleTextEditors = [];
    await run(s, 'paste');
    assert.match(s.log.opened.at(-1).getText(), /^\| Item +\| Qty \|/);
    // empty clipboard: a plain message
    s.clipboard.text = '   ';
    s.openEditor('x');
    await run(s, 'paste');
    assert.match(s.log.info.at(-1).text, /clipboard has no text/);
    // 10,000 columns: a warning, nothing inserted
    s.clipboard.text = Array.from({ length: 10000 }, (_, i) => 'c' + i).join('\t');
    const ed2 = s.openEditor('keep');
    await run(s, 'paste');
    assert.match(s.log.warn.at(-1).text, /2,000 columns/);
    assert.strictEqual(ed2.document.text, 'keep');
  } finally { s.restore(); }
});

test('paste follows the settings: alignment, header row, compact, line breaks', async () => {
  const s = boot();
  try {
    Object.assign(s.config, { alignment: 'left', headerRow: 'empty', compact: true, lineBreaks: 'space' });
    s.clipboard.text = 'a\tb\n"x\ny"\t2';
    const ed = s.openEditor('');
    await run(s, 'paste');
    assert.strictEqual(ed.document.text, '|  |  |\n| :-- | :-- |\n| a | b |\n| x y | 2 |\n');
  } finally { s.restore(); }
});

test('CSV selection -> Markdown in place; a whole .csv file -> a new Markdown file; no selection elsewhere -> a hint', async () => {
  const s = boot();
  try {
    const ed = s.openEditor('before\nid,name\n1,"Smith, Ann"\nafter');
    ed.selection = new Selection(1, 0, 2, 15);
    await run(s, 'csvToTable');
    assert.strictEqual(ed.document.text, 'before\n|  id | name       |\n| --: | ---------- |\n|   1 | Smith, Ann |\nafter');
    const csv = s.openEditor('a\tb\n1\t2\n', { fileName: '/tmp/x.tsv', languageId: 'plaintext' });
    await run(s, 'csvToTable');
    assert.strictEqual(s.log.opened.at(-1).getText(), '|   a |   b |\n| --: | --: |\n|   1 |   2 |\n');
    assert.strictEqual(csv.document.text, 'a\tb\n1\t2\n', 'the CSV file itself is not changed');
    s.openEditor('plain markdown');
    await run(s, 'csvToTable');
    assert.match(s.log.info.at(-1).text, /Select the CSV or TSV text/);
  } finally { s.restore(); }
});

test('Markdown table under the cursor -> CSV and TSV; Format lines it up; outside a table -> a hint', async () => {
  const s = boot();
  try {
    const doc = `Text above\n\n|Name|Note|\n|-|-|\n|Tea|a\\|b ${CUR}<br>two|\n|Cake||\n\nText below`;
    const ed = s.openEditor(doc);
    await run(s, 'tableToCsv');
    assert.strictEqual(ed.document.text, 'Text above\n\nName,Note\nTea,"a|b \ntwo"\nCake,\n\nText below');
    const ed2 = s.openEditor(`|a|b|\n|-|-|\n|1|${CUR}x|`);
    await run(s, 'tableToTsv');
    assert.strictEqual(ed2.document.text, 'a\tb\n1\tx');
    const ed3 = s.openEditor(`intro\n|Item|Qty|\n|:-|-:|\n|Green tea|3|${CUR}\n|Cake|12|\noutro`);
    await run(s, 'formatTable');
    assert.strictEqual(ed3.document.text, 'intro\n| Item      | Qty |\n| :-------- | --: |\n| Green tea |   3 |\n| Cake      |  12 |\noutro');
    s.openEditor(`no table ${CUR}here`);
    await run(s, 'formatTable');
    assert.match(s.log.info.at(-1).text, /Put the cursor in a Markdown table/);
  } finally { s.restore(); }
});

test('preview: sheets, range, header, formats of output; Copy writes the full text; Insert goes to the last editor', async () => {
  const s = boot();
  try {
    const target = s.openEditor('# Report\n');
    const v = await openPreview(s, SAMPLE);
    const init = v.last('init');
    assert.deepStrictEqual(init.sheets.map((x) => x.name), ['Prices', 'Quarterly', 'Team']);
    let out = v.last('output');
    assert.strictEqual(out.rows, 5);
    assert.match(out.text, /\| Oolong \\\| first flush \| +14\.25 \|/);
    assert.deepStrictEqual(out.grid[0], ['Product', 'Price', 'In stock', 'Margin', 'Launched', 'Page']);
    await v.send({ type: 'render', range: 'A1:B3' });
    out = v.last('output');
    assert.strictEqual(out.range, 'A1:B3');
    assert.strictEqual(out.text.trim().split('\n').length, 4);
    await v.send({ type: 'render', sheet: 2, range: '' });
    assert.match(v.last('output').text, /\| 山田 太郎 +\| 東京 +\|/);
    await v.send({ type: 'render', format: 'html' });
    assert.match(v.last('output').text, /^<table>/);
    await v.send({ type: 'render', format: 'markdown', header: false });
    assert.match(v.last('output').text, /^\| +\| +\| +\| +\|\n/);
    await v.send({ type: 'render', header: true, range: 'bad range!' });
    assert.match(v.last('output').error, /not a range/);
    await v.send({ type: 'render', range: '' });
    await v.send({ type: 'copy' });
    assert.match(s.log.clipboardWrites.at(-1), /^\| Name +\| Office/);
    // the preview is a custom editor, so the active text editor is gone; Insert uses the last one
    s.vscode.window.activeTextEditor = undefined;
    await v.send({ type: 'insert' });
    assert.match(target.document.text, /^# Report\n\| Name +\| Office/);
    // the text editor is hidden behind the preview: Insert brings it back and inserts at its cursor
    s.vscode.window.activeTextEditor = undefined; s.vscode.window.visibleTextEditors = [];
    await v.send({ type: 'insert' });
    assert.strictEqual((target.document.text.match(/\| Name +\| Office/g) || []).length, 2);
    assert.strictEqual(s.log.opened.at(-1), target.document);
    await v.send({ type: 'rendered' });
  } finally { s.restore(); }
});

test('preview of a broken workbook shows a message; the host keeps working', async () => {
  const s = boot();
  try {
    const v = await openPreview(s, BAD);
    assert.match(v.last('error').message, /not a zip archive/);
    const v2 = await openPreview(s, SAMPLE);
    assert.strictEqual(v2.last('init').sheets.length, 3);
  } finally { s.restore(); }
});

const PANEL_BUTTONS = ['Pay by card (USD 19)', 'Get Zovo Pro with GitHub Sponsors', 'Pay with PayPal', 'Enter a key'];

test('Pro "Keep Excel formats": a try counts after the formatted table is shown; the 4th turn-on shows the panel', async () => {
  const s = boot();
  try {
    const v = await openPreview(s, SAMPLE);
    for (let i = 1; i <= 3; i++) {
      await v.send({ type: 'tool', tool: 'formats', on: true });
      const out = v.last('output');
      assert.ok(out.state.formats);
      assert.match(out.text, /6\.50 USD/);
      assert.match(out.text, /\[tea\]\(https:\/\/example\.com\/shop\/tea\)/);
      assert.match(s.log.status.at(-1), new RegExp(`free try ${i} of 3`));
      await v.send({ type: 'tool', tool: 'formats', on: false });
      assert.ok(!v.last('output').state.formats);
    }
    // merged cells filled while it was on (checked on try 3 via the Quarterly sheet)
    const before = s.log.info.length;
    s.ui.info.push(undefined); // Escape on the panel
    await v.send({ type: 'tool', tool: 'formats', on: true });
    const panel = s.log.info[before];
    assert.strictEqual(panel.text, 'Keep Excel formats, links and merged cells is part of Zovo Pro.');
    assert.ok(panel.opts.modal);
    assert.deepStrictEqual(panel.buttons, PANEL_BUTTONS);
    assert.ok(!/refund/i.test(panel.opts.detail));
    assert.strictEqual(v.last('pro').formats, false);
    assert.ok(!v.last('output').state.formats);
  } finally { s.restore(); }
});

test('Pro "whole workbook": Escape on the save dialog keeps the try; 3 saves; the 4th click shows the panel; Sponsors opens tier 679878', async () => {
  const s = boot();
  try {
    const v = await openPreview(s, SAMPLE);
    s.ui.save.push(undefined);
    await v.send({ type: 'tool', tool: 'workbook' });
    assert.deepStrictEqual(s.log.saved, []);
    assert.strictEqual(s.state.get('excelToMarkdown.proTrial'), undefined, 'Escape uses no try');
    for (let i = 1; i <= 3; i++) {
      const out = path.join(dir, `book-${i}.md`);
      s.ui.save.push(s.vscode.Uri.file(out));
      await v.send({ type: 'tool', tool: 'workbook' });
      const md = fs.readFileSync(out, 'utf8');
      assert.match(md, /^# sample\.xlsx\n\n## Prices\n\n\| Product/);
      assert.match(md, /## Quarterly/);
      assert.match(md, /## Team/);
      assert.strictEqual(s.state.get('excelToMarkdown.proTrial').batch, i);
    }
    const n = s.log.info.length;
    s.ui.info.push('Get Zovo Pro with GitHub Sponsors');
    await v.send({ type: 'tool', tool: 'workbook' });
    assert.strictEqual(s.log.info[n].text, 'Convert whole workbooks to Markdown files is part of Zovo Pro.');
    assert.deepStrictEqual(s.log.info[n].buttons, PANEL_BUTTONS);
    assert.match(s.log.info[n].opts.detail, /USD 19/);
    assert.ok(s.log.external.some((u) => u === 'https://github.com/sponsors/theluckystrike/sponsorships?tier_id=679878'), JSON.stringify(s.log.external));
    assert.strictEqual(fs.existsSync(path.join(dir, 'book-4.md')), false);
  } finally { s.restore(); }
});

test('Pro "whole workbook" on a folder: every workbook gets a .md file; existing files ask first; a bad file is reported', async () => {
  const s = boot();
  try {
    const root = path.join(dir, 'folder');
    put('folder/a.xlsx', book([['x', 'y'], [1, 2]], 'First'));
    put('folder/sub/b.ods', fs.readFileSync(SAMPLE));
    put('folder/sub/b.ods', XLSX.write(XLSX.read(fs.readFileSync(SAMPLE)), { type: 'buffer', bookType: 'ods' }));
    put('folder/node_modules/skip.xlsx', book([['no']]));
    put('folder/broken.xlsx', Buffer.from('broken'));
    put('folder/~$a.xlsx', Buffer.from('lock file'));
    await run(s, 'workbookToMarkdown', s.vscode.Uri.file(root));
    assert.match(fs.readFileSync(path.join(root, 'a.md'), 'utf8'), /^# a\.xlsx\n\n## First\n\n\| +x \| +y \|/);
    assert.match(fs.readFileSync(path.join(root, 'sub', 'b.md'), 'utf8'), /## Quarterly/);
    assert.ok(!fs.existsSync(path.join(root, 'node_modules', 'skip.md')));
    assert.match(s.log.info.at(-1).text, /Wrote 2 Markdown files \(4 sheets\); 1 workbook could not be read \(broken\.xlsx/);
    assert.strictEqual(s.state.get('excelToMarkdown.proTrial').batch, 1);
    // second run: the .md files exist; Escape on the question uses no try
    s.ui.warn.push(undefined);
    await run(s, 'workbookToMarkdown', s.vscode.Uri.file(root));
    assert.match(s.log.warn.at(-1).text, /2 of the 3 Markdown files exist already/);
    assert.strictEqual(s.state.get('excelToMarkdown.proTrial').batch, 1);
  } finally { s.restore(); }
});

test('Pro "whole workbook" from the palette with a preview open: it asks for files or a folder (starting next to the open workbook); Escape uses no try; a picked folder converts', async () => {
  const s = boot();
  try {
    const root = path.join(dir, 'palette');
    put('palette/q1.xlsx', book([['a', 'b'], [1, 2]], 'Q1'));
    put('palette/q2.xlsx', book([['c'], [3]], 'Q2'));
    await openPreview(s, SAMPLE);
    s.ui.open.push(undefined);
    await run(s, 'workbookToMarkdown');
    assert.strictEqual(s.log.openDialogs.length, 1, 'the palette command asks instead of taking the open workbook');
    assert.strictEqual(s.log.openDialogs[0].defaultUri.fsPath, path.dirname(SAMPLE));
    assert.ok(s.log.openDialogs[0].canSelectFolders);
    assert.ok(!fs.existsSync(SAMPLE.replace(/\.xlsx$/, '.md')));
    assert.strictEqual(s.state.get('excelToMarkdown.proTrial'), undefined, 'Escape uses no try');
    s.ui.open.push([s.vscode.Uri.file(root)]);
    await run(s, 'workbookToMarkdown');
    assert.match(fs.readFileSync(path.join(root, 'q1.md'), 'utf8'), /## Q1/);
    assert.match(fs.readFileSync(path.join(root, 'q2.md'), 'utf8'), /## Q2/);
    assert.match(s.log.info.at(-1).text, /^Wrote 2 Markdown files \(2 sheets\)\.$/);
    assert.strictEqual(s.state.get('excelToMarkdown.proTrial').batch, 1);
  } finally { s.restore(); }
});

test('Pro "Markdown tables to .xlsx": no table uses no try, Escape keeps it, 3 saves, the 4th shows the panel', async () => {
  const s = boot();
  try {
    s.openEditor('no tables here');
    await run(s, 'tablesToXlsx');
    assert.match(s.log.info.at(-1).text, /No Markdown table/);
    const md = '# Prices\n\n| Item | Price | Share |\n|---|--:|--:|\n| Tea | 6.5 | 25% |\n| A\\|B | 12 | |\n\n## Team\n\n| Name |\n|---|\n| <b>Ann</b> |\n';
    s.openEditor(md, { fileName: path.join(dir, 'notes.md'), uri: s.vscode.Uri.file(path.join(dir, 'notes.md')) });
    s.ui.save.push(undefined);
    await run(s, 'tablesToXlsx');
    assert.strictEqual(s.state.get('excelToMarkdown.proTrial'), undefined);
    for (let i = 1; i <= 3; i++) {
      const out = path.join(dir, `tables-${i}.xlsx`);
      s.ui.save.push(s.vscode.Uri.file(out));
      await run(s, 'tablesToXlsx');
      const wb = XLSX.read(fs.readFileSync(out));
      assert.deepStrictEqual(wb.SheetNames, ['Prices', 'Team']);
      assert.strictEqual(wb.Sheets.Prices.B2.v, 6.5);
      assert.strictEqual(wb.Sheets.Prices.A3.v, 'A|B');
      assert.strictEqual(wb.Sheets.Team.A2.v, '<b>Ann</b>');
      assert.strictEqual(s.state.get('excelToMarkdown.proTrial').xlsx, i);
    }
    const n = s.log.info.length;
    await run(s, 'tablesToXlsx');
    assert.strictEqual(s.log.info[n].text, 'Save Markdown tables as an Excel file is part of Zovo Pro.');
    assert.match(s.log.info[n].opts.detail, /Convert whole workbooks to Markdown files/);
  } finally { s.restore(); }
});

test('showProTools off hides the Pro buttons in the preview state', async () => {
  const s = boot();
  try {
    s.config.showProTools = false;
    const v = await openPreview(s, SAMPLE);
    assert.strictEqual(v.last('init').pro.show, false);
  } finally { s.restore(); }
});

test('the webview HTML has a nonce CSP, no remote origin and no inline script', () => {
  const html = buildHtml({ cspSource: 'vscode-resource:', scriptUri: 'vscode-resource:/m.js', styleUri: 'vscode-resource:/m.css', title: '<b>x</b>.xlsx', nonce: 'abc' });
  assert.match(html, /default-src 'none'; style-src vscode-resource: 'nonce-abc'; script-src 'nonce-abc'/);
  assert.ok(!/https?:\/\//.test(html.replace(/<template id="zovo-footer">[\s\S]*?<\/template>/, '')), 'no remote URL outside the footer links');
  assert.ok(!/unsafe-inline|unsafe-eval/.test(html));
  assert.match(html, /<title>&lt;b&gt;x&lt;\/b&gt;\.xlsx<\/title>/);
  assert.strictEqual((html.match(/<script/g) || []).length, 1);
});

test('test clipboard file: only with an absolute path in ZOVO_E2M_TEST_CLIPBOARD; otherwise vscode.env.clipboard', async () => {
  const f = path.join(dir, 'clip.txt');
  const c = makeClipboard({ env: { clipboard: { readText: async () => 'SYSTEM', writeText: async () => { throw new Error('must not write'); } } } }, { ZOVO_E2M_TEST_CLIPBOARD: f });
  assert.strictEqual(c.test, true);
  assert.strictEqual(await c.readText(), '');
  await c.writeText('a\tb');
  assert.strictEqual(fs.readFileSync(f, 'utf8'), 'a\tb');
  assert.strictEqual(await c.readText(), 'a\tb');
  const real = makeClipboard({ env: { clipboard: { readText: async () => 'SYSTEM', writeText: async () => {} } } }, { ZOVO_E2M_TEST_CLIPBOARD: 'relative.txt' });
  assert.strictEqual(real.test, false);
  assert.strictEqual(await real.readText(), 'SYSTEM');
});

test('package.json: no default keybinding on Cmd/Ctrl+V, Pro commands hidden by the setting, engines 1.90', () => {
  const pk = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'package.json'), 'utf8'));
  assert.ok(!pk.contributes.keybindings || !pk.contributes.keybindings.some((k) => /(^|\+)v$/i.test(k.key || '') || /(^|\+)v$/i.test(k.mac || '')));
  assert.strictEqual(pk.engines.vscode, '^1.90.0');
  assert.strictEqual(pk.devDependencies['@types/vscode'], '~1.90.0');
  assert.ok(!pk.dependencies);
  const pal = Object.fromEntries(pk.contributes.menus.commandPalette.map((m) => [m.command, m.when]));
  assert.strictEqual(pal['excelToMarkdown.workbookToMarkdown'], 'config.excelToMarkdown.showProTools');
  assert.strictEqual(pal['excelToMarkdown.tablesToXlsx'], 'config.excelToMarkdown.showProTools');
  assert.ok(pk.displayName.startsWith('Excel to Markdown Table '));
  assert.notStrictEqual(pk.displayName.toLowerCase(), 'excel to markdown table');
  assert.ok(pk.description.startsWith('Excel to Markdown table'));
  const steps = pk.contributes.walkthroughs[0].steps;
  assert.strictEqual(steps.at(-1).id, 'zovo-related');
  for (const st of steps) assert.ok(fs.existsSync(path.join(__dirname, '..', st.media.markdown)), st.media.markdown);
});
