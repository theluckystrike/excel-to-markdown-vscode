'use strict';
// Runs media/main.js in jsdom: the controls, the grid and the table text from host messages, HTML in cells kept as
// text, the Pro buttons (hidden by the setting), and no style attributes (the CSP drops them). jsdom is not a
// dependency of this extension: the test uses JSDOM_PATH or the copy in a sibling product, and is skipped without one.
const test = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');

const candidates = [process.env.JSDOM_PATH, path.join(__dirname, '..', '..', 'docx-viewer-vscode', 'node_modules', 'jsdom')].filter(Boolean);
let JSDOM = null;
for (const c of candidates) { try { ({ JSDOM } = require(c)); break; } catch { /* next */ } }
const script = fs.readFileSync(path.join(__dirname, '..', 'media', 'main.js'), 'utf8');

function boot(pro = { show: true, licensed: false, tries: { batch: 3, xlsx: 3, formats: 2 } }) {
  const dom = new JSDOM('<!DOCTYPE html><body><div id="app"></div><template id="zovo-footer"><div id="zovo-support">Rate</div></template></body>', { runScripts: 'outside-only', pretendToBeVisual: true });
  const sent = [];
  dom.window.acquireVsCodeApi = () => ({ postMessage: (m) => sent.push(m) });
  dom.window.eval(script);
  const send = (m) => dom.window.dispatchEvent(new dom.window.MessageEvent('message', { data: m }));
  send({ type: 'init', name: 'book.xlsx', sheets: [{ name: 'Prices', rows: 3, cols: 2 }, { name: '<img src=x>', rows: 0, cols: 0 }], state: { sheet: 0, range: '', header: true, format: 'markdown', formats: false }, pro });
  send({ type: 'output', state: { sheet: 0, range: '', header: true, format: 'markdown', formats: false }, range: 'A1:B3', rows: 3, cols: 2, r0: 0, c0: 0,
    grid: [['Item', 'Note'], ['Tea', '<img src=x onerror=alert(1)>'], ['Cake', '']], text: '| Item | Note |\n', textLen: 15 });
  return { dom, doc: dom.window.document, sent, send };
}

test('webview: controls, grid and table text; HTML in cells and sheet names stays text; no style attributes', { skip: !JSDOM && 'jsdom not found' }, () => {
  const { doc, sent } = boot();
  assert.strictEqual(sent[0].type, 'ready');
  assert.ok(sent.some((m) => m.type === 'rendered'));
  assert.strictEqual(doc.querySelector('.fname').textContent, 'book.xlsx');
  assert.deepStrictEqual([...doc.querySelectorAll('select.sheet option')].map((o) => o.textContent), ['Prices', '<img src=x> (empty)']);
  assert.deepStrictEqual([...doc.querySelectorAll('.grid thead th')].map((t) => t.textContent), ['', 'A', 'B']);
  assert.strictEqual(doc.querySelectorAll('.grid tbody tr').length, 3);
  assert.strictEqual(doc.querySelectorAll('#app img').length, 0);
  assert.match(doc.querySelector('.grid').textContent, /<img src=x onerror=alert\(1\)>/);
  assert.strictEqual(doc.querySelector('pre.out').textContent, '| Item | Note |\n');
  assert.match(doc.querySelector('.info').textContent, /Prices, A1:B3: 3 rows, 2 columns/);
  assert.strictEqual(doc.querySelectorAll('[style]').length, 0);
  assert.ok(doc.getElementById('zovo-support'));
  assert.match(doc.querySelector('[data-tool=formats]').title, /2 free tries left/);
});

test('webview: changing the sheet, range, header and format posts render messages; Copy, Insert and Pro post theirs', { skip: !JSDOM && 'jsdom not found' }, async () => {
  const { dom, doc, sent } = boot();
  const w = dom.window;
  const sel = doc.querySelector('select.sheet'); sel.value = '1'; sel.dispatchEvent(new w.Event('change'));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(sent.at(-1))), { type: 'render', sheet: 1, range: '' });
  const r = doc.querySelector('input.range'); r.value = 'A1:B2'; r.dispatchEvent(new w.KeyboardEvent('keydown', { key: 'Enter' }));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(sent.at(-1))), { type: 'render', range: 'A1:B2' });
  const h = doc.querySelector('input.hdr'); h.checked = false; h.dispatchEvent(new w.Event('change'));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(sent.at(-1))), { type: 'render', header: false });
  const f = doc.querySelector('select.format'); f.value = 'jira'; f.dispatchEvent(new w.Event('change'));
  assert.deepStrictEqual(JSON.parse(JSON.stringify(sent.at(-1))), { type: 'render', format: 'jira' });
  doc.querySelector('button.copy').click(); assert.deepStrictEqual(JSON.parse(JSON.stringify(sent.at(-1))), { type: 'copy' });
  doc.querySelector('button.insert').click(); assert.deepStrictEqual(JSON.parse(JSON.stringify(sent.at(-1))), { type: 'insert' });
  doc.querySelector('[data-tool=formats]').click(); assert.deepStrictEqual(JSON.parse(JSON.stringify(sent.at(-1))), { type: 'tool', tool: 'formats', on: true });
  doc.querySelector('[data-tool=workbook]').click(); assert.deepStrictEqual(JSON.parse(JSON.stringify(sent.at(-1))), { type: 'tool', tool: 'workbook' });
});

test('webview: showProTools off hides the Pro buttons; an error message replaces the view', { skip: !JSDOM && 'jsdom not found' }, () => {
  const { doc, send } = boot({ show: false, licensed: false, tries: {} });
  assert.ok(doc.querySelector('.protools').classList.contains('hidden'));
  send({ type: 'error', message: 'The part xl/a.xml unpacks to more than it claims.' });
  assert.strictEqual(doc.querySelector('.error h2').textContent, 'This file could not be read');
  assert.match(doc.querySelector('.error p').textContent, /unpacks to more/);
});
