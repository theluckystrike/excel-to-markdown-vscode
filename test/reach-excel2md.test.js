'use strict';
// NEWSHIP: REACH1 wiring of Excel to Markdown (preview row on its own line, Explorer entries, the Keep formats command).
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { buildHtml } = require('../src/html');
const pkg = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'package.json'), 'utf8'));
const cfg = require('../src/reach-config');

test('preview HTML: the REACH1 head (style + boot script) goes into <head> with the page nonce, before main.js', () => {
  const html = buildHtml({ cspSource: 'vscode-resource:', scriptUri: 'x/main.js', styleUri: 'x/main.css', title: 't', nonce: 'N0nce', head: '<style nonce="N0nce">.zovo-reach1{}</style><script nonce="N0nce">1</script>' });
  const head = html.split('</head>')[0];
  assert.ok(head.includes('<script nonce="N0nce">1</script>'));
  assert.ok(html.indexOf('<script nonce="N0nce">1</script>') < html.indexOf('x/main.js'));
  assert.match(html, /script-src 'nonce-N0nce'/);
});

test('the row slot is its own line under the preview toolbar', () => {
  const js = fs.readFileSync(path.join(__dirname, '..', 'media', 'main.js'), 'utf8');
  assert.equal(cfg.slot, '.reach1-line');
  assert.ok(js.includes("app.appendChild(el('div', 'reach1-line'))"));
  assert.ok(js.indexOf("app.appendChild(top);") < js.indexOf("app.appendChild(el('div', 'reach1-line'))"));
});

test('Explorer: one "Zovo Pro:" entry per tool and file type, all hidden with the Pro tools setting off', () => {
  const ex = pkg.contributes.menus['explorer/context'].filter((m) => m.command.startsWith('excelToMarkdown.zovoPro.'));
  assert.deepEqual([...new Set(ex.map((m) => m.command))].sort(), ['excelToMarkdown.zovoPro.batch', 'excelToMarkdown.zovoPro.formats', 'excelToMarkdown.zovoPro.xlsx']);
  for (const m of ex) assert.ok(m.when.includes('config.excelToMarkdown.showProTools'), m.when);
  assert.ok(ex.some((m) => m.command.endsWith('.xlsx') && /\.md\b/.test(m.when)));
  assert.ok(!pkg.contributes.menus['explorer/context'].some((m) => m.command === 'excelToMarkdown.workbookToMarkdown'), 'old entry replaced');
  const kf = pkg.contributes.commands.find((c) => c.command === 'excelToMarkdown.keepFormats');
  assert.ok(kf && / \(Pro\)$/.test(kf.title));
});
