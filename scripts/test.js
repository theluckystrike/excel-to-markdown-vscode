'use strict';
// npm test: rebuilds dist/worker.js (when the sibling table-viewer-vscode tree is there), then runs every
// test/*.test.js with the Node test runner.
const { spawnSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
if (fs.existsSync(path.join(root, '..', 'table-viewer-vscode', 'node_modules', 'esbuild'))) {
  const b = spawnSync(process.execPath, [path.join(__dirname, 'build.js')], { stdio: 'inherit' });
  if (b.status !== 0) process.exit(1);
}
const dir = path.join(root, 'test');
const files = fs.readdirSync(dir).filter((f) => f.endsWith('.test.js')).map((f) => path.join(dir, f));
const env = { ...process.env, ZOVO_SPONSORS_MODULE: path.join(root, 'src', 'license', 'sponsors.js') };
const r = spawnSync(process.execPath, ['--test', ...files], { stdio: 'inherit', env });
process.exit(r.status === null ? 1 : r.status);
