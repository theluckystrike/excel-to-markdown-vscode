'use strict';
// The clipboard the extension reads and writes. Normally vscode.env.clipboard. Real-app tests start the editor with
// ZOVO_E2M_TEST_CLIPBOARD=<absolute file path>: then this file stands in for the clipboard, so a test run never reads
// or overwrites the system clipboard of the person at the computer. Nothing else changes.
const fs = require('fs');
const path = require('path');

const MAX_CLIP = 64 * 1024 * 1024;

function makeClipboard(vscode, env = process.env) {
  const f = env.ZOVO_E2M_TEST_CLIPBOARD;
  if (f && path.isAbsolute(f)) {
    return {
      test: true,
      readText: async () => { try { const st = fs.statSync(f); if (st.size > MAX_CLIP) return ''; return fs.readFileSync(f, 'utf8'); } catch { return ''; } },
      writeText: async (t) => { fs.writeFileSync(f, String(t)); },
    };
  }
  return { test: false, readText: () => vscode.env.clipboard.readText(), writeText: (t) => vscode.env.clipboard.writeText(t) };
}

module.exports = { makeClipboard };
