'use strict';
// Bundles src/worker.js with SheetJS into dist/worker.js. esbuild and SheetJS come from the node_modules tree of the
// sibling product table-viewer-vscode (TABLE_VIEWER_DIR overrides it): this extension has no node_modules of its own.
const path = require('path');
const root = path.join(__dirname, '..');
const tv = process.env.TABLE_VIEWER_DIR || path.join(root, '..', 'table-viewer-vscode');
const esbuild = require(path.join(tv, 'node_modules', 'esbuild'));
esbuild.build({
  entryPoints: [path.join(root, 'src', 'worker.js')],
  outfile: path.join(root, 'dist', 'worker.js'),
  bundle: true, platform: 'node', format: 'cjs', target: 'node18', minify: true, legalComments: 'none',
  nodePaths: [path.join(tv, 'node_modules')], logLevel: 'warning',
  banner: { js: '/* Excel to Markdown worker. Bundles SheetJS Community Edition 0.20.3 (Apache-2.0, see THIRD-PARTY-NOTICES.txt). */' },
}).then(() => console.log('built dist/worker.js'), () => process.exit(1));
