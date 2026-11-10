'use strict';
// Excel to Markdown Table: paste cells from Excel, Google Sheets, Numbers or LibreOffice as a Markdown table, convert
// CSV/TSV and Markdown tables both ways, format tables, and preview workbooks (.xlsx, .xls, .ods, .csv) to copy a sheet
// or range. Workbooks are read in a worker thread (runner.js) with time and memory caps. No network.
const vscode = require('vscode');
const fs = require('fs');
const path = require('path');
const T = require('./core/table');
const { sheetRows } = require('./core/sheet');
const { runJob } = require('./runner');
const { buildHtml } = require('./html');
const { createPro } = require('./pro');
const { makeClipboard } = require('./clip');
const { createRatingPrompt } = require('./rate');
const { registerSupport, commandUris, showLinks } = require('./support');

const EXT_ID = 'zovo.zovo-excel-to-markdown';
const PREFIX = 'excelToMarkdown';
const VIEW_TYPE = 'excelToMarkdown.preview';
const MAX_REMOTE_BYTES = 100 * 1024 * 1024;
const WORKBOOK_RE = /\.(xlsx|xlsm|xlsb|xls|ods)$/i;
const SHEET_FILE_RE = /\.(xlsx|xlsm|xlsb|xls|ods|csv|tsv)$/i;
const SKIP_DIRS = new Set(['node_modules', '.git', '.svn', '.hg', '__pycache__', '.venv', '.cache']);
const GRID_ROWS = 200, GRID_COLS = 50, SHOW_CHARS = 200000;

let rating = { success() {}, failure() {} };
let pro;
let clip;
let lastText = null; // { doc, sel, column } of the last text editor the user worked in
const sampleDocs = new WeakSet(); // the walkthrough's try-it file: work done there does not count for the rating ask
const ratingArg = (ed) => (ed && sampleDocs.has(ed.document) ? { sample: true } : undefined);
const viewers = new Set();
let active = null;

function setting(key, dflt) { try { return vscode.workspace.getConfiguration(PREFIX).get(key, dflt); } catch { return dflt; } }
const baseName = (uri) => path.posix.basename(uri.path) || uri.toString();

/** Table options from the settings. */
function tableOptions(over = {}) {
  return {
    header: setting('headerRow', 'first') === 'empty' ? 'empty' : 'first',
    align: setting('alignment', 'auto'),
    maxWidth: Math.max(0, Number(setting('maxColumnWidth', 40)) || 0),
    compact: !!setting('compact', false),
    escapeHtml: setting('escapeHtml', true) !== false,
    escapeMarkdown: !!setting('escapeMarkdown', false),
    lineBreak: setting('lineBreaks', '<br>') === 'space' ? 'space' : '<br>',
    ...over,
  };
}

// ---------------- text editor commands ----------------

function remember(ed) { if (ed && ed.document) lastText = { doc: ed.document, sel: ed.selection, column: ed.viewColumn }; }

/** The editor to insert into: the active text editor, else the last one (brought back to the front when `reveal`). */
async function targetEditor(reveal) {
  const ed = vscode.window.activeTextEditor;
  if (ed) return ed;
  if (!lastText || lastText.doc.isClosed) return null;
  const vis = (vscode.window.visibleTextEditors || []).find((e) => e.document === lastText.doc);
  if (vis) return vis;
  if (!reveal) return null;
  return vscode.window.showTextDocument(lastText.doc, { viewColumn: lastText.column, selection: lastText.sel, preview: false });
}

/** Inserts a block (a table) at every cursor of the editor, on lines of its own. Opens a new Markdown file if none. */
async function insertBlock(text, editor, reveal = true) {
  const ed = editor || (reveal ? await targetEditor(true) : null);
  if (!ed) {
    const doc = await vscode.workspace.openTextDocument({ language: 'markdown', content: text });
    await vscode.window.showTextDocument(doc, { preview: false });
    return 'new';
  }
  const doc = ed.document;
  const ok = await ed.edit((eb) => {
    for (const sel of ed.selections) {
      const before = doc.lineAt(sel.start.line).text.slice(0, sel.start.character);
      const after = doc.lineAt(sel.end.line).text.slice(sel.end.character);
      let t = text.endsWith('\n') ? text : text + '\n';
      if (before.trim()) t = '\n\n' + t;
      if (after.trim()) t = t + '\n';
      eb.replace(sel, t);
    }
  });
  if (!ok) throw new Error('The editor did not accept the change (it may be read-only).');
  return 'inserted';
}

async function pasteAsTable() {
  const text = await clip.readText();
  if (!text || !text.trim()) {
    vscode.window.showInformationMessage('The clipboard has no text to turn into a table. Copy cells in Excel, Google Sheets, Numbers or LibreOffice first.');
    return;
  }
  let parsed;
  try { parsed = T.parseDelimited(text.replace(/\r?\n$/, '')); } catch (e) { vscode.window.showWarningMessage(e.message); return; }
  const md = T.toMarkdown(parsed.rows, tableOptions());
  if (!md) { vscode.window.showInformationMessage('The clipboard text has no cells.'); return; }
  const ed = await targetEditor(false);
  await insertBlock(md, ed, false);
  vscode.window.setStatusBarMessage(`Markdown table: ${parsed.columns} column${parsed.columns === 1 ? '' : 's'}, ${parsed.rows.length} row${parsed.rows.length === 1 ? '' : 's'}.`, 5000);
  rating.success(ratingArg(ed));
}

/** The Markdown table around the cursor, or the selection: { range, text } or null. */
function tableRange(ed) {
  const doc = ed.document;
  const sel = ed.selection;
  if (!sel.isEmpty) {
    const range = new vscode.Range(sel.start.line, 0, sel.end.line, doc.lineAt(sel.end.line).text.length);
    return { range, text: doc.getText(range) };
  }
  const lines = [];
  const from = Math.max(0, sel.active.line - 20000), to = Math.min(doc.lineCount - 1, sel.active.line + 20000);
  for (let i = from; i <= to; i++) lines.push(doc.lineAt(i).text);
  const t = T.findTableAt(lines, sel.active.line - from);
  if (!t) return null;
  const range = new vscode.Range(t.start + from, 0, t.end + from, doc.lineAt(t.end + from).text.length);
  return { range, text: doc.getText(range) };
}

const NO_TABLE = 'Put the cursor in a Markdown table (a header row, a --- row, then the rows), or select one.';

async function tableTo(delimiter) {
  const ed = vscode.window.activeTextEditor;
  if (!ed) { vscode.window.showInformationMessage('Open a Markdown file and put the cursor in a table first.'); return; }
  const loc = tableRange(ed);
  const t = loc && T.parseMarkdownTable(loc.text);
  if (!t) { vscode.window.showInformationMessage(NO_TABLE); return; }
  const out = T.toDelimited(t.rows, delimiter).replace(/\n$/, '');
  await ed.edit((eb) => eb.replace(loc.range, out));
  vscode.window.setStatusBarMessage(`Converted the table to ${delimiter === '\t' ? 'TSV' : 'CSV'}: ${t.rows.length} rows.`, 5000);
}

async function formatTable() {
  const ed = vscode.window.activeTextEditor;
  if (!ed) { vscode.window.showInformationMessage('Open a Markdown file and put the cursor in a table first.'); return; }
  const loc = tableRange(ed);
  const out = loc && T.formatMarkdownTable(loc.text, tableOptions());
  if (!out) { vscode.window.showInformationMessage(NO_TABLE); return; }
  const text = out.replace(/\n$/, '');
  if (text !== loc.text) await ed.edit((eb) => eb.replace(loc.range, text));
  vscode.window.setStatusBarMessage('Table formatted.', 3000);
}

async function csvToTable() {
  const ed = vscode.window.activeTextEditor;
  if (!ed) { vscode.window.showInformationMessage('Open a file and select CSV or TSV text first.'); return; }
  const doc = ed.document;
  const sel = ed.selection;
  const wholeFile = sel.isEmpty && (/\.(csv|tsv|tab|txt)$/i.test(doc.fileName || '') || /^(csv|tsv|plaintext)/i.test(doc.languageId));
  if (sel.isEmpty && !wholeFile) { vscode.window.showInformationMessage('Select the CSV or TSV text to convert.'); return; }
  const text = wholeFile ? doc.getText() : doc.getText(sel);
  let parsed;
  try { parsed = T.parseDelimited(text.replace(/\r?\n$/, ''), { delimiter: /\.tsv$/i.test(doc.fileName || '') ? '\t' : undefined }); } catch (e) { vscode.window.showWarningMessage(e.message); return; }
  const md = T.toMarkdown(parsed.rows, tableOptions());
  if (!md) { vscode.window.showInformationMessage('The selection has no cells.'); return; }
  if (wholeFile) {
    const out = await vscode.workspace.openTextDocument({ language: 'markdown', content: md });
    await vscode.window.showTextDocument(out, { preview: false, viewColumn: vscode.ViewColumn.Beside });
  } else {
    await ed.edit((eb) => eb.replace(sel, md.replace(/\n$/, '')));
  }
  vscode.window.setStatusBarMessage(`Markdown table: ${parsed.columns} columns, ${parsed.rows.length} rows.`, 5000);
  rating.success(ratingArg(ed));
}

// ---------------- workbook preview ----------------

async function readModel(uri) {
  if (uri.scheme === 'file') return runJob({ op: 'read', file: uri.fsPath, name: uri.fsPath });
  const st = await vscode.workspace.fs.stat(uri);
  if (st.size > MAX_REMOTE_BYTES) throw new Error(`This file is ${Math.round(st.size / 1048576)} MB on a remote or virtual file system; files there are read up to 100 MB.`);
  const data = await vscode.workspace.fs.readFile(uri);
  const buf = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
  return runJob({ op: 'read', buffer: buf, name: uri.path }, { transfer: [buf] });
}

const MEDIA_ERRORS = new Set(['TOO_BIG', 'EMPTY', 'ENOENT', 'NOT_FILE', 'PASSWORD']);

class Viewer {
  constructor(uri, panel) {
    this.uri = uri; this.panel = panel; this.name = baseName(uri); this.model = null; this.counted = false;
    this.state = { sheet: 0, range: '', header: setting('headerRow', 'first') !== 'empty', format: 'markdown', formats: false };
    this.output = '';
  }
  post(msg) { return this.panel.webview.postMessage(msg).then(undefined, () => false); }
  async load() {
    try {
      this.model = await readModel(this.uri);
      const sheets = this.model.sheets.map((s) => ({ name: s.name, ref: s.ref, rows: s.rows.length, cols: s.rows.reduce((m, r) => Math.max(m, r.length), 0), truncated: s.truncated }));
      const first = sheets.findIndex((s) => s.rows > 0);
      this.state.sheet = first < 0 ? 0 : first;
      await this.post({ type: 'init', name: this.name, sheets, state: this.state, kind: this.model.kind, pro: proState() });
      await this.render();
    } catch (e) {
      if (!MEDIA_ERRORS.has(e.code)) rating.failure();
      await this.post({ type: 'error', message: e.message || String(e), pro: proState() });
    }
  }
  rows() {
    const sheet = this.model.sheets[this.state.sheet] || this.model.sheets[0];
    return { sheet, ...sheetRows(sheet, { range: this.state.range, formats: this.state.formats }) };
  }
  /** The table text of the current sheet and range in the chosen format. */
  text(format = this.state.format) {
    const { rows } = this.rows();
    return T.render(format, rows, tableOptions({ header: this.state.header ? 'first' : 'empty' }));
  }
  async render() {
    if (!this.model) return;
    let r;
    try { r = this.rows(); } catch (e) { await this.post({ type: 'output', error: e.message, state: this.state }); return; }
    const out = T.render(this.state.format, r.rows, tableOptions({ header: this.state.header ? 'first' : 'empty' }));
    this.output = out;
    const cols = r.rows.reduce((m, x) => Math.max(m, x.length), 0);
    const grid = r.rows.slice(0, GRID_ROWS).map((x) => x.slice(0, GRID_COLS).map((c) => T.cellText(c)));
    const m = /^([A-Z]+)(\d+)/.exec(r.range || '');
    await this.post({
      type: 'output', state: this.state, range: r.range, rows: r.rows.length, cols, grid,
      r0: m ? Number(m[2]) - 1 : 0, c0: m ? T.colIndex(m[1]) : 0,
      text: out.length > SHOW_CHARS ? out.slice(0, SHOW_CHARS) : out, textLen: out.length, truncated: r.sheet.truncated || '',
    });
  }
  async onMessage(msg) {
    if (!msg || typeof msg.type !== 'string') return undefined;
    switch (msg.type) {
      case 'ready': return this.load();
      case 'rendered':
        if (!this.counted && this.model) { this.counted = true; rating.success(this.uri); }
        return undefined;
      case 'render': {
        const s = this.state;
        if (Number.isInteger(msg.sheet) && msg.sheet >= 0 && this.model && msg.sheet < this.model.sheets.length) s.sheet = msg.sheet;
        if (typeof msg.range === 'string') s.range = msg.range.slice(0, 40);
        if (typeof msg.header === 'boolean') s.header = msg.header;
        if (['markdown', 'html', 'asciidoc', 'jira', 'csv', 'tsv'].includes(msg.format)) s.format = msg.format;
        return this.render();
      }
      case 'copy': {
        if (!this.output) return undefined;
        await clip.writeText(this.output);
        vscode.window.setStatusBarMessage(`Copied the ${formatName(this.state.format)} table (${T.fmtInt(this.output.length)} characters).`, 4000);
        return this.post({ type: 'done', what: 'copy' });
      }
      case 'insert': {
        if (!this.output) return undefined;
        const r = await insertBlock(this.output);
        vscode.window.setStatusBarMessage(r === 'new' ? 'The table is in a new Markdown file.' : 'Table inserted at the cursor.', 4000);
        return this.post({ type: 'done', what: 'insert' });
      }
      case 'tool':
        if (msg.tool === 'formats') return toggleFormats(this, !!msg.on);
        if (msg.tool === 'workbook') return workbookFromViewer(this);
        return undefined;
      case 'settings': return vscode.commands.executeCommand('workbench.action.openSettings', `@ext:${EXT_ID}`);
      default: return undefined;
    }
  }
}

const formatName = (f) => ({ markdown: 'Markdown', html: 'HTML', asciidoc: 'AsciiDoc', jira: 'Jira', csv: 'CSV', tsv: 'TSV' }[f] || 'Markdown');

function proState() {
  return { show: setting('showProTools', true) !== false, licensed: !!(pro && pro.licensed), tries: pro ? pro.tries() : {} };
}
function refreshPro() { for (const v of viewers) v.post({ type: 'pro', pro: proState() }); }

class Provider {
  constructor(context) { this.context = context; }
  openCustomDocument(uri) { return { uri, dispose() {} }; }
  resolveCustomEditor(document, panel) {
    const media = vscode.Uri.joinPath(this.context.extensionUri, 'media');
    panel.webview.options = { enableScripts: true, enableCommandUris: commandUris(PREFIX), localResourceRoots: [media] };
    panel.webview.html = buildHtml({
      cspSource: panel.webview.cspSource,
      scriptUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'main.js')).toString(),
      styleUri: panel.webview.asWebviewUri(vscode.Uri.joinPath(media, 'main.css')).toString(),
      title: baseName(document.uri),
      showSupportLinks: showLinks(vscode, PREFIX),
    });
    const v = new Viewer(document.uri, panel);
    viewers.add(v);
    if (panel.active) active = v;
    panel.onDidChangeViewState(() => { if (panel.active) active = v; else if (active === v && !panel.visible) active = null; });
    panel.onDidDispose(() => { viewers.delete(v); if (active === v) active = null; });
    panel.webview.onDidReceiveMessage((m) => v.onMessage(m).then(undefined, (e) => v.post({ type: 'notice', text: String((e && e.message) || e) })));
  }
}

// ---------------- Pro tools ----------------

/** Pro 3: the preview shows cells as Excel shows them. One try per turn-on, counted after the table is on screen. */
async function toggleFormats(v, on) {
  if (!on) { v.state.formats = false; await v.render(); return; }
  if (!(await pro.allowed('formats'))) { v.state.formats = false; await v.post({ type: 'pro', pro: proState(), formats: false }); return; }
  v.state.formats = true;
  await v.render();
  await pro.produced('formats');
  refreshPro();
}

/** Markdown for every sheet of a workbook model. */
function workbookMarkdown(model, title, o = {}) {
  const parts = [`# ${title.replace(/[\r\n]+/g, ' ')}`, ''];
  let n = 0;
  for (const s of model.sheets) {
    const { rows } = sheetRows(s, { formats: !!o.formats });
    if (!rows.length) continue;
    n++;
    if (model.sheets.length > 1 || model.kind !== 'text') parts.push(`## ${s.name.replace(/[\r\n]+/g, ' ')}`, '');
    parts.push(T.toMarkdown(rows, tableOptions(o.header === false ? { header: 'empty' } : {})).replace(/\n$/, ''), '');
    if (s.truncated) parts.push(`_Only the ${s.truncated} of this sheet were converted._`, '');
  }
  return { text: parts.join('\n'), sheets: n };
}

async function workbookFromViewer(v) {
  if (!v.model) return;
  if (!(await pro.allowed('batch'))) return;
  const base = v.uri.scheme === 'file' ? vscode.Uri.file(v.uri.fsPath.replace(/\.[^./\\]+$/, '') + '.md') : undefined;
  const target = await vscode.window.showSaveDialog({ defaultUri: base, filters: { Markdown: ['md'] }, title: `Save every sheet of ${v.name} as Markdown` });
  if (!target) return; // cancelled: no try used
  const { text, sheets } = workbookMarkdown(v.model, v.name, { formats: v.state.formats, header: v.state.header });
  await vscode.workspace.fs.writeFile(target, Buffer.from(text, 'utf8'));
  await pro.produced('batch');
  refreshPro();
  vscode.window.setStatusBarMessage(`Saved ${sheets} sheet${sheets === 1 ? '' : 's'} to ${path.posix.basename(target.path)}.`, 6000);
  await vscode.window.showTextDocument(target, { preview: false, viewColumn: vscode.ViewColumn.Beside });
}

/** Workbooks under a folder: at most 1,000 files, 20,000 entries looked at, 12 levels deep. */
async function listWorkbooks(root) {
  const out = [];
  let seen = 0;
  const stack = [[root, 0]];
  while (stack.length && seen < 20000 && out.length < 1000) {
    const [dir, depth] = stack.pop();
    let entries;
    try { entries = await fs.promises.readdir(dir, { withFileTypes: true }); } catch { continue; }
    entries.sort((a, b) => a.name.localeCompare(b.name));
    for (const e of entries) {
      if (++seen > 20000) break;
      const p = path.join(dir, e.name);
      if (e.isDirectory()) { if (!SKIP_DIRS.has(e.name) && !e.name.startsWith('.') && depth < 12) stack.push([p, depth + 1]); continue; }
      if (e.isFile() && WORKBOOK_RE.test(e.name) && !e.name.startsWith('~$')) out.push(p);
    }
  }
  return out.sort();
}

/** Pro 1 from the Explorer or the palette: one workbook, several, or every workbook of a folder -> .md files.
 *  From the palette it always asks which files or folder (the preview's own button converts the open workbook); the
 *  dialog starts in the folder of the open workbook, if one is open. */
async function workbookToMarkdown(uri, uris) {
  let picks = Array.isArray(uris) && uris.length ? uris : uri && uri.fsPath ? [uri] : null;
  if (!(await pro.allowed('batch'))) return;
  if (!picks) {
    const near = active && active.uri.scheme === 'file' ? vscode.Uri.file(path.dirname(active.uri.fsPath)) : (vscode.workspace.workspaceFolders || [])[0] && vscode.workspace.workspaceFolders[0].uri;
    const p = await vscode.window.showOpenDialog({ defaultUri: near || undefined, canSelectFiles: true, canSelectFolders: true, canSelectMany: true, openLabel: 'Convert', title: 'Workbooks or a folder to convert to Markdown', filters: { Workbooks: ['xlsx', 'xlsm', 'xlsb', 'xls', 'ods'] } });
    if (!p || !p.length) return; // cancelled: no try used
    picks = p;
  }
  const files = [];
  for (const u of picks) {
    if (u.scheme !== 'file') continue;
    let st;
    try { st = await fs.promises.stat(u.fsPath); } catch { continue; }
    if (st.isDirectory()) files.push(...(await listWorkbooks(u.fsPath)));
    else if (WORKBOOK_RE.test(u.fsPath)) files.push(u.fsPath);
  }
  if (!files.length) { vscode.window.showInformationMessage('No .xlsx, .xls or .ods workbook was found there.'); return; }
  const targets = files.map((f) => f.replace(/\.[^./\\]+$/, '') + '.md');
  const existing = targets.filter((t) => fs.existsSync(t));
  let replace = false;
  if (existing.length) {
    const a = await vscode.window.showWarningMessage(`${existing.length} of the ${targets.length} Markdown files exist already (for example ${path.basename(existing[0])}).`, { modal: true }, 'Replace them', 'Skip them');
    if (!a) return; // cancelled: no try used
    replace = a === 'Replace them';
  }
  const written = [], failed = [];
  let sheetsDone = 0;
  await vscode.window.withProgress({ location: vscode.ProgressLocation.Notification, title: 'Converting workbooks to Markdown', cancellable: true }, async (progress, token) => {
    for (let i = 0; i < files.length && !token.isCancellationRequested; i++) {
      if (!replace && existing.includes(targets[i])) continue;
      progress.report({ message: `${i + 1} of ${files.length}: ${path.basename(files[i])}` });
      try {
        const model = await runJob({ op: 'read', file: files[i], name: files[i] });
        const { text, sheets } = workbookMarkdown(model, path.basename(files[i]));
        await fs.promises.writeFile(targets[i], text, 'utf8');
        written.push(targets[i]); sheetsDone += sheets;
      } catch (e) { failed.push(`${path.basename(files[i])}: ${e.message}`); }
    }
  });
  if (written.length) { await pro.produced('batch'); refreshPro(); }
  const msg = `Wrote ${written.length} Markdown file${written.length === 1 ? '' : 's'} (${sheetsDone} sheet${sheetsDone === 1 ? '' : 's'})${failed.length ? `; ${failed.length} workbook${failed.length === 1 ? '' : 's'} could not be read (${failed[0].slice(0, 160)})` : ''}.`;
  if (!written.length) { vscode.window.showWarningMessage(msg); return; }
  const open = await vscode.window.showInformationMessage(msg, 'Open the first one');
  if (open) await vscode.window.showTextDocument(vscode.Uri.file(written[0]), { preview: false });
}

/** Pro 2: the Markdown tables of the document (or the selection) -> one .xlsx file. */
async function tablesToXlsx() {
  const ed = vscode.window.activeTextEditor;
  if (!ed) { vscode.window.showInformationMessage('Open a Markdown file with tables first.'); return; }
  const sel = ed.selection;
  const text = sel.isEmpty ? ed.document.getText() : ed.document.getText(new vscode.Range(sel.start.line, 0, sel.end.line, ed.document.lineAt(sel.end.line).text.length));
  const found = T.findAllTables(text, 255);
  const tables = [];
  for (const f of found) { const t = T.parseMarkdownTable(f.text); if (t) tables.push({ name: f.title || `Table ${tables.length + 1}`, rows: t.rows }); }
  if (!tables.length) { vscode.window.showInformationMessage(sel.isEmpty ? 'No Markdown table was found in this file.' : 'No Markdown table was found in the selection.'); return; }
  if (!(await pro.allowed('xlsx'))) return;
  const src = ed.document.uri;
  const base = src.scheme === 'file' ? vscode.Uri.file(src.fsPath.replace(/\.[^./\\]+$/, '') + '.xlsx') : undefined;
  const target = await vscode.window.showSaveDialog({ defaultUri: base, filters: { 'Excel workbook': ['xlsx'] }, title: `Save ${tables.length} Markdown table${tables.length === 1 ? '' : 's'} as an Excel file` });
  if (!target) return; // cancelled: no try used
  let buf;
  try { buf = await runJob({ op: 'writeXlsx', tables }); } catch (e) { vscode.window.showWarningMessage(`The Excel file could not be made: ${e.message}`); return; }
  await vscode.workspace.fs.writeFile(target, Buffer.from(buf));
  await pro.produced('xlsx');
  refreshPro();
  const pick = await vscode.window.showInformationMessage(`Saved ${tables.length} table${tables.length === 1 ? '' : 's'} to ${path.posix.basename(target.path)}, one sheet each.`, 'Preview it');
  if (pick) await vscode.commands.executeCommand('vscode.openWith', target, VIEW_TYPE);
}

// ---------------- activation ----------------

function activate(context) {
  clip = makeClipboard(vscode);
  rating = createRatingPrompt({ vscode, context, extensionId: EXT_ID, section: PREFIX, name: 'Excel to Markdown' });
  registerSupport({ vscode, context, prefix: PREFIX, selfId: EXT_ID });
  pro = createPro({ vscode, context });
  context.subscriptions.push(pro.onDidChange(() => refreshPro()));
  context.subscriptions.push(vscode.window.registerCustomEditorProvider(VIEW_TYPE, new Provider(context), { supportsMultipleEditorsPerDocument: true, webviewOptions: { retainContextWhenHidden: true } }));
  remember(vscode.window.activeTextEditor);
  context.subscriptions.push(vscode.window.onDidChangeActiveTextEditor((e) => remember(e)));
  context.subscriptions.push(vscode.window.onDidChangeTextEditorSelection((e) => { if (lastText && e.textEditor.document === lastText.doc) lastText.sel = e.selections[0]; }));
  const reg = (id, fn) => context.subscriptions.push(vscode.commands.registerCommand(id, (...a) => Promise.resolve(fn(...a)).catch((e) => vscode.window.showWarningMessage(String((e && e.message) || e)))));
  reg(`${PREFIX}.paste`, pasteAsTable);
  reg(`${PREFIX}.csvToTable`, csvToTable);
  reg(`${PREFIX}.tableToCsv`, () => tableTo(','));
  reg(`${PREFIX}.tableToTsv`, () => tableTo('\t'));
  reg(`${PREFIX}.formatTable`, formatTable);
  reg(`${PREFIX}.openSpreadsheet`, async (uri) => {
    let target = uri && uri.scheme ? uri : null;
    if (!target && vscode.window.activeTextEditor && SHEET_FILE_RE.test(vscode.window.activeTextEditor.document.fileName || '')) target = vscode.window.activeTextEditor.document.uri;
    if (!target) {
      const p = await vscode.window.showOpenDialog({ canSelectMany: false, openLabel: 'Open', title: 'Open a spreadsheet', filters: { Spreadsheets: ['xlsx', 'xlsm', 'xlsb', 'xls', 'ods', 'csv', 'tsv'] } });
      if (!p || !p[0]) return;
      target = p[0];
    }
    await vscode.commands.executeCommand('vscode.openWith', target, VIEW_TYPE);
  });
  reg(`${PREFIX}.openSample`, async (which) => {
    if (which === 'csv') {
      const content = fs.readFileSync(path.join(context.extensionPath || context.extensionUri.fsPath, 'samples', 'try-it.md'), 'utf8');
      const doc = await vscode.workspace.openTextDocument({ language: 'markdown', content });
      sampleDocs.add(doc);
      await vscode.window.showTextDocument(doc, { preview: false });
      return;
    }
    await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.joinPath(context.extensionUri, 'samples', 'sample.xlsx'), VIEW_TYPE);
  });
  reg(`${PREFIX}.addKeybinding`, () => vscode.commands.executeCommand('workbench.action.openGlobalKeybindings', `${PREFIX}.paste`));
  reg(`${PREFIX}.openWalkthrough`, () => vscode.commands.executeCommand('workbench.action.openWalkthrough', `${EXT_ID}#${PREFIX}.start`, false));
  reg(`${PREFIX}.workbookToMarkdown`, workbookToMarkdown);
  reg(`${PREFIX}.tablesToXlsx`, tablesToXlsx);
  context.subscriptions.push(vscode.window.registerUriHandler({ handleUri: (uri) => pro.handleUri(uri) }));
  context.subscriptions.push(vscode.workspace.onDidChangeConfiguration((e) => {
    if (e.affectsConfiguration(`${PREFIX}.showProTools`)) refreshPro();
    if (e.affectsConfiguration(PREFIX)) for (const v of viewers) v.render();
  }));
}

function deactivate() {}

module.exports = { activate, deactivate, workbookMarkdown, listWorkbooks, tableOptions };
