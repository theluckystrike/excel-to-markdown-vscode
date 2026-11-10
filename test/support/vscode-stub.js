'use strict';
// A small stand-in for the vscode module: enough for activation, text editors with selections and edits, the custom
// editor, the Pro tools and the license facade. Tests script the dialogs through `ui` and read what happened from `log`.
const fs = require('fs');
const os = require('os');
const path = require('path');
const Module = require('module');

class Position { constructor(line, character) { this.line = line; this.character = character; } }
class Range {
  constructor(a, b, c, d) {
    if (a instanceof Position) { this.start = a; this.end = b; } else { this.start = new Position(a, b); this.end = new Position(c, d); }
  }
  get isEmpty() { return this.start.line === this.end.line && this.start.character === this.end.character; }
}
class Selection extends Range {
  constructor(a, b, c, d) { super(a, b, c, d); this.anchor = this.start; this.active = this.end; }
}

class Doc {
  constructor(text, { fileName = '', languageId = 'markdown', uri } = {}) { this.text = text; this.fileName = fileName; this.languageId = languageId; this.uri = uri || { scheme: 'untitled', path: fileName || 'Untitled-1', fsPath: fileName, toString: () => 'untitled:' + (fileName || 'Untitled-1') }; }
  get lines() { return this.text.split('\n'); }
  get lineCount() { return this.lines.length; }
  lineAt(i) { return { text: this.lines[i] }; }
  offsetAt(p) { const ls = this.lines; let o = 0; for (let i = 0; i < p.line; i++) o += ls[i].length + 1; return o + Math.min(p.character, (ls[p.line] || '').length); }
  getText(r) { if (!r) return this.text; return this.text.slice(this.offsetAt(r.start), this.offsetAt(r.end)); }
}

class Editor {
  constructor(doc, sel) { this.document = doc; this.selections = [sel || new Selection(0, 0, 0, 0)]; this.readOnly = false; }
  get selection() { return this.selections[0]; }
  set selection(s) { this.selections = [s]; }
  async edit(cb) {
    if (this.readOnly) return false;
    const edits = [];
    cb({ replace: (r, t) => edits.push([r, t]), insert: (p, t) => edits.push([new Range(p, p), t]) });
    edits.sort((a, b) => this.document.offsetAt(b[0].start) - this.document.offsetAt(a[0].start));
    for (const [r, t] of edits) { const s = this.document.offsetAt(r.start), e = this.document.offsetAt(r.end); this.document.text = this.document.text.slice(0, s) + t + this.document.text.slice(e); }
    return true;
  }
}

function create() {
  const log = { info: [], warn: [], error: [], status: [], opened: [], external: [], commands: [], posted: [], saved: [], progress: [], clipboardWrites: [], openDialogs: [] };
  const ui = { open: [], save: [], input: [], quick: [], info: [], warn: [] };
  const commands = new Map();
  const providers = new Map();
  const state = new Map();
  const config = { showProTools: true, showSupportLinks: true, askForRating: true };
  const clipboard = { text: '' };
  class EventEmitter { constructor() { this.l = new Set(); this.event = (f) => { this.l.add(f); return { dispose: () => this.l.delete(f) }; }; } fire(v) { for (const f of this.l) f(v); } dispose() { this.l.clear(); } }
  class Uri {
    constructor(scheme, p) { this.scheme = scheme; this.path = p; this.fsPath = p; }
    static file(p) { return new Uri('file', path.resolve(p)); }
    static parse(s) { const u = new Uri(s.split(':')[0], s); u.str = s; return u; }
    static joinPath(u, ...parts) { return new Uri(u.scheme, path.join(u.path, ...parts)); }
    toString() { return this.str || `${this.scheme}://${this.path}`; }
    with() { return this; }
  }
  const nextOf = (q, dflt) => (q.length ? q.shift() : dflt);
  const msg = (kind) => async (text, ...rest) => {
    const opts = rest[0] && typeof rest[0] === 'object' && !Array.isArray(rest[0]) ? rest.shift() : {};
    log[kind].push({ text, opts, buttons: rest.map((b) => (typeof b === 'string' ? b : b.title)) });
    if (kind === 'info') return nextOf(ui.info, undefined);
    if (kind === 'warn') return nextOf(ui.warn, undefined);
    return undefined;
  };
  const changeEditor = new EventEmitter();
  const vscode = {
    version: '1.90.2',
    EventEmitter, Uri, Range, Position, Selection,
    ViewColumn: { Beside: -2, Active: -1, One: 1 },
    ProgressLocation: { Notification: 15, Window: 10 },
    QuickPickItemKind: { Separator: -1, Default: 0 },
    env: { appName: 'Visual Studio Code', uriScheme: 'vscode', machineId: 'test', clipboard: { readText: async () => clipboard.text, writeText: async (t) => { clipboard.text = t; log.clipboardWrites.push(t); } }, openExternal: async (u) => { log.external.push(String(u)); return true; }, asExternalUri: async (u) => u },
    workspace: {
      workspaceFolders: undefined,
      getConfiguration: () => ({ get: (k, d) => (k in config ? config[k] : d), update: async () => {} }),
      onDidChangeConfiguration: () => ({ dispose() {} }),
      fs: {
        stat: async (u) => ({ size: fs.statSync(u.fsPath).size }),
        readFile: async (u) => fs.readFileSync(u.fsPath),
        writeFile: async (u, b) => { fs.writeFileSync(u.fsPath, b); log.saved.push(u.fsPath); },
      },
      openTextDocument: async (o) => new Doc((o && o.content) || '', { languageId: (o && o.language) || 'plaintext' }),
      isTrusted: true,
    },
    window: {
      state: { focused: true },
      activeTextEditor: undefined,
      visibleTextEditors: [],
      onDidChangeWindowState: () => ({ dispose() {} }),
      onDidChangeActiveTextEditor: changeEditor.event,
      registerCustomEditorProvider: (vt, p) => { providers.set(vt, p); return { dispose() {} }; },
      registerUriHandler: () => ({ dispose() {} }),
      showInformationMessage: msg('info'), showWarningMessage: msg('warn'), showErrorMessage: msg('error'),
      setStatusBarMessage: (t) => { log.status.push(t); return { dispose() {} }; },
      showOpenDialog: async (o) => { log.openDialogs.push(o); return nextOf(ui.open, undefined); },
      showSaveDialog: async () => nextOf(ui.save, undefined),
      showInputBox: async () => nextOf(ui.input, undefined),
      showQuickPick: async (items) => { const f = nextOf(ui.quick, undefined); return typeof f === 'function' ? f(await items) : f; },
      showTextDocument: async (d, o) => {
        log.opened.push(d);
        if (!(d instanceof Doc)) return {};
        const ed = new Editor(d, o && o.selection ? new Selection(o.selection.start.line, o.selection.start.character, o.selection.end.line, o.selection.end.character) : undefined);
        vscode.window.activeTextEditor = ed; vscode.window.visibleTextEditors = [ed];
        return ed;
      },
      onDidChangeTextEditorSelection: () => ({ dispose() {} }),
      withProgress: async (o, fn) => { log.progress.push(o.title); return fn({ report: (r) => log.progress.push(r.message) }, { isCancellationRequested: false, onCancellationRequested: () => ({ dispose() {} }) }); },
      createOutputChannel: () => ({ appendLine() {}, dispose() {} }),
    },
    commands: {
      registerCommand: (id, fn) => { commands.set(id, fn); return { dispose() {} }; },
      executeCommand: async (id, ...a) => { log.commands.push([id, ...a]); if (commands.has(id)) return commands.get(id)(...a); return undefined; },
      getCommands: async () => [...commands.keys()],
    },
    extensions: { all: [], getExtension: () => undefined, onDidChange: () => ({ dispose() {} }) },
    authentication: { getSession: async () => undefined, onDidChangeSessions: () => ({ dispose() {} }) },
    l10n: { t: (s) => s },
  };
  const secrets = new Map();
  const root = path.join(__dirname, '..', '..');
  const context = {
    subscriptions: [],
    extensionUri: Uri.file(root),
    extensionPath: root,
    extension: undefined,
    globalStorageUri: Uri.file(fs.mkdtempSync(path.join(os.tmpdir(), 'e2m-gs-'))),
    globalState: { get: (k, d) => (state.has(k) ? state.get(k) : d), update: async (k, v) => { state.set(k, v); }, keys: () => [...state.keys()], setKeysForSync() {} },
    secrets: { get: async (k) => secrets.get(k), store: async (k, v) => { secrets.set(k, v); }, delete: async (k) => { secrets.delete(k); }, onDidChange: () => ({ dispose() {} }) },
  };
  /** Opens a text editor with `text`; the cursor goes to the first | in the text (removed) or to the end. */
  function openEditor(text, opts = {}) {
    const k = text.indexOf('\u2038');
    const clean = k >= 0 ? text.slice(0, k) + text.slice(k + 1) : text;
    const doc = new Doc(clean, opts);
    let pos = doc.lines.length - 1, ch = doc.lines[pos].length;
    if (k >= 0) { const before = clean.slice(0, k).split('\n'); pos = before.length - 1; ch = before[pos].length; }
    const ed = new Editor(doc, opts.selection || new Selection(pos, ch, pos, ch));
    vscode.window.activeTextEditor = ed;
    vscode.window.visibleTextEditors = [ed];
    changeEditor.fire(ed);
    return ed;
  }
  return { vscode, context, log, ui, commands, providers, state, config, clipboard, openEditor, Doc, Editor };
}

/** Load src/extension.js with a fresh stub; returns the stub and the module. */
function loadExtension() {
  const stub = create();
  const orig = Module._load;
  Module._load = function (req, parent, isMain) { if (req === 'vscode') return stub.vscode; return orig.call(this, req, parent, isMain); };
  const root = path.join(__dirname, '..', '..', 'src');
  for (const k of Object.keys(require.cache)) if (k.startsWith(root)) delete require.cache[k];
  try { stub.ext = require(path.join(root, 'extension.js')); } finally { Module._load = orig; }
  Module._load = function (req, parent, isMain) { if (req === 'vscode') return stub.vscode; return orig.call(this, req, parent, isMain); };
  stub.restore = () => { Module._load = orig; };
  return stub;
}

/** Open a file in the preview and wait for the init/output or error message. */
async function openPreview(stub, file) {
  const posted = [];
  let onMsg;
  const panel = {
    active: true, visible: true,
    webview: { options: {}, html: '', cspSource: 'vscode-resource:', asWebviewUri: (u) => u, postMessage: async (m) => { posted.push(m); stub.log.posted.push(m); return true; }, onDidReceiveMessage: (f) => { onMsg = f; return { dispose() {} }; } },
    onDidChangeViewState: () => ({ dispose() {} }), onDidDispose: () => ({ dispose() {} }), reveal() {},
  };
  const p = stub.providers.get('excelToMarkdown.preview');
  const doc = await p.openCustomDocument(stub.vscode.Uri.file(file));
  p.resolveCustomEditor(doc, panel);
  await onMsg({ type: 'ready' });
  return { panel, posted, send: (m) => onMsg(m), last: (type) => [...posted].reverse().find((m) => m.type === type) };
}

module.exports = { create, loadExtension, openPreview, Range, Position, Selection };
