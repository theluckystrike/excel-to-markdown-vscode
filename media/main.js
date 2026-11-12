// Spreadsheet preview webview: sheet picker, range, header and format options, a grid of the first rows, and the
// table text. Every value is set with textContent; there are no inline styles (the CSP has no 'unsafe-inline').
(function () {
  'use strict';
  const vscode = acquireVsCodeApi();
  const app = document.getElementById('app');
  let sheets = [], state = { sheet: 0, range: '', header: true, format: 'markdown', formats: false }, pro = { show: true, licensed: false, tries: {} };
  let built = false, rendered = false, rangeTimer = null;
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };
  const fmtInt = (n) => String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const colName = (i) => { let s = ''; i += 1; while (i > 0) { const m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; };

  function footer() {
    const t = document.getElementById('zovo-footer');
    if (t && t.content && t.content.firstElementChild) document.body.appendChild(t.content.cloneNode(true));
  }

  const ui = {};
  function build(name) {
    app.textContent = '';
    const top = el('div', 'top');
    ui.fname = el('span', 'fname', name);
    top.appendChild(ui.fname);
    ui.sheet = el('select', 'sheet'); ui.sheet.title = 'Sheet';
    sheets.forEach((s, i) => { const o = el('option', null, s.name + (s.rows ? '' : ' (empty)')); o.value = String(i); ui.sheet.appendChild(o); });
    ui.sheet.addEventListener('change', () => send({ sheet: Number(ui.sheet.value), range: '' }, true));
    top.appendChild(labelled('Sheet', ui.sheet));
    ui.range = el('input', 'range'); ui.range.type = 'text'; ui.range.placeholder = 'whole sheet, or A1:D20'; ui.range.spellcheck = false; ui.range.title = 'Range in A1 notation: A1:D20, B:D or 2:10. Empty = the whole sheet.';
    ui.range.addEventListener('input', () => { clearTimeout(rangeTimer); rangeTimer = setTimeout(() => send({ range: ui.range.value }), 350); });
    ui.range.addEventListener('keydown', (e) => { if (e.key === 'Enter') { clearTimeout(rangeTimer); send({ range: ui.range.value }); } });
    top.appendChild(labelled('Range', ui.range));
    ui.header = el('input', 'hdr'); ui.header.type = 'checkbox'; ui.header.checked = !!state.header;
    ui.header.addEventListener('change', () => send({ header: ui.header.checked }));
    const hl = el('label', 'check'); hl.appendChild(ui.header); hl.appendChild(document.createTextNode(' First row is the header'));
    top.appendChild(hl);
    ui.format = el('select', 'format'); ui.format.title = 'Output format';
    for (const [v, t] of [['markdown', 'Markdown'], ['html', 'HTML'], ['asciidoc', 'AsciiDoc'], ['jira', 'Jira'], ['csv', 'CSV'], ['tsv', 'TSV']]) { const o = el('option', null, t); o.value = v; ui.format.appendChild(o); }
    ui.format.value = state.format;
    ui.format.addEventListener('change', () => send({ format: ui.format.value }));
    top.appendChild(labelled('Format', ui.format));
    ui.copy = el('button', 'primary copy', 'Copy'); ui.copy.title = 'Copy the table text';
    ui.copy.addEventListener('click', () => vscode.postMessage({ type: 'copy' }));
    ui.insert = el('button', 'insert', 'Insert at cursor'); ui.insert.title = 'Insert the table into the last text editor you used (or a new Markdown file)';
    ui.insert.addEventListener('click', () => vscode.postMessage({ type: 'insert' }));
    top.appendChild(ui.copy); top.appendChild(ui.insert);
    ui.pro = el('span', 'protools');
    ui.formats = el('button', 'toggle'); ui.formats.dataset.tool = 'formats'; ui.formats.textContent = 'Excel formats, links, merged cells';
    ui.formats.setAttribute('aria-pressed', 'false');
    ui.formats.addEventListener('click', () => vscode.postMessage({ type: 'tool', tool: 'formats', on: !state.formats }));
    ui.workbook = el('button', 'workbook'); ui.workbook.dataset.tool = 'workbook'; ui.workbook.textContent = 'Whole workbook to .md';
    ui.workbook.addEventListener('click', () => vscode.postMessage({ type: 'tool', tool: 'workbook' }));
    ui.pro.appendChild(ui.formats); ui.pro.appendChild(ui.workbook);
    top.appendChild(ui.pro);
    app.appendChild(top);
    app.appendChild(el('div', 'reach1-line')); // REACH1: the Zovo Pro tools row (filled by the extension's boot script)
    ui.info = el('div', 'info muted');
    app.appendChild(ui.info);
    const split = el('div', 'split');
    ui.gridWrap = el('div', 'grid');
    ui.outWrap = el('div', 'outwrap');
    ui.out = el('pre', 'out');
    ui.outNote = el('div', 'muted outnote');
    ui.outWrap.appendChild(ui.out); ui.outWrap.appendChild(ui.outNote);
    split.appendChild(ui.gridWrap); split.appendChild(ui.outWrap);
    app.appendChild(split);
    built = true;
    footer();
    applyPro();
  }

  function labelled(text, control) { const l = el('label', 'field'); l.appendChild(el('span', 'muted', text + ' ')); l.appendChild(control); return l; }

  function send(change, resetRange) {
    Object.assign(state, change);
    if (resetRange && ui.range) ui.range.value = '';
    vscode.postMessage({ type: 'render', ...change });
  }

  function applyPro() {
    if (!built) return;
    ui.pro.classList.toggle('hidden', !pro.show);
    const t = pro.tries || {};
    const note = (k, what) => (pro.licensed ? `${what} (Zovo Pro)` : `${what} (Zovo Pro tool, ${t[k] || 0} free tr${t[k] === 1 ? 'y' : 'ies'} left)`);
    ui.formats.title = note('formats', 'Copy cells as Excel shows them: number and date formats, hyperlinks as Markdown links, merged cells filled');
    ui.workbook.title = note('batch', 'Save every sheet of this workbook to one .md file');
    ui.formats.classList.toggle('on', !!state.formats);
    ui.formats.setAttribute('aria-pressed', state.formats ? 'true' : 'false');
  }

  function grid(msg) {
    ui.gridWrap.textContent = '';
    const table = el('table');
    const thead = el('thead'); const hr = el('tr');
    hr.appendChild(el('th', 'corner', ''));
    const cols = msg.grid.reduce((m, r) => Math.max(m, r.length), 0);
    for (let c = 0; c < cols; c++) hr.appendChild(el('th', null, colName(msg.c0 + c)));
    thead.appendChild(hr); table.appendChild(thead);
    const tb = el('tbody');
    msg.grid.forEach((r, i) => {
      const tr = el('tr');
      tr.appendChild(el('th', 'rownum', String(msg.r0 + i + 1)));
      for (let c = 0; c < cols; c++) { const td = el('td', null, r[c] == null ? '' : r[c]); if (i === 0 && state.header) td.className = 'head'; tr.appendChild(td); }
      tb.appendChild(tr);
    });
    table.appendChild(tb);
    ui.gridWrap.appendChild(table);
    if (msg.rows > msg.grid.length || msg.cols > cols) ui.gridWrap.appendChild(el('div', 'muted gridnote', `The grid shows the first ${fmtInt(msg.grid.length)} rows and ${cols} columns; the table text has all ${fmtInt(msg.rows)} rows and ${msg.cols} columns.`));
  }

  function output(msg) {
    if (msg.state) { state = Object.assign(state, msg.state); }
    if (msg.error) { ui.info.textContent = msg.error; ui.info.classList.add('err'); return; }
    ui.info.classList.remove('err');
    const s = sheets[state.sheet] || {};
    ui.sheet.value = String(state.sheet);
    if (document.activeElement !== ui.range) ui.range.value = state.range || '';
    const parts = [`${s.name || 'Sheet'}${msg.range ? ', ' + msg.range : ''}: ${fmtInt(msg.rows)} row${msg.rows === 1 ? '' : 's'}, ${msg.cols} column${msg.cols === 1 ? '' : 's'}`];
    if (state.formats) parts.push('Excel formats, links and merged cells kept');
    if (msg.truncated) parts.push(`only the ${msg.truncated} were read`);
    ui.info.textContent = parts.join(' \u00b7 ');
    if (!msg.rows) { ui.gridWrap.textContent = ''; ui.gridWrap.appendChild(el('p', 'muted', 'This sheet or range has no cells.')); }
    else grid(msg);
    ui.out.textContent = msg.text || '';
    ui.outNote.textContent = msg.textLen > (msg.text || '').length ? `Showing the first ${fmtInt((msg.text || '').length)} of ${fmtInt(msg.textLen)} characters. Copy and Insert use all of it.` : '';
    applyPro();
    if (!rendered) { rendered = true; vscode.postMessage({ type: 'rendered' }); }
  }

  function error(msg) {
    app.textContent = '';
    const box = el('div', 'error');
    box.appendChild(el('h2', null, 'This file could not be read'));
    box.appendChild(el('p', null, msg.message || 'Unknown error.'));
    box.appendChild(el('p', 'muted', 'Excel to Markdown reads .xlsx, .xlsm, .xlsb, .xls, .ods, .csv and .tsv files.'));
    app.appendChild(box);
    footer();
  }

  window.addEventListener('message', (e) => {
    const msg = e.data;
    if (!msg || typeof msg.type !== 'string') return;
    if (msg.type === 'init') { sheets = msg.sheets || []; state = Object.assign(state, msg.state || {}); pro = msg.pro || pro; build(msg.name || ''); }
    else if (msg.type === 'output' && built) output(msg);
    else if (msg.type === 'error') error(msg);
    else if (msg.type === 'pro') { pro = msg.pro || pro; if (typeof msg.formats === 'boolean') state.formats = msg.formats; applyPro(); }
    else if (msg.type === 'notice' && built) { ui.info.textContent = msg.text; ui.info.classList.add('err'); }
    else if (msg.type === 'done' && built) { const b = msg.what === 'copy' ? ui.copy : ui.insert; b.classList.add('flash'); setTimeout(() => b.classList.remove('flash'), 900); }
  });
  vscode.postMessage({ type: 'ready' });
}());
