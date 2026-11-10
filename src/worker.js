'use strict';
// Reads and writes workbooks in a worker thread (bundled with SheetJS into dist/worker.js by scripts/build.js), so a
// hostile or huge file can never block or crash the extension host. runner.js kills this worker on a time or memory cap.
const fs = require('fs');
const { parentPort } = require('worker_threads');
const XLSX = require('xlsx');
const { readWorkbook, writeXlsx, SHEET_LIMITS } = require('./core/sheet');

function readFileCapped(file) {
  const st = fs.statSync(file);
  if (!st.isFile()) { const e = new Error('This is not a file.'); e.code = 'NOT_FILE'; throw e; }
  if (st.size > SHEET_LIMITS.maxFileBytes) { const e = new Error(`The file is ${Math.round(st.size / 1048576)} MB; files up to ${Math.round(SHEET_LIMITS.maxFileBytes / 1048576)} MB are read.`); e.code = 'TOO_BIG'; throw e; }
  return fs.readFileSync(file);
}

function run(job) {
  if (job.op === 'read') {
    const buf = job.buffer ? Buffer.from(job.buffer) : readFileCapped(job.file);
    return readWorkbook(XLSX, buf, job.name || job.file);
  }
  if (job.op === 'writeXlsx') {
    const out = writeXlsx(XLSX, job.tables || []);
    return out.buffer.slice(out.byteOffset, out.byteOffset + out.byteLength);
  }
  if (job.op === 'hang') { for (;;) { /* test hook: the watchdog must stop this */ } }
  throw new Error(`unknown job ${job.op}`);
}

parentPort.on('message', ({ job }) => {
  try {
    const result = run(job);
    parentPort.postMessage({ ok: true, result }, result instanceof ArrayBuffer ? [result] : []);
  } catch (e) {
    parentPort.postMessage({ ok: false, error: String((e && e.message) || e), code: e && e.code });
  }
});
