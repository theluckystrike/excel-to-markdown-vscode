'use strict';
// Starts a worker thread per job (dist/worker.js, SheetJS inside), with a memory cap and a watchdog. A job that takes
// longer than `ms` is killed and rejected with code 'TIMEOUT'; a worker that runs out of memory is rejected with code
// 'MEMORY'. The extension host keeps running either way.
const path = require('path');
const { Worker } = require('worker_threads');

const WORKER = process.env.EXCEL_TO_MD_WORKER || path.join(__dirname, '..', 'dist', 'worker.js');
const MEMORY_MB = 1536;

function runJob(job, { ms = 30000, transfer } = {}) {
  return new Promise((resolve, reject) => {
    let w;
    try { w = new Worker(WORKER, { resourceLimits: { maxOldGenerationSizeMb: MEMORY_MB, maxYoungGenerationSizeMb: 64 } }); } catch (e) { reject(e); return; }
    let done = false;
    const finish = (fn, v) => { if (done) return; done = true; clearTimeout(timer); w.terminate().catch(() => {}); fn(v); };
    const timer = setTimeout(() => { const e = new Error(`Reading took longer than ${Math.round(ms / 1000)} seconds and was stopped.`); e.code = 'TIMEOUT'; finish(reject, e); }, ms);
    w.on('message', (m) => {
      if (m.ok) finish(resolve, m.result);
      else { const e = new Error(m.error); e.code = m.code; finish(reject, e); }
    });
    w.on('error', (err) => {
      const oom = err && err.code === 'ERR_WORKER_OUT_OF_MEMORY';
      const e = new Error(oom ? `Reading needed more than ${MEMORY_MB} MB of memory and was stopped.` : String((err && err.message) || err));
      e.code = oom ? 'MEMORY' : 'WORKER';
      finish(reject, e);
    });
    w.on('exit', (code) => { if (!done) { const e = new Error(`The reader stopped (exit code ${code}).`); e.code = 'WORKER'; finish(reject, e); } });
    w.postMessage({ job }, transfer || []);
  });
}

module.exports = { runJob, WORKER, MEMORY_MB };
