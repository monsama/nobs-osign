/*
 * The audit log: who did what, when, from where. Sign-ins (and failed ones), publishing, bringing
 * back a version, the server's connections and settings, the setup script, the add-in's upload,
 * backups, approvals. Never a password, a token or a signature's content.
 *
 * In the data volume: audit.log, one JSON line each; at MAX bytes it becomes audit.1.log (the one
 * before goes), so the two hold the last few thousand entries.
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');

const MAX = 2 * 1024 * 1024;
const clip = (v, n) => String(v == null ? '' : v).replace(/[\u0000-\u001f]/g, ' ').slice(0, n);

function createAudit({ dataDir, now = Date.now }) {
  const file = path.join(dataDir, 'audit.log');
  const older = path.join(dataDir, 'audit.1.log');

  // what: a short name (signin, publish, …); ok: false for what was refused.
  function add({ who, ip, what, detail, ok = true }) {
    const line = JSON.stringify({ at: new Date(now()).toISOString(), who: clip(who, 120), ip: clip(ip, 60), what: clip(what, 40), detail: clip(detail, 300), ok: !!ok }) + '\n';
    try {
      let size = 0; try { size = fs.statSync(file).size; } catch (e) { /* none yet */ }
      if (size + line.length > MAX) fs.renameSync(file, older);
      fs.appendFileSync(file, line, { mode: 0o600 });
    } catch (e) { /* the log never stops what it logs */ }
  }

  // The newest first, at most limit, optionally only one kind or only refused ones.
  function list({ limit = 200, what = '', failed = false } = {}) {
    const read = (f) => { try { return fs.readFileSync(f, 'utf8').split('\n').filter(Boolean); } catch (e) { return []; } };
    const out = [];
    for (const l of read(older).concat(read(file)).reverse()) {
      let e; try { e = JSON.parse(l); } catch (x) { continue; }
      if (what && e.what !== what) continue;
      if (failed && e.ok) continue;
      out.push(e);
      if (out.length >= Math.min(Math.max(1, Number(limit) || 200), 2000)) break;
    }
    return out;
  }

  return { add, list };
}

module.exports = { createAudit };
