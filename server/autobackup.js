/*
 * Backups made by themselves: every day at a time set by the local account admin, locked with a
 * passphrase of theirs (server/backup.js), kept for a number of days in backups/ of the data volume
 * (or OSIGN_BACKUP_DIR, e.g. another disk). The passphrase is kept sealed with the server's key, so
 * the server can lock the backups; opening one still takes the passphrase.
 *
 * autobackup.json: { at: '02:30', keep: 30, pass: <sealed>, last: { at, ok, name|error } }
 * Neither it nor backups/ goes into a backup (server/backup.js SKIP).
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { exportBackup, MIN_PASS } = require('./backup');
const { writeAtomic } = require('./store');
const { nextAt } = require('./users');

const NAME = /^nobs-osign-backup-\d{4}-\d{2}-\d{2}(?:-\d{4})?\.osignbak$/;

// The automatic backups: their settings, the daily timer, making, pruning and handing out the files.
function createAutoBackup({ dataDir, dir, seal, open, rawKey, version, log }) {
  const file = path.join(dataDir, 'autobackup.json');
  // A folder of its own inside the data volume would go into every backup and be emptied when one
  // is brought back: then the usual backups/ (left out of both), and said.
  const inside = (d) => { const r = path.relative(path.resolve(dataDir), path.resolve(d)); return !r.startsWith('..') && !path.isAbsolute(r); };
  const warning = dir && inside(dir) && path.resolve(dir) !== path.resolve(dataDir, 'backups')
    ? 'OSIGN_BACKUP_DIR is inside the data volume (' + dir + '): the backups go into backups/ of the data volume instead. Set a folder outside it, best on another disk.' : '';
  if (warning && log) log.warn('automatic backup: ' + warning);
  const where = dir && !warning ? dir : path.join(dataDir, 'backups');
  let timer = null;
  const read = () => { try { const s = JSON.parse(fs.readFileSync(file, 'utf8')); return s && typeof s === 'object' ? s : {}; } catch (e) { return {}; } };
  const write = (s) => writeAtomic(file, JSON.stringify(s, null, 2));
  const passphrase = (s) => (s.pass ? open('autobackupPassphrase', s.pass) : '');

  function list() {
    let names = [];
    try { names = fs.readdirSync(where).filter((n) => NAME.test(n)); } catch (e) { /* none yet */ }
    return names.map((n) => { const st = fs.statSync(path.join(where, n)); return { name: n, size: st.size, at: st.mtime.toISOString() }; })
      .sort((a, b) => (a.at < b.at ? 1 : -1));
  }
  function answer() {
    const s = read();
    return { on: !!(s.at && s.pass), at: s.at || '', keep: s.keep || 30, passSet: !!passphrase(s), passLost: !!s.pass && !passphrase(s), where, warning, last: s.last || null, backups: list() };
  }
  // Older ones than the days kept go; the newest always stays.
  function prune(keep) {
    const all = list(), cut = Date.now() - keep * 86400000;
    all.slice(1).filter((b) => Date.parse(b.at) < cut).forEach((b) => { try { fs.rmSync(path.join(where, b.name)); } catch (e) { /* next time */ } });
  }
  // One backup made now, with the passphrase kept: written, the older ones pruned, the outcome kept.
  function run(now = new Date()) {
    const s = read(), pass = passphrase(s);
    if (!pass) return { error: 'No passphrase for the automatic backups: set it (again) in Backup.' };
    let b;
    // A file gone while it is read (a request withdrawn, history pruned): the next try.
    try { b = exportBackup({ dataDir, passphrase: pass, appKey: rawKey(), version, now }); } catch (e) { b = { error: 'Reading the data: ' + e.message }; }
    if (b.error) { s.last = { at: now.toISOString(), ok: false, error: b.error }; write(s); return b; }
    const name = b.name.replace(/\.osignbak$/, '-' + now.toISOString().slice(11, 16).replace(':', '') + '.osignbak');
    try {
      fs.mkdirSync(where, { recursive: true, mode: 0o700 });
      writeAtomic(path.join(where, name), Buffer.from(b.data, 'base64'));
      prune(s.keep || 30);
    } catch (e) { s.last = { at: now.toISOString(), ok: false, error: e.message }; write(s); return { error: 'Writing the backup: ' + e.message }; }
    s.last = { at: now.toISOString(), ok: true, name }; write(s);
    return { name };
  }
  // The next backup armed at its time of day; none while it is off.
  function schedule() {
    if (timer) { clearTimeout(timer); timer = null; }
    const s = read();
    if (!s.at || !s.pass) return;
    const next = nextAt(s.at, Date.now());
    if (!next) return;
    timer = setTimeout(() => {
      try {
        const r = run();
        if (r.error) log.warn('automatic backup: ' + r.error); else log.info('automatic backup: ' + r.name);
      } catch (e) { log.error('automatic backup: ' + (e && e.message)); }
      schedule();
    }, Math.max(1000, next - Date.now()));
    if (timer.unref) timer.unref();
  }
  // { at: '' (off) | 'HH:MM', keep: days, passphrase: new one (or left out to keep it) }
  function save(d) {
    const s = read();
    const at = String(d.at || '').trim();
    if (at && !/^([01]?\d|2[0-3]):[0-5]\d$/.test(at)) return { error: 'The time of the daily backup is a time of day, like 02:30.' };
    const keep = Math.round(Number(d.keep || 30));
    if (!(keep >= 1 && keep <= 365)) return { error: 'Keep the backups for 1 to 365 days.' };
    if (typeof d.passphrase === 'string' && d.passphrase) {
      if (d.passphrase.length < MIN_PASS) return { error: 'The passphrase has at least ' + MIN_PASS + ' characters.' };
      s.pass = seal('autobackupPassphrase', d.passphrase);
    }
    if (at && !passphrase(s)) return { error: 'Set a passphrase for the backups first.' };
    s.at = at; s.keep = keep;
    write(s); schedule();
    return answer();
  }
  // One of them, to download: { name, data (base64) }.
  function get(name) {
    if (!NAME.test(String(name || ''))) return { error: 'There is no such backup.' };
    try { return { name, data: fs.readFileSync(path.join(where, name)).toString('base64') }; } catch (e) { return { error: 'There is no such backup.' }; }
  }
  function stop() { if (timer) clearTimeout(timer); timer = null; }
  return { answer, save, run, get, schedule, stop, list };
}

module.exports = { createAutoBackup };
