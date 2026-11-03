/*
 * The backups, for the local account admin: one downloaded now (locked with a passphrase), one
 * brought back, and those made by themselves every day (server/autobackup.js). Routes for
 * server/app.js, given what they need from it.
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { exportBackup, openBackup, restoreFiles } = require('./backup');
const { createAutoBackup } = require('./autobackup');

function createBackupRoutes({ dataDir, env, connections, store, configure, adminOnly, version, log }) {
  const VERSION = version;
  // Backups by themselves, every day.
  const autoBackup = createAutoBackup({ dataDir, dir: env.OSIGN_BACKUP_DIR || '', seal: connections.seal, open: connections.open, rawKey: connections.rawKey, version, log });
  autoBackup.schedule();
  const routes = {
    // A backup of the data (server/backup.js), for the local account admin.
    'backup/export': (d, s) => {
      if (adminOnly(s)) return { error: 'Only the local account admin can make a backup.' };
      return exportBackup({ dataDir, passphrase: d.passphrase, appKey: connections.rawKey(), version: VERSION });
    },
    'backup/restore': (d, s) => {
      if (adminOnly(s)) return { error: 'Only the local account admin can bring back a backup.' };
      const b = openBackup(d.data, d.passphrase);
      if (b.error) return b;
      const r = restoreFiles(dataDir, b.files);
      if (r.error) return r;
      fs.mkdirSync(store.assetDir, { recursive: true });
      fs.mkdirSync(path.join(dataDir, 'history', 'assets'), { recursive: true });
      // Sealed with the key of the server it came from: sealed again with this one's.
      connections.reseal(b.appKey);
      configure();
      log.info('backup: brought back the backup of ' + b.at + ' (version ' + b.version + ', ' + r.files + ' files) by ' + s.user.login);
      return { at: b.at, version: b.version, files: r.files };
    },
    // Backups by themselves: set, made now, downloaded. The local account admin's.
    'backup/auto/get': (d, s) => (adminOnly(s) ? { error: 'Only the local account admin sets the automatic backups.' } : autoBackup.answer()),
    'backup/auto/save': (d, s) => (adminOnly(s) ? { error: 'Only the local account admin sets the automatic backups.' } : autoBackup.save(d)),
    'backup/auto/run': (d, s) => {
      if (adminOnly(s)) return { error: 'Only the local account admin sets the automatic backups.' };
      const r = autoBackup.run(); return r.error ? r : autoBackup.answer();
    },
    'backup/auto/download': (d, s) => (adminOnly(s) ? { error: 'Only the local account admin downloads backups.' } : autoBackup.get(d.name))
  };
  const audit = {
    'backup/export': () => 'backup downloaded',
    'backup/restore': () => 'backup restored',
    'backup/auto/save': (d) => 'automatic backups ' + (d.at ? 'daily at ' + d.at + ', kept ' + (d.keep || 30) + ' days' : 'off') + (d.passphrase ? ', passphrase set' : ''),
    'backup/auto/run': (d, r) => 'automatic backup made now' + (r && r.last && r.last.name ? ': ' + r.last.name : ''),
    'backup/auto/download': (d) => 'backup downloaded: ' + String(d.name || '').slice(0, 80)
  };
  return { routes, audit, autoBackup };
}

module.exports = { createBackupRoutes };
