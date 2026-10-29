/*
 * The settings the editor keeps, the same keys as editor/OSign.ps1 (see docs/API.md, settings/get).
 *
 * In the container two of them are the server's, not the editor's: folder (where Publish writes,
 * OSIGN_PUBLISH_DIR) and photosFolder (its photos/). They are answered, never saved, so nobody
 * signed in to the editor can point the server at another place on its disk.
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { writeAtomic, readText } = require('./store');

const KEYS = ['folder', 'photosFolder', 'addinUrl', 'clientId', 'tenantId', 'spSite', 'spFolder', 'addinFolder', 'addinHost', 'setupDone'];
const SERVER_KEYS = ['folder', 'photosFolder', 'addinFolder'];
const DEFAULTS = { folder: '', photosFolder: '', addinUrl: '', clientId: '', tenantId: '', spSite: 'root', spFolder: 'OSign', addinFolder: '', addinHost: '', setupDone: '' };
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// The same rules as Test-AddinSettings in OSign.ps1 and checkSetup in addin/fill.mjs. Empty values pass.
function check(s) {
  if (s.addinUrl && !/^https:\/\/[^\s"<>'&]+$/.test(s.addinUrl)) return 'The add-in address is an https:// address, like https://osign.contoso.com/.';
  if (s.clientId && !GUID.test(s.clientId)) return "The client ID is the app registration's Application (client) ID, a GUID.";
  if (s.tenantId && !GUID.test(s.tenantId)) return 'The tenant ID is a Directory (tenant) ID, a GUID.';
  if (s.spSite && !/^root$|^[A-Za-z0-9.-]+\.sharepoint\.com:\/(sites|teams|personal)\/[^:/"\\\s]+:$/.test(s.spSite)) return 'The SharePoint site is root, <host>.sharepoint.com:/sites/<name>: or, for a OneDrive, <host>-my.sharepoint.com:/personal/<name>:';
  if (s.addinHost && !['cloudflare', 'azure', 'iis', 'apache', 'other'].includes(s.addinHost)) return 'The web server is cloudflare, azure, iis, apache or other.';
  const f = String(s.spFolder || '').replace(/^\/+|\/+$/g, '');
  if (s.spFolder && (/["\\\u0000-\u001f]/.test(f) || f.split('/').some((p) => !p || p === '.' || p === '..'))) return 'The SharePoint folder is a folder like OSign.';
  return null;
}

function createSettings(dataDir, publishDir) {
  const file = path.join(dataDir, 'settings.json');
  function get() {
    const s = Object.assign({}, DEFAULTS);
    try {
      const saved = JSON.parse(readText(file) || '{}');
      for (const k of KEYS) if (saved[k] !== undefined && saved[k] !== null) s[k] = String(saved[k]);
    } catch (e) { /* defaults */ }
    s.folder = publishDir; s.photosFolder = ''; s.addinFolder = '';
    return s;
  }
  const isDir = (p) => { try { return fs.statSync(p).isDirectory(); } catch (e) { return false; } };
  function answer() {
    const s = get();
    return { settings: s, folderFound: isDir(publishDir), photosFound: isDir(path.join(publishDir, 'photos')) };
  }
  // Only what is sent changes; the rest stays as it was.
  function save(sent) {
    const s = get();
    for (const k of KEYS) {
      if (SERVER_KEYS.includes(k) || !sent || sent[k] === undefined || sent[k] === null) continue;
      s[k] = String(sent[k]).trim().replace(/[\\/]+$/, '');
    }
    const err = check(s); if (err) return { error: err };
    const keep = {};
    for (const k of KEYS) if (!SERVER_KEYS.includes(k)) keep[k] = s[k];
    writeAtomic(file, JSON.stringify(keep));
    return answer();
  }
  return { get, answer, save, photosDir: () => path.join(publishDir, 'photos') };
}

module.exports = { createSettings, check };
