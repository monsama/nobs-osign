/*
 * The settings the editor keeps, the same keys as editor/OSign.ps1 (see docs/API.md, settings/get).
 *
 * In the container the publish target is not a setting of its own: it is the SharePoint site and
 * folder the add-in reads (spSite, spFolder), written through Microsoft Graph (server/graph.js),
 * or the server's folder when Graph is not set up. folder, photosFolder and addinFolder are
 * answered (where publishing goes, as words), never saved, so nobody signed in to the editor can
 * point the server at another place on its disk.
 */
'use strict';
const path = require('node:path');
const { writeAtomic, readText } = require('./store');

const KEYS = ['folder', 'photosFolder', 'addinUrl', 'clientId', 'tenantId', 'spSite', 'spFolder', 'addinFolder', 'addinHost', 'setupDone'];
const SERVER_KEYS = ['folder', 'photosFolder', 'addinFolder'];
const DEFAULTS = { folder: '', photosFolder: '', addinUrl: '', clientId: '', tenantId: '', spSite: 'root', spFolder: 'OSign', addinFolder: '', addinHost: '', setupDone: '' };
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
// How long what was found in SharePoint (is the folder there?) is believed.
const FOUND_MS = 60 * 1000;

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

// targetFor(settings) → where publishing goes for these settings (server/graph.js).
function createSettings(dataDir, targetFor) {
  const file = path.join(dataDir, 'settings.json');
  let found = null;
  function get() {
    const s = Object.assign({}, DEFAULTS);
    try {
      const saved = JSON.parse(readText(file) || '{}');
      for (const k of KEYS) if (saved[k] !== undefined && saved[k] !== null) s[k] = String(saved[k]);
    } catch (e) { /* defaults */ }
    s.folder = targetFor(s).describe(); s.photosFolder = ''; s.addinFolder = '';
    return s;
  }
  const target = () => targetFor(get());
  // Whether the folder and its photos/ are there; a target that cannot be asked says why.
  async function where(t) {
    const key = t.describe();
    if (found && found.key === key && Date.now() - found.at < FOUND_MS) return found.value;
    let value;
    try { value = { folderFound: await t.exists(''), photosFound: await t.exists('photos'), folderError: '' }; }
    catch (e) { value = { folderFound: false, photosFound: false, folderError: e.message }; }
    found = { key, at: Date.now(), value };
    return value;
  }
  async function answer() {
    const s = get();
    return Object.assign({ settings: s }, await where(targetFor(s)));
  }
  // Only what is sent changes; the rest stays as it was.
  async function save(sent) {
    const s = get();
    for (const k of KEYS) {
      if (SERVER_KEYS.includes(k) || !sent || sent[k] === undefined || sent[k] === null) continue;
      s[k] = String(sent[k]).trim().replace(/[\\/]+$/, '');
    }
    const err = check(s); if (err) return { error: err };
    const keep = {};
    for (const k of KEYS) if (!SERVER_KEYS.includes(k)) keep[k] = s[k];
    writeAtomic(file, JSON.stringify(keep));
    found = null;
    return answer();
  }
  return { get, answer, save, target, forget: () => { found = null; } };
}

module.exports = { createSettings, check };
