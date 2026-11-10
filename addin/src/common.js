/*
 * What the event handler and the task pane share: the signatures and their images (from the
 * company's own SharePoint), the user's details and groups (Microsoft Graph), the user's own choice,
 * and putting a signature into the item being written.
 *
 * Nothing of a company is in this code or where it is served: everything is read with the user's
 * own sign-in from a folder in their tenant's SharePoint (see docs/SETUP.md).
 *
 * Runs in two places. In new Outlook, Outlook on the web and Outlook for Mac, in a browser
 * runtime. In classic Outlook on Windows, the event handler runs in a JavaScript-only runtime:
 * no DOM, no localStorage, no window, no btoa - so nothing here may need them.
 */
import { LANG } from './lang.js';
import core from '../../shared/signature-core.js';
import { getGraphToken } from './auth.js';

/* global Office, OfficeRuntime, OSIGN_SETUP */
export { core };

// Where the editor publishes: a folder in the document library of a SharePoint site. "root" is the
// tenant's home site (https://<tenant>.sharepoint.com), which every Microsoft 365 tenant has.
// OSIGN_SETUP is the first line of the bundle, filled in by the editor (see addin/build.mjs).
const SITE = OSIGN_SETUP.site;
const FOLDER = OSIGN_SETUP.folder;
// The folder by its ids, as the server found it when it made these files: read through them, a
// user needs to be able to read the folder only, nothing of the site around it. Older files have
// none: the folder is then found through its site.
const DRIVE = /^[A-Za-z0-9!_.-]+$/.test(OSIGN_SETUP.drive || '') ? OSIGN_SETUP.drive : '';
const ITEM = /^[A-Za-z0-9!_.-]+$/.test(OSIGN_SETUP.item || '') ? OSIGN_SETUP.item : '';
// The add-in's own address (https://…/), where the server also puts the published images (img/).
const BASE = /^(https:\/\/[^\s"'<>]+|http:\/\/localhost(:\d+)?)\/$/.test(OSIGN_SETUP.base || '') ? OSIGN_SETUP.base : '';
// The version of the signatures loaded last (their "updated"): the images next to the add-in's
// files are used only when they are of the same publish, else the copies kept with that version.
let loadedUpdated = '';
const GRAPH = 'https://graph.microsoft.com/v1.0';
// Reading the folder needs Sites.Read.All; groups need GroupMember.Read.All. Both delegated, so the
// add-in reads only what the user may read anyway, and both with admin consent, so nobody is asked.
const BASE_SCOPES = ['User.Read', 'Sites.Read.All'];
const GROUP_SCOPE = 'GroupMember.Read.All';

const PROFILE_KEY = 'osign.profile';
const PREFS_KEY = 'osign.prefs';
const CONFIG_KEY = 'osign.config';
// The user's details are read again after this long; within it the copy is used as it is.
const PROFILE_FRESH_MS = 15 * 60 * 1000;
// Images are kept between messages up to this size each (they are logos, icons, small photos).
const IMAGE_CACHE_MAX = 300 * 1024;
const IMAGE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif' };
// What the add-in attaches to show its images starts with this, so it knows its own again.
const ATTACHMENT_PREFIX = 'osign-';

// ctrl: an AbortController, stopped when the time is up (so a request does not go on unseen).
function withTimeout(promise, ms, what, ctrl) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => {
      reject(new Error(what + ' took too long.'));
      try { if (ctrl) ctrl.abort(); } catch (e) { /* already over */ }
    }, ms);
    promise.then((v) => { clearTimeout(t); resolve(v); }, (e) => { clearTimeout(t); reject(e); });
  });
}

// A request and the reading of its answer (read(response)), both within ms: a server that sends
// the headers and then stalls in the body is cut off too. Where there is no AbortController
// (classic Outlook's runtime may have none), the request goes on, but nobody waits for it.
// No network at all (fetch fails with a TypeError) is said as that.
function fetchWithin(url, init, ms, what, read) {
  const ctrl = typeof AbortController === 'function' ? new AbortController() : null;
  const p = (async () => {
    let r;
    try { r = await fetch(url, ctrl ? Object.assign({}, init, { signal: ctrl.signal }) : init); }
    catch (e) { throw unreachable(e, what); }
    try { return await read(r); }
    catch (e) { throw unreachable(e, what); }
  })();
  return withTimeout(p, ms, what, ctrl);
}
function unreachable(e, what) {
  return e && (e.name === 'TypeError' || e instanceof TypeError) ? new Error(what + ': cannot be reached.') : e;
}

// ---------------------------------------------------------------------------------------------
// Storage between messages: localStorage where there is one, OfficeRuntime.storage in the
// JavaScript-only runtime, memory otherwise.
// ---------------------------------------------------------------------------------------------
const memory = {};
// Written to both where both are there: classic Outlook runs the event handler in a runtime of its
// own without localStorage, and reads there (through OfficeRuntime.storage) what the pane kept, for
// when the sign-in is refused to it.
const runtimeStore = () => { try { return typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage ? OfficeRuntime.storage : null; } catch (e) { return null; } };
const local = () => { try { return typeof localStorage !== 'undefined' ? localStorage : null; } catch (e) { return null; } };
const store = {
  async get(key) {
    try { const l = local(); const v = l && l.getItem(key); if (v) return v; } catch (e) { /* none */ }
    try { const r = runtimeStore(); const v = r && await r.getItem(key); if (v) return v; } catch (e) { /* none */ }
    return memory[key] || null;
  },
  async set(key, value) {
    memory[key] = value;
    try { const l = local(); if (l) l.setItem(key, value); } catch (e) { /* full, or none */ }
    try { const r = runtimeStore(); if (r) await r.setItem(key, value); } catch (e) { /* none */ }
  }
};
async function readJson(key) { try { const v = await store.get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; } }

// ---------------------------------------------------------------------------------------------
// Microsoft Graph
// ---------------------------------------------------------------------------------------------
// What went wrong, in German for a German Outlook: the messages the pane and the notice show.
const WHAT_DE = {
  'Finding the signatures': 'Signaturen suchen', 'Loading the signatures': 'Signaturen laden', 'Finding the site': 'Website suchen',
  'Reading your details': 'Ihre Angaben lesen', 'Reading your groups': 'Ihre Gruppen lesen', 'Finding your entry': 'Ihren Eintrag suchen',
  'Loading your entry': 'Ihren Eintrag laden', 'Signing in': 'Anmelden', 'Listing the images': 'Bilder auflisten', 'Finding the photo': 'Foto suchen'
};
const MSG_DE = [
  [/SharePoint does not let you read the folder (.+?) \((\d+)\)\. It has to be shared with everyone who gets a signature: .*$/, 'SharePoint lässt Sie den Ordner $1 nicht lesen ($2). Er muss für alle freigegeben sein, die eine Signatur bekommen: in SharePoint beim Ordner → Zugriff verwalten → Zugriff gewähren: «Jeder ausser externen Benutzern», Anzeigen.'],
  [/not allowed \((\d+)\)\. The app registration needs admin consent for (.+?) \(.*$/, 'nicht erlaubt ($1). Der App-Registrierung fehlt die Administratorzustimmung für $2 (Entra ID → Unternehmensanwendungen → die App des Add-ins → Berechtigungen; das Setup-Skript erteilt sie).'],
  [/You are not in the AD export, so your groups are not known\.$/, 'Sie sind nicht im AD-Export, darum sind Ihre Gruppen nicht bekannt.'],
  [/You are not in the AD export \((.+?)\)\. Ask IT to export the users again\.$/, 'Sie sind nicht im AD-Export ($1). Bitten Sie die IT, die Benutzer neu zu exportieren.'],
  [/No signatures are published yet: (.+) is not in SharePoint\.$/, 'Noch keine Signaturen veröffentlicht: $1 fehlt in SharePoint.'],
  [/^In classic Outlook, open OSign → Signatures once; from then on it goes in by itself\.$/, 'Im klassischen Outlook einmal OSign → Signaturen öffnen; danach fügt es sie selbst ein.'],
  [/^These add-in files have no app registration \(client ID\)\.$/, 'Diesen Add-in-Dateien fehlt die App-Registrierung (Client-ID). Bitten Sie die IT, sie neu zu erstellen.'],
  [/^This Outlook cannot sign in add-ins \(nested app authentication\)\. Update Outlook\.$/, 'Dieses Outlook kann Add-ins nicht anmelden (Nested App Authentication). Aktualisieren Sie Outlook.'],
  [/^The add-in was made for the SharePoint site "(.*)": make its files again in NOBS OSign\.$/, 'Das Add-in wurde für die SharePoint-Website «$1» erstellt: Die IT muss seine Dateien in NOBS OSign neu erstellen.'],
  [/^Finding the site: Microsoft Graph did not name its id\.$/, 'Website suchen: Microsoft Graph nannte keine ID.'],
  [/^The published signatures are not valid: /, 'Die veröffentlichten Signaturen sind ungültig: '],
  [/^Outlook refused\.$/, 'Outlook hat abgelehnt.'],
  [/ took too long\.$/, ' dauerte zu lange.'],
  [/: not found in SharePoint \(404\)\.$/, ': in SharePoint nicht gefunden (404).'],
  [/: (Microsoft Graph|SharePoint) answered (\d+)\.$/, ': $1 antwortete $2.'],
  [/: cannot be reached\.$/, ': nicht erreichbar.'],
  [/^your details could not be read\.$/, 'Ihre Angaben konnten nicht gelesen werden.']
];
// The language of Outlook's display: German, French, Italian, else English.
export function langOf(code) { const m = /^(de|fr|it)/i.exec(String(code || '')); return m ? m[1].toLowerCase() : 'en'; }
// A message in that language (true, as before, is German).
export function localize(message, lang) {
  let m = String(message || '');
  if (lang === true) lang = 'de';
  if (!lang || lang === 'en') return m;
  const L = lang === 'de' ? { what: WHAT_DE, msg: MSG_DE.map((x) => x[1]) } : LANG[lang];
  if (!L) return m;
  for (const [en, to] of Object.entries(L.what)) m = m.split(en + ':').join(to + ':').split(en + ' took').join(to + ' took');
  MSG_DE.forEach(([rx], i) => { if (L.msg[i]) m = m.replace(rx, L.msg[i]); });
  return m;
}

// The permissions a token carries (its scp claim). Decoded by hand: classic Outlook's runtime has
// no atob.
const B64URL = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
// The permissions a token carries (its scp claim), or null.
function scopesOf(token) {
  try {
    const part = String(token).split('.')[1].replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    let bits = 0, n = 0, bytes = [];
    for (const ch of part) { n = (n << 6) | B64URL.indexOf(ch); bits += 6; if (bits >= 8) { bits -= 8; bytes.push((n >> bits) & 255); } }
    const json = decodeURIComponent(bytes.map((b) => '%' + ('0' + b.toString(16)).slice(-2)).join(''));
    return String(JSON.parse(json).scp || '').split(' ').filter(Boolean);
  } catch (e) { return null; }
}

// One GET to Graph with the user's token: its JSON, null for a 404 when allowed, or an error in words.
async function graphGet(token, url, what, allowNotFound) {
  return fetchWithin(url, { headers: { Authorization: 'Bearer ' + token } }, 8000, what, (r) => graphAnswer(r, token, url, what, allowNotFound));
}
function graphAnswer(r, token, url, what, allowNotFound) {
  if (r.status === 404 && allowNotFound) return null;
  if (r.status === 401 || r.status === 403) {
    const need = /memberOf/i.test(url) ? GROUP_SCOPE : /\/(sites|drives)\//.test(url) ? 'Sites.Read.All' : 'User.Read';
    // The consent is there (the token carries it): then SharePoint refused you, not the app.
    const scp = scopesOf(token);
    if (need === 'Sites.Read.All' && scp && scp.indexOf(need) >= 0) {
      throw new Error(what + ': SharePoint does not let you read the folder ' + FOLDER + ' (' + r.status + '). It has to be shared with everyone who gets a signature: in SharePoint, the folder → Manage access → Grant access: Everyone except external users, Can view.');
    }
    throw new Error(what + ': not allowed (' + r.status + '). The app registration needs admin consent for ' + need + ' (Entra ID → Enterprise applications → the add-in\'s app → Permissions; the setup script gives it).');
  }
  if (r.status === 404) throw new Error(what + ': not found in SharePoint (404).');
  if (!r.ok) throw new Error(what + ': Microsoft Graph answered ' + r.status + '.');
  return r.json();
}

// MSAL keeps the token and renews it; this only saves asking it again for each call of one message.
// Kept well below a token's life (an hour at least): a runtime Outlook reuses never sends an expired one.
const TOKEN_REUSE_MS = 10 * 60 * 1000;
// Outlook's sign-in answers in a second or two. When it does not answer at all, the event handler
// would wait until Outlook gives up on it, silently: so it waits this long, then says so.
const SIGN_IN_MS = 20 * 1000;
// A sign-in without a window that failed is not tried again for this long (the pane's own, with a
// window, always is): else one message would wait for it once for each thing it reads.
const SIGN_IN_RETRY_MS = 5 * 60 * 1000;
let tokens = {};
// A Graph token for the add-in (nested app authentication), cached until shortly before it runs out.
async function token(scopes, interactive) {
  const key = scopes.join(' ');
  const had = tokens[key];
  if (!had || interactive || Date.now() - had.at > (had.failed ? SIGN_IN_RETRY_MS : TOKEN_REUSE_MS)) {
    const entry = { at: Date.now() };
    const p = getGraphToken(interactive, scopes);
    entry.p = (interactive ? p : withTimeout(p, SIGN_IN_MS, 'Signing in')).catch((e) => {
      if (tokens[key] === entry) {
        if (interactive) delete tokens[key];
        else { entry.failed = true; entry.at = Date.now(); }
      }
      throw e;
    });
    tokens[key] = entry;
  }
  return tokens[key].p;
}

// The site, by its id: Graph does not take a path into a library after a site's own path
// (/sites/host:/sites/x:/drive/root:/… answers 400), so a site named by its address is looked up
// once, and its id kept.
const SITE_KEY = 'osign.site.' + SITE;
let siteId = null;
function siteUrl() {
  // As addin/fill.mjs accepts it: "Sites" or "sites" alike.
  const m = /^([A-Za-z0-9.-]+):\/(sites|teams|personal)\/([^:/]+):$/i.exec(SITE);
  if (!m) throw new Error('The add-in was made for the SharePoint site "' + SITE + '": make its files again in NOBS OSign.');
  return GRAPH + '/sites/' + m[1] + ':/' + m[2].toLowerCase() + '/' + encodeURIComponent(m[3]) + ':';
}
// The Graph address of the site's drive the signatures are in, found once.
async function siteBase(tok) {
  if (!SITE || /^root$/i.test(SITE)) return GRAPH + '/sites/root';
  if (!siteId) {
    siteId = (async () => {
      const kept = await store.get(SITE_KEY);
      if (kept) return kept;
      const j = await graphGet(tok, siteUrl() + '?$select=id', 'Finding the site');
      if (!j || !j.id) throw new Error('Finding the site: Microsoft Graph did not name its id.');
      await store.set(SITE_KEY, j.id);
      return j.id;
    })();
    siteId.catch(() => { siteId = null; });
  }
  return GRAPH + '/sites/' + encodeURIComponent(await siteId);
}

// A file in the OSign folder, by its path in there ("signatures.json", "assets/logo.png").
async function itemUrl(tok, rel) {
  if (DRIVE && ITEM) return GRAPH + '/drives/' + encodeURIComponent(DRIVE) + '/items/' + encodeURIComponent(ITEM) + ':/' + String(rel).split('/').filter(Boolean).map(encodeURIComponent).join('/');
  const parts = (FOLDER + '/' + rel).split('/').filter(Boolean).map(encodeURIComponent);
  return (await siteBase(tok)) + '/drive/root:/' + parts.join('/');
}

// A file of the signatures' folder, as text or as bytes.
async function download(meta, tok, what, asText) {
  // The address Graph hands out is signed for a short while: no token is sent there.
  const url = meta['@microsoft.graph.downloadUrl'];
  const target = url || (await itemUrl(tok, meta.__rel)) + ':/content';
  return fetchWithin(target, url ? {} : { headers: { Authorization: 'Bearer ' + tok } }, 10000, what, async (r) => {
    if (!r.ok) throw new Error(what + ': SharePoint answered ' + r.status + '.');
    return asText ? r.text() : new Uint8Array(await r.arrayBuffer());
  });
}

// ---------------------------------------------------------------------------------------------
// The configuration the editor published
// ---------------------------------------------------------------------------------------------
// interactive: the task pane may show a sign-in window; the event handler never can.
export async function loadConfig(opts) {
  opts = opts || {};
  const cached = await readJson(CONFIG_KEY);
  try {
    const tok = await token(BASE_SCOPES, !!opts.interactive);
    const meta = await graphGet(tok, await itemUrl(tok, 'signatures.json'), 'Finding the signatures', true);
    if (!meta) throw new Error('No signatures are published yet: ' + FOLDER + '/signatures.json is not in SharePoint.');
    meta.__rel = 'signatures.json';
    // Unchanged since the last message: no need to fetch it again.
    const text = cached && cached.etag && cached.etag === meta.eTag ? cached.text : await download(meta, tok, 'Loading the signatures', true);
    const cfg = JSON.parse(text);
    const errs = core.validate(cfg);
    if (errs.length) throw new Error('The published signatures are not valid: ' + errs[0]);
    if (!cached || cached.etag !== meta.eTag) await store.set(CONFIG_KEY, JSON.stringify({ etag: meta.eTag || '', text }));
    // A pilot (published by 0.9.57-0.9.60): not any more; everyone gets the version for all.
    delete cfg.pilot;
    loadedUpdated = String(cfg.updated || '');
    return cfg;
  } catch (e) {
    if (cached && cached.text) {
      try {
        const cfg = JSON.parse(cached.text);
        console.log('OSign: signatures from the copy kept before (' + (e && e.message ? e.message : e) + ').');
        delete cfg.pilot;
        loadedUpdated = String(cfg.updated || '');
        return cfg;
      } catch (e2) { /* below */ }
    }
    const classic = classicCopy();
    if (classic) { console.log('OSign: signatures from the copy in the mailbox (' + (e && e.message ? e.message : e) + ').'); loadedUpdated = String(classic.cfg.updated || ''); return classic.cfg; }
    throw e;
  }
}

// ---------------------------------------------------------------------------------------------
// Roaming settings: kept in the mailbox, so they follow the user to every Outlook.
// ---------------------------------------------------------------------------------------------
function roaming() { return Office.context && Office.context.roamingSettings; }
function readRoaming(key) { try { const v = roaming() && roaming().get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
// Answers whether Outlook kept it: the mailbox holds 32 KB of these at most. Refused, the value
// before is put back: Outlook saves all settings at once, so one too large left in there would
// have every later save refused too.
function writeRoaming(key, value) {
  return new Promise((resolve) => {
    const rs = roaming(); if (!rs) { resolve(false); return; }
    let had = null;
    try { had = rs.get(key); } catch (e) { had = null; }
    const undo = () => { try { if (had !== undefined && had !== null) rs.set(key, had); else rs.remove(key); } catch (e) { /* nothing to put back */ } };
    try {
      rs.set(key, JSON.stringify(value));
      rs.saveAsync((r) => {
        const ok = !r || !r.status || r.status === 'succeeded' || (Office.AsyncResultStatus && r.status === Office.AsyncResultStatus.Succeeded);
        if (!ok) undo();
        resolve(!!ok);
      });
    } catch (e) { undo(); resolve(false); }
  });
}
// How much a setting takes in the mailbox: Outlook keeps them as JSON, so the text kept (itself
// JSON) is quoted once more, and counted in UTF-8 bytes.
export function storedSize(text) {
  const s = JSON.stringify(typeof text === 'string' ? text : JSON.stringify(text));
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80) n += 1;
    else if (c < 0x800) n += 2;
    else if (c >= 0xd800 && c < 0xdc00 && i + 1 < s.length) { n += 4; i++; }
    else n += 3;
  }
  return n;
}

export function getPrefs() { return readRoaming(PREFS_KEY) || {}; }
// The user's own variants of the company signatures (see core.withVariants), in the mailbox.
const VARIANTS_KEY = 'osign.variants';
export function getVariants() { const v = readRoaming(VARIANTS_KEY); return Array.isArray(v) ? v : []; }
export function setVariants(list) { return writeRoaming(VARIANTS_KEY, (list || []).slice(0, core.VARIANT_MAX)); }
export function setPrefs(p) { return writeRoaming(PREFS_KEY, { new: p.new || '', reply: p.reply || '' }); }

// ---------------------------------------------------------------------------------------------
// The user's details and groups, from Microsoft Graph (/me)
// ---------------------------------------------------------------------------------------------
// The user's groups, nested ones included, as the names the configuration uses: only those are
// kept (a user can be in hundreds of groups, and the copy goes into the mailbox settings).
async function readGroups(tok, wanted) {
  const found = [];
  let url = GRAPH + '/me/transitiveMemberOf/microsoft.graph.group?$select=id,displayName,onPremisesSamAccountName&$top=999';
  for (let page = 0; url && page < 20; page++) {
    const j = await graphGet(tok, url, 'Reading your groups');
    (j.value || []).forEach((g) => {
      [g.id, g.displayName, g.onPremisesSamAccountName].forEach((n) => {
        const l = String(n || '').toLowerCase();
        if (l && wanted.indexOf(l) >= 0 && found.indexOf(l) < 0) found.push(l);
      });
    });
    url = j['@odata.nextLink'];
  }
  return found;
}

// The user's own file of values exported from AD (users/<mail>.json in the OSign folder), or null.
async function readAdFile(tok, mail) {
  const name = core.userFileName(mail);
  if (!name) return null;
  const meta = await graphGet(tok, await itemUrl(tok, 'users/' + name), 'Finding your entry', true);
  if (!meta) return null;
  meta.__rel = 'users/' + name;
  const j = JSON.parse(await download(meta, tok, 'Loading your entry', true));
  return { values: j.values || {}, groups: (j.groups || []).map((g) => String(g).toLowerCase()) };
}

// The user's details from where the configuration says (options.valueSource): Entra ID through
// Graph (/me), the AD export in SharePoint, or each field from where it has a source.
async function readGraph(cfg, interactive, wanted) {
  const mode = core.valueSource(cfg);
  const tok = await token(BASE_SCOPES, interactive);
  let graphValues = {};
  const gFields = core.graphFields(cfg);
  if (gFields.length) {
    const select = core.graphSelect(gFields);
    const me = await graphGet(tok, GRAPH + '/me' + (select.length ? '?$select=' + encodeURIComponent(select.join(',')) : ''), 'Reading your details');
    graphValues = core.valuesFromGraph(gFields, me);
  }
  let groups = [], groupError = null;
  if (mode === 'graph') {
    // Groups that cannot be read (no consent yet) take only the group signatures out of reach.
    if (wanted.length) {
      try { groups = await readGroups(await token([GROUP_SCOPE], interactive), wanted); }
      catch (e) { groups = null; groupError = e; }
    }
    return { values: core.mergeValues(cfg, graphValues, {}), groups, groupError };
  }
  // From AD: the file of the address this mailbox has, as the editor exported it.
  const mail = (Office.context.mailbox.userProfile || {}).emailAddress || '';
  const ad = await readAdFile(tok, mail);
  if (!ad && mode === 'ad') throw new Error('You are not in the AD export (users/' + (core.userFileName(mail) || mail) + '). Ask IT to export the users again.');
  return {
    values: core.mergeValues(cfg, graphValues, ad ? ad.values : {}),
    groups: ad ? ad.groups.filter((g) => wanted.indexOf(g) >= 0) : null,
    groupError: ad ? null : new Error('You are not in the AD export, so your groups are not known.')
  };
}

// Returns { values, groups, source: 'graph'|'cache'|'outlook', error } - or values null when there
// are none. groups is null when they could not be read: no signature limited to groups applies then.
export async function loadUserValues(cfg, opts) {
  opts = opts || {};
  const wanted = core.groupsUsed(cfg);
  const key = core.valueSource(cfg) + '\n' + wanted.slice().sort().join('\n');
  const cached = readRoaming(PROFILE_KEY);
  // A copy read for other groups than the configuration now names says nothing about the new ones.
  const cacheFits = cached && cached.values && (cached.groupsFor || '') === key;
  if (!opts.refresh && cacheFits && Date.now() - cached.at < PROFILE_FRESH_MS) return { values: cached.values, groups: cached.groups || [], source: 'cache' };
  let error = null;
  try {
    const r = await readGraph(cfg, !!opts.interactive, wanted);
    // Only from a configuration that names its fields (not ever an empty set read from too little).
    if ((cfg.fields || []).length) await writeRoaming(PROFILE_KEY, { at: Date.now(), values: r.values, groups: r.groups, groupsFor: r.groups ? key : '\u0000' });
    return { values: r.values, groups: r.groups, source: 'graph', error: r.groupError };
  } catch (e) { error = e; }
  if (cached && cached.values) return { values: cached.values, groups: cacheFits ? cached.groups || [] : null, source: 'cache', error };
  if (cfg.options && cfg.options.insertWithoutProfile) {
    // What Outlook itself knows: the address, nothing else that fits a field reliably.
    const v = {};
    (cfg.fields || []).forEach((f) => { v[f.key] = ''; });
    const up = Office.context.mailbox && Office.context.mailbox.userProfile;
    if (up && 'mail' in v) v.mail = up.emailAddress || '';
    return { values: v, groups: null, source: 'outlook', error };
  }
  return { values: null, groups: null, source: 'none', error };
}

// ---------------------------------------------------------------------------------------------
// Images: assets/<name> (logos, icons) and photos/<name> (one per person), from the OSign folder.
// ---------------------------------------------------------------------------------------------
const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
export function base64(bytes) {
  let out = '', i = 0;
  for (; i + 2 < bytes.length; i += 3) {
    const n = (bytes[i] << 16) | (bytes[i + 1] << 8) | bytes[i + 2];
    out += B64[n >> 18 & 63] + B64[n >> 12 & 63] + B64[n >> 6 & 63] + B64[n & 63];
  }
  if (i < bytes.length) {
    const n = (bytes[i] << 16) | ((i + 1 < bytes.length ? bytes[i + 1] : 0) << 8);
    out += B64[n >> 18 & 63] + B64[n >> 12 & 63] + (i + 1 < bytes.length ? B64[n >> 6 & 63] : '=') + '=';
  }
  return out;
}

function typeOf(name) { return IMAGE_TYPES[String(name).split('.').pop().toLowerCase()] || ''; }
// The attachment's name, without its osign-<tag>- prefix (insertSignature puts that in front).
function fileOf(w) { return w.dir + '-' + w.name.replace(/[^A-Za-z0-9._-]+/g, '-'); }
// An image as it was kept before (this Outlook's copy, else the one in the mailbox), or null: for
// when SharePoint cannot be read now.
async function keptImage(w) {
  const cached = await readJson('osign.img.' + w.key.toLowerCase());
  if (cached && cached.b64) return { b64: cached.b64, type: typeOf(w.name), file: fileOf(w) };
  const classic = classicCopy();
  return classic && classic.images && classic.images[w.key] ? classic.images[w.key] : null;
}

// The published images next to the add-in's files (img/, with img/index.json naming them): read
// without a sign-in, so classic Outlook's event handler, refused one, still puts them in as they
// are, not the smaller copies kept in the mailbox. The index is looked at once a minute at most;
// an image is cached by its file (a new image is a new file). null when not there (no index, the
// network gone, a photo): the copies kept before then.
const HOSTED_KEY = 'osign.hosted';
// At most this much of them kept in this Outlook (its storage holds about 10 MB in all), and an
// image downloaded only up to the size the server takes (1 MB).
const HOSTED_CACHE_MAX = 2 * 1024 * 1024, HOSTED_FILE_MAX = 1024 * 1024;
let hosted = null, hostedAt = 0;
// The index ({ updated, files: { name: 'img/<hash>.png' } }), looked at once a minute at most,
// also when it could not be had (no network, a host that answers something else): then {}.
async function hostedIndex() {
  if (!BASE) return null;
  if (hostedAt && Date.now() - hostedAt < 60 * 1000) return hosted;
  hostedAt = Date.now();
  try {
    const j = await fetchWithin(BASE + 'img/index.json', { cache: 'no-cache' }, 8000, 'The images', (r) => (r.ok ? r.json() : null));
    hosted = j && typeof j === 'object' && !Array.isArray(j) && j.files && typeof j.files === 'object' && !Array.isArray(j.files) ? j : {};
  } catch (e) { console.log('OSign: hosted images: ' + (e && e.message ? e.message : e)); hosted = {}; }
  return hosted;
}
// An image's file in the index (its name as published, in any case), or '': only for the
// signatures of that same publish.
function hostedFile(index, name) {
  const files = index && index.files;
  if (!files || !index.updated || String(index.updated) !== loadedUpdated) return '';
  const f = files[name] || files[Object.keys(files).find((k) => k.toLowerCase() === String(name).toLowerCase())];
  return typeof f === 'string' && /^img\/[0-9a-f]{24}\.(png|jpg|gif)$/.test(f) ? f : '';
}
// Whether the bytes are the image the name says (an error page is not).
function isImage(b, file) {
  const ext = file.split('.').pop();
  if (ext === 'png') return b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47;
  if (ext === 'jpg') return b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff;
  return b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x38;
}
async function hostedImage(w) {
  if (w.dir !== 'assets') return null;
  const index = await hostedIndex();
  const file = hostedFile(index, w.name);
  if (!file) return null;
  // Kept in one place, by file, only for the files the index names now (an image no longer
  // published goes from this Outlook's storage too), at most HOSTED_CACHE_MAX of them.
  const cache = (await readJson(HOSTED_KEY)) || {};
  if (cache[file]) return { b64: cache[file], type: typeOf(w.name), file: fileOf(w) };
  try {
    const bytes = await fetchWithin(BASE + file, {}, 15000, 'Loading ' + w.name, async (r) => {
      if (!r.ok) throw new Error('Loading ' + w.name + ': ' + r.status);
      const n = Number(r.headers && r.headers.get ? r.headers.get('content-length') : 0);
      if (n > HOSTED_FILE_MAX) throw new Error('Loading ' + w.name + ': too large');
      return new Uint8Array(await r.arrayBuffer());
    });
    if (bytes.length > HOSTED_FILE_MAX || !isImage(bytes, file)) throw new Error('Loading ' + w.name + ': not an image');
    const b64 = base64(bytes);
    const now = new Set(Object.values(index.files));
    const keep = {}; let total = b64.length;
    Object.keys(cache).forEach((f) => { if (now.has(f) && total + cache[f].length <= HOSTED_CACHE_MAX) { keep[f] = cache[f]; total += cache[f].length; } });
    if (b64.length <= HOSTED_CACHE_MAX) keep[file] = b64;
    await store.set(HOSTED_KEY, JSON.stringify(keep));
    return { b64, type: typeOf(w.name), file: fileOf(w) };
  } catch (e) { console.log('OSign: hosted image ' + w.name + ': ' + (e && e.message ? e.message : e)); return null; }
}
// An image when SharePoint cannot be read now: next to the add-in's files as it is, else a copy kept before.
async function fallbackImage(w) { return (await hostedImage(w)) || keptImage(w); }

// { 'assets/logo.png': { b64, type, file } } for what the signature shows; what is not there is
// left out of the map (and so out of the signature).
async function loadImages(refs, interactive) {
  const out = {};
  const wanted = [];
  refs.forEach((r) => { const k = r.dir + '/' + r.name; if (typeOf(r.name) && !wanted.some((w) => w.key === k)) wanted.push({ key: k, dir: r.dir, name: r.name }); });
  if (!wanted.length) return out;
  let tok;
  try { tok = await token(BASE_SCOPES, interactive); } catch (e) {
    // No sign-in (classic Outlook's event runtime is refused it): the images next to the add-in's
    // files, else the copies kept before.
    // All at once, and not longer than 20 seconds together (Outlook waits for the signature).
    const got = await withTimeout(Promise.all(wanted.map((w) => fallbackImage(w).catch(() => null))), 20000, 'The images').catch(() => null);
    if (got) wanted.forEach((w, i) => { if (got[i]) out[w.key] = got[i]; });
    else for (const w of wanted) { const k = await keptImage(w); if (k) out[w.key] = k; }
    if (!Object.keys(out).length) throw e;
    console.log('OSign: images without SharePoint (sign-in: ' + (e && e.message ? e.message : e) + ').');
    return out;
  }
  // The logos and icons: one listing tells which are there and whether they changed. Not to be
  // had (offline, SharePoint busy): null, and the copies kept before are used.
  let assets = null;
  if (wanted.some((w) => w.dir === 'assets')) {
    try {
      const j = await graphGet(tok, (await itemUrl(tok, 'assets')) + ':/children?$top=999', 'Listing the images', true);
      assets = {};
      ((j && j.value) || []).forEach((it) => { assets[String(it.name).toLowerCase()] = it; });
    } catch (e) { console.log('OSign: images: ' + (e && e.message ? e.message : e)); }
  }
  // The placeholder photo, for whoever has none of their own: looked for once per extension.
  const placeholders = {};
  // The placeholder photo, for people without one of their own.
  const placeholder = async (ext) => {
    if (!(ext in placeholders)) {
      placeholders[ext] = null;
      try {
        const m = await graphGet(tok, await itemUrl(tok, 'photos/' + core.PHOTO_DEFAULT + '.' + ext), 'Finding the photo', true);
        if (m) { m.__rel = 'photos/' + core.PHOTO_DEFAULT + '.' + ext; placeholders[ext] = m; }
      } catch (e) { /* none */ }
    }
    return placeholders[ext];
  };
  for (const w of wanted) {
    try {
      if (w.dir === 'assets' && !assets) { const k = await fallbackImage(w); if (k) out[w.key] = k; continue; }
      let meta = w.dir === 'assets' ? assets[w.name.toLowerCase()] : await graphGet(tok, await itemUrl(tok, 'photos/' + w.name), 'Finding the photo', true);
      if (meta) meta.__rel = w.key;
      else if (w.dir === 'photos') meta = await placeholder(w.name.split('.').pop().toLowerCase());
      // Not in SharePoint (it said so): not in the signature either.
      if (!meta) continue;
      const ck = 'osign.img.' + w.key.toLowerCase();
      const cached = await readJson(ck);
      let b64 = cached && cached.etag && cached.etag === meta.eTag ? cached.b64 : null;
      if (!b64) {
        b64 = base64(await download(meta, tok, 'Loading ' + w.name));
        if (b64.length < IMAGE_CACHE_MAX * 4 / 3) await store.set(ck, JSON.stringify({ etag: meta.eTag || '', b64 }));
      }
      out[w.key] = { b64, type: typeOf(w.name), file: fileOf(w) };
    } catch (e) {
      // Not to be read now: the copy kept before, if there is one.
      console.log('OSign: image ' + w.key + ': ' + (e && e.message ? e.message : e));
      try { const k = await fallbackImage(w); if (k) out[w.key] = k; } catch (e2) { /* without it */ }
    }
  }
  return out;
}

// Each of the signature's own images becomes either an inline attachment (cid:, in a mail) or a
// data: address (the pane's preview). One that could not be loaded goes.
function placeImages(html, images, mode) {
  return core.localImages(html).reduce((h, r) => {
    const img = images[r.dir + '/' + r.name];
    if (!img) return h.split(r.tag).join('');
    const src = mode === 'cid' ? 'cid:' + img.file : 'data:' + img.type + ';base64,' + img.b64;
    return h.split(r.tag).join(r.tag.replace(r.src, src));
  }, html);
}

function renderHtml(sig, values) {
  return core.render(sig.compiled || sig.html, core.valuesFor(sig, values), { removeEmptyLines: sig.removeEmptyLines !== false });
}

// For the pane: the signature as the mail will show it, images inside.
export async function previewHtml(sig, values) {
  const html = renderHtml(sig, values);
  let images = {};
  try { images = await loadImages(core.localImages(html), false); } catch (e) { /* shown without */ }
  return placeImages(html, images, 'data');
}

// For a signature of one's own: the signature filled in, its images shown (data:) but each
// remembering where it is from (data-osign-src), so that is what is kept.
export async function editableHtml(sig, values) {
  const html = renderHtml(sig, values);
  let images = {};
  try { images = await loadImages(core.localImages(html), false); } catch (e) { /* shown without */ }
  return core.localImages(html).reduce((h, r) => {
    const img = images[r.dir + '/' + r.name];
    if (!img) return h;
    const tag = r.tag.replace(r.src, 'data:' + img.type + ';base64,' + img.b64).replace(/^<img\b/i, '<img data-osign-src="' + r.src.replace(/"/g, '&quot;') + '"');
    return h.split(r.tag).join(tag);
  }, html);
}

// ---------------------------------------------------------------------------------------------
// Classic Outlook on Windows: its event runtime is refused the sign-in, and shares with the pane
// nothing but the add-in's settings in the mailbox (roaming settings, 32 KB in all). So the pane
// keeps there the signatures that would go in by themselves, with their images made as small as
// they are shown, as many as fit; the event handler takes them from there.
// ---------------------------------------------------------------------------------------------
const CLASSIC_KEY = 'osign.classic';
// What the mailbox holds for all of the add-in's settings, less some room to spare.
const ROAMING_ROOM = 30 * 1024;
export function isClassicOutlook() {
  try { const d = Office.context.diagnostics; return !!d && d.platform === 'PC'; } catch (e) { return false; }
}
// The steps an image is made smaller by when it does not fit into classic Outlook's copy: its
// width against the width shown, and a JPEG's quality (a PNG only gets narrower). The same in the
// editor's estimate (10-publish.js, CLASSIC_SQUEEZE).
export const CLASSIC_SQUEEZE = [{ scale: 1, quality: 0.7 }, { scale: 1, quality: 0.55 }, { scale: 0.8, quality: 0.5 }, { scale: 0.65, quality: 0.45 }];
// shrink(image, shownWidth, step) → a smaller { b64, type } or the same; only the pane has the canvas for it.
export async function keepForClassic(cfg, sigs, values, shrink) {
  const used = [PREFS_KEY, VARIANTS_KEY, PROFILE_KEY].reduce((n, k) => { try { const v = roaming() && roaming().get(k); return n + (v ? storedSize(String(v)) + k.length + 4 : 0); } catch (e) { return n; } }, 0);
  const room = ROAMING_ROOM - used;
  const keep = ['id', 'name', 'compiled', 'removeEmptyLines', 'defaultNew', 'defaultReply', 'audience', 'validFrom', 'validUntil', 'variantOf', 'userValues'];
  // The fields as far as reading the user's details and their variants needs them: without them
  // the details read from here would come out empty, and be kept so for every Outlook.
  const fields = (cfg.fields || []).map((f) => { const o = { key: f.key }; ['graph', 'ad', 'userEditable'].forEach((k) => { if (f[k]) o[k] = f[k]; }); return o; });
  const copy = {
    at: cfg.updated || '',
    cfg: { schemaVersion: cfg.schemaVersion, updated: cfg.updated, options: cfg.options, fields, signatures: sigs.map((x) => {
      const o = {}; keep.forEach((k) => { if (x[k] !== undefined) o[k] = x[k]; }); if (!o.compiled) o.compiled = x.html; return o;
    }), campaigns: classicCampaigns(cfg, values) },
    images: {}
  };
  // sigs come most needed first: the later ones go when not all fit (the banners first).
  const size = () => storedSize(copy) + CLASSIC_KEY.length + 4;
  if (size() > room) copy.cfg.campaigns = [];
  while (copy.cfg.signatures.length > 1 && size() > room) copy.cfg.signatures.pop();
  if (size() > room) { console.log('OSign: the signatures do not fit into the mailbox settings for classic Outlook.'); return { kept: false, images: 0, left: 0 }; }
  sigs = sigs.slice(0, copy.cfg.signatures.length);
  // The images, each as small as it is shown; the smallest first, as many as fit.
  const wanted = [];
  const banners = (copy.cfg.campaigns || []).flatMap((c) => [c].concat(c.more || []).map((x) => ({ compiled: x.compiled })));
  for (const sg of sigs.concat(banners)) {
    const html = renderHtml(sg, values);
    for (const r of core.localImages(html)) {
      const key = r.dir + '/' + r.name;
      if (wanted.some((w) => w.key === key)) continue;
      // A width in pixels only: width="100%" says nothing about how many.
      const wm = /\bwidth\s*=\s*["']?(\d+)(?![\d%])/i.exec(r.tag);
      wanted.push({ key, ref: r, width: wm ? Number(wm[1]) : 0 });
    }
  }
  let images = {};
  try { images = await loadImages(wanted.map((w) => w.ref), false); } catch (e) { /* the text alone */ }
  const small = [];
  for (const w of wanted) {
    const img = images[w.key]; if (!img) continue;
    let s2 = img;
    try { if (shrink) s2 = Object.assign({}, img, await shrink(img, w.width)); } catch (e) { /* as it is */ }
    small.push([w.key, s2, w]);
  }
  small.sort((a, b) => a[1].b64.length - b[1].b64.length);
  let left = 0, squeezed = 0;
  // Whether the copy still fits with this image in it (counted as Outlook stores it).
  const fits = (key, img) => { copy.images[key] = { b64: img.b64, type: img.type, file: img.file }; const ok = size() <= room; if (!ok) delete copy.images[key]; return ok; };
  for (const [key, img, w] of small) {
    // Too large as it is: made smaller step by step, for this copy only (the mail from the pane,
    // and every other Outlook, has the image as it is), until it fits or the last step is done.
    let cur = img;
    for (let l = 0; shrink && !fits(key, cur) && l < CLASSIC_SQUEEZE.length; l++) {
      try { const x = Object.assign({}, images[key], await shrink(images[key], w.width, CLASSIC_SQUEEZE[l])); if (x.b64 && x.b64.length < cur.b64.length) cur = x; } catch (e) { break; }
    }
    if (!fits(key, cur)) { left++; continue; }
    if (cur !== img) squeezed++;
  }
  let had = '';
  try { had = roaming() ? String(roaming().get(CLASSIC_KEY) || '') : ''; } catch (e) { had = ''; }
  // Refused (too large after all: the mailbox counts in its own way): tried again without the
  // largest image, a few times, before nothing is kept. writeRoaming puts the copy before back,
  // so the settings saved later (the user's choices, their variants) are not refused with it.
  for (let tries = 0; ; tries++) {
    if (had === JSON.stringify(copy) || await writeRoaming(CLASSIC_KEY, copy)) break;
    const keys = Object.keys(copy.images);
    if (tries >= 3 || !keys.length) {
      console.log('OSign: the mailbox did not keep the signatures for classic Outlook.');
      return { kept: false, images: 0, left: 0 };
    }
    const largest = keys.reduce((a, b) => (copy.images[b].b64.length > copy.images[a].b64.length ? b : a));
    delete copy.images[largest];
    left++;
  }
  // Of those left out (or made smaller) here, the ones classic Outlook loads as they are from next
  // to the add-in's files: nothing is missing there for them while it is online.
  const index = await hostedIndex();
  const hostedAll = wanted.filter((w) => w.ref.dir === 'assets' && hostedFile(index, w.ref.name)).length;
  const leftOut = wanted.filter((w) => !copy.images[w.key] && !(w.ref.dir === 'assets' && hostedFile(index, w.ref.name))).length;
  return { kept: true, images: Object.keys(copy.images).length, left: Math.min(left, leftOut), squeezed, hosted: hostedAll, signatures: copy.cfg.signatures.length };
}
// The banner campaigns not yet over, as far as the event handler needs them (the banner in the
// user's language only).
function classicCampaigns(cfg, values) {
  const day = core.today(), lang = String((values || {}).preferredLanguage || '').split(/[-_]/)[0].toLowerCase();
  return (cfg.campaigns || []).filter((c) => !c.validUntil || c.validUntil >= day).map((c) => {
    const o = { id: c.id, compiled: core.campaignHtml(c, values, 0) };
    ['validFrom', 'validUntil', 'signatures', 'position', 'newOnly', 'audience', 'days'].forEach((k) => { if (c[k] !== undefined) o[k] = c[k]; });
    // The banner in the user's language, if it has one; else the banners that take turns, each
    // with its HTML (with a language of its own, that one only).
    if (lang && c.langs && c.langs[lang]) o.langs = { [lang]: { compiled: c.langs[lang].compiled || c.langs[lang].html } };
    if (!o.langs && (c.more || []).length) o.more = c.more.map((x) => ({ compiled: x.compiled || x.html || '' }));
    return o;
  });
}
function classicCopy() { const c = readRoaming(CLASSIC_KEY); return c && c.cfg && Array.isArray(c.cfg.signatures) ? c : null; }

// ---------------------------------------------------------------------------------------------
// The item being written
// ---------------------------------------------------------------------------------------------
function call(fn, ...args) {
  return new Promise((resolve, reject) => {
    fn(...args, (res) => (res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(res.error || new Error('Outlook refused.'))));
  });
}

export function isMessage(item) { return item && item.itemType === Office.MailboxEnums.ItemType.Message; }

// 'new' or 'reply', and whether the body is HTML.
export async function composeInfo(item) {
  if (isMessage(item) && item.getComposeTypeAsync) {
    try {
      const v = await call(item.getComposeTypeAsync.bind(item));
      return { kind: v.composeType === 'newMail' ? 'new' : 'reply', html: v.coercionType !== Office.CoercionType.Text };
    } catch (e) { /* below */ }
  }
  let html = true;
  try { html = (await call(item.body.getTypeAsync.bind(item.body))) !== Office.CoercionType.Text; } catch (e) { /* keep html */ }
  return { kind: 'new', html };
}

export async function fromAddress(item) {
  try { if (item.from && item.from.getAsync) { const v = await call(item.from.getAsync.bind(item.from)); if (v && v.emailAddress) return v.emailAddress; } } catch (e) { /* below */ }
  try { if (item.organizer && item.organizer.getAsync) { const v = await call(item.organizer.getAsync.bind(item.organizer)); if (v && v.emailAddress) return v.emailAddress; } } catch (e) { /* below */ }
  return (Office.context.mailbox.userProfile || {}).emailAddress || '';
}

// A reply or forward also carries the images of the mail it quotes, and that may be signed by NOBS OSign
// too (osign-assets-logo.png). So this message's own images are named osign-<tag>-…, with a tag of
// its own kept with the message (sessionData, which the event handler and the pane share): only
// those are ever taken out, and none has the name, and so the cid, of an image in the quoted mail.
// Without sessionData there is no tag, and every osign- image goes, as before.
const TAG_KEY = 'osign.tag';
async function attachmentPrefix(item) {
  const sd = item.sessionData;
  if (!sd || !sd.getAsync || !sd.setAsync) return ATTACHMENT_PREFIX;
  try { const v = await call(sd.getAsync.bind(sd), TAG_KEY); if (/^[0-9a-f]{8}$/.test(v)) return ATTACHMENT_PREFIX + v + '-'; } catch (e) { /* none yet */ }
  const tag = (Math.random().toString(16).slice(2) + '00000000').slice(0, 8);
  try { await call(sd.setAsync.bind(sd), TAG_KEY, tag); } catch (e) { return ATTACHMENT_PREFIX; }
  return ATTACHMENT_PREFIX + tag + '-';
}

// The images an earlier signature attached go when another one is put in (a changed sender, a
// signature picked in the pane), so a mail never carries a logo it no longer shows.
async function removeOwnAttachments(item, prefix) {
  if (!item.getAttachmentsAsync) return;
  let list = [];
  try { list = await call(item.getAttachmentsAsync.bind(item)); } catch (e) { return; }
  for (const a of list || []) {
    if (a.isInline && String(a.name || '').indexOf(prefix) === 0) {
      try { await call(item.removeAttachmentAsync.bind(item), a.id); } catch (e) { /* leave it */ }
    }
  }
}

// Puts the signature where Outlook keeps one: a later call replaces it, nothing else in the body changes.
export async function insertSignature(item, cfg, sig, values, asHtml, opts) {
  if (cfg.options && cfg.options.disableOutlookSignature && isMessage(item) && item.disableClientSignatureAsync) {
    try { await call(item.disableClientSignatureAsync.bind(item)); } catch (e) { /* insert ours anyway */ }
  }
  const prefix = await attachmentPrefix(item);
  await removeOwnAttachments(item, prefix);
  let content = '';
  if (sig) {
    const html = renderHtml(sig, values);
    if (asHtml) {
      let images = {};
      try { images = await loadImages(core.localImages(html), !!(opts && opts.interactive)); } catch (e) { console.log('OSign: images: ' + e.message); }
      for (const key of Object.keys(images)) {
        images[key] = Object.assign({}, images[key], { file: prefix + images[key].file });
        try { await call(item.addFileAttachmentFromBase64Async.bind(item), images[key].b64, images[key].file, { isInline: true }); }
        catch (e) { delete images[key]; }
      }
      content = placeImages(html, images, 'cid');
    } else content = core.htmlToText(html);
  }
  await call(item.body.setSignatureAsync.bind(item.body), content, { coercionType: asHtml ? Office.CoercionType.Html : Office.CoercionType.Text });
}
