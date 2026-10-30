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

function withTimeout(promise, ms, what) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error(what + ' took too long.')), ms);
    promise.then((v) => { clearTimeout(t); resolve(v); }, (e) => { clearTimeout(t); reject(e); });
  });
}

// ---------------------------------------------------------------------------------------------
// Storage between messages: localStorage where there is one, OfficeRuntime.storage in the
// JavaScript-only runtime, memory otherwise.
// ---------------------------------------------------------------------------------------------
const memory = {};
const store = {
  async get(key) {
    try { if (typeof localStorage !== 'undefined') return localStorage.getItem(key); } catch (e) { /* none */ }
    try { if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) return await OfficeRuntime.storage.getItem(key); } catch (e) { /* none */ }
    return memory[key] || null;
  },
  async set(key, value) {
    memory[key] = value;
    try { if (typeof localStorage !== 'undefined') { localStorage.setItem(key, value); return; } } catch (e) { /* none */ }
    try { if (typeof OfficeRuntime !== 'undefined' && OfficeRuntime.storage) await OfficeRuntime.storage.setItem(key, value); } catch (e) { /* none */ }
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
  'Loading your entry': 'Ihren Eintrag laden', 'Listing the images': 'Bilder auflisten', 'Finding the photo': 'Foto suchen'
};
const MSG_DE = [
  [/SharePoint does not let you read the folder (.+?) \((\d+)\)\. It has to be shared with everyone who gets a signature: .*$/, 'SharePoint lässt Sie den Ordner $1 nicht lesen ($2). Er muss für alle freigegeben sein, die eine Signatur bekommen: in SharePoint beim Ordner → Zugriff verwalten → Zugriff gewähren: «Jeder außer externen Benutzern», Anzeigen.'],
  [/not allowed \((\d+)\)\. The app registration needs admin consent for (.+?) \(.*$/, 'nicht erlaubt ($1). Der App-Registrierung fehlt die Administratorzustimmung für $2 (Entra ID → Unternehmensanwendungen → die App des Add-ins → Berechtigungen; das Setup-Skript erteilt sie).'],
  [/You are not in the AD export, so your groups are not known\.$/, 'Sie sind nicht im AD-Export, darum sind Ihre Gruppen nicht bekannt.'],
  [/You are not in the AD export \((.+?)\)\. Ask IT to export the users again\.$/, 'Sie sind nicht im AD-Export ($1). Bitten Sie die IT, die Benutzer neu zu exportieren.'],
  [/No signatures are published yet: (.+) is not in SharePoint\.$/, 'Noch keine Signaturen veröffentlicht: $1 fehlt in SharePoint.'],
  [/ took too long\.$/, ' dauerte zu lange.'],
  [/: not found in SharePoint \(404\)\.$/, ': in SharePoint nicht gefunden (404).'],
  [/: (Microsoft Graph|SharePoint) answered (\d+)\.$/, ': $1 antwortete $2.']
];
export function localize(message, german) {
  let m = String(message || '');
  if (!german) return m;
  for (const [en, de] of Object.entries(WHAT_DE)) m = m.split(en + ':').join(de + ':').split(en + ' took').join(de + ' took');
  for (const [rx, de] of MSG_DE) m = m.replace(rx, de);
  return m;
}

// The permissions a token carries (its scp claim). Decoded by hand: classic Outlook's runtime has
// no atob.
const B64URL = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
function scopesOf(token) {
  try {
    const part = String(token).split('.')[1].replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
    let bits = 0, n = 0, bytes = [];
    for (const ch of part) { n = (n << 6) | B64URL.indexOf(ch); bits += 6; if (bits >= 8) { bits -= 8; bytes.push((n >> bits) & 255); } }
    const json = decodeURIComponent(bytes.map((b) => '%' + ('0' + b.toString(16)).slice(-2)).join(''));
    return String(JSON.parse(json).scp || '').split(' ').filter(Boolean);
  } catch (e) { return null; }
}

async function graphGet(token, url, what, allowNotFound) {
  const r = await withTimeout(fetch(url, { headers: { Authorization: 'Bearer ' + token } }), 8000, what);
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
let tokens = {};
async function token(scopes, interactive) {
  const key = scopes.join(' ');
  const had = tokens[key];
  if (!had || interactive || Date.now() - had.at > TOKEN_REUSE_MS) {
    const entry = { at: Date.now() };
    entry.p = getGraphToken(interactive, scopes).catch((e) => { if (tokens[key] === entry) delete tokens[key]; throw e; });
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
  const m = /^([A-Za-z0-9.-]+):\/(sites|teams|personal)\/([^:/]+):$/.exec(SITE);
  if (!m) throw new Error('The add-in was made for the SharePoint site "' + SITE + '": make its files again in NOBS OSign.');
  return GRAPH + '/sites/' + m[1] + ':/' + m[2] + '/' + encodeURIComponent(m[3]) + ':';
}
async function siteBase(tok) {
  if (!SITE || SITE === 'root') return GRAPH + '/sites/root';
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

async function download(meta, tok, what, asText) {
  // The address Graph hands out is signed for a short while: no token is sent there.
  const url = meta['@microsoft.graph.downloadUrl'];
  const r = await withTimeout(url ? fetch(url) : fetch((await itemUrl(tok, meta.__rel)) + ':/content', { headers: { Authorization: 'Bearer ' + tok } }), 10000, what);
  if (!r.ok) throw new Error(what + ': SharePoint answered ' + r.status + '.');
  return asText ? r.text() : new Uint8Array(await r.arrayBuffer());
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
    return cfg;
  } catch (e) {
    if (cached && cached.text) { try { return JSON.parse(cached.text); } catch (e2) { /* below */ } }
    throw e;
  }
}

// ---------------------------------------------------------------------------------------------
// Roaming settings: kept in the mailbox, so they follow the user to every Outlook.
// ---------------------------------------------------------------------------------------------
function roaming() { return Office.context && Office.context.roamingSettings; }
function readRoaming(key) { try { const v = roaming() && roaming().get(key); return v ? JSON.parse(v) : null; } catch (e) { return null; } }
function writeRoaming(key, value) {
  return new Promise((resolve) => {
    const rs = roaming(); if (!rs) { resolve(); return; }
    try { rs.set(key, JSON.stringify(value)); rs.saveAsync(() => resolve()); } catch (e) { resolve(); }
  });
}

export function getPrefs() { return readRoaming(PREFS_KEY) || {}; }
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
    await writeRoaming(PROFILE_KEY, { at: Date.now(), values: r.values, groups: r.groups, groupsFor: r.groups ? key : '\u0000' });
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

// { 'assets/logo.png': { b64, type, file } } for what the signature shows; what is not there is
// left out of the map (and so out of the signature).
async function loadImages(refs, interactive) {
  const out = {};
  const wanted = [];
  refs.forEach((r) => { const k = r.dir + '/' + r.name; if (typeOf(r.name) && !wanted.some((w) => w.key === k)) wanted.push({ key: k, dir: r.dir, name: r.name }); });
  if (!wanted.length) return out;
  const tok = await token(BASE_SCOPES, interactive);
  // The logos and icons: one listing tells which are there and whether they changed.
  let assets = null;
  if (wanted.some((w) => w.dir === 'assets')) {
    const j = await graphGet(tok, (await itemUrl(tok, 'assets')) + ':/children?$top=999', 'Listing the images', true);
    assets = {};
    ((j && j.value) || []).forEach((it) => { assets[String(it.name).toLowerCase()] = it; });
  }
  for (const w of wanted) {
    try {
      const meta = w.dir === 'assets' ? assets[w.name.toLowerCase()] : await graphGet(tok, await itemUrl(tok, 'photos/' + w.name), 'Finding the photo', true);
      if (!meta) continue;
      meta.__rel = w.key;
      const ck = 'osign.img.' + w.key.toLowerCase();
      const cached = await readJson(ck);
      let b64 = cached && cached.etag && cached.etag === meta.eTag ? cached.b64 : null;
      if (!b64) {
        b64 = base64(await download(meta, tok, 'Loading ' + w.name));
        if (b64.length < IMAGE_CACHE_MAX * 4 / 3) await store.set(ck, JSON.stringify({ etag: meta.eTag || '', b64 }));
      }
      // The attachment's name, without its osign-<tag>- prefix (insertSignature puts that in front).
      out[w.key] = { b64, type: typeOf(w.name), file: w.dir + '-' + w.name.replace(/[^A-Za-z0-9._-]+/g, '-') };
    } catch (e) {
      console.log('OSign: image ' + w.key + ': ' + (e && e.message ? e.message : e));
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
  return core.render(sig.compiled || sig.html, values, { removeEmptyLines: sig.removeEmptyLines !== false });
}

// For the pane: the signature as the mail will show it, images inside.
export async function previewHtml(sig, values) {
  const html = renderHtml(sig, values);
  let images = {};
  try { images = await loadImages(core.localImages(html), false); } catch (e) { /* shown without */ }
  return placeImages(html, images, 'data');
}

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
