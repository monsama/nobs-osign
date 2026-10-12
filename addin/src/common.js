/*
 * What the event handler and the task pane share: the published configuration, the user's
 * details, the user's own choice, and putting a signature into the item being written.
 *
 * Runs in two places. In new Outlook, Outlook on the web and Outlook for Mac, in a browser
 * runtime. In classic Outlook on Windows, the event handler runs in a JavaScript-only runtime:
 * no DOM, no localStorage, no window - so nothing here may need them.
 */
import core from '../../shared/signature-core.js';
import { getGraphToken } from './auth.js';

/* global Office, OfficeRuntime, __BASE_URL__, __CONFIG_PATH__ */
export const BASE_URL = __BASE_URL__;
export const CONFIG_URL = BASE_URL + __CONFIG_PATH__ + '/signatures.json';
export { core };

const PROFILE_KEY = 'osign.profile';
const PREFS_KEY = 'osign.prefs';
const CONFIG_KEY = 'osign.config';
// The user's details are read again after this long; within it the copy is used as it is.
const PROFILE_FRESH_MS = 15 * 60 * 1000;

function withTimeout(promise, ms, what) {
  return new Promise((resolve, reject) => {
    const t = setTimeout(() => reject(new Error(what + ' took too long.')), ms);
    promise.then((v) => { clearTimeout(t); resolve(v); }, (e) => { clearTimeout(t); reject(e); });
  });
}

// ---------------------------------------------------------------------------------------------
// Storage for the last good configuration: localStorage where there is one, OfficeRuntime.storage
// in the JavaScript-only runtime, memory otherwise.
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

// ---------------------------------------------------------------------------------------------
// The configuration the editor published
// ---------------------------------------------------------------------------------------------
export async function loadConfig() {
  try {
    // A new address every minute: neither GitHub Pages' cache nor the browser's keeps an old one longer.
    const r = await withTimeout(fetch(CONFIG_URL + '?m=' + Math.floor(Date.now() / 60000), { cache: 'no-store' }), 6000, 'Loading the signatures');
    if (!r.ok) throw new Error('The signatures could not be loaded (' + r.status + ').');
    const text = await r.text();
    const cfg = JSON.parse(text);
    const errs = core.validate(cfg);
    if (errs.length) throw new Error('The published signatures are not valid: ' + errs[0]);
    store.set(CONFIG_KEY, text);
    return cfg;
  } catch (e) {
    const cached = await store.get(CONFIG_KEY);
    if (cached) { try { return JSON.parse(cached); } catch (e2) { /* fall through */ } }
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
// The user's details, from Microsoft Graph (/me)
// ---------------------------------------------------------------------------------------------
// Groups need GroupMember.Read.All (admin consent); only asked for when a signature is limited to groups.
const GROUP_SCOPE = 'GroupMember.Read.All';

async function graphGet(token, url, what) {
  const r = await withTimeout(fetch(url, { headers: { Authorization: 'Bearer ' + token } }), 8000, what);
  if (r.status === 403) throw new Error(what + ': not allowed (403). The app registration needs admin consent for ' + (/memberOf/i.test(url) ? GROUP_SCOPE : 'User.Read') + '.');
  if (!r.ok) throw new Error(what + ': Microsoft Graph answered ' + r.status + '.');
  return r.json();
}

// The user's groups, nested ones included, as the names the configuration uses: only those are
// kept (a user can be in hundreds of groups, and the copy goes into the mailbox settings).
async function readGroups(token, wanted) {
  const found = [];
  let url = 'https://graph.microsoft.com/v1.0/me/transitiveMemberOf/microsoft.graph.group?$select=id,displayName,onPremisesSamAccountName&$top=999';
  for (let page = 0; url && page < 20; page++) {
    const j = await graphGet(token, url, 'Reading your groups');
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

async function readGraph(cfg, interactive, wanted) {
  const scopes = wanted.length ? ['User.Read', GROUP_SCOPE] : ['User.Read'];
  const token = await getGraphToken(cfg, interactive, scopes);
  const select = core.graphSelect(cfg.fields);
  const me = await graphGet(token, 'https://graph.microsoft.com/v1.0/me' + (select.length ? '?$select=' + encodeURIComponent(select.join(',')) : ''), 'Reading your details');
  const groups = wanted.length ? await readGroups(token, wanted) : [];
  return { values: core.valuesFromGraph(cfg.fields, me), groups };
}

// Returns { values, groups, source: 'graph'|'cache'|'outlook', error } - or values null when there
// are none. groups is null when they could not be read: no signature limited to groups applies then.
// interactive: the task pane may show a sign-in window; the event handler never can.
export async function loadUserValues(cfg, opts) {
  opts = opts || {};
  const wanted = core.groupsUsed(cfg);
  const key = wanted.slice().sort().join('\n');
  const cached = readRoaming(PROFILE_KEY);
  // A copy read for other groups than the configuration now names says nothing about the new ones.
  const cacheFits = cached && cached.values && (cached.groupsFor || '') === key;
  if (!opts.refresh && cacheFits && Date.now() - cached.at < PROFILE_FRESH_MS) return { values: cached.values, groups: cached.groups || [], source: 'cache' };
  let error = null;
  try {
    const r = await readGraph(cfg, !!opts.interactive, wanted);
    await writeRoaming(PROFILE_KEY, { at: Date.now(), values: r.values, groups: r.groups, groupsFor: key });
    return { values: r.values, groups: r.groups, source: 'graph' };
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

export function renderFor(cfg, sig, values, asHtml) {
  const html = core.render(sig.compiled || sig.html, values, { assetBase: cfg.assetBaseUrl || BASE_URL + __CONFIG_PATH__ + '/', removeEmptyLines: sig.removeEmptyLines !== false });
  return asHtml ? html : core.htmlToText(html);
}

// Puts the signature where Outlook keeps one: a later call replaces it, nothing else in the body changes.
export async function insertSignature(item, cfg, sig, values, asHtml) {
  if (cfg.options && cfg.options.disableOutlookSignature && isMessage(item) && item.disableClientSignatureAsync) {
    try { await call(item.disableClientSignatureAsync.bind(item)); } catch (e) { /* insert ours anyway */ }
  }
  const content = sig ? renderFor(cfg, sig, values, asHtml) : '';
  await call(item.body.setSignatureAsync.bind(item.body), content, { coercionType: asHtml ? Office.CoercionType.Html : Office.CoercionType.Text });
}
