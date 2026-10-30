/*
 * The task pane: the signatures open to this user, each with a preview to insert, and - when the
 * editor allows it - which one is used for new mail and which for replies.
 */
import { core, loadConfig, loadUserValues, getPrefs, setPrefs, composeInfo, fromAddress, insertSignature, previewHtml } from './common.js';

/* global Office, OSIGN_VERSION */
const $ = (id) => document.getElementById(id);

// In Outlook's own display language: German, or English.
const DE = {
  'Signatures': 'Signaturen',
  'For new mail': 'Für neue Mails',
  'For replies and forwards': 'Für Antworten und Weiterleitungen',
  'Read my details again': 'Meine Angaben neu lesen',
  'Company default': 'Firmenstandard',
  '(none)': '(keine)',
  'No signature': 'Keine Signatur',
  'Saved. It applies to the next new message.': 'Gespeichert. Gilt ab der nächsten neuen Mail.',
  'Saved. It applies to the next reply.': 'Gespeichert. Gilt ab der nächsten Antwort.',
  'No signature is set up for you.': 'Für Sie ist keine Signatur eingerichtet.',
  'new mail': 'neue Mails',
  'replies': 'Antworten',
  'Insert': 'Einfügen',
  'Outlook did not take the signature: {0}': 'Outlook hat die Signatur nicht übernommen: {0}',
  'Your details could not be read.': 'Ihre Angaben konnten nicht gelesen werden.',
  'Your details could not be read: {0}': 'Ihre Angaben konnten nicht gelesen werden: {0}',
  'Showing your details as last read; reading them now failed: {0}': 'Angezeigt werden Ihre zuletzt gelesenen Angaben; sie jetzt zu lesen ist fehlgeschlagen: {0}',
  'Your groups could not be read, so signatures meant for groups are not shown.': 'Ihre Gruppen konnten nicht gelesen werden, darum fehlen die Signaturen für Gruppen.',
  'Your groups could not be read, so signatures meant for groups are not shown: {0}': 'Ihre Gruppen konnten nicht gelesen werden, darum fehlen die Signaturen für Gruppen: {0}',
  'details not available, fields stay empty': 'Angaben nicht verfügbar, Felder bleiben leer',
  'signatures published {0}': 'Signaturen veröffentlicht {0}'
};
let german = false;
function t(s, ...args) {
  const out = german && DE[s] ? DE[s] : s;
  return out.replace(/\{(\d+)\}/g, (all, i) => (args[i] !== undefined ? String(args[i]) : all));
}
function translatePage() {
  german = /^de/i.test((Office.context && Office.context.displayLanguage) || navigator.language || '');
  document.documentElement.lang = german ? 'de' : 'en';
  document.querySelectorAll('h1, label, #refresh').forEach((el) => {
    el.childNodes.forEach((n) => { if (n.nodeType === 3 && n.nodeValue.trim() && DE[n.nodeValue.trim()] && german) n.nodeValue = n.nodeValue.replace(n.nodeValue.trim(), DE[n.nodeValue.trim()]); });
  });
}
let cfg = null, user = null, from = '', info = { kind: 'new', html: true };

function message(text, isError) {
  const m = $('msg');
  m.hidden = !text; m.textContent = text || ''; m.className = 'msg' + (isError ? ' error' : '');
}

function option(value, text, selected) {
  const o = document.createElement('option'); o.value = value; o.textContent = text; o.selected = !!selected; return o;
}

function renderChoice(list) {
  if (!cfg.options || cfg.options.allowUserChoice === false) { $('choice').hidden = true; return; }
  const prefs = getPrefs();
  ['new', 'reply'].forEach((kind) => {
    const sel = $(kind === 'new' ? 'prefNew' : 'prefReply');
    const central = core.pick(cfg, user.values, from, kind, null, user.groups);
    sel.replaceChildren(
      option('', t('Company default') + (central ? ' (' + central.name + ')' : ' ' + t('(none)')), !prefs[kind]),
      option('none', t('No signature'), prefs[kind] === 'none'),
      ...list.map((s) => option(s.id, s.name, prefs[kind] === s.id)));
    sel.onchange = async () => {
      const p = getPrefs(); p[kind] = sel.value; await setPrefs(p);
      message(kind === 'new' ? t('Saved. It applies to the next new message.') : t('Saved. It applies to the next reply.'));
    };
  });
  $('choice').hidden = false;
}

function renderList(list) {
  const box = $('list'); box.textContent = '';
  if (!list.length) { box.textContent = t('No signature is set up for you.'); return; }
  list.forEach((s) => {
    const card = document.createElement('div'); card.className = 'card';
    const head = document.createElement('div'); head.className = 'card-head';
    const name = document.createElement('b'); name.textContent = s.name;
    const tags = document.createElement('small');
    tags.textContent = [s.defaultNew ? t('new mail') : '', s.defaultReply ? t('replies') : ''].filter(Boolean).join(', ');
    const left = document.createElement('div'); left.append(name, document.createElement('br'), tags);
    const btn = document.createElement('button'); btn.textContent = t('Insert');
    btn.onclick = async () => {
      btn.disabled = true;
      try { await insertSignature(Office.context.mailbox.item, cfg, s, user.values, info.html, { interactive: true }); message(''); }
      catch (e) { message(t('Outlook did not take the signature: {0}', e.message || e), true); }
      finally { btn.disabled = false; }
    };
    head.append(left, btn);
    const frame = document.createElement('iframe');
    frame.setAttribute('sandbox', '');
    frame.title = s.name;
    const page = (body) => '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:8px 10px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222;zoom:.85}p{margin:0}</style></head><body>' + body + '</body></html>';
    frame.srcdoc = page('');
    previewHtml(s, user.values).then((h) => { frame.srcdoc = page(h); });
    card.append(head, frame);
    box.append(card);
  });
}

async function load(refresh) {
  message('');
  try { cfg = await loadConfig({ interactive: true }); } catch (e) { message(e.message, true); showVersion(); return; }
  showVersion();
  const item = Office.context.mailbox.item;
  if (item) { info = await composeInfo(item); from = await fromAddress(item); }
  user = await loadUserValues(cfg, { interactive: true, refresh });
  if (!user.values) {
    message(user.error ? t('Your details could not be read: {0}', user.error.message) : t('Your details could not be read.'), true);
    return;
  }
  const up = Office.context.mailbox.userProfile || {};
  $('who').textContent = (up.displayName || '') + (from ? ' · ' + from : '') +
    (user.source === 'outlook' ? ' · ' + t('details not available, fields stay empty') : '');
  if (user.error && user.source === 'cache') message(t('Showing your details as last read; reading them now failed: {0}', user.error.message));
  const list = core.available(cfg, user.values, from, user.groups);
  if (user.groups === null && core.groupsUsed(cfg).length) message(user.error ? t('Your groups could not be read, so signatures meant for groups are not shown: {0}', user.error.message) : t('Your groups could not be read, so signatures meant for groups are not shown.'), true);
  renderChoice(list);
  renderList(list);
}

// NOBS OSign 0.9.14 · signatures published 30.10.2026, 14:02
function showVersion() {
  const at = cfg && cfg.updated ? new Date(cfg.updated) : null;
  const when = at && !isNaN(at) ? at.toLocaleString(german ? 'de-CH' : 'en-GB', { dateStyle: 'short', timeStyle: 'short' }) : '';
  $('ver').textContent = 'NOBS OSign ' + OSIGN_VERSION + (when ? ' · ' + t('signatures published {0}', when) : '');
}

Office.onReady(() => {
  translatePage();
  showVersion();
  $('refresh').addEventListener('click', (e) => { e.preventDefault(); load(true); });
  load(false);
});
