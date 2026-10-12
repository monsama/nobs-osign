/*
 * The task pane: the signatures open to this user, each with a preview to insert, and - when the
 * editor allows it - which one is used for new mail and which for replies.
 */
import { core, loadConfig, loadUserValues, getPrefs, setPrefs, composeInfo, fromAddress, insertSignature, renderFor } from './common.js';

/* global Office */
const $ = (id) => document.getElementById(id);
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
    const central = core.pick(cfg, user.values, from, kind, null);
    sel.replaceChildren(
      option('', 'Company default' + (central ? ' (' + central.name + ')' : ' (none)'), !prefs[kind]),
      option('none', 'No signature', prefs[kind] === 'none'),
      ...list.map((s) => option(s.id, s.name, prefs[kind] === s.id)));
    sel.onchange = async () => {
      const p = getPrefs(); p[kind] = sel.value; await setPrefs(p);
      message('Saved. It applies to the next ' + (kind === 'new' ? 'new message' : 'reply') + '.');
    };
  });
  $('choice').hidden = false;
}

function renderList(list) {
  const box = $('list'); box.textContent = '';
  if (!list.length) { box.textContent = 'No signature is set up for you.'; return; }
  list.forEach((s) => {
    const card = document.createElement('div'); card.className = 'card';
    const head = document.createElement('div'); head.className = 'card-head';
    const name = document.createElement('b'); name.textContent = s.name;
    const tags = document.createElement('small');
    tags.textContent = [s.defaultNew ? 'new mail' : '', s.defaultReply ? 'replies' : ''].filter(Boolean).join(', ');
    const left = document.createElement('div'); left.append(name, document.createElement('br'), tags);
    const btn = document.createElement('button'); btn.textContent = 'Insert';
    btn.onclick = async () => {
      btn.disabled = true;
      try { await insertSignature(Office.context.mailbox.item, cfg, s, user.values, info.html); message(''); }
      catch (e) { message('Outlook did not take the signature: ' + (e.message || e), true); }
      finally { btn.disabled = false; }
    };
    head.append(left, btn);
    const frame = document.createElement('iframe');
    frame.setAttribute('sandbox', '');
    frame.title = s.name;
    frame.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:8px 10px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222;zoom:.85}p{margin:0}</style></head><body>' + renderFor(cfg, s, user.values, true) + '</body></html>';
    card.append(head, frame);
    box.append(card);
  });
}

async function load(refresh) {
  message('');
  try { cfg = await loadConfig(); } catch (e) { message(e.message, true); return; }
  const item = Office.context.mailbox.item;
  if (item) { info = await composeInfo(item); from = await fromAddress(item); }
  user = await loadUserValues(cfg, { interactive: true, refresh });
  if (!user.values) {
    message('Your details could not be read' + (user.error ? ': ' + user.error.message : '.'), true);
    return;
  }
  const up = Office.context.mailbox.userProfile || {};
  $('who').textContent = (up.displayName || '') + (from ? ' · ' + from : '') +
    (user.source === 'outlook' ? ' · details not available, fields stay empty' : '');
  if (user.error && user.source === 'cache') message('Showing your details as last read; reading them now failed: ' + user.error.message);
  const list = core.available(cfg, user.values, from);
  renderChoice(list);
  renderList(list);
}

Office.onReady(() => {
  $('refresh').addEventListener('click', (e) => { e.preventDefault(); load(true); });
  load(false);
});
