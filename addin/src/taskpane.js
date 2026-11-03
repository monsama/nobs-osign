/*
 * The task pane: the signatures open to this user, each with a preview to insert, and - when the
 * editor allows it - which one is used for new mail and which for replies.
 */
import { LANG } from './lang.js';
import { core, localize, langOf, loadConfig, loadUserValues, getPrefs, setPrefs, getVariants, setVariants, isClassicOutlook, keepForClassic, composeInfo, fromAddress, insertSignature, previewHtml, editableHtml } from './common.js';

/* global Office, OSIGN_VERSION */
const $ = (id) => document.getElementById(id);

// In Outlook's own display language: German here, French and Italian in lang.js, else English.
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
  'Inserted: {0}': 'Eingefügt: {0}',
  'Outlook did not take the signature: {0}': 'Outlook hat die Signatur nicht übernommen: {0}',
  'Your details could not be read.': 'Ihre Angaben konnten nicht gelesen werden.',
  'Your details could not be read: {0}': 'Ihre Angaben konnten nicht gelesen werden: {0}',
  'Showing your details as last read; reading them now failed: {0}': 'Angezeigt werden Ihre zuletzt gelesenen Angaben; sie jetzt zu lesen ist fehlgeschlagen: {0}',
  'Your groups could not be read, so signatures meant for groups are not shown.': 'Ihre Gruppen konnten nicht gelesen werden, darum fehlen die Signaturen für Gruppen.',
  'Your groups could not be read, so signatures meant for groups are not shown: {0}': 'Ihre Gruppen konnten nicht gelesen werden, darum fehlen die Signaturen für Gruppen: {0}',
  'details not available, fields stay empty': 'Angaben nicht verfügbar, Felder bleiben leer',
  'signatures published {0}': 'Signaturen veröffentlicht {0}',
  'Own variant': 'Eigene Variante',
  'Yours to change: {0}': 'Änderbar: {0}',
  'Yours to change: {0}: in the boxes below, or right in the signature underneath (marked). The rest stays the company signature.': 'Änderbar: {0} - in den Feldern unten oder direkt in der Signatur darunter (markiert). Alles andere bleibt die Firmensignatur.',
  'Your signature from {0}': 'Ihre Signatur aus {0}',
  'Your own signature from {0}': 'Ihre eigene Signatur aus {0}',
  "Change it as you like right here: click into it and type. Ctrl+B, Ctrl+I and Ctrl+U format; the company's images stay.": 'Ändern Sie sie direkt hier, wie Sie möchten: hineinklicken und schreiben. Strg+B, Strg+I und Strg+U formatieren; die Bilder der Firma bleiben.',
  'Your signature': 'Ihre Signatur',
  'Loading…': 'Wird geladen…',
  'Company signature again': 'Wieder die Firmensignatur',
  'Bold (Ctrl+B)': 'Fett (Strg+B)',
  'Italic (Ctrl+I)': 'Kursiv (Strg+I)',
  'Underline (Ctrl+U)': 'Unterstrichen (Strg+U)',
  'Your signature is too long to be kept in your mailbox: shorten it.': 'Ihre Signatur ist zu lang, um im Postfach gespeichert zu werden: Kürzen Sie sie.',
  'Your signature is saved. Pick it above for new mail or replies, or insert it here.': 'Ihre Signatur ist gespeichert. Wählen Sie sie oben für neue Mails oder Antworten, oder fügen Sie sie hier ein.',
  'Your own signature, from {0}': 'Ihre eigene Signatur, aus {0}',
  'Your variant of {0}': 'Ihre Variante von {0}',
  'Edit': 'Bearbeiten',
  'Delete': 'Löschen',
  'Save': 'Speichern',
  'Cancel': 'Abbrechen',
  'Name': 'Name',
  'New variant of {0}': 'Neue Variante von {0}',
  'Variant of {0}': 'Variante von {0}',
  'Give the variant a name.': 'Geben Sie der Variante einen Namen.',
  'Variant saved. Pick it above for new mail or replies, or insert it here.': 'Variante gespeichert. Wählen Sie sie oben für neue Mails oder Antworten, oder fügen Sie sie hier ein.',
  'Really delete?': 'Wirklich löschen?',
  'Classic Outlook: these signatures are too large to be kept in your mailbox, so they are not put in by themselves there. Use Insert.': 'Klassisches Outlook: Diese Signaturen sind zu gross, um im Postfach abgelegt zu werden; sie werden dort darum nicht selbst eingefügt. Nutzen Sie Einfügen.',
  'Classic Outlook: {0} image(s) did not fit into your mailbox settings and are left out when the signature goes in by itself there. Insert puts in everything.': 'Klassisches Outlook: {0} Bild(er) passen nicht in die Postfach-Einstellungen und fehlen, wenn die Signatur dort selbst eingefügt wird. Einfügen fügt alles ein.',
  'Outlook did not keep the variant: the room for add-in settings in your mailbox is full. Shorten or delete other variants.': 'Outlook hat die Variante nicht gespeichert: Der Platz für Add-in-Einstellungen in Ihrem Postfach ist voll. Kürzen oder löschen Sie andere Varianten.',
  'You have {0} variants, the most there can be.': 'Sie haben {0} Varianten, mehr sind nicht möglich.',
  'Only the fields below are yours to change; the rest stays the company signature.': 'Nur die Felder unten können Sie ändern; der Rest bleibt die Firmensignatur.',
  'How it works': 'So funktioniert es',
  'Your signature goes into new mail and replies by itself, with your details from the company directory. If they are wrong, IT changes them there; "Read my details again" fetches them anew.': 'Ihre Signatur kommt von selbst in neue Mails und Antworten, mit Ihren Angaben aus dem Firmenverzeichnis. Stimmen sie nicht, ändert die IT sie dort; «Meine Angaben neu lesen» holt sie neu.',
  'Above, choose which signature goes into new mail and which into replies. "Company default" follows what the company sets.': 'Oben wählen Sie, welche Signatur in neue Mails und welche in Antworten kommt. «Firmenstandard» folgt dem, was die Firma festlegt.',
  'Insert puts a signature into the message open now, e.g. another one than usual.': 'Einfügen setzt eine Signatur in die gerade offene Mail, z. B. eine andere als sonst.',
  'Own variant: a copy of a company signature that is yours to change, all of it. It is kept in your mailbox, also on your other devices with Outlook.': 'Eigene Variante: eine Kopie einer Firmensignatur, die Sie ganz ändern können. Sie wird in Ihrem Postfach gespeichert, auch für Ihre anderen Geräte mit Outlook.',
  'Own variant: a copy of a company signature in which you change the marked fields (e.g. a greeting); the rest stays the company signature. It is kept in your mailbox.': 'Eigene Variante: eine Kopie einer Firmensignatur, in der Sie die markierten Felder ändern (z. B. einen Gruss); der Rest bleibt die Firmensignatur. Sie wird in Ihrem Postfach gespeichert.',
  'After saving, choose your variant above for new mail or replies.': 'Nach dem Speichern wählen Sie Ihre Variante oben für neue Mails oder Antworten.'
};
let lang = 'en';
const dict = () => (lang === 'de' ? DE : LANG[lang] ? LANG[lang].ui : null);
function t(s, ...args) {
  const d = dict();
  const out = d && d[s] ? d[s] : s;
  return out.replace(/\{(\d+)\}/g, (all, i) => (args[i] !== undefined ? String(args[i]) : all));
}
function translatePage() {
  lang = langOf((Office.context && Office.context.displayLanguage) || navigator.language || '');
  document.documentElement.lang = lang;
  document.querySelectorAll('h1 .sub, label, #refresh, #help summary').forEach((el) => {
    el.childNodes.forEach((n) => { const d = dict(), k = n.nodeType === 3 && n.nodeValue.trim(); if (k && d && d[k]) n.nodeValue = n.nodeValue.replace(k, d[k]); });
  });
}
let cfg = null, full = null, user = null, from = '', info = { kind: 'new', html: true };

function message(text, isError) {
  const m = $('msg');
  m.hidden = !text; m.textContent = localize(text || '', lang); m.className = 'msg' + (isError ? ' error' : '');
}

function option(value, text, selected) {
  const o = document.createElement('option'); o.value = value; o.textContent = text; o.selected = !!selected; return o;
}

function renderChoice(list) {
  if (!cfg.options || cfg.options.allowUserChoice === false) { $('choice').hidden = true; return; }
  const prefs = getPrefs();
  ['new', 'reply'].forEach((kind) => {
    const sel = $(kind === 'new' ? 'prefNew' : 'prefReply');
    const central = core.pick(full, user.values, from, kind, null, user.groups);
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

const PAGE = (body) => '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:8px 10px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222;zoom:.85}p{margin:0}a{pointer-events:none;cursor:default}</style></head><body>' + body + '</body></html>';
function previewFrame(sig, title) {
  const frame = document.createElement('iframe');
  frame.setAttribute('sandbox', '');
  frame.title = title;
  frame.srcdoc = PAGE('');
  // mark: the field keys a variant may change, shown marked in the preview (where they are, even
  // when empty: their name then).
  frame.show = (s, mark) => {
    let shown = s;
    if (mark && mark.length) {
      const vals = Object.assign({}, user.values, s.userValues || {}), uv = {};
      mark.forEach((k) => { uv[k] = '\u2063' + (vals[k] || '') + '\u2064' + k + '\u2065'; });
      shown = Object.assign({}, s, { userValues: Object.assign({}, s.userValues || {}, uv) });
    }
    return previewHtml(shown, user.values).then((h) => {
      if (mark && mark.length) {
        const labels = {}; (cfg.fields || []).forEach((f) => { labels[f.key] = f.label || f.key; });
        // In a tag (href="mailto:@@mail"): the value only; in the text: marked.
        h = h.replace(/<[^>]*>/g, (tag) => tag.replace(/\u2063([^\u2064]*)\u2064[^\u2065]*\u2065/g, '$1'));
        h = h.replace(/\u2063([^\u2064]*)\u2064([^\u2065]*)\u2065/g, (all, v, k) =>
          '<span style="background:#fff3bf;outline:1px dashed #d4a017;border-radius:2px;padding:0 1px" title="' + (labels[k] || k).replace(/"/g, '&quot;') + '">' + (v || '<i style="color:#8a6d00">' + (labels[k] || k).replace(/</g, '&lt;') + '</i>') + '</span>');
      }
      frame.srcdoc = PAGE(h);
    });
  };
  frame.show(sig);
  return frame;
}
function button(text, onclick, secondary) {
  const b = document.createElement('button'); b.type = 'button'; b.textContent = text; b.onclick = onclick;
  if (secondary) b.className = 'secondary';
  return b;
}

function renderList(list) {
  const box = $('list'); box.textContent = '';
  if (!list.length) { box.textContent = t('No signature is set up for you.'); return; }
  const names = {}; (cfg.signatures || []).forEach((s) => { names[s.id] = s.name; });
  list.forEach((s) => {
    const card = document.createElement('div'); card.className = 'card';
    const head = document.createElement('div'); head.className = 'card-head';
    const name = document.createElement('b'); name.textContent = s.name;
    const tags = document.createElement('small');
    tags.textContent = s.own ? t('Your own signature, from {0}', names[s.variantOf] || '') : s.variantOf ? t('Your variant of {0}', names[s.variantOf] || '')
      : [s.defaultNew ? t('new mail') : '', s.defaultReply ? t('replies') : ''].filter(Boolean).join(', ');
    const left = document.createElement('div'); left.append(name, document.createElement('br'), tags);
    const btn = button(t('Insert'), async () => {
      btn.disabled = true;
      try { await insertSignature(Office.context.mailbox.item, cfg, s, user.values, info.html, { interactive: true }); message(t('Inserted: {0}', s.name)); }
      catch (e) { message(t('Outlook did not take the signature: {0}', e.message || e), true); }
      finally { btn.disabled = false; }
    });
    const actions = document.createElement('div'); actions.className = 'actions';
    if (s.variantOf) {
      actions.append(button(t('Edit'), () => editVariant(card, s, getVariants().find((v) => v.id === s.id)), true),
        button(t('Delete'), (e) => deleteVariant(s, e.currentTarget), true));
    } else if (core.variantsAllowed(cfg) && (core.editableIn(cfg, s).length || core.ownAllowed(cfg))) {
      // Which fields a variant of it may change (all of it, when own signatures are allowed).
      if (!core.ownAllowed(cfg)) {
        const lab = {}; (cfg.fields || []).forEach((f) => { lab[f.key] = f.label || f.key; });
        const ed = document.createElement('small'); ed.className = 'editable';
        ed.textContent = t('Yours to change: {0}', core.editableIn(cfg, s).map((k) => lab[k] || k).join(', '));
        left.append(document.createElement('br'), ed);
      }
      actions.append(button(t('Own variant'), () => editVariant(card, s, null), true));
    }
    actions.append(btn);
    head.append(left, actions);
    card.append(head, previewFrame(s, s.name));
    box.append(card);
  });
}

// What is kept of a signature of one's own: its markup, with nothing that runs or loads from
// elsewhere; the company's images by their place (assets/, photos/), any other image left out.
function ownHtml(box) {
  const c = box.cloneNode(true);
  c.querySelectorAll('script,style,iframe,object,embed,form,input,button,select,textarea,link,meta,base,svg,video,audio').forEach((n) => n.remove());
  c.querySelectorAll('*').forEach((n) => {
    [...n.attributes].forEach((a) => {
      if (/^on/i.test(a.name) || a.name === 'contenteditable') n.removeAttribute(a.name);
      if (/^(href|src)$/i.test(a.name) && /^\s*(javascript|vbscript):/i.test(a.value)) n.removeAttribute(a.name);
    });
  });
  c.querySelectorAll('img').forEach((im) => {
    const from = im.getAttribute('data-osign-src');
    if (from && /^(assets|photos)\//.test(from)) { im.setAttribute('src', from); im.removeAttribute('data-osign-src'); }
    else if (!/^https:\/\//i.test(im.getAttribute('src') || '')) im.remove();
  });
  return c.innerHTML.trim();
}
// A signature of one's own: all of the company signature to change, filled in with the user's
// details (and the company's images), kept in the mailbox.
async function editOwn(card, base, existing) {
  const form = document.createElement('form'); form.className = 'variant';
  const title = document.createElement('b'); title.textContent = existing ? t('Your signature from {0}', base.name) : t('Your own signature from {0}', base.name);
  const hint = document.createElement('small'); hint.textContent = t('Change it as you like right here: click into it and type. Ctrl+B, Ctrl+I and Ctrl+U format; the company\'s images stay.');
  const l = document.createElement('label'); l.append(t('Name'));
  const nameIn = document.createElement('input'); nameIn.maxLength = core.VARIANT_NAME_MAX;
  nameIn.value = existing ? existing.name : base.name + ' - ' + t('Own variant');
  l.append(nameIn);
  const box = document.createElement('div'); box.className = 'own-edit'; box.contentEditable = 'true'; box.spellcheck = true;
  box.setAttribute('role', 'textbox'); box.setAttribute('aria-multiline', 'true'); box.setAttribute('aria-label', t('Your signature'));
  box.textContent = t('Loading…');
  const tools = document.createElement('div'); tools.className = 'own-tools';
  const tool = (label, title2, cmd) => { const b = button(label, () => { box.focus(); document.execCommand(cmd); }, true); b.title = title2; b.addEventListener('mousedown', (e) => e.preventDefault()); return b; };
  const reset = button(t('Company signature again'), async () => { box.innerHTML = await editableHtml(base, user.values); }, true);
  tools.append(tool('B', t('Bold (Ctrl+B)'), 'bold'), tool('I', t('Italic (Ctrl+I)'), 'italic'), tool('U', t('Underline (Ctrl+U)'), 'underline'), reset);
  const row = document.createElement('div'); row.className = 'actions';
  row.append(button(t('Cancel'), () => render(), true), button(t('Save'), () => form.requestSubmit()));
  form.append(title, hint, l, tools, box, row);
  form.onsubmit = async (e) => {
    e.preventDefault();
    const name = nameIn.value.trim();
    if (!name) { message(t('Give the variant a name.'), true); nameIn.focus(); return; }
    const own = core.compile(ownHtml(box), document);
    if (own.length > core.VARIANT_OWN_MAX) { message(t('Your signature is too long to be kept in your mailbox: shorten it.'), true); return; }
    const v = { id: existing ? existing.id : 'my-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), name, base: base.id, own };
    const before = getVariants();
    const list = before.filter((x) => x.id !== v.id); list.push(v);
    if (!(await setVariants(list))) {
      await setVariants(before);
      message(t('Outlook did not keep the variant: the room for add-in settings in your mailbox is full. Shorten or delete other variants.'), true);
      return;
    }
    message(t('Your signature is saved. Pick it above for new mail or replies, or insert it here.'));
    render();
  };
  card.replaceChildren(form);
  form.scrollIntoView({ block: 'nearest' });
  const ownSig = existing ? (full.signatures || []).find((x) => x.id === existing.id) : null;
  box.innerHTML = await editableHtml(ownSig || base, user.values);
  box.focus();
}

// A variant: its name and the editable fields its signature uses, with the preview as it changes.
function editVariant(card, sig, existing) {
  // All of it to change, when the editor allows that.
  if (core.ownAllowed(cfg) && (!existing || typeof existing.own === 'string')) {
    const base0 = (cfg.signatures || []).find((x) => x.id === (sig.variantOf || sig.id));
    if (base0) { editOwn(card, base0, existing).catch((e) => message(e.message || String(e), true)); return; }
  }
  const base = (cfg.signatures || []).find((x) => x.id === (sig.variantOf || sig.id));
  if (!base) return;
  if (!existing && getVariants().length >= core.VARIANT_MAX) { message(t('You have {0} variants, the most there can be.', core.VARIANT_MAX), true); return; }
  const labels = {}; (cfg.fields || []).forEach((f) => { labels[f.key] = f.label || f.key; });
  const form = document.createElement('form'); form.className = 'variant';
  const title = document.createElement('b'); title.textContent = existing ? t('Variant of {0}', base.name) : t('New variant of {0}', base.name);
  const hint = document.createElement('small'); hint.textContent = t('Yours to change: {0}: in the boxes below, or right in the signature underneath (marked). The rest stays the company signature.', core.editableIn(cfg, base).map((k) => labels[k] || k).join(', '));
  const field = (label, value, max) => {
    const l = document.createElement('label'); l.append(label);
    const i = document.createElement('input'); i.value = value || ''; i.maxLength = max;
    l.append(i); form.append(l); return i;
  };
  form.append(title, hint);
  const nameIn = field(t('Name'), existing ? existing.name : base.name + ' - ' + t('Own variant'), core.VARIANT_NAME_MAX);
  const inputs = {};
  core.editableIn(cfg, base).forEach((k) => {
    const had = existing && existing.values && typeof existing.values[k] === 'string' ? existing.values[k] : (user.values[k] || '');
    inputs[k] = field(labels[k], had, core.VARIANT_VALUE_MAX);
  });
  const draft = () => {
    const values = {}; Object.keys(inputs).forEach((k) => { values[k] = inputs[k].value; });
    return { id: existing ? existing.id : 'my-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6), name: nameIn.value.trim(), base: base.id, values };
  };
  // The signature as it will look, the fields that are the user's marked and written in right
  // there (or in the boxes above: the two stay the same).
  const marked = core.editableIn(cfg, base);
  const live = document.createElement('div'); live.className = 'own-edit live';
  const spans = (k) => [...live.querySelectorAll('.fld')].filter((x) => x.dataset.k === k);
  (async () => {
    const s0 = core.withVariants(cfg, [draft()]).signatures.slice(-1)[0] || base;
    const vals = Object.assign({}, user.values, s0.userValues || {}), uv = {};
    marked.forEach((k) => { uv[k] = '\u2063' + (vals[k] || '') + '\u2064' + k + '\u2065'; });
    let h = await editableHtml(Object.assign({}, s0, { userValues: Object.assign({}, s0.userValues || {}, uv) }), user.values);
    h = h.replace(/<[^>]*>/g, (tag) => tag.replace(/\u2063([^\u2064]*)\u2064[^\u2065]*\u2065/g, '$1'));
    h = h.replace(/\u2063([^\u2064]*)\u2064([^\u2065]*)\u2065/g, (all, v, k) => '<span class="fld" data-k="' + k + '">' + v + '</span>');
    live.innerHTML = h;
    live.querySelectorAll('a[href]').forEach((a) => a.removeAttribute('href'));
    live.querySelectorAll('.fld').forEach((sp) => {
      const k = sp.dataset.k;
      sp.contentEditable = 'plaintext-only';
      if (sp.contentEditable !== 'plaintext-only') sp.contentEditable = 'true';
      sp.spellcheck = true; sp.title = labels[k] || k; sp.dataset.ph = labels[k] || k;
      sp.setAttribute('role', 'textbox'); sp.setAttribute('aria-label', labels[k] || k);
      sp.addEventListener('keydown', (e) => { if (e.key === 'Enter') e.preventDefault(); });
      sp.addEventListener('input', () => {
        const v = sp.textContent.replace(/\s+/g, ' ').slice(0, core.VARIANT_VALUE_MAX);
        if (inputs[k]) inputs[k].value = v;
        spans(k).forEach((o) => { if (o !== sp) o.textContent = v; });
      });
    });
  })().catch((e) => message(e.message || String(e), true));
  form.oninput = (e) => { const k = Object.keys(inputs).find((x) => inputs[x] === e.target); if (k) spans(k).forEach((o) => { o.textContent = inputs[k].value; }); };
  const row = document.createElement('div'); row.className = 'actions';
  row.append(button(t('Cancel'), () => render(), true), button(t('Save'), () => form.requestSubmit()));
  form.append(row);
  form.onsubmit = async (e) => {
    e.preventDefault();
    const v = draft();
    if (!v.name) { message(t('Give the variant a name.'), true); nameIn.focus(); return; }
    const before = getVariants();
    const list = before.filter((x) => x.id !== v.id); list.push(v);
    if (!(await setVariants(list))) {
      // Not kept (the mailbox's room for add-in settings is full): as it was before.
      await setVariants(before);
      message(t('Outlook did not keep the variant: the room for add-in settings in your mailbox is full. Shorten or delete other variants.'), true);
      return;
    }
    message(t('Variant saved. Pick it above for new mail or replies, or insert it here.'));
    render();
  };
  card.replaceChildren(form, live);
  form.scrollIntoView({ block: 'nearest' });
  nameIn.focus();
}

// Outlook's pane has no confirm(): the button asks once more itself.
async function deleteVariant(sig, btn) {
  if (btn.dataset.sure !== '1') {
    btn.dataset.sure = '1'; btn.textContent = t('Really delete?'); btn.classList.add('danger');
    setTimeout(() => { btn.dataset.sure = ''; btn.textContent = t('Delete'); btn.classList.remove('danger'); }, 4000);
    return;
  }
  await setVariants(getVariants().filter((v) => v.id !== sig.id));
  // A choice that named it falls back to the company's.
  const p = getPrefs(); let changed = false;
  ['new', 'reply'].forEach((k) => { if (p[k] === sig.id) { p[k] = ''; changed = true; } });
  if (changed) await setPrefs(p);
  message('');
  render();
}

// How it works, in a few lines: only what this setup allows (choosing, variants, own signatures).
function renderHelp() {
  const o = cfg.options || {}, items = [t('Your signature goes into new mail and replies by itself, with your details from the company directory. If they are wrong, IT changes them there; "Read my details again" fetches them anew.')];
  if (o.allowUserChoice !== false) items.push(t('Above, choose which signature goes into new mail and which into replies. "Company default" follows what the company sets.'));
  items.push(t('Insert puts a signature into the message open now, e.g. another one than usual.'));
  if (core.variantsAllowed(cfg)) {
    items.push(core.ownAllowed(cfg)
      ? t('Own variant: a copy of a company signature that is yours to change, all of it. It is kept in your mailbox, also on your other devices with Outlook.')
      : t('Own variant: a copy of a company signature in which you change the marked fields (e.g. a greeting); the rest stays the company signature. It is kept in your mailbox.'));
    items.push(t('After saving, choose your variant above for new mail or replies.'));
  }
  $('helpList').replaceChildren(...items.map((x) => { const li = document.createElement('li'); li.textContent = x; return li; }));
  $('help').hidden = false;
}

function render() {
  renderHelp();
  full = core.withVariants(cfg, getVariants());
  const list = core.available(full, user.values, from, user.groups);
  renderChoice(list);
  renderList(list);
  if (isClassicOutlook()) keepClassic();
}

// An image as small as it is shown (twice that, for sharp screens, at most), in its own type:
// a PNG stays a PNG (it may be transparent), a JPEG is saved again at good quality; a GIF as it is.
async function shrink(img, shown) {
  if (!/^image\/(png|jpeg)$/.test(img.type)) return img;
  const el = new Image();
  await new Promise((res, rej) => { el.onload = res; el.onerror = rej; el.src = 'data:' + img.type + ';base64,' + img.b64; });
  const want = Math.min(el.naturalWidth, Math.max(1, (shown || el.naturalWidth) * 1.5));
  const w = Math.round(want), h = Math.round(el.naturalHeight * want / el.naturalWidth);
  const c = document.createElement('canvas'); c.width = w; c.height = h;
  const g = c.getContext('2d'); g.imageSmoothingQuality = 'high'; g.drawImage(el, 0, 0, w, h);
  const b64 = c.toDataURL(img.type, 0.82).split(',')[1] || '';
  return b64 && b64.length < img.b64.length ? { b64 } : img;
}
// Classic Outlook: what would go in by itself, kept in the mailbox for its event handler.
let classicBusy = false;
async function keepClassic() {
  if (classicBusy) return;
  classicBusy = true;
  try {
    // What goes in today first; then what goes in once today's campaigns end, and on the first
    // days of those still to come: the event handler picks among them by the day.
    const prefs = getPrefs(), sigs = [];
    const add = (x) => { if (x && !sigs.some((y) => y.id === x.id)) sigs.push(x); };
    const plain = Object.assign({}, full, { signatures: full.signatures.filter((s) => !core.isCampaign(s)) });
    const today = new Date().toISOString().slice(0, 10);
    const later = full.signatures.map((s) => String(s.validFrom || '')).filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d) && d > today).sort().slice(0, 3);
    ['new', 'reply'].forEach((kind) => add(core.pick(full, user.values, from, kind, prefs, user.groups)));
    ['new', 'reply'].forEach((kind) => add(core.pick(plain, user.values, from, kind, prefs, user.groups)));
    later.forEach((d) => ['new', 'reply'].forEach((kind) => add(core.pick(full, user.values, from, kind, prefs, user.groups, new Date(d + 'T12:00:00')))));
    if (!sigs.length) return;
    const r = await keepForClassic(cfg, sigs, user.values, shrink);
    if (!r.kept) message(t('Classic Outlook: these signatures are too large to be kept in your mailbox, so they are not put in by themselves there. Use Insert.'), true);
    else if (r.left) message(t('Classic Outlook: {0} image(s) did not fit into your mailbox settings and are left out when the signature goes in by itself there. Insert puts in everything.', r.left));
  } catch (e) { console.log('OSign: classic copy: ' + (e && e.message ? e.message : e)); } finally { classicBusy = false; }
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
  if (user.groups === null && core.groupsUsed(cfg).length) message(user.error ? t('Your groups could not be read, so signatures meant for groups are not shown: {0}', user.error.message) : t('Your groups could not be read, so signatures meant for groups are not shown.'), true);
  render();
}

// OSign 0.9.14 · signatures published 30.10.2026, 14:02
function showVersion() {
  const at = cfg && cfg.updated ? new Date(cfg.updated) : null;
  const when = at && !isNaN(at) ? at.toLocaleString(lang === 'de' ? 'de-CH' : LANG[lang] ? LANG[lang].locale : 'en-GB', { dateStyle: 'short', timeStyle: 'short' }) : '';
  $('ver').textContent = 'OSign ' + OSIGN_VERSION + (when ? ' · ' + t('signatures published {0}', when) : '');
}

Office.onReady(() => {
  translatePage();
  showVersion();
  $('refresh').addEventListener('click', (e) => { e.preventDefault(); load(true); });
  load(false);
});
