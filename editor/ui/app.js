/*
 * OSign editor - the web UI.
 *
 * Talks to its backend only through the HTTP API in docs/API.md (today OSign.ps1; a container
 * later). Everything about signatures themselves - placeholders, the lines left out when empty,
 * who gets which signature - is in shared/signature-core.js, the same file the add-in runs.
 */
'use strict';
(function () {
  const core = window.SignatureCore;
  const $ = (id) => document.getElementById(id);

  // ---------------------------------------------------------------------------------------------
  // Backend
  // ---------------------------------------------------------------------------------------------
  // The token comes after # in the address the window was opened with (never sent to the server
  // in a URL, never in a log). Kept for this tab, so a reload still works.
  const TOKEN = (function () {
    const m = /(?:^#|&)t=([0-9a-f]{64})/.exec(location.hash);
    try { if (m) sessionStorage.setItem('osign.token', m[1]); return m ? m[1] : sessionStorage.getItem('osign.token') || ''; }
    catch (e) { return m ? m[1] : ''; }
  })();

  let serverDown = false;
  async function api(path, body, opts) {
    let r;
    try {
      r = await fetch('/api/' + path, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + TOKEN },
        body: JSON.stringify(body || {}),
        keepalive: !!(opts && opts.keepalive)
      });
    } catch (e) {
      if (!serverDown) { serverDown = true; toast(t('The OSign server is not running. Start OSign.ps1 again.'), true); }
      throw new Error(t('The OSign server is not running.'));
    }
    serverDown = false;
    let j;
    try { j = await r.json(); } catch (e) { throw new Error(t('The server answered {0}.', r.status)); }
    if (r.status === 401) throw new Error(t('This window is not signed in to the server: open the address OSign printed.'));
    if (!j.ok) throw new Error(t(j.error || 'Failed.'));
    return j;
  }

  // ---------------------------------------------------------------------------------------------
  // State
  // ---------------------------------------------------------------------------------------------
  const S = {
    config: null, settings: null, info: null, assets: [], dirty: false, publishedAt: 0,
    view: 'sig', sigId: null, mode: 'design', adValues: null, adName: '', 
  };
  const sig = () => (S.config.signatures || []).find((s) => s.id === S.sigId) || null;
  // Signatures are compiled in a document of their own: one that is not shown loads no images.
  const scratchDoc = document.implementation.createHTMLDocument('');
  const fieldKeys = () => S.config.fields.map((f) => f.key);

  function newId(name) {
    const base = String(name || 'signature').toLowerCase().normalize('NFKD').replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 30) || 'signature';
    let id = base, i = 2;
    while (S.config.signatures.some((s) => s.id === id)) id = base + '-' + i++;
    return id;
  }

  // The configuration as published: every signature compiled, the image address set.
  function prepared() {
    const c = JSON.parse(JSON.stringify(S.config));
    c.signatures.forEach((s) => { s.html = sanitize(s.html); s.compiled = core.compile(s.html, scratchDoc); });
    return c;
  }

  // ---------------------------------------------------------------------------------------------
  // Saving: every change goes to the draft on its own, a moment after the last keystroke.
  // ---------------------------------------------------------------------------------------------
  let saveTimer = null, saving = Promise.resolve(), changeSeq = 0;
  function changed(opts) {
    S.dirty = true; changeSeq++; renderStatus(t('Saving…'));
    clearTimeout(saveTimer);
    saveTimer = setTimeout(save, 600);
    if (!opts || opts.list !== false) renderList();
    if (!opts || opts.preview !== false) schedulePreview();
  }
  function save(opts) {
    clearTimeout(saveTimer); saveTimer = null;
    const json = JSON.stringify(prepared()), seq = changeSeq;
    saving = saving.then(() => api('draft/save', { configJson: json }, opts))
      // The answer to an older save says nothing about changes made since: those are still coming.
      .then((r) => { if (seq === changeSeq) { S.dirty = r.dirty; renderStatus(); } })
      .catch((e) => { renderStatus(t('Not saved')); toast(t(e.message), true); });
    return saving;
  }
  window.addEventListener('beforeunload', () => { if (saveTimer) save({ keepalive: true }); });
  setInterval(() => { api('ping').catch(() => {}); }, 60000);

  function renderStatus(text) {
    const el = $('status');
    if (text) { el.textContent = text; el.className = 'status'; return; }
    if (S.dirty) { el.textContent = t('Unpublished changes'); el.className = 'status dirty'; }
    else { el.textContent = S.publishedAt ? t('Published') : t('Saved'); el.className = 'status clean'; }
    el.title = S.publishedAt ? t('Last published {0}', new Date(S.publishedAt).toLocaleString(I18N.lang)) : '';
  }

  // ---------------------------------------------------------------------------------------------
  // HTML hygiene: what goes into a signature is markup Outlook shows, nothing that runs.
  // ---------------------------------------------------------------------------------------------
  const DROP = 'script,style,meta,link,title,iframe,object,embed,form,input,button,select,textarea,base,xml,o\\:p,svg,video,audio';
  function sanitize(html) {
    const t = document.createElement('template');
    t.innerHTML = String(html || '');
    t.content.querySelectorAll(DROP).forEach((n) => {
      // Word's <o:p> holds the text of the paragraph's end: keep that, lose the tag.
      if (n.nodeName.toLowerCase() === 'o:p') n.replaceWith(...n.childNodes); else n.remove();
    });
    const walker = document.createTreeWalker(t.content, NodeFilter.SHOW_ELEMENT | NodeFilter.SHOW_COMMENT);
    const comments = [];
    while (walker.nextNode()) {
      const n = walker.currentNode;
      if (n.nodeType === 8) { comments.push(n); continue; }
      [...n.attributes].forEach((a) => {
        const name = a.name.toLowerCase(), v = a.value.trim().toLowerCase();
        if (name.startsWith('on')) n.removeAttribute(a.name);
        else if ((name === 'href' || name === 'src' || name === 'background') && /^(javascript|vbscript|data:text)/.test(v.replace(/\s+/g, ''))) n.removeAttribute(a.name);
        else if (name === 'class' && /^mso/i.test(a.value)) n.removeAttribute(a.name);
        else if (name === 'contenteditable') n.removeAttribute(a.name);
      });
    }
    comments.forEach((c) => c.remove());
    return t.innerHTML.trim();
  }

  // ---------------------------------------------------------------------------------------------
  // Sidebar and views
  // ---------------------------------------------------------------------------------------------
  function el(tag, attrs, ...kids) {
    const e = document.createElement(tag);
    Object.entries(attrs || {}).forEach(([k, v]) => {
      if (v === undefined || v === null || v === false) return;
      if (k === 'class') e.className = v; else if (k.startsWith('on')) e.addEventListener(k.slice(2), v); else e.setAttribute(k, v === true ? '' : v);
    });
    kids.flat().forEach((k) => { if (k !== null && k !== undefined) e.append(k.nodeType ? k : document.createTextNode(String(k))); });
    return e;
  }

  function audienceText(s) {
    const a = s.audience || {}, parts = [];
    if ((a.groups || []).length) parts.push(a.groups.join(', '));
    if ((a.departments || []).length) parts.push(a.departments.join(', '));
    if ((a.domains || []).length) parts.push('@' + a.domains.join(', @'));
    return parts.length ? parts.join(' · ') : t('everyone');
  }

  function renderList() {
    const ul = $('sigList'); ul.textContent = '';
    if (!S.config.signatures.length) ul.append(el('li', { class: 'none' }, t('None yet')));
    S.config.signatures.forEach((s) => {
      const a = el('a', { href: '#', class: S.view === 'sig' && s.id === S.sigId ? 'active' : '', onclick: (e) => { e.preventDefault(); openSig(s.id); } },
        s.name || t('(no name)'),
        el('div', { class: 'badges' },
          s.defaultNew ? el('span', { class: 'badge on' }, t('New')) : null,
          s.defaultReply ? el('span', { class: 'badge on' }, t('Reply')) : null,
          audienceText(s) !== t('everyone') ? el('span', { class: 'badge', title: audienceText(s) }, t('Limited')) : null));
      ul.append(el('li', null, a));
    });
    document.querySelectorAll('.nav-list a').forEach((a) => a.classList.toggle('active', a.dataset.view === S.view));
  }

  function show(view) {
    S.view = view;
    $('viewSig').hidden = view !== 'sig';
    $('viewFields').hidden = view !== 'fields';
    $('viewOptions').hidden = view !== 'options';
    $('viewUsers').hidden = view !== 'users';
    $('viewEmpty').hidden = view !== 'empty';
    renderList();
  }

  function openSig(id) {
    flushSource();
    S.sigId = id;
    if (!sig()) { show(S.config.signatures.length ? 'sig' : 'empty'); if (S.config.signatures.length) openSig(S.config.signatures[0].id); return; }
    show('sig');
    const s = sig();
    $('sigName').value = s.name;
    $('sigNew').checked = !!s.defaultNew;
    $('sigReply').checked = !!s.defaultReply;
    $('sigRemoveEmpty').checked = s.removeEmptyLines !== false;
    $('sigGroups').value = (s.audience.groups || []).join(', ');
    $('sigDeps').value = (s.audience.departments || []).join(', ');
    $('sigDoms').value = (s.audience.domains || []).join(', ');
    $('audienceSummary').textContent = audienceText(s);
    loadDesign(s.html);
    $('source').value = s.html;
    schedulePreview();
  }

  const STARTER = [
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:10pt;color:#222222">Freundliche Grüsse</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:10pt">&nbsp;</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:10pt;color:#222222"><b>@@givenName @@sn</b></p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:9pt;color:#555555">@@description</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:9pt;color:#555555">@@title, @@department</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:10pt">&nbsp;</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:9pt;color:#222222">Direkt @@homePhone</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:9pt;color:#222222">Mobile @@mobile</p>',
    '<p style="margin:0;font-family:Arial,sans-serif;font-size:9pt"><a href="mailto:@@mail" style="color:#2f6fde">@@mail</a></p>'
  ].join('\n');

  function addSig(from) {
    const name = from ? from.name + ' ' + t('(copy)') : (S.config.signatures.length ? t('New signature') : t('Standard'));
    const s = from ? JSON.parse(JSON.stringify(from)) : core.newSignature('', name);
    s.id = newId(name); s.name = name;
    if (from) { s.defaultNew = false; s.defaultReply = false; } else { s.html = STARTER; s.defaultNew = !S.config.signatures.some((x) => x.defaultNew); s.defaultReply = !S.config.signatures.some((x) => x.defaultReply); }
    S.config.signatures.push(s);
    changed();
    openSig(s.id);
    if (!from) { $('sigName').focus(); $('sigName').select(); }
  }

  // ---------------------------------------------------------------------------------------------
  // The design editor: an iframe of its own, so the app's styles never reach the signature
  // ---------------------------------------------------------------------------------------------
  const design = $('design');
  let ddoc = null, designReady = null;
  const EDITOR_CSS = 'html{background:#fff}body{margin:14px 16px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222;min-height:380px;outline:none;word-wrap:break-word}' +
    'img{max-width:100%;cursor:pointer}img.sel{outline:2px solid #2f6fde}table{border-collapse:collapse}td{vertical-align:top}' +
    'table:not([border]) td{outline:1px dashed #d6dbe3}p{margin:0}a{color:#2f6fde}';

  function initDesign() {
    designReady = new Promise((resolve) => {
      design.addEventListener('load', () => {
        ddoc = design.contentDocument;
        ddoc.body.contentEditable = 'true';
        try { ddoc.execCommand('styleWithCSS', false, true); ddoc.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) { }
        ddoc.addEventListener('input', onDesignInput);
        ddoc.addEventListener('keyup', updateToolbarState);
        ddoc.addEventListener('mouseup', updateToolbarState);
        ddoc.addEventListener('click', onDesignClick);
        ddoc.addEventListener('paste', onPaste);
        resolve();
      }, { once: true });
    });
    // The base makes "assets/logo.png" load from this server, as it will from the published one.
    design.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><base href="/"><style>' + EDITOR_CSS + '</style></head><body></body></html>';
  }

  async function loadDesign(html) {
    await designReady;
    ddoc.body.innerHTML = html || '';
    hideImgBar();
  }

  function onDesignInput() {
    const s = sig(); if (!s) return;
    s.html = cleanEditorHtml(ddoc.body.innerHTML);
    changed({ list: false });
  }

  // The editor's own marks (a selected image's outline) never go into the signature.
  function cleanEditorHtml(html) {
    return html.replace(/\sclass="sel"/g, '').replace(/(<img[^>]*?)\sclass="([^"]*)\bsel\b([^"]*)"/g, '$1 class="$2$3"');
  }

  function onPaste(e) {
    const html = e.clipboardData && e.clipboardData.getData('text/html');
    if (!html) return;
    e.preventDefault();
    const body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
    const clean = sanitize((body ? body[1] : html).replace(/<!--StartFragment-->|<!--EndFragment-->/g, ''));
    ddoc.execCommand('insertHTML', false, clean);
  }

  function exec(cmd, val) {
    if (S.mode !== 'design') return;
    design.contentWindow.focus();
    ddoc.execCommand(cmd, false, val === undefined ? null : val);
    onDesignInput(); updateToolbarState();
  }

  // execCommand only knows sizes 1-7: size 7 is set, and what it made gets the real size.
  function setFontSize(pt) {
    exec('fontSize', '7');
    ddoc.querySelectorAll('font[size="7"]').forEach((f) => { const sp = ddoc.createElement('span'); sp.style.fontSize = pt + 'pt'; sp.append(...f.childNodes); f.replaceWith(sp); });
    ddoc.querySelectorAll('[style*="xxx-large"]').forEach((n) => { n.style.fontSize = pt + 'pt'; });
    onDesignInput();
  }

  // Inserted as it is written. execCommand('insertHTML') would copy the editor's own font onto
  // every element it inserts (an image with a font-size), which then goes out in the signature.
  function insertHtml(html) {
    if (S.mode !== 'design') return;
    design.contentWindow.focus();
    const frag = ddoc.createRange().createContextualFragment(html);
    const last = frag.lastChild;
    const sel = ddoc.getSelection();
    if (!sel.rangeCount || !ddoc.body.contains(sel.getRangeAt(0).commonAncestorContainer)) ddoc.body.append(frag);
    else {
      const r = sel.getRangeAt(0); r.deleteContents(); r.insertNode(frag);
      if (last) { r.setStartAfter(last); r.collapse(true); sel.removeAllRanges(); sel.addRange(r); }
    }
    onDesignInput();
  }

  const ICON = {
    bold: '<path d="M7 5h6a3.5 3.5 0 0 1 0 7H7zM7 12h7a3.5 3.5 0 0 1 0 7H7z"/>',
    italic: '<path d="M14 5h-4M14 19h-4M14 5l-4 14"/>',
    underline: '<path d="M7 4v7a5 5 0 0 0 10 0V4M5 20h14"/>',
    left: '<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>',
    center: '<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>',
    right: '<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>',
    link: '<path d="M10 14a4 4 0 0 0 6 0l3-3a4 4 0 0 0-6-6l-1 1M14 10a4 4 0 0 0-6 0l-3 3a4 4 0 0 0 6 6l1-1"/>',
    unlink: '<path d="M10 14a4 4 0 0 0 6 0l3-3a4 4 0 0 0-6-6l-1 1M14 10a4 4 0 0 0-6 0l-3 3a4 4 0 0 0 6 6l1-1M4 4l16 16"/>',
    image: '<rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-8 8"/>',
    table: '<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M3 12h18M12 5v14"/>',
    hr: '<path d="M4 12h16"/>',
    clear: '<path d="M6 5h12M12 5l-3 14M4 20l16-16"/>',
    undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>',
    redo: '<path d="M15 14l5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>'
  };
  const svg = (name) => '<svg viewBox="0 0 24 24" aria-hidden="true">' + ICON[name] + '</svg>';

  function buildToolbar() {
    const tb = $('toolbar');
    const btn = (name, title, fn, cmd) => { title = t(title); const b = el('button', { type: 'button', title, 'aria-label': title, 'data-cmd': cmd || '' }); b.innerHTML = svg(name); b.addEventListener('mousedown', (e) => e.preventDefault()); b.addEventListener('click', fn); return b; };
    const sep = () => el('span', { class: 'sep' });
    const fonts = ['Arial', 'Aptos', 'Calibri', 'Segoe UI', 'Verdana', 'Tahoma', 'Helvetica', 'Georgia', 'Times New Roman'];
    const fontSel = el('select', { title: t('Font'), 'aria-label': t('Font') }, el('option', { value: '' }, t('Font')), fonts.map((f) => el('option', { value: f }, f)));
    fontSel.addEventListener('change', () => { if (fontSel.value) exec('fontName', fontSel.value + (fontSel.value === 'Georgia' || fontSel.value === 'Times New Roman' ? ', serif' : ', sans-serif')); fontSel.value = ''; });
    const sizeSel = el('select', { title: t('Size'), 'aria-label': t('Size') }, el('option', { value: '' }, t('Size')), [8, 9, 10, 11, 12, 14, 16, 18, 24].map((n) => el('option', { value: n }, n + ' pt')));
    sizeSel.addEventListener('change', () => { if (sizeSel.value) setFontSize(sizeSel.value); sizeSel.value = ''; });
    const color = el('input', { type: 'color', title: t('Text colour'), 'aria-label': t('Text colour'), value: '#222222' });
    color.addEventListener('input', () => exec('foreColor', color.value));
    const fieldSel = el('select', { class: 'field-pick', title: t('Insert a field'), 'aria-label': t('Insert a field') });
    fieldSel.addEventListener('focus', () => fillFieldSelect(fieldSel));
    fieldSel.addEventListener('change', () => {
      const k = fieldSel.value; fieldSel.value = '';
      if (!k) return;
      if (S.mode === 'html') insertAtCursor($('source'), '@@' + k); else exec('insertText', '@@' + k);
    });
    fillFieldSelect(fieldSel);
    tb.append(
      fieldSel, sep(), fontSel, sizeSel, color, sep(),
      btn('bold', 'Bold (Ctrl+B)', () => exec('bold'), 'bold'), btn('italic', 'Italic (Ctrl+I)', () => exec('italic'), 'italic'), btn('underline', 'Underline (Ctrl+U)', () => exec('underline'), 'underline'), sep(),
      btn('left', 'Align left', () => exec('justifyLeft'), 'justifyLeft'), btn('center', 'Centre', () => exec('justifyCenter'), 'justifyCenter'), btn('right', 'Align right', () => exec('justifyRight'), 'justifyRight'), sep(),
      btn('link', 'Link', makeLink), btn('unlink', 'Remove link', () => exec('unlink')), btn('image', 'Image', openImageDialog), btn('table', 'Two columns (logo beside the text)', insertTable), btn('hr', 'Line', () => exec('insertHorizontalRule')), sep(),
      btn('clear', 'Clear formatting', () => exec('removeFormat')), btn('undo', 'Undo (Ctrl+Z)', () => exec('undo')), btn('redo', 'Redo (Ctrl+Y)', () => exec('redo'))
    );
  }

  function fillFieldSelect(sel) {
    sel.textContent = '';
    sel.append(el('option', { value: '' }, t('@@ Insert field')));
    S.config && S.config.fields.forEach((f) => sel.append(el('option', { value: f.key }, (f.label || f.key) + '  @@' + f.key)));
  }

  function insertAtCursor(ta, text) {
    const a = ta.selectionStart, b = ta.selectionEnd;
    ta.setRangeText(text, a, b, 'end'); ta.focus(); ta.dispatchEvent(new Event('input'));
  }

  function updateToolbarState() {
    if (!ddoc) return;
    document.querySelectorAll('#toolbar button[data-cmd]').forEach((b) => {
      const c = b.dataset.cmd; if (!c) return;
      let on = false; try { on = ddoc.queryCommandState(c); } catch (e) { }
      b.classList.toggle('on', on);
    });
  }

  function makeLink() {
    const url = prompt(t('Link address (https://…, mailto:…, tel:…). Fields work here too: mailto:@@mail'), 'https://');
    if (!url || url === 'https://') return;
    if (!/^(https?:|mailto:|tel:)/i.test(url.trim())) { toast(t('A link starts with https://, mailto: or tel:.'), true); return; }
    exec('createLink', url.trim());
  }

  function insertTable() {
    insertHtml('<table cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tr>' +
      '<td style="padding:0 14px 0 0;vertical-align:top">Logo</td>' +
      '<td style="vertical-align:top;font-family:Arial,sans-serif;font-size:10pt"><b>@@givenName @@sn</b><br>@@title<br>Mobile @@mobile<br>@@mail</td>' +
      '</tr></table><p><br></p>');
  }

  // Selected image: width, alt text and a link, in a bar above the editor.
  let selImg = null;
  function onDesignClick(e) {
    if (e.target && e.target.nodeName === 'IMG') {
      hideImgBar(); selImg = e.target; selImg.classList.add('sel');
      $('imgW').value = selImg.getAttribute('width') || selImg.naturalWidth || '';
      $('imgAlt').value = selImg.getAttribute('alt') || '';
      const a = selImg.closest('a'); $('imgHref').value = a ? a.getAttribute('href') : '';
      $('imgBar').hidden = false;
    } else hideImgBar();
  }
  function hideImgBar() {
    if (selImg) { selImg.classList.remove('sel'); if (!selImg.className) selImg.removeAttribute('class'); }
    selImg = null; $('imgBar').hidden = true;
  }
  function applyImg() {
    if (!selImg) return;
    const w = parseInt($('imgW').value, 10);
    if (w > 0) { selImg.setAttribute('width', w); selImg.removeAttribute('height'); selImg.style.width = ''; selImg.style.height = ''; if (!selImg.getAttribute('style')) selImg.removeAttribute('style'); }
    selImg.setAttribute('alt', $('imgAlt').value);
    const href = $('imgHref').value.trim(), a = selImg.closest('a');
    if (href && /^(https?:|mailto:|tel:)/i.test(href)) {
      if (a) a.setAttribute('href', href); else { const na = ddoc.createElement('a'); na.href = href; selImg.replaceWith(na); na.append(selImg); }
    } else if (!href && a) a.replaceWith(...a.childNodes);
    onDesignInput();
  }

  // Images: pick one already in the draft, or upload one.
  function openImageDialog() {
    $('imgErr').textContent = '';
    renderImageList();
    $('dlgImage').showModal();
  }
  function renderImageList() {
    const box = $('imgList'); box.textContent = '';
    if (!S.assets.length) box.append(el('p', { class: 'lead' }, t('No images yet.')));
    S.assets.forEach((a) => {
      const del = el('span', { class: 'del', title: t('Remove from the draft'), role: 'button', onclick: (e) => { e.stopPropagation(); deleteAsset(a.name); } }, t('remove'));
      box.append(el('button', { type: 'button', onclick: () => { $('dlgImage').close(); insertImage('assets/' + a.name); } },
        el('img', { src: '/assets/' + encodeURIComponent(a.name), alt: '' }), a.name, del));
    });
  }
  function insertImage(src) {
    const i = new Image();
    i.onload = () => { const w = Math.min(i.naturalWidth, 240); insertHtml('<img src="' + core.escapeHtml(src) + '" width="' + w + '" alt="" style="border:0">'); };
    i.onerror = () => insertHtml('<img src="' + core.escapeHtml(src) + '" alt="" style="border:0">');
    i.src = '/' + src;
  }
  async function deleteAsset(name) {
    const used = S.config.signatures.filter((s) => (s.html || '').includes('assets/' + name)).map((s) => s.name);
    if (used.length && !confirm(t('{0} is used in: {1}. Remove it anyway?', name, used.join(', ')))) return;
    try { await api('assets/delete', { name }); S.assets = S.assets.filter((a) => a.name !== name); renderImageList(); changed({ list: false }); } catch (e) { $('imgErr').textContent = t(e.message); }
  }
  $('imgFile').addEventListener('change', async () => {
    const f = $('imgFile').files[0]; $('imgFile').value = '';
    if (!f) return;
    $('imgErr').textContent = '';
    if (f.size > 1024 * 1024) { $('imgErr').textContent = t('The image is larger than 1 MB.'); return; }
    const data = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = rej; r.readAsDataURL(f); });
    try {
      const r = await api('assets/upload', { name: f.name, data });
      S.assets = S.assets.filter((a) => a.name !== r.name).concat([{ name: r.name, size: r.size }]).sort((a, b) => a.name.localeCompare(b.name));
      $('dlgImage').close();
      insertImage(r.path);
    } catch (e) { $('imgErr').textContent = t(e.message); }
  });

  // Design / HTML
  function setMode(mode) {
    if (mode === S.mode) return;
    const s = sig();
    if (mode === 'html') { $('source').value = s ? s.html : ''; }
    else { flushSource(); loadDesign(s ? s.html : ''); }
    S.mode = mode;
    document.querySelectorAll('.tab').forEach((t) => t.classList.toggle('active', t.dataset.mode === mode));
    design.hidden = mode !== 'design';
    $('source').hidden = mode !== 'html';
    document.querySelectorAll('#toolbar > *').forEach((n) => { n.style.display = mode === 'html' && !n.classList.contains('field-pick') ? 'none' : ''; });
    hideImgBar();
  }
  // What was typed in the HTML tab, made safe, once the tab is left.
  function flushSource() {
    if (S.mode !== 'html') return;
    const s = sig(); if (!s) return;
    const clean = sanitize($('source').value);
    if (clean !== s.html) { s.html = clean; changed({ list: false }); }
  }
  $('source').addEventListener('input', () => { const s = sig(); if (!s) return; s.html = $('source').value; changed({ list: false }); });
  $('source').addEventListener('blur', flushSource);

  // ---------------------------------------------------------------------------------------------
  // Preview
  // ---------------------------------------------------------------------------------------------
  let pvTimer = null;
  function schedulePreview() { clearTimeout(pvTimer); pvTimer = setTimeout(renderPreview, 150); }

  function previewValues() {
    const src = $('pvSource').value, v = {};
    S.config.fields.forEach((f) => { v[f.key] = src === 'sample' ? (f.sample || '') : src === 'ad' && S.adValues ? (S.adValues[f.key] || '') : ''; });
    return v;
  }

  function renderPreview() {
    const s = sig(); if (!s || S.view !== 'sig') return;
    const html = core.render(core.compile(sanitize(s.html), scratchDoc), previewValues(), { assetBase: '/', removeEmptyLines: s.removeEmptyLines !== false });
    const asText = $('pvText').checked;
    const greet = asText ? '' : '<p>' + core.escapeHtml(t('Hello,')) + '</p><p>' + core.escapeHtml(t('this is how the signature looks under a message.')) + '</p><p>&nbsp;</p>';
    const body = asText ? '<pre style="font:11pt/1.4 Consolas,monospace;white-space:pre-wrap">' + core.escapeHtml(core.htmlToText(html)) + '</pre>' : greet + html;
    $('preview').srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:14px 16px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222}p{margin:0}table{border-collapse:collapse}</style></head><body>' + body + '</body></html>';

    // What to know about this signature before it goes out.
    const notes = [], keys = fieldKeys(), used = core.placeholdersIn(s.html);
    const unknown = used.filter((k) => keys.indexOf(k) < 0);
    const noGraph = used.filter((k) => { const f = S.config.fields.find((x) => x.key === k); return f && core.fieldSource(S.config, f) === 'none'; });
    if (unknown.length) notes.push(el('div', { class: 'warn' }, t('Not a field: {0} - left as typed.', unknown.map((k) => '@@' + k).join(', '))));
    if (noGraph.length) notes.push(el('div', { class: 'warn' }, core.valueSource(S.config) === 'graph'
      ? t('No Graph source, so empty in Outlook: {0}. See Fields.', noGraph.map((k) => '@@' + k).join(', '))
      : t('No source (neither Graph nor AD attribute), so empty in Outlook: {0}. See Fields.', noGraph.map((k) => '@@' + k).join(', '))));
    if ($('pvSource').value === 'ad' && S.adName) {
      // What this user would get, as the add-in decides it (the From address being their own).
      const v = S.adValues || {}, from = v.mail || '', g = S.adGroups || [];
      const nm = (x) => (x ? x.name : t('none'));
      const open = core.isAvailable(s, v, from, g);
      notes.push(el('div', { class: 'pv-groups' }, t('Values of {0} from Active Directory.', S.adName) + ' ' +
        (open ? t('This signature is open to them.') : t('This signature is NOT open to them.')) + ' ' +
        t('New mail: {0}, replies: {1}.', nm(core.pick(S.config, v, from, 'new', null, g)), nm(core.pick(S.config, v, from, 'reply', null, g)))));
    }
    $('pvNote').replaceChildren(...notes);
  }

  $('pvSource').addEventListener('change', () => {
    const ad = $('pvSource').value === 'ad';
    $('pvAd').hidden = !ad;
    if (ad && !(S.info && S.info.features && S.info.features.adLookup)) toast(t('Active Directory is not reachable from here: this needs Windows on a machine in the domain.'), true);
    if (ad && !S.adValues) $('pvUser').focus();
    renderPreview();
  });
  $('pvText').addEventListener('change', renderPreview);
  async function adLookup() {
    try {
      const r = await api('ad/lookup', { query: $('pvUser').value, fields: S.config.fields.map((f) => ({ key: f.key, ad: f.ad })) });
      S.adValues = r.values; S.adName = r.displayName; S.adGroups = r.groups || []; renderPreview();
    } catch (e) { toast(t(e.message), true); }
  }
  $('pvLookup').addEventListener('click', adLookup);
  $('pvUser').addEventListener('keydown', (e) => { if (e.key === 'Enter') adLookup(); });

  // ---------------------------------------------------------------------------------------------
  // Signature properties
  // ---------------------------------------------------------------------------------------------
  const splitList = (v) => v.split(/[,;]/).map((x) => x.trim()).filter(Boolean);
  $('sigName').addEventListener('input', () => { const s = sig(); if (s) { s.name = $('sigName').value; changed({ preview: false }); } });
  $('sigName').addEventListener('blur', () => { const s = sig(); if (s && !s.name.trim()) { s.name = t('Signature'); $('sigName').value = s.name; changed(); } });
  // One default for new mail and one for replies, unless the audiences differ: a second one with
  // the same audience could never be picked.
  function setDefault(flag, on) {
    const s = sig(); if (!s) return;
    if (on) S.config.signatures.forEach((x) => { if (x !== s && x[flag] && audienceText(x) === audienceText(s)) x[flag] = false; });
    s[flag] = on; changed();
  }
  $('sigNew').addEventListener('change', () => setDefault('defaultNew', $('sigNew').checked));
  $('sigReply').addEventListener('change', () => setDefault('defaultReply', $('sigReply').checked));
  $('sigRemoveEmpty').addEventListener('change', () => { const s = sig(); if (s) { s.removeEmptyLines = $('sigRemoveEmpty').checked; changed({ list: false }); } });
  $('sigGroups').addEventListener('input', () => { const s = sig(); if (s) { s.audience.groups = splitList($('sigGroups').value); $('audienceSummary').textContent = audienceText(s); changed({ preview: false }); } });
  // Group names from AD to pick from, for what is typed after the last comma.
  let grpTimer = null;
  $('sigGroups').addEventListener('input', () => {
    clearTimeout(grpTimer);
    if (!(S.info && S.info.features && S.info.features.adLookup)) return;
    const q = $('sigGroups').value.split(/[,;]/).pop().trim();
    if (q.length < 2) return;
    grpTimer = setTimeout(async () => {
      try {
        const r = await api('ad/groups', { query: q });
        const before = $('sigGroups').value.replace(/[^,;]*$/, '');
        $('adGroups').replaceChildren(...r.groups.map((g) => el('option', { value: (before ? before.trim() + ' ' : '') + g })));
      } catch (e) { /* the list stays as it was */ }
    }, 250);
  });
  $('sigDeps').addEventListener('input', () => { const s = sig(); if (s) { s.audience.departments = splitList($('sigDeps').value); $('audienceSummary').textContent = audienceText(s); changed({ preview: false }); } });
  $('sigDoms').addEventListener('input', () => { const s = sig(); if (s) { s.audience.domains = splitList($('sigDoms').value).map((d) => d.replace(/^@/, '').toLowerCase()); $('audienceSummary').textContent = audienceText(s); changed({ preview: false }); } });
  $('btnNew').addEventListener('click', () => addSig());
  $('btnNew2').addEventListener('click', () => addSig());
  $('btnDup').addEventListener('click', () => { flushSource(); const s = sig(); if (s) addSig(s); });
  $('btnDel').addEventListener('click', () => {
    const s = sig(); if (!s || !confirm(t('Delete the signature "{0}"? It goes from Outlook with the next publish.', s.name))) return;
    S.config.signatures = S.config.signatures.filter((x) => x !== s);
    changed(); openSig(S.config.signatures[0] ? S.config.signatures[0].id : null);
  });
  document.querySelectorAll('.tab').forEach((t) => t.addEventListener('click', () => setMode(t.dataset.mode)));
  ['imgW', 'imgAlt', 'imgHref'].forEach((id) => $(id).addEventListener('change', applyImg));
  $('imgDone').addEventListener('click', () => { applyImg(); hideImgBar(); });

  // ---------------------------------------------------------------------------------------------
  // Import: templates (.htm) with their images, from a folder or picked files. CI-Sign keeps them
  // as UTF-16 files that load "images/<name>"; classic Outlook as "<name>_files/<name>".
  // ---------------------------------------------------------------------------------------------
  let impItems = [], impImages = new Map();
  const IMG_EXT = /\.(png|jpe?g|gif)$/i;

  function relativeSrcs(doc) {
    return [...doc.querySelectorAll('img[src]')].map((i) => i.getAttribute('src')).filter((src) => !/^(https?:|cid:|data:)/i.test(src));
  }

  async function readImportFiles(files) {
    impItems = []; impImages = new Map();
    const list = [...files];
    list.filter((f) => IMG_EXT.test(f.name)).forEach((f) => { if (!impImages.has(f.name.toLowerCase())) impImages.set(f.name.toLowerCase(), f); });
    for (const f of list.filter((x) => /\.html?$/i.test(x.name))) {
      // A folder holds more than templates (CI-Sign's own editor, for one): those stay unticked.
      const path = f.webkitRelativePath || f.name;
      const text = core.decodeHtmlFile(new Uint8Array(await f.arrayBuffer()));
      const d = new DOMParser().parseFromString(text, 'text/html');
      const body = d.body ? d.body.innerHTML : text;
      const name = f.name.replace(/\.html?$/i, '');
      const fields = core.placeholdersIn(body);
      const blocks = /\{##[^}]+\}/.test(text);
      const srcs = relativeSrcs(d);
      const perPerson = srcs.filter((x) => x.indexOf('@@') >= 0);
      const imgs = srcs.filter((x) => x.indexOf('@@') < 0).map((x) => decodeURIComponent(x.split(/[\\/]/).pop()));
      const missing = imgs.filter((x) => !impImages.has(x.toLowerCase()));
      const looksLikeSig = fields.length > 0 && !blocks && !/\\(editor|ci-sign-update|AutoBackup-[^\\/]*|Templates)[\\/]/i.test(path.replace(/\//g, '\\'));
      impItems.push({ file: f, path, name, body, fields, blocks, perPerson, imgs, missing, checked: looksLikeSig });
    }
    impItems.sort((a, b) => (b.checked - a.checked) || a.name.localeCompare(b.name));
    // The same template twice (CI-Sign keeps a copy next to the program): the first one is ticked.
    const seen = new Set();
    impItems.forEach((it) => { const k = it.name.toLowerCase(); if (it.checked && seen.has(k)) it.checked = false; if (it.checked) seen.add(k); });
    renderImportList();
  }

  function renderImportList() {
    const ul = $('impList'); ul.textContent = '';
    impItems.forEach((it) => {
      const box = el('input', { type: 'checkbox' }); box.checked = it.checked;
      box.addEventListener('change', () => { it.checked = box.checked; $('impGo').disabled = !impItems.some((x) => x.checked); });
      const exists = S.config.signatures.some((x) => x.name.toLowerCase() === it.name.toLowerCase());
      const meta = [];
      if (it.blocks) meta.push(el('span', { class: 'warn' }, t('CI-Sign text blocks ({##…}), not a signature. ')));
      meta.push(it.fields.length ? t('{0} fields', it.fields.length) : t('no fields'));
      meta.push(t(' · {0} image(s)', it.imgs.length));
      if (it.missing.length) meta.push(el('span', { class: 'warn' }, t(' · not picked: {0}', [...new Set(it.missing)].join(', '))));
      if (it.perPerson.length) meta.push(t(' · photo per person {0} → photos\\{1}', it.perPerson[0], it.perPerson[0].split(/[\\/]/).pop()));
      if (exists) meta.push(t(' · replaces the signature of the same name'));
      ul.append(el('li', { class: it.checked ? '' : 'skip' }, box, el('div', null, el('div', { class: 'imp-name' }, it.name), el('div', { class: 'imp-meta' }, el('span', null, it.path), el('br'), ...meta))));
    });
    const n = impItems.length;
    $('impNote').textContent = n ? t('{0} template file(s) found, {1} image(s) picked.', n, impImages.size) : t('No .htm files among what was picked.');
    $('impGo').disabled = !impItems.some((x) => x.checked);
  }

  async function uploadImage(file) {
    const data = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = rej; r.readAsDataURL(file); });
    const r = await api('assets/upload', { name: file.name, data });
    S.assets = S.assets.filter((a) => a.name !== r.name).concat([{ name: r.name, size: r.size }]).sort((a, b) => a.name.localeCompare(b.name));
    return r.name;
  }

  async function runImport() {
    $('impErr').textContent = ''; $('impGo').disabled = true;
    const uploaded = new Map(); // picked file name (lower case) -> asset name
    let first = null, count = 0;
    try {
      for (const it of impItems.filter((x) => x.checked)) {
        const t = document.createElement('template');
        t.innerHTML = sanitize(it.body);
        for (const img of t.content.querySelectorAll('img[src]')) {
          const src = img.getAttribute('src');
          if (/^(https?:|cid:|data:)/i.test(src)) continue;
          // One image per person ("images/@@sAMAccountName.jpg"): from the photos folder.
          if (src.indexOf('@@') >= 0) { img.setAttribute('src', 'photos/' + src.split(/[\\/]/).pop()); continue; }
          const base = decodeURIComponent(src.split(/[\\/]/).pop()), key = base.toLowerCase();
          if (impImages.has(key) && !uploaded.has(key)) uploaded.set(key, await uploadImage(impImages.get(key)));
          // Not picked: it points where the image will be once uploaded under that name.
          img.setAttribute('src', 'assets/' + (uploaded.get(key) || core.assetName(base)));
        }
        let s = S.config.signatures.find((x) => x.name.toLowerCase() === it.name.toLowerCase());
        if (!s) { s = core.newSignature(newId(it.name), it.name); S.config.signatures.push(s); }
        s.html = t.innerHTML.trim();
        first = first || s; count++;
      }
      // With no default yet, the one called Standard (or the first) becomes it.
      if (first && !S.config.signatures.some((x) => x.defaultNew)) {
        const std = S.config.signatures.find((x) => /^standard$/i.test(x.name)) || first;
        std.defaultNew = true; std.defaultReply = true;
      }
      changed();
      $('dlgImport').close();
      if (first) openSig(first.id);
      toast(uploaded.size ? t('{0} template(s) imported, with {1} image(s).', count, uploaded.size) : t('{0} template(s) imported.', count));
    } catch (e) { $('impErr').textContent = t(e.message); $('impGo').disabled = false; }
  }

  $('btnImportOpen').addEventListener('click', () => {
    impItems = []; impImages = new Map(); renderImportList(); $('impNote').textContent = ''; $('impErr').textContent = '';
    $('dlgImport').showModal();
  });
  ['impFiles', 'impFolder'].forEach((id) => $(id).addEventListener('change', async () => {
    const files = $(id).files; if (!files || !files.length) return;
    try { await readImportFiles(files); } catch (e) { $('impErr').textContent = t(e.message); }
    $(id).value = '';
  }));
  $('impGo').addEventListener('click', (e) => { e.preventDefault(); runImport(); });

  // ---------------------------------------------------------------------------------------------
  // Fields
  // ---------------------------------------------------------------------------------------------
  const GRAPH_PROPS = ['givenName', 'surname', 'displayName', 'jobTitle', 'department', 'companyName', 'officeLocation', 'mobilePhone', 'businessPhones[0]', 'faxNumber', 'mail', 'userPrincipalName', 'streetAddress', 'postalCode', 'city', 'state', 'country', 'employeeId']
    .concat(Array.from({ length: 15 }, (_, i) => 'onPremisesExtensionAttributes.extensionAttribute' + (i + 1)));
  $('graphProps').append(...GRAPH_PROPS.map((p) => el('option', { value: p })));

  function renderFields() {
    const tb = $('fieldRows'); tb.textContent = '';
    S.config.fields.forEach((f, i) => {
      const inp = (prop, attrs) => {
        const x = el('input', Object.assign({ value: f[prop] || '' }, attrs || {}));
        x.addEventListener('input', () => { f[prop] = x.value.trim(); changed({ list: false }); if (prop === 'graph') x.closest('td').classList.toggle('nograph', !f.graph); });
        return x;
      };
      const key = el('input', { value: f.key, 'aria-label': t('Placeholder name') });
      let oldKey = f.key;
      key.addEventListener('change', () => {
        const nk = key.value.trim().replace(/^@@/, '');
        if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(nk)) { toast(t('A placeholder name is letters, digits and _, starting with a letter.'), true); key.value = oldKey; return; }
        if (S.config.fields.some((x) => x !== f && x.key === nk)) { toast(t('@@{0} is already a field.', nk), true); key.value = oldKey; return; }
        const uses = S.config.signatures.filter((s) => core.placeholdersIn(s.html).indexOf(oldKey) >= 0);
        if (uses.length && confirm(t('Change @@{0} to @@{1} in {2} signature(s) too?', oldKey, nk, uses.length))) {
          const re = new RegExp('@@' + oldKey + '(?![A-Za-z0-9_])', 'g');
          uses.forEach((s) => { s.html = s.html.replace(re, '@@' + nk); });
        }
        f.key = nk; oldKey = nk; changed({ list: false });
      });
      const del = el('button', { class: 'ghost small danger', title: t('Remove field'), 'aria-label': t('Remove field'), onclick: () => {
        const uses = S.config.signatures.filter((s) => core.placeholdersIn(s.html).indexOf(f.key) >= 0).map((s) => s.name);
        if (uses.length && !confirm(t('@@{0} is used in: {1}. Remove the field anyway? The placeholder then stays as typed.', f.key, uses.join(', ')))) return;
        S.config.fields.splice(i, 1); changed({ list: false }); renderFields();
      } }, '✕');
      tb.append(el('tr', null,
        el('td', null, inp('label', { 'aria-label': t('Field') })),
        el('td', null, el('div', { class: 'ph' }, el('span', null, '@@'), key)),
        el('td', null, inp('ad', { 'aria-label': t('AD attribute') })),
        el('td', { class: f.graph ? '' : 'nograph' }, inp('graph', { list: 'graphProps', placeholder: t('not available'), 'aria-label': t('Graph source') })),
        el('td', null, inp('sample', { 'aria-label': t('Sample value') })),
        el('td', null, del)));
    });
  }
  $('btnAddField').addEventListener('click', () => {
    let n = 1; while (S.config.fields.some((f) => f.key === 'field' + n)) n++;
    S.config.fields.push({ key: 'field' + n, label: t('New field'), ad: '', graph: '', sample: '' });
    changed({ list: false }); renderFields();
    const rows = $('fieldRows').querySelectorAll('tr'); rows[rows.length - 1].querySelector('input').focus();
  });

  // ---------------------------------------------------------------------------------------------
  // Add-in options
  // ---------------------------------------------------------------------------------------------
  const OPTS = { optDisable: 'disableOutlookSignature', optChoice: 'allowUserChoice', optMeetings: 'meetings', optNoProfile: 'insertWithoutProfile' };
  function renderOptions() {
    Object.entries(OPTS).forEach(([id, k]) => { $(id).checked = !!S.config.options[k]; });
    const mode = core.valueSource(S.config);
    document.querySelectorAll('input[name=vsrc]').forEach((r) => { r.checked = r.value === mode; });
    $('vsrcNote').hidden = mode === 'graph';
  }
  document.querySelectorAll('input[name=vsrc]').forEach((r) => r.addEventListener('change', () => {
    if (!r.checked) return;
    S.config.options.valueSource = r.value; changed({ list: false }); renderOptions();
  }));
  Object.entries(OPTS).forEach(([id, k]) => $(id).addEventListener('change', () => { S.config.options[k] = $(id).checked; changed({ list: false }); }));
  $('lnkSetup').addEventListener('click', (e) => { e.preventDefault(); $('dlgHelp').showModal(); });

  document.querySelectorAll('.nav-list a').forEach((a) => a.addEventListener('click', (e) => {
    e.preventDefault(); flushSource();
    show(a.dataset.view);
    if (a.dataset.view === 'fields') renderFields(); else if (a.dataset.view === 'users') renderUsers(); else renderOptions();
  }));

  // ---------------------------------------------------------------------------------------------
  // Users from AD: read, look at, and only then publish
  // ---------------------------------------------------------------------------------------------
  const adRequest = (c) => ({ fields: core.adFields(c).map((f) => ({ key: f.key, ad: f.ad })), groups: core.groupsUsed(c) });
  S.snap = null; S.snapMatches = false;
  async function loadSnapshot() {
    try { const r = await api('ad/snapshot', adRequest(S.config)); S.snap = r.snapshot; S.snapMatches = !!r.matches; }
    catch (e) { S.snap = null; S.snapMatches = false; }
  }
  const when = (ms) => new Date(ms).toLocaleString(I18N.lang);

  function renderUsers() {
    const mode = core.valueSource(S.config);
    const fields = core.adFields(S.config);
    const off = mode === 'graph' || !fields.length;
    $('usersOff').hidden = !off;
    $('usersOff').textContent = mode === 'graph'
      ? t('The values come from Entra ID (Add-in → Where the users\' values come from): there is nothing to read from AD.')
      : t('No field takes its value from AD (Fields: an AD attribute, and no Graph source in the mixed mode).');
    $('btnAdRead').disabled = off;
    const snap = S.snap;
    const head = $('usersHead'), body = $('usersRows');
    head.textContent = ''; body.textContent = ''; $('usersMore').textContent = '';
    if (off || !snap) { $('usersInfo').textContent = off ? '' : t('Not read yet.'); return; }
    const c = snap.counts;
    $('usersInfo').textContent = t('Read {0}: {1} users - {2} new, {3} changed, {4} removed.', when(snap.readAt), snap.users.length, c.new, c.changed, c.removed) +
      (!S.snapMatches ? ' ' + t('The fields or groups changed since: read again.') : snap.written ? ' ' + t('Published already.') : ' ' + t('Not published yet.'));
    const showGroups = core.groupsUsed(S.config).length > 0;
    head.append(el('tr', null, el('th', null, ''), el('th', null, t('Mail')), fields.map((f) => el('th', null, (f.label || f.key))), showGroups ? el('th', null, t('Groups')) : null));
    const q = $('usersFilter').value.trim().toLowerCase();
    const only = $('usersChanges').checked;
    const rows = snap.users.filter((u) => (!only || u.status !== 'unchanged') &&
      (!q || u.mail.toLowerCase().includes(q) || Object.values(u.values || {}).some((v) => String(v).toLowerCase().includes(q))));
    const LIMIT = 500;
    rows.slice(0, LIMIT).forEach((u) => {
      body.append(el('tr', null,
        el('td', null, el('span', { class: 'st ' + u.status }, t(u.status))),
        el('td', null, u.mail),
        fields.map((f) => {
          const v = (u.values || {})[f.key] || '', was = u.old ? (u.old[f.key] || '') : null;
          const chg = u.status === 'changed' && was !== null && was !== v;
          return el('td', { class: chg ? 'chg' : '' }, v, chg ? el('span', { class: 'was' }, was || t('(empty)')) : null);
        }),
        showGroups ? el('td', null, (u.groups || []).join(', ')) : null));
    });
    snap.removed.filter((m) => !q || m.includes(q)).forEach((m) => {
      body.append(el('tr', { class: 'removed' }, el('td', null, el('span', { class: 'st removed' }, t('removed'))), el('td', { colspan: fields.length + (showGroups ? 2 : 1) }, m)));
    });
    if (rows.length > LIMIT) $('usersMore').textContent = t('{0} more - narrow the search to see them.', rows.length - LIMIT);
    if (!rows.length && !snap.removed.length) body.append(el('tr', null, el('td', { colspan: fields.length + 3 }, only ? t('No changes: everything is as published.') : t('No users found.'))));
  }
  $('btnAdRead').addEventListener('click', async () => {
    $('btnAdRead').disabled = true; $('usersInfo').textContent = t('Reading…');
    try { await saving; const r = await api('ad/read', adRequest(prepared())); S.snap = r; S.snapMatches = true; }
    catch (e) { toast(t(e.message), true); }
    finally { $('btnAdRead').disabled = false; renderUsers(); }
  });
  $('usersChanges').addEventListener('change', renderUsers);
  $('usersFilter').addEventListener('input', renderUsers);

  // ---------------------------------------------------------------------------------------------
  // Settings
  // ---------------------------------------------------------------------------------------------
  function renderSettingsNotes(r) {
    $('setFolderNote').textContent = !S.settings.folder ? t('Not set yet.') : r.folderFound ? t('Found.') : t('Not found on this PC: is the library synced?');
    $('setPhotosNote').textContent = r.photosFound ? t('Found.') : t('No photos folder there (yet): the preview shows no photos.');
  }
  let lastSettings = {};
  const ADDIN_FIELDS = [['setAddinUrl', 'addinUrl'], ['setAddinFolder', 'addinFolder'], ['setAddinHost', 'addinHost'], ['setClientId', 'clientId'], ['setTenantId', 'tenantId'], ['setSpSite', 'spSite'], ['setSpFolder', 'spFolder']];
  function settingsFromDialog() {
    const s = { folder: $('setFolder').value, photosFolder: $('setPhotos').value };
    for (const [id, k] of ADDIN_FIELDS) s[k] = $(id).value;
    return s;
  }
  $('btnSettings').addEventListener('click', () => {
    $('setFolder').value = S.settings.folder || '';
    $('setPhotos').value = S.settings.photosFolder || '';
    for (const [id, k] of ADDIN_FIELDS) $(id).value = S.settings[k] || '';
    $('setAddinBuild').disabled = !(S.info.features && S.info.features.addinFiles);
    $('setAddinOk').textContent = $('setAddinBuild').disabled ? t('The add-in files are not next to OSign.ps1: use the release zip.') : '';
    $('setErr').textContent = '';
    renderSettingsNotes(lastSettings);
    $('setWhere').textContent = '';
    $('dlgSettings').showModal();
  });
  $('setSave').addEventListener('click', async (e) => {
    e.preventDefault();
    try {
      const r = await api('settings/save', { settings: settingsFromDialog() });
      // Another language: the page starts again in it (the draft is saved, nothing is lost).
      if ($('setLang').value !== I18N.lang) { await saving; I18N.setLang($('setLang').value); return; }
      S.settings = r.settings; lastSettings = r;
      $('dlgSettings').close(); toast(r.folderFound || !r.settings.folder ? t('Settings saved.') : t('Saved, but the folder is not there on this PC.'), !r.folderFound && !!r.settings.folder);
      schedulePreview();
    } catch (err) { $('setErr').textContent = t(err.message); }
  });

  $('setAddinBuild').addEventListener('click', async () => {
    $('setErr').textContent = ''; $('setAddinOk').textContent = '';
    try {
      const saved = await api('settings/save', { settings: settingsFromDialog() });
      S.settings = saved.settings; lastSettings = saved;
      const r = await api('addin/build');
      $('setAddinOk').textContent = t('{0} files (version {1}) in {2}. Copy them to the web server, add the redirect URI {3} to the app registration, and upload {4} in the admin center.', r.files, r.version, r.folder, r.redirectUri, r.manifestUrl);
    } catch (err) { $('setErr').textContent = t(err.message); }
  });

  // A site's or OneDrive's address as copied from the browser, or a sharing link (Share → Copy
  // link), becomes what Graph wants, e.g. contoso-my.sharepoint.com:/personal/me_contoso_com:.
  // Where the address also names the folder (…?id=/personal/me/Documents/OSign, or a path in the
  // library), that comes back too: the path within the default library ("OSign"); else null.
  //   https://contoso.sharepoint.com/sites/IT/Shared Documents/Forms/AllItems.aspx?id=…
  //   https://contoso-my.sharepoint.com/personal/me_contoso_com/_layouts/15/onedrive.aspx?id=…
  //   https://contoso-my.sharepoint.com/:f:/g/personal/me_contoso_com/EaBc…?e=…
  //   https://contoso.sharepoint.com/:f:/s/IT/EaBc…  (/s/ a site, /t/ a team, /p/ a OneDrive)
  const LIBRARIES = ['shared documents', 'documents', 'freigegebene dokumente', 'dokumente'];
  function parseSpAddress(v) {
    let u;
    try { u = new URL(String(v).trim()); } catch (e) { return { site: String(v).trim(), folder: null }; }
    const host = u.hostname.toLowerCase();
    if (u.protocol !== 'https:' || !/^[a-z0-9-]+(-my)?\.sharepoint\.com$/.test(host)) return { site: String(v).trim(), folder: null };
    let parts = u.pathname.split('/').filter(Boolean).map((x) => { try { return decodeURIComponent(x); } catch (e) { return x; } });
    // A sharing link: /:f:/<kind>/…, with the site after it; its last part is a code, not a folder.
    let shared = false;
    if (parts.length && /^:[a-z]:$/i.test(parts[0])) {
      parts = parts.slice(1);
      const k = (parts[0] || '').toLowerCase();
      // /r/ links carry the real path, folder included; the others end in a code.
      shared = k !== 'r';
      if (k === 'r' || k === 'g') parts = parts.slice(1);
      else if (k === 's' || k === 't' || k === 'p') parts = [{ s: 'sites', t: 'teams', p: 'personal' }[k]].concat(parts.slice(1));
    }
    let site = 'root', rest = parts;
    if (parts.length >= 2 && /^(sites|teams|personal)$/i.test(parts[0])) {
      site = host + ':/' + parts[0].toLowerCase() + '/' + parts[1] + ':';
      rest = parts.slice(2);
    }
    // The folder: from ?id= (the server-relative path), else from the path itself.
    const sitePath = site === 'root' ? [] : parts.slice(0, 2).map((x) => x.toLowerCase());
    const id = u.searchParams.get('id') || u.searchParams.get('RootFolder');
    let path = id ? id.split('/').filter(Boolean) : (shared ? [] : rest);
    if (id && sitePath.length && path.slice(0, 2).map((x) => x.toLowerCase()).join('/') === sitePath.join('/')) path = path.slice(2);
    let folder = null;
    const lib = path.findIndex((x) => LIBRARIES.includes(x.toLowerCase()));
    if (lib === 0) {
      const inLib = path.slice(1);
      const forms = inLib.findIndex((x) => /^forms$/i.test(x) || /\.aspx$/i.test(x));
      const f = (forms < 0 ? inLib : inLib.slice(0, forms)).join('/');
      if (f) folder = f;
    }
    return { site, folder };
  }
  [['setSpSite', 'setSpFolder'], ['wizSpSite', 'wizSpFolder']].forEach(([siteId, folderId]) => $(siteId).addEventListener('change', () => {
    const r = parseSpAddress($(siteId).value);
    $(siteId).value = r.site;
    if (r.folder) $(folderId).value = r.folder;
  }));

  // ---------------------------------------------------------------------------------------------
  // Help: a ? next to a label opens a short explanation (the texts are in #helpTexts, so they are
  // translated with the page). A popover, so it shows above everything; a click elsewhere or Esc
  // closes it, a second click on the same ? as well.
  // ---------------------------------------------------------------------------------------------
  const helpPop = $('helpPop');
  let helpFor = null, helpClosedAt = 0;
  // beforetoggle, not toggle: it comes at once, before the click that closed it.
  helpPop.addEventListener('beforetoggle', (e) => { if (e.newState === 'closed') helpClosedAt = Date.now(); });
  document.addEventListener('click', (e) => {
    const b = e.target.closest && e.target.closest('button.help');
    if (!b) return;
    e.preventDefault(); e.stopPropagation();
    // The click that closed it (outside the popover, on this same ?) does not open it again.
    if (helpFor === b && Date.now() - helpClosedAt < 400) { helpFor = null; return; }
    const src = document.querySelector('[data-help-text="' + b.dataset.help + '"]');
    if (!src) return;
    if (helpPop.matches(':popover-open')) helpPop.hidePopover();
    // Inside the open dialog, if the ? is in one: outside it, an open modal dialog makes it inert.
    const home = b.closest('dialog') || document.body;
    if (helpPop.parentNode !== home) home.append(helpPop);
    helpPop.innerHTML = src.innerHTML;
    helpPop.showPopover();
    helpFor = b;
    const r = b.getBoundingClientRect(), w = Math.min(380, window.innerWidth - 24);
    helpPop.style.width = w + 'px';
    helpPop.style.left = Math.max(12, Math.min(r.left - 12, window.innerWidth - w - 12)) + 'px';
    const h = helpPop.offsetHeight;
    helpPop.style.top = (r.bottom + 8 + h < window.innerHeight ? r.bottom + 8 : Math.max(12, r.top - 8 - h)) + 'px';
  }, true);

  // ---------------------------------------------------------------------------------------------
  // Folders: the Windows folder dialog behind each "Choose folder…", and where a synced folder is
  // in SharePoint, found out from OneDrive's own notes (sp/resolve) and said in plain words.
  // ---------------------------------------------------------------------------------------------
  document.querySelectorAll('button.pick-btn').forEach((b) => b.addEventListener('click', async () => {
    const input = $(b.dataset.for);
    b.disabled = true;
    try {
      const r = await api('pick-folder', { start: input.value.trim(), title: t('Choose the folder') });
      if (r.path) { input.value = r.path; input.dispatchEvent(new Event('input')); input.dispatchEvent(new Event('change')); }
    } catch (err) { toast(t(err.message), true); }
    finally { b.disabled = false; }
  }));
  function describeWhere(loc) {
    const site = loc.site === 'root' ? t('Home site') : /:\/personal\//.test(loc.site) ? 'OneDrive' : loc.siteName;
    return [site, loc.defaultLibrary ? t('Documents') : loc.library].concat(loc.folder ? loc.folder.split('/') : []).join(' › ');
  }
  // Fills the SharePoint site and folder fields from where the folder at `path` is synced from,
  // and says so in `out`; when it cannot tell, opens `details` (the fields to set by hand).
  async function detectWhere(path, siteId, folderId, out, details) {
    out.textContent = ''; out.className = 'where';
    if (!String(path || '').trim()) return null;
    let r;
    try { r = await api('sp/resolve', { path }); } catch (e) { return null; }
    if (!r.found) {
      out.textContent = t('OSign cannot tell where this folder is in SharePoint. Is it connected with OneDrive (Add shortcut to My files, or Sync)? Otherwise set the place by hand.');
      out.classList.add('warn');
      if (details) details.open = true;
      return null;
    }
    const loc = r.location;
    $(siteId).value = loc.site; $(folderId).value = loc.folder;
    if (!loc.defaultLibrary) { out.textContent = t('In SharePoint: {0}. The add-in reads only the main library (Documents): move the folder there.', describeWhere(loc)); out.classList.add('warn'); }
    else if (!loc.folder) { out.textContent = t('This is a whole library, not a folder in it: make a folder OSign in it and choose that.'); out.classList.add('warn'); }
    else { out.textContent = '✓ ' + t('In SharePoint: {0}', describeWhere(loc)); out.classList.add('good'); }
    return loc;
  }
  $('setFolder').addEventListener('change', () => detectWhere($('setFolder').value, 'setSpSite', 'setSpFolder', $('setWhere'), null));
  $('wizFolder').addEventListener('change', () => detectWhere($('wizFolder').value, 'wizSpSite', 'wizSpFolder', $('wizWhereFound'), $('wizWhereEdit')));

  // ---------------------------------------------------------------------------------------------
  // Setup assistant: opens by itself until it was finished once, while OSign is not set up
  // (no publish folder or no app registration). Each step saves what it asks for.
  // ---------------------------------------------------------------------------------------------
  const WIZ = Array.from(document.querySelectorAll('#dlgWizard .wiz-step')).map((sec) => sec.dataset.step);
  const WIZ_KEY = 'osign.wizStep';
  const wiz = { step: 0, built: false, warned: '' };
  const stored = (k, v) => { try { if (v === undefined) return localStorage.getItem(k); if (v === null) localStorage.removeItem(k); else localStorage.setItem(k, v); } catch (e) { /* no storage */ } return null; };
  const radio = (name) => (document.querySelector('input[name=' + name + ']:checked') || {}).value || '';
  const setRadio = (name, v) => document.querySelectorAll('input[name=' + name + ']').forEach((r) => { r.checked = r.value === v; });
  // The address as typed, if it is one: https://host/path/
  const addinBase = () => {
    const u = $('wizUrl').value.trim().replace(/\/+$/, '') + '/';
    try { return /^https:\/\//.test(u) && new URL(u).host ? u : ''; } catch (e) { return ''; }
  };
  const URL_HINTS = { cloudflare: '(e.g. https://company-osign.pages.dev/ or https://osign.company.ch/)', azure: '(e.g. https://<name>.azurestaticapps.net/ or https://osign.company.ch/)', apache: '(e.g. https://osign.company.ch/)', iis: '(e.g. https://osign.company.ch/)', other: '(https://…/)' };

  function wizFill() {
    const st = S.settings;
    $('wizFolder').value = st.folder || ''; $('wizPhotos').value = st.photosFolder || '';
    // Nothing made yet: the recommended one; otherwise what is set.
    setRadio('wvsrc', S.config.signatures.length || S.settings.setupDone === '1' ? core.valueSource(S.config) : 'mixed');
    setRadio('whost', st.addinHost || 'cloudflare');
    $('wizUrl').value = st.addinUrl || '';
    $('wizClientId').value = st.clientId || ''; $('wizTenantId').value = st.tenantId || '';
    $('wizSpSite').value = st.spSite || 'root'; $('wizSpFolder').value = st.spFolder || 'OSign';
    $('wizAddinFolder').value = st.addinFolder || (S.info.backend === 'powershell' && navigator.platform.startsWith('Win') ? 'C:\\OSign-addin' : '');
  }
  function wizRender() {
    const name = WIZ[wiz.step];
    document.querySelectorAll('#dlgWizard .wiz-step').forEach((sec) => { sec.hidden = sec.dataset.step !== name; });
    const sec = document.querySelector('#dlgWizard .wiz-step[data-step=' + name + ']');
    $('wizTitle').textContent = t(sec.dataset.title);
    $('wizCount').textContent = wiz.step && wiz.step < WIZ.length - 1 ? t('Step {0} of {1}', wiz.step, WIZ.length - 2) : '';
    $('wizBar').style.width = Math.round(100 * wiz.step / (WIZ.length - 1)) + '%';
    $('wizBack').hidden = wiz.step === 0;
    $('wizNext').textContent = name === 'welcome' ? t('Start') : name === 'done' ? t('Finish') : t('Next');
    $('wizLater').hidden = name === 'done';
    $('wizErr').textContent = '';
    const host = radio('whost'), base = addinBase();
    $('wizUrlHint').textContent = t(URL_HINTS[host] || URL_HINTS.other);
    $('wizRedirect').textContent = base ? 'brk-multihub://' + new URL(base).host : t('(enter the address in the step before)');
    $('wizManifest').textContent = base ? base + 'manifest.xml' : '';
    const perms = ['User.Read', 'Sites.Read.All'];
    if (radio('wvsrc') === 'graph') perms.push('GroupMember.Read.All');
    $('wizPerms').textContent = perms.join(', ');
    $('wizUpload').hidden = !wiz.built;
    document.querySelectorAll('#wizUpload [data-host]').forEach((p) => { p.hidden = p.dataset.host !== (host || 'other'); });
  }
  function wizOpen(step) {
    wizFill();
    wiz.step = Math.max(0, Math.min(WIZ.length - 1, step || 0)); wiz.built = false; wiz.warned = '';
    $('wizMakeOk').textContent = ''; $('wizChecks').replaceChildren();
    $('wizWhereEdit').open = false; $('wizWhereFound').textContent = '';
    if ($('wizFolder').value) detectWhere($('wizFolder').value, 'wizSpSite', 'wizSpFolder', $('wizWhereFound'), null);
    wizRender();
    if (!$('dlgWizard').open) $('dlgWizard').showModal();
  }
  async function wizSave(settings) {
    const r = await api('settings/save', { settings });
    S.settings = r.settings; lastSettings = r;
    return r;
  }
  // What a step does on Next; throws to stay on it.
  const wizLeave = {
    async folder() {
      if (!$('wizFolder').value.trim()) throw new Error(t('Enter the path of the synced OSign folder.'));
      const where = parseSpAddress($('wizSpSite').value);
      $('wizSpSite').value = where.site;
      if (where.folder) $('wizSpFolder').value = where.folder;
      const r = await wizSave({ folder: $('wizFolder').value, photosFolder: $('wizPhotos').value, spSite: $('wizSpSite').value || 'root', spFolder: $('wizSpFolder').value || 'OSign' });
      if (!r.folderFound && wiz.warned !== 'folder') {
        wiz.warned = 'folder';
        $('wizFolderNote').textContent = t('Not found on this PC: is the library synced? Next again to go on anyway.');
        throw new Error('');
      }
      schedulePreview();
    },
    async values() {
      const v = radio('wvsrc') || 'graph';
      if (core.valueSource(S.config) !== v) { S.config.options.valueSource = v; changed({ list: false }); renderOptions(); }
    },
    async host() {
      if (!$('wizUrl').value.trim()) throw new Error(t('Enter the address the add-in will be served from.'));
      await wizSave({ addinHost: radio('whost'), addinUrl: $('wizUrl').value });
    },
    async app() {
      if (!$('wizClientId').value.trim()) throw new Error(t('Enter the app registration\'s client ID: without it the add-in cannot sign in.'));
      await wizSave({ clientId: $('wizClientId').value, tenantId: $('wizTenantId').value });
    },
    async files() {
      if (!wiz.built) throw new Error(t('Make the add-in files first.'));
    },
    async done() {
      await wizSave({ setupDone: '1' });
    }
  };
  $('wizNext').addEventListener('click', async () => {
    $('wizErr').textContent = '';
    const name = WIZ[wiz.step];
    $('wizNext').disabled = true;
    try {
      if (wizLeave[name]) await wizLeave[name]();
      if (name === 'done') { $('dlgWizard').close(); toast(t('OSign is set up.')); return; }
      wiz.step++; wiz.warned = ''; wizRender();
    } catch (err) { $('wizErr').textContent = err.message ? t(err.message) : ''; }
    finally { $('wizNext').disabled = false; }
  });
  $('wizBack').addEventListener('click', () => { if (wiz.step) { wiz.step--; wizRender(); } });
  $('wizLater').addEventListener('click', () => $('dlgWizard').close());
  // Closed in any way (Later, Esc, finished): it opens by itself again at the next start only if
  // OSign is still not set up.
  $('dlgWizard').addEventListener('close', () => stored(WIZ_KEY, null));
  document.querySelectorAll('input[name=whost], input[name=wvsrc]').forEach((r) => r.addEventListener('change', wizRender));
  $('wizUrl').addEventListener('input', wizRender);
  $('wizFolder').addEventListener('input', () => { $('wizFolderNote').textContent = ''; wiz.warned = ''; });
  $('wizMake').addEventListener('click', async () => {
    $('wizErr').textContent = ''; $('wizMakeOk').textContent = '';
    try {
      await wizSave({ addinFolder: $('wizAddinFolder').value });
      const r = await api('addin/build');
      wiz.built = true;
      $('wizMakeOk').textContent = t('{0} files (version {1}) in {2}.', r.files, r.version, r.folder);
      wizRender();
    } catch (err) { $('wizErr').textContent = t(err.message); }
  });
  $('wizCheck').addEventListener('click', async () => {
    $('wizErr').textContent = '';
    const ul = $('wizChecks');
    ul.replaceChildren(el('li', null, t('Checking…')));
    try {
      const r = await api('addin/check');
      ul.replaceChildren(...r.checks.map((c) => el('li', { class: c.ok ? 'good' : c.what === 'Cache-Control' ? 'warn' : 'bad' }, c.what + ': ' + t(c.detail))));
    } catch (err) { ul.replaceChildren(); $('wizErr').textContent = t(err.message); }
  });
  document.querySelectorAll('#dlgWizard button.copy').forEach((b) => b.addEventListener('click', async () => {
    try { await navigator.clipboard.writeText($(b.dataset.copy).textContent); b.textContent = t('Copied'); setTimeout(() => { b.textContent = t('Copy'); }, 1500); } catch (e) { /* select it by hand */ }
  }));
  // Another language: the page starts again in it, and the assistant with it, on this step.
  $('wizLang').addEventListener('change', async () => { await saving; stored(WIZ_KEY, String(wiz.step)); I18N.setLang($('wizLang').value); });
  $('wizImport').addEventListener('click', async () => { await wizLeave.done(); $('dlgWizard').close(); $('btnImportOpen').click(); });
  $('wizNewSig').addEventListener('click', async () => { await wizLeave.done(); $('dlgWizard').close(); $('btnNew').click(); });
  $('btnWizard').addEventListener('click', () => { $('dlgSettings').close(); wizOpen(0); });

  // ---------------------------------------------------------------------------------------------
  // Publish and load
  // ---------------------------------------------------------------------------------------------
  function publishChecks() {
    const c = S.config, out = [];
    const errs = core.validate(prepared());
    errs.forEach((e) => out.push(['bad', t(e)]));
    if (!c.signatures.length) out.push(['bad', t('There is no signature.')]);
    const plain = c.signatures.filter((s) => !(s.audience.departments || []).length && !(s.audience.domains || []).length && !(s.audience.groups || []).length);
    if (!c.signatures.some((s) => s.defaultNew)) out.push(['warn', t('No signature is the default for new mail: users get none unless they pick one.')]);
    else if (!plain.some((s) => s.defaultNew)) out.push(['warn', t('The default for new mail is limited to some users: everyone else gets none.')]);
    if (!c.signatures.some((s) => s.defaultReply)) out.push(['warn', t('No signature is the default for replies.')]);
    if (!S.settings.folder) out.push(['bad', t('No folder to publish to: set the synced OSign folder in Settings.')]);
    if (core.groupsUsed(c).length && core.valueSource(c) === 'graph') out.push(['warn', t('Signatures limited to AD groups: the app registration needs GroupMember.Read.All (delegated, with admin consent), and the groups have to be synced to Entra ID.')]);
    const perPerson = c.signatures.filter((s) => /<img[^>]+src="photos\//i.test(s.html || '')).map((s) => s.name);
    if (perPerson.length) out.push(['good', t('Photos per person ({0}) come from photos\\<username>.jpg in the OSign folder; who has none gets the signature without.', perPerson.join(', '))]);
    const lost = c.signatures.filter((s) => /<img[^>]+src="(?!assets\/|photos\/|https:)[^"]*"/i.test(s.html || '')).map((s) => s.name);
    if (lost.length) out.push(['warn', t('Images that are no file of the signatures are left out in: {0}.', lost.join(', '))]);
    const keys = fieldKeys();
    const used = [...new Set(c.signatures.flatMap((s) => core.placeholdersIn(s.html)))];
    const unknown = used.filter((k) => keys.indexOf(k) < 0);
    if (unknown.length) out.push(['warn', t('Not fields, left as typed: {0}', unknown.map((k) => '@@' + k).join(', '))]);
    const noGraph = used.filter((k) => { const f = c.fields.find((x) => x.key === k); return f && core.fieldSource(c, f) === 'none'; });
    if (noGraph.length) out.push(['warn', t('Empty in Outlook (no source): {0}', noGraph.map((k) => '@@' + k).join(', '))]);
    if (core.usesAdExport(c)) {
      const sn = S.snap;
      if (sn && S.snapMatches && !sn.written) out.push(['good', t('The users as read {0} are written: {1} new, {2} changed, {3} removed.', when(sn.readAt), sn.counts.new, sn.counts.changed, sn.counts.removed)]);
      else if (sn && S.snapMatches && sn.written) out.push(['warn', t('The users stay as published {0}. To update them, read and check them under Users first.', when(sn.readAt))]);
      else out.push(['warn', t('The users are not written: read and check them under Users first. Until then they stay as they are in the folder.')]);
    }
    const missing = [...new Set(c.signatures.flatMap((s) => [...(s.html || '').matchAll(/src="assets\/([^"]+)"/g)].map((m) => m[1])))].filter((n) => !S.assets.some((a) => a.name === n));
    if (missing.length) out.push(['warn', t('Images not in the draft, left out until added: {0}', missing.join(', '))]);
    if (!out.some((x) => x[0] !== 'good')) out.push(['good', t('{0} signature(s), ready.', c.signatures.length)]);
    return out;
  }

  $('btnPublish').addEventListener('click', async () => {
    flushSource();
    if (core.usesAdExport(S.config)) { await saving; await loadSnapshot(); }
    $('pubWhere').textContent = S.settings.folder
      ? t('To {0}. OneDrive uploads it to SharePoint; new messages use it a minute or two later.', S.settings.folder)
      : t('No folder set yet (Settings).');
    const checks = publishChecks();
    $('pubChecks').replaceChildren(...checks.map(([k, text]) => el('li', { class: k }, text)));
    $('pubGo').disabled = checks.some((c) => c[0] === 'bad');
    $('pubErr').textContent = ''; $('pubOk').textContent = ''; $('pubMsg').value = '';
    $('dlgPublish').showModal();
  });
  $('pubGo').addEventListener('click', async (e) => {
    e.preventDefault();
    $('pubGo').disabled = true; $('pubErr').textContent = ''; $('pubOk').textContent = t('Publishing…');
    try {
      // A change still waiting for its save goes first, so the draft and what is published match.
      if (saveTimer) save();
      await saving;
      const c = prepared(); c.updated = new Date().toISOString();
      // Values from AD: the users first, so a signature never asks for a file that is not there yet.
      // The users as read and checked under Users (never a fresh read nobody has seen), first, so a
      // signature never asks for a file that is not there yet.
      let exported = '';
      if (core.usesAdExport(c) && S.snap && S.snapMatches && !S.snap.written) {
        try {
          const x = await api('ad/export', adRequest(c));
          exported = ' ' + t('{0} users written ({1} changed, {2} removed).', x.users, x.written, x.removed);
          S.snap.written = true;
        } catch (err) { exported = ' ' + t('The users could not be written: {0}', t(err.message)); }
      }
      const r = await api('publish', { configJson: JSON.stringify(c), message: $('pubMsg').value });
      S.config.updated = c.updated; S.dirty = false; S.publishedAt = Date.now(); renderStatus();
      $('pubOk').textContent = (r.changed ? t('Published to {0}.', r.folder) : t('Nothing had changed.')) + exported;
    } catch (err) { $('pubOk').textContent = ''; $('pubErr').textContent = t(err.message); }
    finally { $('pubGo').disabled = false; }
  });

  $('btnLoad').addEventListener('click', async () => {
    if (S.dirty && !confirm(t('Replace the draft with what is published? Changes not published are lost (the draft is kept once as draft.json.bak).'))) return;
    try { applyDraft(await api('load-published')); toast(t('Loaded the published signatures.')); }
    catch (e) { toast(t(e.message), true); }
  });

  // ---------------------------------------------------------------------------------------------
  // Start
  // ---------------------------------------------------------------------------------------------
  // A configuration from an older editor or another hand gets whatever it lacks, so the rest of
  // the UI can rely on every part being there.
  function normalize(c) {
    const d = core.defaultConfig();
    c = c && typeof c === 'object' ? c : d;
    c.schemaVersion = core.SCHEMA_VERSION;
    c.options = Object.assign({}, d.options, c.options || {});
    delete c.auth; delete c.assetBaseUrl;
    if (!Array.isArray(c.fields)) c.fields = d.fields;
    c.signatures = (Array.isArray(c.signatures) ? c.signatures : []).map((s) => {
      const n = Object.assign(core.newSignature(s.id, s.name), s);
      n.audience = Object.assign({ departments: [], domains: [], groups: [] }, s.audience || {});
      return n;
    });
    return c;
  }

  function applyDraft(r) {
    let c = null;
    try { c = r.configJson ? JSON.parse(r.configJson) : null; } catch (e) { toast(t('The saved draft could not be read; starting from an empty one.'), true); }
    S.config = normalize(c);
    S.assets = r.assets || []; S.dirty = !!r.dirty; S.publishedAt = +r.publishedAt || 0;
    renderStatus();
    if (S.config.signatures.length) openSig(S.sigId && S.config.signatures.some((s) => s.id === S.sigId) ? S.sigId : S.config.signatures[0].id);
    else show('empty');
  }

  let toastTimer = null;
  function toast(msg, isError) {
    const t = $('toast'); t.textContent = msg; t.className = 'toast show' + (isError ? ' error' : '');
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { t.className = 'toast'; }, isError ? 6000 : 3000);
  }

  async function start() {
    I18N.apply(document.body);
    const langSel = $('setLang');
    Object.entries(I18N.langs).forEach(([k, v]) => langSel.append(el('option', { value: k }, v)));
    langSel.value = I18N.lang;
    initDesign();
    try {
      const [info, settings, draft] = await Promise.all([api('info'), api('settings/get'), api('draft/get')]);
      S.info = info; S.settings = settings.settings; lastSettings = settings;
      S.config = normalize(null);
      buildToolbar();
      applyDraft(draft);
      loadSnapshot();
      document.querySelectorAll('button.pick-btn').forEach((b) => { b.hidden = !info.features.pickFolder; });
      const wl = $('wizLang');
      Object.entries(I18N.langs).forEach(([k, v]) => wl.append(el('option', { value: k }, v)));
      wl.value = I18N.lang;
      // A step stored: the assistant was open when the page started again (another language).
      const resume = stored(WIZ_KEY);
      if (resume !== null || (S.settings.setupDone !== '1' && (!S.settings.folder || !S.settings.clientId))) wizOpen(+resume || 0);
      if (!info.features.adLookup) $('pvSource').querySelector('option[value=ad]').textContent = t('User from Active Directory (not reachable)');
    } catch (e) {
      document.body.classList.add('failed');
      $('main').replaceChildren(el('section', { class: 'view empty' }, el('h1', null, t('Not connected')), el('p', null, t(e.message))));
    }
  }
  start();
})();
