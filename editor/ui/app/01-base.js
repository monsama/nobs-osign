/*
 * NOBS OSign editor - the web UI.
 *
 * One script for the browser, kept here as parts by topic: editor/ui/app/01-base.js (the API,
 * the state, saving, the sidebar) … 11-start.js (the start), in the order of their numbers. The
 * server puts them together as /ui/app.js; they share one function scope, opened here and closed
 * at the end of 11-start.js, so each part uses what the others declare.
 *
 * Talks to its server (server/) only through the HTTP API in docs/API.md. Everything about signatures themselves - placeholders, the lines left out when empty,
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
      // Without a start token this is the server version: it is not started by hand.
      if (!serverDown) { serverDown = true; toast(TOKEN ? t('The NOBS OSign server is not running. Start OSign.ps1 again.') : t('The NOBS OSign server cannot be reached.'), true); }
      throw new Error(TOKEN ? t('The NOBS OSign server is not running.') : t('The NOBS OSign server cannot be reached.'));
    }
    serverDown = false;
    // The server version signs in with a page of its own: a session that ended goes back there.
    if (r.status === 401 && !TOKEN) { location.replace('/login'); throw new Error(t('Not signed in.')); }
    let j;
    try { j = await r.json(); } catch (e) { throw new Error(t('The server answered {0}.', r.status)); }
    if (r.status === 401) throw new Error(t('This window is not signed in to the server: open the address NOBS OSign printed.'));
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
  // The server version (server/): publishing goes to SharePoint through Graph, or, without it, into
  // the server's own folder. No synced folder, no Windows.
  const onServer = () => !!(S.info && S.info.backend === 'node');
  const toSharePoint = () => onServer() && S.info.features.publishTarget === 'sharepoint';
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
    histNote();
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
      .then((r) => { if (seq === changeSeq) { S.dirty = r.dirty; S.savedAt = Date.now(); renderStatus(); } })
      .catch((e) => { renderStatus(t('Not saved')); toast(t(e.message), true); });
    return saving;
  }
  window.addEventListener('beforeunload', () => { if (saveTimer) save({ keepalive: true }); });
  setInterval(() => { api('ping').catch(() => {}); }, 60000);

  // Two states, side by side: the draft (saved by itself on the server) and what Outlook gets
  // (only what was published). A change not published yet: one click to Publish….
  function renderStatus(text) {
    const d = $('stDraft'), p = $('stPub');
    const time = (at) => new Date(at).toLocaleTimeString(I18N.lang === 'de' ? 'de-CH' : 'en-GB', { hour: '2-digit', minute: '2-digit' });
    if (text) { d.textContent = text; d.className = 'status' + (text === t('Not saved') ? ' bad' : ' busy'); }
    else { d.textContent = S.savedAt ? t('Draft saved {0}', time(S.savedAt)) : t('Draft saved'); d.className = 'status clean'; }
    d.title = t('Every change is saved on the server by itself (Ctrl+S: at once). Outlook gets it only when you publish.');
    const when = S.publishedAt ? new Date(S.publishedAt).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB') : '';
    if (S.dirty) {
      p.textContent = S.publishedAt ? t('Unpublished changes') : t('Not published yet'); p.className = 'status pub dirty'; p.disabled = false;
      p.title = (S.publishedAt ? t('Outlook still uses what was published {0}.', when) : t('Outlook gets no signatures from here yet.')) + ' ' + t('Click to publish.');
    } else {
      p.textContent = S.publishedAt ? t('Published') : t('Not published yet'); p.className = 'status pub' + (S.publishedAt ? ' clean' : ''); p.disabled = true;
      p.title = S.publishedAt ? t('Last published {0}: Outlook uses the draft as it is.', when) : '';
    }
    // Nothing published yet: nothing to load.
    $('btnLoad').disabled = !S.publishedAt;
  }

  // ---------------------------------------------------------------------------------------------
  // HTML hygiene: what goes into a signature is markup Outlook shows, nothing that runs.
  // ---------------------------------------------------------------------------------------------
  const DROP = 'script,style,meta,link,title,iframe,object,embed,form,input,button,select,textarea,base,xml,o\\:p,svg,video,audio';
  // Parsed in a document of its own (DOMParser), without a window: nothing in it runs or loads.
  // Event handler attributes are taken out of the text first as well, so the browser never even
  // weighs one against the page's Content-Security-Policy; the walk below takes out any left.
  const HANDLER_ATTR = /(<[^>]*?\s)on[a-z]+\s*=\s*("[^"]*"|'[^']*'|[^\s>]+)/gi;
  function stripHandlers(html) {
    let s = String(html || ''), before;
    do { before = s; s = s.replace(HANDLER_ATTR, '$1'); } while (s !== before);
    return s;
  }
  // A <template> in an inert document: what it holds is a fragment, as a signature is.
  function inertTemplate(html) {
    const doc = new DOMParser().parseFromString('<!doctype html><body><template>' + stripHandlers(html).replace(/<\/template/gi, '&lt;/template') + '</template></body>', 'text/html');
    return doc.querySelector('template');
  }
  function sanitize(html) {
    const t = inertTemplate(html);
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

  // A campaign's days, as words; '' when it has none.
  function campaignText(s) {
    if (!core.isCampaign(s)) return '';
    return s.validFrom && s.validUntil ? t('{0} to {1}', day(s.validFrom), day(s.validUntil)) : s.validFrom ? t('from {0}', day(s.validFrom)) : t('until {0}', day(s.validUntil));
  }

  function renderList() {
    const ul = $('sigList'); ul.textContent = '';
    if (!S.config.signatures.length) ul.append(el('li', { class: 'none' }, t('None yet')));
    S.config.signatures.forEach((s) => {
      const a = el('a', { href: '#', class: S.view === 'sig' && s.id === S.sigId ? 'active' : '', 'aria-current': S.view === 'sig' && s.id === S.sigId ? 'page' : null, onclick: (e) => { e.preventDefault(); openSig(s.id); } },
        s.name || t('(no name)'),
        el('div', { class: 'badges' },
          s.defaultNew ? el('span', { class: 'badge on' }, t('New')) : null,
          s.defaultReply ? el('span', { class: 'badge on' }, t('Reply')) : null,
          audienceText(s) !== t('everyone') ? el('span', { class: 'badge', title: audienceText(s) }, t('Limited')) : null,
          core.isCampaign(s) ? el('span', { class: 'badge' + (core.inPeriod(s) ? ' on' : ''), title: campaignText(s) }, core.inPeriod(s) ? t('Campaign') : t('Campaign, not now')) : null));
      a.addEventListener('contextmenu', (e) => { e.preventDefault(); sigMenu(s, e.clientX, e.clientY); });
      a.addEventListener('keydown', (e) => {
        if (e.key === 'ContextMenu' || (e.shiftKey && e.key === 'F10')) { e.preventDefault(); const r = a.getBoundingClientRect(); sigMenu(s, r.left + 24, r.bottom); }
        else if (e.key === 'F2') { e.preventDefault(); openSig(s.id); $('sigName').focus(); $('sigName').select(); }
        else if (e.key === 'Delete') { e.preventDefault(); openSig(s.id); $('btnDel').click(); }
      });
      ul.append(el('li', null, a));
    });
    document.querySelectorAll('.nav-list a').forEach((a) => a.classList.toggle('active', a.dataset.view === S.view));
    renderStatusBar();
  }
  // The status bar: how many signatures, where publishing goes, which version runs where.
  function renderStatusBar() {
    const n = S.config.signatures.length;
    $('stSigs').textContent = n === 1 ? t('1 signature') : t('{0} signatures', n);
    const f = (S.info && S.info.features) || {};
    $('stTarget').textContent = !S.settings ? '' : onServer()
      ? (f.publishTarget === 'sharepoint' ? t('Publishes to {0}', S.settings.folder) : t('Publishes into the server\'s own folder (not SharePoint)'))
      : S.settings.folder ? t('Publishes to {0}', S.settings.folder) : t('No folder to publish to yet');
    $('verFoot').textContent = S.info ? 'NOBS OSign ' + S.info.version + ' · ' + (onServer() ? t('Server') : t('Windows')) : '';
    // The server's client secret, in time: from 30 days before it runs out.
    const left = f.secretDaysLeft;
    $('stSecret').hidden = !(typeof left === 'number' && left <= 30);
    // A newer version, when the server may look; a publish waiting for approval.
    const up = f.update;
    // The add-in's files older than this server (after an update): for the local admin, who makes them.
    const f0s = S.info.features || {}, admin = !!(S.info.user && S.info.user.local);
    const stale = admin && f0s.addinVersion && f0s.addinVersion !== S.info.version;
    $('stAddin').hidden = !stale;
    if (stale) { $('stAddin').textContent = t('Add-in files from {0}: make them again', f0s.addinVersion); $('stAddin').title = t('Settings → Server'); }
    $('stUpdate').hidden = !up;
    if (up) { $('stUpdate').textContent = t('Version {0} is out', up.version); $('stUpdate').href = up.url || '#'; $('stUpdate').title = t('Update with sudo ./scripts/update.sh {0}', up.version); }
    const ap = S.approval && S.approval.pending;
    $('stApproval').hidden = !ap;
    if (ap) $('stApproval').textContent = ap.mine ? t('Waiting for approval') : t('{0} asks for approval', ap.by);
    const cert = f.graphCredential === 'certificate';
    $('stSecret').textContent = typeof left !== 'number' ? '' : left < 0
      ? t(cert ? 'The server\'s certificate ran out on {0}: publishing fails. Run the setup script again.' : 'The client secret ran out on {0}: publishing fails. Run the setup script again.', day(f.secretExpires))
      : t(cert ? 'The server\'s certificate runs out on {0} (in {1} days): run the setup script again.' : 'The client secret runs out on {0} (in {1} days): run the setup script again.', day(f.secretExpires), left);
  }

  // Back and forward (the mouse's side buttons, Alt+← and Alt+→, the browser's own): each page
  // and each signature opened is a step in the browser's history. The address stays as it is.
  let navRestoring = false;
  function navStep() {
    if (navRestoring) return;
    const st = { osign: 1, view: S.view, sigId: S.view === 'sig' ? S.sigId : null };
    const cur = history.state;
    if (cur && cur.osign && cur.view === st.view && cur.sigId === st.sigId) return;
    if (cur && cur.osign) history.pushState(st, ''); else history.replaceState(st, '');
  }
  function goView(view) {
    show(view);
    if (view === 'fields') renderFields(); else if (view === 'users') renderUsers(); else if (view === 'options') renderOptions(); else if (view === 'photos') renderPhotos();
  }
  addEventListener('popstate', (e) => {
    const st = e.state;
    if (!st || !st.osign || !S.config) return;
    closeMenu();
    navRestoring = true;
    try {
      if (st.view === 'sig' && S.config.signatures.some((x) => x.id === st.sigId)) openSig(st.sigId);
      else if (st.view === 'sig' || st.view === 'empty') openSig(null);
      else { flushSource(); goView(st.view); }
    } finally { navRestoring = false; }
    // What history held may be gone (a deleted signature): the step now says what is shown.
    history.replaceState({ osign: 1, view: S.view, sigId: S.view === 'sig' ? S.sigId : null }, '');
  });

  function show(view) {
    S.view = view;
    $('viewSig').hidden = view !== 'sig';
    $('viewFields').hidden = view !== 'fields';
    $('viewOptions').hidden = view !== 'options';
    $('viewUsers').hidden = view !== 'users';
    $('viewPhotos').hidden = view !== 'photos';
    $('viewEmpty').hidden = view !== 'empty';
    renderList();
    navStep();
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
    $('sigFrom').value = s.validFrom || '';
    $('sigUntil').value = s.validUntil || '';
    $('campaignSummary').textContent = campaignText(s) || t('always');
    loadDesign(s.html);
    $('source').value = prettyHtml(s.html);
    histReset(s);
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
    const name = from ? from.name + ' ' + t('(copy)') : (S.config.signatures.length ? t('New signature') : t('Main signature'));
    const s = from ? JSON.parse(JSON.stringify(from)) : core.newSignature('', name);
    s.id = newId(name); s.name = name;
    if (from) { s.defaultNew = false; s.defaultReply = false; } else { s.html = STARTER; s.defaultNew = !S.config.signatures.some((x) => x.defaultNew); s.defaultReply = !S.config.signatures.some((x) => x.defaultReply); }
    S.config.signatures.push(s);
    changed();
    openSig(s.id);
    if (!from) { $('sigName').focus(); $('sigName').select(); }
  }

