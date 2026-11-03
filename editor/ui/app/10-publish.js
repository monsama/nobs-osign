  // ---------------------------------------------------------------------------------------------
  // Publish and load
  // ---------------------------------------------------------------------------------------------
  // What a recipient would see go wrong: images too large, without a text when images are blocked,
  // too wide, unencrypted links, a signature too heavy, a dark logo gone in dark mode.
  const dark = {};
  // The images the signatures use, made as small as they are shown: twice the width they have in a
  // signature (sharp on high-resolution screens), uploaded as a new file and put in their place;
  // the old file stays until removed. Without a width in the signature an image stays as it is.
  async function shrinkUsed() {
    const all = S.config.signatures.map((s) => s.html || '').concat((S.config.blocks || []).map((b) => b.html || ''));
    const shown = new Map();
    for (const h of all) for (const m of h.matchAll(/<img\b[^>]*>/gi)) {
      const src = (/\ssrc\s*=\s*["']?assets\/([^"'\s>]+)/i.exec(m[0]) || [])[1], w = Number((/\swidth\s*=\s*["']?(\d+)/i.exec(m[0]) || [])[1]);
      if (src && w) shown.set(src, Math.max(shown.get(src) || 0, w));
    }
    let n = 0, from = 0, to = 0;
    for (const [name, w] of shown) {
      try {
        const blob = await (await fetch('/assets/' + encodeURIComponent(name), { credentials: 'same-origin' })).blob();
        const f = new File([blob], name, { type: blob.type || (/\.png$/i.test(name) ? 'image/png' : 'image/jpeg') });
        const small = await shrinkUpload(f, Math.max(48, 2 * w));
        if (!small.shrunk || small.size >= f.size) continue;
        const data = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = rej; r.readAsDataURL(small); });
        const r = await api('assets/upload', { name, data });
        if (r.name !== name) {
          const re = new RegExp('assets/' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=["\'\\s>?#])', 'g');
          S.config.signatures.forEach((s) => { s.html = (s.html || '').replace(re, 'assets/' + r.name); });
          (S.config.blocks || []).forEach((b) => { b.html = (b.html || '').replace(re, 'assets/' + r.name); });
        }
        S.assets = S.assets.filter((a) => a.name !== r.name).concat([{ name: r.name, size: r.size }]);
        n++; from += f.size; to += small.size;
      } catch (e) { /* that one stays */ }
    }
    if (n) { changed(); if (S.view === 'sig' && S.sigId && !S.blockEdit) openSig(S.sigId); }
    toast(n ? t('{0} images made smaller: {1} → {2}.', n, kb(from), kb(to)) : t('The images are as small as they are shown already.'));
  }
  // "3 signatures", or the name when it is one.
  const count = (list) => (list.length === 1 ? String(list[0].name || list[0]) : t('{0} signatures', list.length));
  function signatureChecks(c) {
    const out = [], names = (f) => [...new Set(c.signatures.filter(f).map((s) => s.name))];
    const imgs = (s) => [...(s.html || '').matchAll(/<img\b[^>]*>/gi)].map((m) => m[0]);
    const attr = (tag, a) => (new RegExp('\\s' + a + '\\s*=\\s*["\']?([^"\'\\s>]*)', 'i').exec(tag) || [])[1] || '';
    const size = (n) => { const a = S.assets.find((x) => x.name === n); return a ? a.size || 0 : 0; };
    const big = [...new Set(c.signatures.flatMap((s) => imgs(s).map((i) => (/^assets\/(.+)$/.exec(attr(i, 'src')) || [])[1]).filter((n) => n && size(n) > 100 * 1024)))];
    if (big.length) out.push(['warn', t('Large images (over 100 KB), in every mail: {0}. A smaller file makes mails lighter.', big.map((n) => n + ' (' + Math.round(size(n) / 1024) + ' KB)').join(', ')), { label: t('Make them smaller'), fn: shrinkUsed }, t('Large images: {0}', big.length)]);
    const noAlt = names((s) => imgs(s).some((i) => !/photos\//.test(attr(i, 'src')) && !/\salt\s*=\s*["'][^"']+["']/i.test(i)));
    if (noAlt.length) out.push(['warn', t('Images without a text for when images are blocked (alt): {0}. Image bar → alt text, e.g. the company name for the logo.', noAlt.join(', ')), null, t('Images without alt text: {0}', count(noAlt))]);
    const tooWide = (s) => imgs(s).some((i) => Number(attr(i, 'width')) > 600) ||
      [...(s.html || '').replace(/<img\b[^>]*>/gi, '').matchAll(/[\s;"']width\s*[=:]\s*["']?(\d+)(?:px)?(?![\d%])/gi)].some((m) => Number(m[1]) > 600);
    const wide = names(tooWide);
    if (wide.length) out.push(['warn', t('Wider than 600 pixels (too wide on phones): {0}.', wide.join(', ')), null, t('Wider than 600 pixels: {0}', count(wide))]);
    const http = names((s) => /href\s*=\s*["']http:\/\//i.test(s.html || '') || /src\s*=\s*["']http:\/\//i.test(s.html || ''));
    if (http.length) out.push(['warn', t('Links or images with http:// (not encrypted, some mail programs warn): {0}. Use https://.', http.join(', ')), null, t('Links with http://: {0}', count(http))]);
    const heavy = c.signatures.map((s) => ({ s, kb: Math.round(((s.html || '').length + [...new Set(imgs(s).map((i) => (/^assets\/(.+)$/.exec(attr(i, 'src')) || [])[1]).filter(Boolean))].reduce((n, x) => n + size(x), 0)) / 1024) })).filter((x) => x.kb > 250);
    if (heavy.length) out.push(['warn', t('Heavy signatures (text and images over 250 KB): {0}.', heavy.map((x) => x.s.name + ' (' + x.kb + ' KB)').join(', ')), null, t('Heavy signatures: {0}', count(heavy.map((x) => x.s)))]);
    const gone = [...new Set(c.signatures.flatMap((s) => imgs(s).map((i) => (/^assets\/(.+)$/.exec(attr(i, 'src')) || [])[1]).filter((n) => n && dark[n] === true)))];
    if (gone.length) out.push(['warn', t('Dark image on a transparent background, hard to see in dark mode: {0}. A light edge, or a background of its own, helps.', gone.join(', ')), null, t('Hard to see in dark mode: {0}', gone.length === 1 ? gone[0] : t('{0} images', gone.length))]);
    return out;
  }
  // Whether an image is dark lines on nothing (a logo that disappears on a dark background):
  // looked at once, in the background; the check above says it the next time it is shown.
  function darkCheck(name) {
    if (name in dark) return; dark[name] = null;
    const img = new Image();
    img.onload = () => {
      try {
        const w = Math.min(120, img.naturalWidth), h = Math.max(1, Math.round(img.naturalHeight * w / Math.max(1, img.naturalWidth)));
        const cv = document.createElement('canvas'); cv.width = w; cv.height = h;
        const g = cv.getContext('2d'); g.drawImage(img, 0, 0, w, h);
        const d = g.getImageData(0, 0, w, h).data;
        let clear = 0, seen = 0, lum = 0;
        for (let i = 0; i < d.length; i += 4) { if (d[i + 3] < 40) { clear++; continue; } seen++; lum += 0.2126 * d[i] + 0.7152 * d[i + 1] + 0.0722 * d[i + 2]; }
        dark[name] = clear > d.length / 4 * 0.2 && seen > 0 && lum / seen < 70;
      } catch (e) { dark[name] = false; }
    };
    img.onerror = () => { dark[name] = false; };
    img.src = '/assets/' + encodeURIComponent(name);
  }
  function publishChecks() {
    (S.assets || []).forEach((a) => { if (/\.png$/i.test(a.name)) darkCheck(a.name); });
    // As it goes out: the building blocks put in.
    const c = prepared(), out = [];
    const errs = core.validate(c);
    errs.forEach((e) => out.push(['bad', t(e)]));
    if (!c.signatures.length) out.push(['bad', t('There is no signature.')]);
    const plain = c.signatures.filter((s) => !(s.audience.departments || []).length && !(s.audience.domains || []).length && !(s.audience.groups || []).length && s.audience.everyone !== false);
    const nobody = c.signatures.filter((s) => core.forNobody(s));
    if (nobody.length) out.push(['warn', t('Not assigned, nobody gets: {0} (Who gets it).', nobody.map((s) => s.name).join(', ')), null, t('Not assigned: {0}', count(nobody))]);
    if (!c.signatures.some((s) => s.defaultNew)) out.push(['warn', t('No signature is the default for new mail: users get none unless they pick one.'), null, t('No default for new mail')]);
    else if (!plain.some((s) => s.defaultNew)) out.push(['warn', t('The default for new mail is limited to some users: everyone else gets none.'), null, t('The default for new mail is for some only')]);
    if (!c.signatures.some((s) => s.defaultReply)) out.push(['warn', t('No signature is the default for replies.')]);
    if (onServer()) {
      if (!toSharePoint()) out.push(['warn', t("Publishing goes into the server's own folder, not to SharePoint: Microsoft Graph is not set up on the server."), null, t('Into the server\'s folder, not SharePoint')]);
      else if (lastSettings.folderError) out.push(['warn', t('SharePoint cannot be reached: {0}', t(lastSettings.folderError)), null, t('SharePoint cannot be reached')]);
    } else if (!S.settings.folder) out.push(['bad', t('No folder to publish to: set the synced OSign folder in Settings.')]);
    if (core.groupsUsed(c).length && core.valueSource(c) === 'graph') out.push(['warn', t('Signatures limited to AD groups: the app registration needs GroupMember.Read.All (delegated, with admin consent), and the groups have to be synced to Entra ID.'), null, t('Groups need GroupMember.Read.All')]);
    const perPerson = c.signatures.filter((s) => /<img[^>]+src="photos\//i.test(s.html || '')).map((s) => s.name);
    if (perPerson.length) out.push(['good', t('Photos per person ({0}) come from photos\\<username>.jpg in the OSign folder; users without a photo get the signature without one.', perPerson.join(', ')), null, t('Photos per person')]);
    const lost = c.signatures.filter((s) => /<img[^>]+src="(?!assets\/|photos\/|https:)[^"]*"/i.test(s.html || '')).map((s) => s.name);
    if (lost.length) out.push(['warn', t('Images not stored with the signatures are left out in: {0}.', lost.join(', ')), null, t('Images left out: {0}', count(lost))]);
    const keys = fieldKeys();
    const used = [...new Set(c.signatures.flatMap((s) => core.placeholdersIn(s.html)))];
    const unknown = used.filter((k) => keys.indexOf(k) < 0);
    if (unknown.length) out.push(['warn', t('Unknown placeholders, left as typed: {0}', unknown.map((k) => '@@' + k).join(', ')), null, t('Unknown placeholders: {0}', unknown.length)]);
    const noBlock = [...new Set(c.signatures.flatMap((s) => core.blocksIn(s.html)))];
    if (noBlock.length) out.push(['warn', t('Unknown building blocks, left as typed: {0}', noBlock.map((k) => '{{' + k + '}}').join(', ')), null, t('Unknown building blocks: {0}', noBlock.length)]);
    const noGraph = used.filter((k) => { const f = c.fields.find((x) => x.key === k); return f && core.fieldSource(c, f) === 'none'; });
    if (noGraph.length) out.push(['warn', t('Empty in Outlook (no source): {0}', noGraph.map(fieldName).join(', ')), null, t('Fields empty in Outlook: {0}', noGraph.length)]);
    if (core.usesAdExport(c)) {
      const sn = S.snap;
      if (sn && S.snapMatches && !sn.written) out.push(['good', t('The users as read {0} are written: {1} new, {2} changed, {3} removed.', when(sn.readAt), sn.counts.new, sn.counts.changed, sn.counts.removed), null, t('Users: {0} new, {1} changed', sn.counts.new, sn.counts.changed)]);
      else if (sn && S.snapMatches && sn.written) out.push(['warn', t('The users stay as published {0}. To update them, read and check them under Users first.', when(sn.readAt)), null, t('The users stay as published')]);
      else out.push(['warn', t('The users are not written: read and check them under Users first. Until then they stay as they are in the folder.'), null, t('The users are not written')]);
    }
    const missing = [...new Set(c.signatures.flatMap((s) => [...(s.html || '').matchAll(/src="assets\/([^"]+)"/g)].map((m) => m[1])))].filter((n) => !S.assets.some((a) => a.name === n));
    if (missing.length) out.push(['warn', t('Images not in the draft, left out until added: {0}', missing.join(', ')), null, t('Images missing in the draft: {0}', missing.length)]);
    signatureChecks(c).forEach((x) => out.push(x));
    // Who would get nothing, as read from AD under Users.
    if (S.snap && S.snap.users && S.snap.users.length) {
      const g = whoGets(S.snap.users), none = S.snap.users.filter((u) => !g.get(u).new).length;
      if (none) out.push(['warn', t('{0} of the {1} users read from AD get no signature for new mail (Users → Only without a signature).', none, S.snap.users.length), null, t('{0} users without a signature', none)]);
    }
    if (!out.some((x) => x[0] !== 'good')) out.push(['good', t('{0} signature(s), ready.', c.signatures.length)]);
    return out;
  }

  // Keyboard shortcuts, in the page and in the design view (its own document).
  function onKeys(e) {
    const mod = e.ctrlKey || e.metaKey;
    const inDialog = document.querySelector('dialog[open]');
    // Undo and redo of the signature: in the design view, and anywhere on its page but a field
    // (a field, and the HTML tab, undo their own typing).
    const k = (e.key || '').toLowerCase();
    // Ctrl+Z / Ctrl+Y: in a text field the browser's own; in the signature editor its typing first
    // (as long as it has steps); else, and on every other page, the whole draft (10-undo.js).
    if (mod && !e.altKey && (k === 'z' || k === 'y') && !inDialog) {
      const inField = e.target && e.target.ownerDocument !== ddoc && /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName || '');
      if (!inField) {
        const dir = k === 'y' || e.shiftKey ? 1 : -1;
        e.preventDefault();
        if (S.view === 'sig' && histCan(dir)) histStep(dir); else ghStep(dir);
        return;
      }
    }
    if (mod && !e.shiftKey && !e.altKey && k === 'k' && !(inDialog && inDialog.id !== 'dlgFind')) { e.preventDefault(); openFind(); return; }
    if (mod && !e.shiftKey && !e.altKey && k === 'h' && !inDialog) { e.preventDefault(); openReplace(); return; }
    if (mod && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 's') { e.preventDefault(); save(); toast(t('Draft saved.')); return; }
    if (mod && e.key === 'Enter' && !inDialog) { e.preventDefault(); $('btnPublish').click(); return; }
    const typing = /^(INPUT|TEXTAREA|SELECT)$/.test((e.target && e.target.tagName) || '') || (e.target && e.target.isContentEditable) || (e.target && e.target.ownerDocument === ddoc);
    if (e.key === '?' && !mod && !typing && !inDialog) { e.preventDefault(); $('dlgKeys').showModal(); }
  }
  document.addEventListener('keydown', onKeys);

  // Ctrl+K: one box to go anywhere. Signatures, the views, the dialogs, the fields, the photos.
  let findItems = [], findAt = 0;
  function findSources() {
    const view = (v) => () => document.querySelector('.nav-list a[data-view="' + v + '"]').click();
    const items = [];
    S.config.signatures.forEach((s) => items.push({ kind: t('Signature'), label: s.name || t('(no name)'), go: () => openSig(s.id) }));
    [[t('All signatures'), () => openAll()], [t('Fields'), view('fields')], [t('Building blocks'), view('blocks')], [t('Add-in'), view('options')], core.usesAdExport(S.config) ? [t('Users'), view('users')] : null, [t('Photos'), view('photos')],
      [t('History'), () => $('navHistory').click()], [t('Settings'), () => $('btnSettings').click()], [t('New signature'), () => $('btnNew').click()],
      [t('Import templates'), () => $('btnImportOpen').click()], [t('Publish…'), () => $('btnPublish').click()], [t('Keyboard shortcuts'), () => $('dlgKeys').showModal()]]
      .filter(Boolean).forEach(([label, go]) => items.push({ kind: t('Go to'), label, go }));
    if (!$('setWizRow').hidden) items.push({ kind: t('Go to'), label: t('Setup assistant'), go: () => wizOpen(0) });
    items.push({ kind: t('Go to'), label: t('Load published'), go: () => $('btnLoad').click() });
    items.push({ kind: t('Go to'), label: t('Replace in all signatures'), go: () => openReplace() });
    if (S.config.signatures.some((s) => core.isCampaign(s))) items.push({ kind: t('Go to'), label: t('Campaigns'), go: () => $('btnCampaigns').click() });
    S.config.fields.forEach((f) => items.push({ kind: t('Field'), label: (f.label || f.key) + (f.label && f.label !== f.key ? ' · @@' + f.key : ''), go: () => { view('fields')(); setTimeout(() => {
      const row = [...document.querySelectorAll('#viewFields tr')].find((tr) => [...tr.querySelectorAll('input')].some((i) => i.value === f.key));
      if (row) { row.scrollIntoView({ block: 'center' }); const i = row.querySelector('input'); if (i) i.focus(); }
    }, 50); } }));
    (typeof photoItems !== 'undefined' ? photoItems : []).forEach((p) => items.push({ kind: t('Photo'), label: p.name, go: () => { view('photos')(); $('photosFilter').value = p.name; $('photosFilter').dispatchEvent(new Event('input')); } }));
    return items;
  }
  function findShow() {
    const q = $('findInput').value.trim().toLowerCase();
    const all = findSources();
    findItems = (q ? all.filter((x) => (x.label + ' ' + x.kind).toLowerCase().includes(q)) : all.filter((x) => x.kind !== t('Field') && x.kind !== t('Photo'))).slice(0, 40);
    findAt = Math.min(findAt, Math.max(0, findItems.length - 1));
    $('findList').replaceChildren(...(findItems.length ? findItems.map((x, i) => el('li', { role: 'option', 'aria-selected': String(i === findAt), class: i === findAt ? 'on' : '',
      onmousedown: (e) => { e.preventDefault(); findGo(i); }, onmousemove: () => { if (findAt !== i) { findAt = i; findShow(); } } },
      el('span', { class: 'fk' }, x.kind), el('span', { class: 'fl' }, x.label)))
      : [el('li', { class: 'none' }, emptyState('search', t('Nothing found.'), t('Try another word: a signature, a view, a field or a photo.')))]));
    const on = $('findList').querySelector('.on'); if (on) on.scrollIntoView({ block: 'nearest' });
  }
  function findGo(i) { const x = findItems[i]; if (!x) return; $('dlgFind').close(); x.go(); }
  function openFind() {
    if ($('dlgFind').open) { $('findInput').select(); return; }
    $('findInput').value = ''; findAt = 0; findShow();
    // Right under the search field, from its left edge (kept inside the window).
    const d = $('dlgFind'), r = $('btnFind').getBoundingClientRect();
    const w = Math.min(620, innerWidth - 16), top = (r.height ? r.bottom : 48) + 6;
    d.style.width = w + 'px';
    d.style.left = Math.max(8, Math.min(r.height ? r.left : (innerWidth - w) / 2, innerWidth - w - 8)) + 'px';
    d.style.top = top + 'px';
    d.style.setProperty('--find-max', Math.max(240, innerHeight - top - 16) + 'px');
    d.showModal(); $('findInput').focus();
  }
  $('btnFind').addEventListener('click', openFind);
  $('findInput').addEventListener('input', () => { findAt = 0; findShow(); });
  $('findInput').addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') { e.preventDefault(); findAt = Math.min(findItems.length - 1, findAt + 1); findShow(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); findAt = Math.max(0, findAt - 1); findShow(); }
    else if (e.key === 'Enter') { e.preventDefault(); findGo(findAt); }
  });
  $('dlgFind').addEventListener('click', (e) => { if (e.target === $('dlgFind')) $('dlgFind').close(); });
  $('keysBtn').addEventListener('click', () => $('dlgKeys').showModal());
  // What is not published yet: the list of it (when something was published), else Publish….
  $('stPub').addEventListener('click', () => { if (S.pub) openChanges(); else $('btnPublish').click(); });
  // What to look at before publishing, short: what stops it first (red), then the notes (one line
  // each, a click for the names and what to do), what is fine in one grey line at the end. With
  // nothing red, the notes are folded under one line.
  function renderChecks(checks) {
    const order = { bad: 0, warn: 1, good: 2 };
    const list = checks.slice().sort((a, b) => order[a[0]] - order[b[0]]);
    const bad = list.filter((c) => c[0] === 'bad'), warn = list.filter((c) => c[0] === 'warn'), good = list.filter((c) => c[0] === 'good');
    const actBtn = (act) => el('button', { type: 'button', class: 'small', onclick: async (e) => { e.currentTarget.disabled = true; await act.fn(); $('dlgPublish').close(); $('btnPublish').click(); } }, act.label);
    const item = ([k, text, act, short]) => el('li', { class: k }, short
      ? el('details', { class: 'chk' }, el('summary', null, short), el('p', null, text), act ? actBtn(act) : null)
      : el('span', null, text, act ? actBtn(act) : null));
    const out = bad.map(item);
    if (warn.length) {
      const items = warn.map(item);
      if (bad.length) out.push(...items);
      else out.push(el('li', { class: 'chk-head' }, el('details', { class: 'chk-all' },
        el('summary', null, el('b', null, t('Ready to publish.')), ' ', warn.length === 1 ? t('1 note') : t('{0} notes', warn.length)),
        el('ul', { class: 'checks chk-sub' }, ...items))));
    } else if (!bad.length) out.push(el('li', { class: 'good' }, el('b', null, t('Ready to publish.'))));
    if (good.length) out.push(el('li', { class: 'good chk-ok' }, good.map((g) => g[3] || g[1]).join(' · '),
      el('details', { class: 'chk' }, el('summary', null, t('details')), ...good.map((g) => el('p', null, g[1])))));
    $('pubChecks').replaceChildren(...out);
  }
  $('btnPublish').addEventListener('click', async () => {
    flushSource();
    if (core.usesAdExport(S.config)) { await saving; await loadSnapshot(); }
    $('pubWhere').textContent = onServer() ? t('To {0}. New messages use it a minute or two later.', S.settings.folder)
      : S.settings.folder ? t('To {0}. OneDrive uploads it to SharePoint; new messages use it a minute or two later.', S.settings.folder)
      : t('No folder set yet (Settings).');
    const checks = publishChecks();
    renderChecks(checks);
    offFor($('pubGo'), checks.some((c) => c[0] === 'bad') ? t('First fix what is marked red above.') : '');
    $('pubErr').textContent = ''; $('pubOk').textContent = ''; $('pubMsg').value = ''; $('pubGo').hidden = false;
    await loadApproval(); renderApproval();
    await loadRelease();
    $('dlgPublish').showModal();
  });
  // Later, at a set time (server/routes-release.js): the server's, for whoever publishes without approval.
  async function loadRelease() {
    const can = onServer() && !(S.info.features && S.info.features.requireApproval);
    $('pubSteps').hidden = !can;
    $('pubLater').checked = false; $('pubAt').hidden = true;
    let x = null;
    if (onServer()) { try { x = (await api('schedule/get')).schedule; } catch (e) { x = null; } }
    const when = (v) => new Date(v).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB');
    $('pubSchedNow').hidden = !x;
    if (x) $('pubSchedWho').textContent = t('Set to go out on {0} ({1}): the draft as it was then.', when(x.at), x.by) +
      (x.error ? ' ' + t('It did not go out: {0}', t(x.error)) : '');
    $('pubSchedWho').classList.toggle('warn', !!(x && x.error));
  }
  $('pubLater').addEventListener('change', () => {
    $('pubAt').hidden = !$('pubLater').checked;
    if ($('pubLater').checked && !$('pubAt').value) { const d = new Date(Date.now() + 86400000); d.setHours(6, 0, 0, 0); $('pubAt').value = new Date(d - d.getTimezoneOffset() * 60000).toISOString().slice(0, 16); }
    $('pubGo').textContent = pubGoText();
  });
  const pubGoText = () => ($('pubLater').checked ? t('Set the time') : t('Publish'));
  $('pubSchedCancel').addEventListener('click', async () => { try { await api('schedule/cancel'); await loadRelease(); $('pubOk').textContent = t('The set time is cancelled.'); } catch (err) { $('pubErr').textContent = t(err.message); } });
  // The note for History: what was typed, or the placeholder's words.
  const pubNote = () => $('pubMsg').value.trim() || $('pubMsg').placeholder;
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
      // With four eyes, the server writes them when the publish is approved.
      if (core.usesAdExport(c) && S.snap && S.snapMatches && !S.snap.written && !S.info.features.requireApproval) {
        try {
          const x = await api('ad/export', adRequest(c));
          exported = ' ' + t('{0} users written ({1} new or changed, {2} removed).', x.users, x.written, x.removed);
          S.snap.written = true;
          if (!$('viewUsers').hidden) renderUsers();
        } catch (err) { exported = ' ' + t('The users could not be written: {0}', t(err.message)); }
      }
      if (S.info.features.requireApproval) {
        await api('approval/request', { configJson: JSON.stringify(c), message: pubNote() });
        S.config.updated = c.updated;
        await loadApproval(); renderApproval();
        $('pubOk').textContent = t('Asked for approval: someone else who signs in to the editor can now publish it.') + exported;
        $('pubGo').hidden = true;
        return;
      }
      if ($('pubLater').checked) {
        const at = new Date($('pubAt').value);
        if (isNaN(at)) throw new Error('The time is in the future, a minute from now at least.');
        await api('schedule/set', { configJson: JSON.stringify(c), at: at.toISOString(), message: pubNote() });
        $('pubOk').textContent = t('Set: it goes out on {0}, as the draft is now. Changes after that wait for the next publish.', at.toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB')) + exported;
        $('pubGo').hidden = true; await loadRelease();
        return;
      }
      const r = await api('publish', { configJson: JSON.stringify(c), message: pubNote() });
      S.config.updated = c.updated; S.dirty = false; S.publishedAt = Date.now(); renderStatus(); loadPublishedCopy();
      // The server version: whether the folder is there now (it is made on the first publish).
      if (onServer()) { try { lastSettings = await api('settings/get'); } catch (e2) { /* as it was */ } }
      $('pubOk').textContent = (r.changed ? t('Published to {0}.', r.folder) : t('Nothing had changed.')) + exported;
      // The users written now: their marks (new, changed) go on the Users page.
      if (core.usesAdExport(S.config)) loadSnapshot().then(() => { if (S.view === 'users') renderUsers(); });
      // Done: only Close is left.
      $('pubGo').hidden = true;
    } catch (err) { $('pubOk').textContent = ''; $('pubErr').textContent = t(err.message); }
    finally { $('pubGo').disabled = false; }
  });

  // The web certificate the proxy serves for this address: what it is, and whether PCs trust it.
  let tlsLast = null;
  async function loadTls() {
    if (!onServer() || location.protocol !== 'https:') return null;
    try { tlsLast = await api('server/tls', { host: location.hostname }); } catch (e) { tlsLast = { error: e.message }; }
    // Not trusted by the PCs is a setup matter: said in Settings → Server, not on every page. Running
    // out soon is the one that needs doing in time, so that stays in the status bar.
    const c = tlsLast, warn = !c.error && c.trusted && c.daysLeft <= 30 ? t('The web certificate runs out on {0}.', day(c.validTo)) : '';
    $('stTls').hidden = !warn; $('stTls').textContent = warn;
    return c;
  }
  function tlsText(c) {
    if (!c) return t('Web certificate: only shown over https.');
    if (c.error) return t('Web certificate: {0}', t(c.error));
    const base = t('Web certificate: {0}, issued by {1}, valid until {2} ({3}).', c.names.length ? c.names.join(', ') : c.subject, c.issuer, day(c.validTo), t('{0} days', c.daysLeft));
    if (!c.nameOk) return base + ' ' + t('It is not for {0}.', c.host);
    if (c.trusted) return base + ' ' + t('Trusted.');
    return base + ' ' + (c.selfSigned || /caddy|local authority/i.test(c.issuer) ? t('It is the proxy\'s own: PCs do not trust it (the browser warns, the setup script cannot send its result).') : t('Not trusted here: PCs that know your CA may trust it.')) +
      ' ' + t('Issue a web server certificate from your CA for {0} and install it in the proxy (docs/SETUP.md).', c.host);
  }

  // Four eyes: turned on and off by the local account admin, at once.
  $('setApproval').addEventListener('change', async () => {
    const on = $('setApproval').checked;
    $('setApprovalNote').textContent = '';
    try {
      const r = await api('settings/save', { settings: { requireApproval: on ? 'on' : '' } });
      S.settings = r.settings; lastSettings = r;
      S.info = await api('info');
      $('setApprovalNote').textContent = on ? (ldapOn() ? t('On: publishing now waits for a second person.') : t('On, but nobody else can sign in yet (Active Directory): set up the sign-in first, or nothing can be published.')) : t('Off: whoever signs in publishes.');
      renderStatusBar();
    } catch (err) { $('setApproval').checked = !on; $('setApprovalNote').textContent = t(err.message); }
  });
  function ldapOn() { return !!(S.info && S.info.features && S.info.features.adSignIn); }

  // What changed, as a list: signatures added, removed, renamed, their HTML line by line, who gets
  // them, the fields and the options.
  function renderChanges(box, changes) {
    if (!changes.length) { box.replaceChildren(el('p', { class: 'muted' }, t('Nothing changed.'))); return; }
    const KIND = { added: 'Added', removed: 'Removed', renamed: 'Renamed', html: 'Changed', rules: 'Who gets it, defaults or days', fields: 'Fields', options: 'Add-in options', imgadd: 'Image added', imgdel: 'Image removed', imgchg: 'Image replaced' };
    box.replaceChildren(...changes.map((c) => el('div', { class: 'change ' + c.kind },
      el('div', { class: 'change-head' }, el('span', { class: 'badge' + (c.kind === 'removed' || c.kind === 'imgdel' ? ' danger' : c.kind === 'added' || c.kind === 'imgadd' ? ' on' : '') }, t(KIND[c.kind] || c.kind)), ' ',
        c.name ? el('b', null, c.name) : null, c.kind === 'renamed' ? ' ' + t('(was {0})', c.detail) : c.detail ? ' ' + c.detail : ''),
      c.lines && c.lines.length ? el('pre', { class: 'diff' }, ...c.lines.map((l) => el('span', { class: l.op === '+' ? 'add' : 'del' }, l.op + ' ' + l.text + '\n'))) : null)));
  }

  // The audit log, for the local account admin: each entry with its action in words, whether it
  // went through, and a search over what is loaded.
  const AUDIT_WHAT = {
    signin: 'Sign-in', signout: 'Sign-out', publish: 'Published', 'history/restore': 'Version brought back',
    'connections/save': 'Server connections', 'connections/cert': 'Certificate', 'settings/save': 'Settings',
    'setup/start': 'Setup script', 'setup/finish': 'Setup script result', 'addin/build': 'Add-in files', 'addin/upload': 'Add-in upload',
    'ad/export': 'Users exported', 'backup/export': 'Backup', 'backup/restore': 'Backup restored', 'backup/check': 'Backup checked', 'sessions/end': 'Session ended', 'notify/save': 'Notifications', 'notify/test': 'Test notification',
    'approval/request': 'Approval asked', 'approval/approve': 'Approved', 'approval/cancel': 'Approval withdrawn',
    'photos/upload': 'Photo', 'photos/delete': 'Photo removed', 'photos/fetch': 'Photo from the web'
  };
  let auditEntries = [], auditShown = [];
  function auditShow() {
    const q = $('auditFind').value.trim().toLowerCase();
    const loc = I18N.lang === 'de' ? 'de-CH' : 'en-GB';
    const since = Number($('auditSince').value) ? Date.now() - Number($('auditSince').value) * 86400000 : 0;
    const rows = auditShown = auditEntries.filter((e) => (!since || Date.parse(e.at) >= since) && (!q || [e.who, e.ip, e.what, t(AUDIT_WHAT[e.what] || e.what), e.detail ? t(e.detail) : ''].join(' ').toLowerCase().includes(q)));
    $('auditCount').textContent = auditEntries.length ? t('{0} of {1}', rows.length, auditEntries.length) : '';
    const tb = $('auditRows');
    if (!rows.length) { tb.replaceChildren(el('tr', null, el('td', { colspan: 5, class: 'muted empty' }, auditEntries.length ? t('Nothing matches.') : t('Nothing yet.')))); return; }
    tb.replaceChildren(...rows.map((e) => {
      const at = new Date(e.at);
      return el('tr', { class: e.ok ? '' : 'refused' },
        el('td', { class: 'when' }, el('span', null, at.toLocaleDateString(loc)), el('small', null, at.toLocaleTimeString(loc))),
        el('td', { class: 'who' }, e.who || '-'),
        el('td', { class: 'ip' }, e.ip || ''),
        el('td', { class: 'what' }, el('span', { class: 'st ' + (e.ok ? 'ok' : 'removed'), title: e.what }, t(AUDIT_WHAT[e.what] || e.what)), e.ok ? null : el('small', { class: 'warn' }, t('refused'))),
        el('td', { class: 'detail' }, e.detail ? t(e.detail) : ''));
    }));
  }
  async function renderAudit() {
    $('auditErr').textContent = '';
    const v = $('auditWhat').value;
    $('auditRows').replaceChildren(el('tr', null, el('td', { colspan: 5, class: 'muted empty' }, t('Loading…'))));
    try {
      const r = await api('audit/list', v === 'failed' ? { failed: true } : { what: v });
      auditEntries = r.entries;
      auditShow();
    } catch (err) { auditEntries = []; $('auditRows').replaceChildren(); $('auditCount').textContent = ''; $('auditErr').textContent = t(err.message); }
  }
  // Who is signed in now (the local admin's): one ended by a click, asked first.
  function renderSessions(list) {
    const loc = I18N.lang === 'de' ? 'de-CH' : 'en-GB', w = (iso) => new Date(iso).toLocaleString(loc);
    $('sessRows').replaceChildren(...list.map((x) => el('tr', null,
      el('td', { class: 'who' }, x.name && x.name !== x.login ? x.name + ' (' + x.login + ')' : x.login, x.mine ? el('small', null, ' ' + t('(you)')) : null),
      el('td', { class: 'ip' }, x.from || ''), el('td', null, w(x.started)), el('td', null, w(x.last)),
      el('td', null, x.mine ? null : el('button', { type: 'button', class: 'small ghost danger', onclick: async () => {
        if (!(await ask(t('Sign {0} out? They sign in again to go on.', x.login), t('Sign out'), true))) return;
        try { const r = await api('sessions/end', { ref: x.ref }); renderSessions(r.sessions); toast(t('{0} is signed out.', r.ended)); } catch (e) { $('sessErr').textContent = t(e.message); }
      } }, t('Sign out'))))));
  }
  // Notifications to a webhook (the local admin's): the address kept sealed, never shown again.
  function notifyRender(r) {
    $('ntUrl').value = ''; $('ntUrl').placeholder = r.on ? t('set (type a new one to replace it)') : 'https://…';
    $('ntUrlNote').textContent = r.lost ? t('The saved address cannot be opened any more: enter it again.') : '';
    $('ntOff').hidden = !r.on;
    const last = r.last ? (r.last.ok ? t('Last message {0}.', new Date(r.last.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB')) : t('The last message did not go out: {0}', t(r.last.error))) : '';
    $('ntNow').textContent = (r.now && r.now.length ? t('To tell now: {0}', r.now.map((x) => t(x)).join(' · ')) : t('Nothing to tell right now.')) + (last ? ' ' + last : '');
    $('setNotifyState').textContent = r.on ? t('On.') : '';
  }
  $('setNotify').addEventListener('click', async () => {
    $('ntErr').textContent = ''; $('ntOk').textContent = '';
    try { notifyRender(await api('notify/get')); $('dlgNotify').showModal(); } catch (e) { toast(t(e.message), true); }
  });
  $('ntSave').addEventListener('click', async () => {
    $('ntErr').textContent = ''; $('ntOk').textContent = '';
    if (!$('ntUrl').value.trim()) { $('ntErr').textContent = t('Enter the webhook\'s address.'); return; }
    try { await api('notify/save', { url: $('ntUrl').value }); notifyRender(await api('notify/get')); $('ntOk').textContent = t('Saved. "Send a test" shows whether it arrives.'); } catch (e) { $('ntErr').textContent = t(e.message); }
  });
  $('ntOff').addEventListener('click', async () => {
    try { await api('notify/save', { url: '' }); notifyRender(await api('notify/get')); $('ntOk').textContent = t('Notifications off.'); } catch (e) { $('ntErr').textContent = t(e.message); }
  });
  $('ntTest').addEventListener('click', async () => {
    $('ntErr').textContent = ''; $('ntOk').textContent = t('Sending…'); $('ntTest').disabled = true;
    try {
      if ($('ntUrl').value.trim()) await api('notify/save', { url: $('ntUrl').value });
      await api('notify/test'); notifyRender(await api('notify/get')); $('ntOk').textContent = '✓ ' + t('Sent: look in the channel.');
    } catch (e) { $('ntOk').textContent = ''; $('ntErr').textContent = t(e.message); }
    finally { $('ntTest').disabled = false; }
  });
  $('setSessions').addEventListener('click', async () => {
    $('sessErr').textContent = ''; $('sessRows').replaceChildren(); $('dlgSessions').showModal();
    try { renderSessions((await api('sessions/list')).sessions); } catch (e) { $('sessErr').textContent = t(e.message); }
  });
  $('setAudit').addEventListener('click', () => { $('auditFind').value = ''; $('auditSince').value = ''; $('dlgAudit').showModal(); renderAudit(); });
  $('auditWhat').addEventListener('change', renderAudit);
  $('auditReload').addEventListener('click', renderAudit);
  $('auditFind').addEventListener('input', auditShow);
  $('auditSince').addEventListener('change', auditShow);
  $('auditCsv').addEventListener('click', () => {
    if (!auditShown.length) { toast(t('Nothing to export.'), true); return; }
    downloadCsv('nobs-osign-audit-' + new Date().toISOString().slice(0, 10) + '.csv', [[t('When'), t('Who'), t('From'), t('What'), t('Went through'), t('Details')]]
      .concat(auditShown.map((e) => [new Date(e.at).toISOString(), e.who || '', e.ip || '', t(AUDIT_WHAT[e.what] || e.what), e.ok ? t('yes') : t('no'), e.detail ? t(e.detail) : ''])));
  });

  // A backup: downloaded locked with a passphrase, or brought back.
  $('setBackup').addEventListener('click', () => {
    for (const id of ['bakPass', 'bakPass2', 'bakRestorePass', 'bakFile']) $(id).value = '';
    fileName('bakFile');
    for (const id of ['bakErr', 'bakExportOk', 'bakRestoreOk', 'abOk']) $(id).textContent = '';
    $('abPass').value = '';
    $('dlgBackup').showModal();
    abLoad();
  });
  // Automatic backups (server/autobackup.js).
  function abShow(r) {
    $('abAt').value = r.at || ''; $('abKeep').value = r.keep || 30;
    $('abWhere').textContent = t('Locked with a passphrase of yours and kept in {0}. Last: {1}.', r.where,
      r.last ? new Date(r.last.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB') + ' ' + (r.last.ok ? '✓' : '✕ ' + t(r.last.error || '')) : t('none yet'));
    if (r.warning) $('abWhere').textContent += ' ' + t(r.warning);
    $('abPassNote').textContent = r.passLost ? t('The one set cannot be opened any more (another server key): set it again.') : r.passSet ? t('(set; type a new one to change it)') : t('(at least 12 characters; needed to open them)');
    offFor($('abRun'), r.passSet ? '' : t('Set a passphrase for the automatic backups first.'));
    $('abList').replaceChildren(...(r.backups || []).slice(0, 10).map((b) => el('li', null,
      el('span', null, b.name), el('small', { class: 'where' }, Math.round(b.size / 1024) + ' KB'),
      el('button', { type: 'button', class: 'small ghost', onclick: () => abDownload(b.name) }, t('Download')))));
  }
  async function abLoad() {
    try { abShow(await api('backup/auto/get')); } catch (e) { $('abWhere').textContent = t(e.message); }
  }
  async function abDownload(name) {
    try {
      const r = await api('backup/auto/download', { name });
      const bytes = Uint8Array.from(atob(r.data), (c) => c.charCodeAt(0));
      const a = el('a', { href: URL.createObjectURL(new Blob([bytes], { type: 'application/octet-stream' })), download: r.name });
      document.body.append(a); a.click(); a.remove();
    } catch (e) { $('bakErr').textContent = t(e.message); }
  }
  $('abSave').addEventListener('click', async () => {
    $('bakErr').textContent = ''; $('abOk').textContent = '';
    try {
      const d = { at: $('abAt').value, keep: Number($('abKeep').value) };
      if ($('abPass').value) d.passphrase = $('abPass').value;
      abShow(await api('backup/auto/save', d)); $('abPass').value = '';
      $('abOk').textContent = d.at ? t('Saved: every day at {0}.', d.at) : t('Saved: off.');
    } catch (e) { $('bakErr').textContent = t(e.message); }
  });
  $('abRun').addEventListener('click', async () => {
    $('bakErr').textContent = ''; $('abOk').textContent = t('Making the backup…'); $('abRun').disabled = true;
    try { abShow(await api('backup/auto/run')); $('abOk').textContent = t('Made.'); }
    catch (e) { $('abOk').textContent = ''; $('bakErr').textContent = t(e.message); $('abRun').disabled = false; }
  });
  $('bakExport').addEventListener('click', async () => {
    $('bakErr').textContent = ''; $('bakExportOk').textContent = '';
    if ($('bakPass').value !== $('bakPass2').value) { $('bakErr').textContent = t('The two passphrases are not the same.'); return; }
    $('bakExport').disabled = true; $('bakExportOk').textContent = t('Making the backup…');
    try {
      const r = await api('backup/export', { passphrase: $('bakPass').value });
      const bytes = Uint8Array.from(atob(r.data), (c) => c.charCodeAt(0));
      const a = el('a', { href: URL.createObjectURL(new Blob([bytes], { type: 'application/octet-stream' })), download: r.name });
      document.body.append(a); a.click(); a.remove();
      $('bakExportOk').textContent = t('Downloaded: {0}. Keep the passphrase apart from it.', r.name);
    } catch (err) { $('bakExportOk').textContent = ''; $('bakErr').textContent = t(err.message); }
    finally { $('bakExport').disabled = false; }
  });
  // Whether the file opens with that passphrase, and what is in it: nothing is changed here.
  const bakRead = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = rej; r.readAsDataURL(f); });
  $('bakCheck').addEventListener('click', async () => {
    $('bakErr').textContent = ''; $('bakRestoreOk').textContent = '';
    const f = $('bakFile').files[0];
    if (!f) { $('bakErr').textContent = t('Choose the backup file.'); return; }
    $('bakCheck').disabled = true; $('bakRestoreOk').textContent = t('Opening it…');
    try {
      const r = await api('backup/check', { data: await bakRead(f), passphrase: $('bakRestorePass').value });
      const when = new Date(r.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB');
      $('bakRestoreOk').className = 'where good';
      $('bakRestoreOk').textContent = '✓ ' + t('It opens: the backup of {0} (version {1}), {2} signatures, {3} published versions, {4} files. Nothing was changed.', when, r.version, r.signatures === null ? '?' : r.signatures, r.versions, r.files);
    } catch (err) { $('bakRestoreOk').textContent = ''; $('bakRestoreOk').className = 'where'; $('bakErr').textContent = t(err.message); }
    finally { $('bakCheck').disabled = false; }
  });
  $('bakRestore').addEventListener('click', async () => {
    $('bakErr').textContent = ''; $('bakRestoreOk').textContent = ''; $('bakRestoreOk').className = 'where';
    const f = $('bakFile').files[0];
    if (!f) { $('bakErr').textContent = t('Choose the backup file.'); return; }
    if (!(await ask(t('Replace the draft, the settings, the connections and the history here with the backup\'s?'), t('Restore'), true))) return;
    $('bakRestore').disabled = true; $('bakRestoreOk').textContent = t('Bringing it back…');
    try {
      const data = await bakRead(f);
      const r = await api('backup/restore', { data, passphrase: $('bakRestorePass').value });
      const [info, st, draft] = await Promise.all([api('info'), api('settings/get'), api('draft/get')]);
      S.info = info; S.settings = st.settings; lastSettings = st;
      applyDraft(draft); renderStatusBar();
      $('bakRestoreOk').textContent = t('Brought back: the backup of {0} (version {1}).', new Date(r.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB'), r.version);
    } catch (err) { $('bakRestoreOk').textContent = ''; $('bakErr').textContent = t(err.message); }
    finally { $('bakRestore').disabled = false; }
  });

  // Four eyes: what waits for approval, shown in Publish….
  async function loadApproval() {
    if (!(S.info && S.info.features && S.info.features.requireApproval)) { S.approval = null; renderStatusBar(); return null; }
    try { S.approval = await api('approval/get'); } catch (e) { S.approval = null; }
    renderStatusBar();
    return S.approval;
  }
  function renderApproval() {
    const need = !!(S.info && S.info.features && S.info.features.requireApproval);
    const p = need && S.approval && S.approval.pending;
    $('pubApproval').hidden = !p;
    $('pubGo').textContent = need ? (p && p.mine ? t('Ask again') : t('Ask for approval')) : t('Publish');
    if (!p) return;
    const when = new Date(p.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB');
    $('pubApprovalWho').textContent = (p.mine ? t('You asked for approval on {0}.', when) : t('{0} asks for approval since {1}.', p.by, when)) + (p.message ? ' «' + p.message + '»' : '') +
      (p.current ? '' : ' ' + t('The draft changed since: it has to be asked for again.'));
    $('pubApprovalWho').className = 'lead' + (p.current ? '' : ' warn');
    renderChanges($('pubApprovalChanges'), p.changes);
    $('pubApprove').hidden = !p.mayApprove;
    $('pubWithdraw').textContent = p.mine ? t('Withdraw') : t('Reject');
  }
  $('stApproval').addEventListener('click', () => $('btnPublish').click());
  // To the place that makes them again.
  // One click: the add-in's files made for this version and uploaded (with Cloudflare Pages);
  // what it does is shown on the button, done with a check that goes by itself.
  $('stAddin').addEventListener('click', async () => {
    const b = $('stAddin'), text = $('stAddinText');
    if (b.classList.contains('busy')) return;
    b.classList.add('busy'); b.classList.remove('failed'); b.disabled = true;
    try {
      text.textContent = t('Making the add-in…');
      const r = await api('addin/build');
      let uploaded = false;
      if (S.info.features.pagesUpload) { text.textContent = t('Uploading…'); await api('addin/upload'); uploaded = true; }
      b.classList.remove('busy'); b.classList.add('done');
      text.textContent = uploaded ? t('Add-in {0} is out', r.version) : t('Add-in {0} made', r.version);
      if (!uploaded) toast(t('Add-in files made for {0}: copy them to the web server.', r.version), false, { label: t('Download as ZIP'), run: () => { location.href = '/addin.zip'; } });
      else if (r.note) toast(t(r.note), true);
      setTimeout(() => { b.classList.remove('done'); b.disabled = false; afterAddinChange(); }, 4000);
    } catch (e) {
      b.classList.remove('busy'); b.classList.add('failed'); b.disabled = false;
      text.textContent = t('Not done: try again');
      toast(t(e.message), true, { label: t('Settings'), run: () => { $('btnSettings').click(); setTab('addin'); } });
    }
  });
  $('pubApprove').addEventListener('click', async () => {
    $('pubApprove').disabled = true; $('pubErr').textContent = ''; $('pubOk').textContent = t('Publishing…');
    try {
      const r = await api('approval/approve', { hash: (S.approval && S.approval.pending && S.approval.pending.hash) || '' });
      S.dirty = false; S.publishedAt = Date.now(); renderStatus(); loadPublishedCopy();
      $('pubOk').textContent = r.changed ? t('Approved and published to {0}.', r.folder) : t('Nothing had changed.');
      await loadApproval(); renderApproval();
    } catch (err) { $('pubOk').textContent = ''; $('pubErr').textContent = t(err.message); }
    finally { $('pubApprove').disabled = false; }
  });
  $('pubWithdraw').addEventListener('click', async () => {
    try { await api('approval/cancel'); await loadApproval(); renderApproval(); $('pubOk').textContent = t('The request is withdrawn.'); }
    catch (err) { $('pubErr').textContent = t(err.message); }
  });

  // History: the versions published, and one of them back into the draft.
  async function renderHistory() {
    $('historyErr').textContent = '';
    const tb = $('historyRows'); tb.replaceChildren(el('tr', null, el('td', { colspan: 5 }, t('Loading…'))));
    let r;
    try { r = await api('history/list'); } catch (err) { tb.replaceChildren(); $('historyErr').textContent = t(err.message); return; }
    if (!r.versions.length) { tb.replaceChildren(el('tr', null, el('td', { colspan: 5 }, emptyState('history', t('Nothing published yet.'), t('Each version you publish is kept here, to compare and to go back to.'), [t('Publish…'), () => { $('dlgHistory').open && $('dlgHistory').close(); $('btnPublish').click(); }])))); return; }
    tb.replaceChildren(...r.versions.map((v, i) => el('tr', null,
      el('td', null, new Date(v.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB'), i === 0 ? el('span', { class: 'badge on', style: 'margin-left:6px' }, t('current')) : null),
      el('td', null, v.by), el('td', null, v.message || '-'), el('td', null, String(v.signatures)),
      el('td', { class: 'nowrap' }, el('button', { type: 'button', class: 'small', onclick: () => historyDiff(v) }, t('Compare')), ' ',
        el('button', { type: 'button', class: 'small', onclick: () => historyRestore(v) }, t('Load into the draft'))))));
  }
  let diffOf = null;
  async function historyDiff(v) {
    if (v) { diffOf = v; document.querySelector('#dlgDiff input[value=previous]').checked = true; $('dlgDiff').showModal(); }
    const against = document.querySelector('#dlgDiff input[name=diffAgainst]:checked').value;
    $('diffTitle').textContent = t('What changed: the version of {0}', new Date(diffOf.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB'));
    $('diffErr').textContent = ''; $('diffChanges').replaceChildren(el('p', null, t('Loading…')));
    try { renderChanges($('diffChanges'), (await api('history/diff', { id: diffOf.id, against })).changes); }
    catch (err) { $('diffChanges').replaceChildren(); $('diffErr').textContent = t(err.message); }
  }
  document.querySelectorAll('#dlgDiff input[name=diffAgainst]').forEach((r) => r.addEventListener('change', () => historyDiff()));
  async function historyRestore(v) {
    if (!(await ask(t('Replace the draft and its images with the version of {0}? What is not published is lost.', new Date(v.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB')), t('Bring back'), true))) return;
    try {
      await api('history/restore', { id: v.id });
      applyDraft(await api('draft/get'));
      $('dlgHistory').close();
      toast(t('Loaded the version of {0} into the draft. Publish it to send it out.', new Date(v.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB')));
    } catch (err) { $('historyErr').textContent = t(err.message); }
  }
  // Small screens: the sidebar opens over the page, and closes once something in it is chosen.
  function sideOpen(on) { document.body.classList.toggle('side-open', on); $('btnSide').setAttribute('aria-expanded', on ? 'true' : 'false'); }
  $('btnSide').addEventListener('click', () => sideOpen(!document.body.classList.contains('side-open')));
  // Something chosen in it (a page, a signature, a new one, the import): the menu closes.
  $('side').addEventListener('click', (e) => { if (e.target.closest('a, .sig-list li, #btnNew, #btnImportOpen')) sideOpen(false); });
  $('navHistory').addEventListener('click', (e) => { e.preventDefault(); $('dlgHistory').showModal(); renderHistory(); });
  $('btnLoad').addEventListener('click', async () => {
    if (S.dirty && !(await ask(t('Replace the draft with what is published? Changes not published are lost (the draft is kept once as draft.json.bak).'), t('Replace'), true))) return;
    // Everything as published again, read anew in this page (no reload): the draft, its images,
    // the state; a change still waiting to be saved is dropped, not written over it afterwards.
    try {
      clearTimeout(saveTimer); saveTimer = null; await saving;
      await api('load-published');
      const where = S.view;
      applyDraft(await api('draft/get'));
      if (where !== 'sig' && where !== 'empty') goView(where);
      renderStatusBar();
      toast(t('Loaded the published signatures.'));
    } catch (e) { toast(t(e.message), true); }
  });

