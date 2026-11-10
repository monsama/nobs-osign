  // ---------------------------------------------------------------------------------------------
  // Publish and load
  // ---------------------------------------------------------------------------------------------
  // What a recipient would see go wrong: images too large, without a text when images are blocked,
  // too wide, unencrypted links, a signature too heavy, a dark logo gone in dark mode.
  // "3 signatures", or the name when it is one.
  const count = (list) => (list.length === 1 ? String(list[0].name || list[0]) : t('{0} signatures', list.length));
  // What a recipient would see go wrong in the signatures (images, width, links, dark mode).
  function signatureChecks(c) {
    const out = [], names = (f) => [...new Set(c.signatures.filter(f).map((s) => s.name))];
    const imgs = (s) => [...(s.html || '').matchAll(/<img\b[^>]*>/gi)].map((m) => m[0]);
    const attr = (tag, a) => (new RegExp('\\s' + a + '\\s*=\\s*["\']?([^"\'\\s>]*)', 'i').exec(tag) || [])[1] || '';
    const size = (n) => { const a = S.assets.find((x) => x.name === n); return a ? a.size || 0 : 0; };
    const big = [...new Set(c.signatures.flatMap((s) => imgs(s).map((i) => (/^assets\/(.+)$/.exec(attr(i, 'src')) || [])[1]).filter((n) => n && size(n) > 100 * 1024)))];
    if (big.length) out.push(['warn', t('Large images (over 100 KB), in every mail: {0}. A smaller file makes mails lighter.', big.map((n) => n + ' (' + Math.round(size(n) / 1024) + ' KB)').join(', ')), { label: t('Make images lighter'), fn: prepareClassic, leave: true }, t('Large images: {0}', big.length)]);
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
  // What would go wrong with the banner campaigns not over yet.
  function campaignChecks(c) {
    const out = [], day = core.today(), list = (c.campaigns || []).filter((x) => !x.validUntil || x.validUntil >= day);
    const parts = (x) => [x].concat(x.more || [], Object.values(x.langs || {}));
    const nm = (xs) => xs.map((x) => '«' + (x.name || x.id) + '»').join(', ');
    const empty = list.filter((x) => parts(x).some((p) => !String(p.html || '').replace(/<[^>]*>/g, '').trim() && !/<img\b/i.test(p.html || '')));
    if (empty.length) out.push(['warn', t('Campaigns with an empty banner (no image, no text): {0}.', nm(empty)), null, t('Empty banner: {0}', empty.length)]);
    const imgs = (h) => [...String(h || '').matchAll(/<img\b[^>]*>/gi)].map((m) => m[0]);
    const wide = list.filter((x) => parts(x).some((p) => imgs(p.html).some((i) => Number((/\swidth\s*=\s*["']?(\d+)/i.exec(i) || [])[1]) > 600)));
    if (wide.length) out.push(['warn', t('Banners wider than 600 pixels (too wide on phones): {0}.', nm(wide)), null, t('Banner too wide: {0}', wide.length)]);
    const badLink = list.filter((x) => parts(x).some((p) => /href\s*=\s*["']http:\/\//i.test(p.html || '') || ((p.design || {}).tpl !== 'html' && (p.design || {}).link && !/^(https:|mailto:|tel:)/i.test(p.design.link.trim()) && p.design.link.trim() !== 'https://')));
    if (badLink.length) out.push(['warn', t('Banner links that are not https:// (or mailto:, tel:): {0}.', nm(badLink)), null, t('Banner links: {0}', badLink.length)]);
    const noImg = [...new Set(list.flatMap((x) => parts(x).flatMap((p) => [...String(p.html || '').matchAll(/src="assets\/([^"]+)"/g)].map((m) => m[1]))))].filter((n) => !S.assets.some((a) => a.name === n));
    if (noImg.length) out.push(['warn', t('Banner images not in the draft: {0}.', noImg.join(', ')), null, t('Banner images missing: {0}', noImg.length)]);
    const ids = new Set(c.signatures.map((x) => x.id));
    const lost = list.filter((x) => (x.signatures || []).length && (x.signatures || []).every((id) => !ids.has(id)));
    if (lost.length) out.push(['warn', t('Campaigns for signatures that are gone (nobody gets them): {0}.', nm(lost)), null, t('Campaigns without signatures: {0}', lost.length)]);
    // Two at the same time for the same signatures and people: one of them is shown.
    const range = (x) => [x.validFrom || '0000', x.validUntil || '9999'];
    const key = (x) => JSON.stringify([(x.signatures || []).slice().sort(), audienceText(x), x.newOnly ? 1 : 0, (x.days || []).join()]);
    const clash = list.filter((x) => list.some((o) => o !== x && key(o) === key(x) && range(x)[0] <= range(o)[1] && range(o)[0] <= range(x)[1]));
    if (clash.length) out.push(['warn', t('Campaigns at the same time for the same signatures and people, only one of them shows: {0}. Change the days, or choose other signatures.', nm(clash)), null, t('Campaigns at the same time: {0}', clash.length)]);
    // A banner's image: a text for blocked images, visible in dark mode, not heavy in every mail.
    const bimgs = (x) => parts(x).flatMap((p) => imgs(p.html));
    const noAlt = list.filter((x) => bimgs(x).some((i) => !/\salt\s*=\s*["'][^"']+["']/i.test(i)));
    if (noAlt.length) out.push(['warn', t('Banner images without a text for when images are blocked (alt): {0}. Without it, those recipients see nothing of the banner.', nm(noAlt)), null, t('Banner without alt text: {0}', noAlt.length)]);
    const nameOf = (i) => (/src="assets\/([^"?#]+)"/.exec(i) || [])[1];
    const size = (n) => { const a = S.assets.find((z) => z.name === n); return a ? a.size || 0 : 0; };
    const heavy = [...new Set(list.flatMap((x) => bimgs(x).map(nameOf)).filter((n) => n && size(n) > 100 * 1024))];
    if (heavy.length) out.push(['warn', t('Large banner images (over 100 KB), in every mail while the campaign runs: {0}.', heavy.map((n) => n + ' (' + Math.round(size(n) / 1024) + ' KB)').join(', ')), { label: t('Make images lighter'), fn: prepareClassic, leave: true }, t('Large banner images: {0}', heavy.length)]);
    const darkOnes = [...new Set(list.flatMap((x) => bimgs(x).map(nameOf)).filter((n) => n && dark[n] === true))];
    if (darkOnes.length) out.push(['warn', t('Dark banner image on a transparent background, hard to see in dark mode: {0}. A background of its own helps.', darkOnes.join(', ')), null, t('Banner hard to see in dark mode: {0}', darkOnes.length)]);
    const over = (c.campaigns || []).filter((x) => x.validUntil && x.validUntil < day);
    if (over.length) out.push(['good', t('Campaigns over: {0}. They can be removed (Campaigns → Over).', nm(over)), null, t('Campaigns over: {0}', over.length)]);
    return out;
  }
  // Everything to look at before publishing: what stops it, then the notes, the ones that matter
  // most first (who gets nothing, images missing), the looks after. What a recipient would see is
  // checked in the active signatures only: the ones someone gets, now or on days still to come.
  // The rest (templates, tests, ones over) are said in one grey line, not checked.
  function publishChecks() {
    (S.assets || []).forEach((a) => { if (/\.png$/i.test(a.name)) darkCheck(a.name); });
    // As it goes out: the building blocks put in.
    const c = prepared(), out = [], later = [], day = core.today();
    const errs = core.validate(c);
    errs.forEach((e) => out.push(['bad', t(e)]));
    // Where it goes: also with no signature left.
    if (onServer()) {
      if (!toSharePoint()) out.push(['warn', t("Publishing goes into the server's own folder, not to SharePoint: Microsoft Graph is not set up on the server."), null, t('Into the server\'s folder, not SharePoint')]);
      else if (lastSettings.folderError) out.push(['warn', t('SharePoint cannot be reached: {0}', t(lastSettings.folderError)), null, t('SharePoint cannot be reached')]);
    } else if (!S.settings.folder) out.push(['bad', t('No folder to publish to: set the synced OSign folder in Settings.')]);
    // None left (all deleted): publishing that takes them out of Outlook, said as the one thing
    // to know (the defaults, the users and the looks say nothing then), in full, not folded (the
    // flag at the end); asked once more on Publish.
    if (!c.signatures.length) {
      out.push(['warn', t('There is no signature: publishing takes all of them out of Outlook. From then on, nobody gets one by OSign.'), null, t('No signature: Outlook gets none'), true]);
      return out;
    }
    const idle = c.signatures.filter((s) => core.forNobody(s) || (core.isCampaign(s) && s.validUntil && s.validUntil < day));
    const active = c.signatures.filter((s) => !idle.includes(s)), ca = Object.assign({}, c, { signatures: active });
    // Who gets nothing for new mail: as read from AD under Users when there are users, else as the
    // defaults say. One note, not two.
    const plain = c.signatures.filter((s) => !(s.audience.departments || []).length && !(s.audience.domains || []).length && !(s.audience.groups || []).length && s.audience.everyone !== false);
    const snapUsers = S.snap && S.snap.users && S.snap.users.length ? S.snap.users : null;
    const g = snapUsers ? whoGets(snapUsers) : null, none = snapUsers ? snapUsers.filter((u) => !g.get(u).new).length : 0;
    if (!c.signatures.some((s) => s.defaultNew)) out.push(['warn', t('No signature is the default for new mail: users get none unless they pick one.'), null, t('No default for new mail')]);
    else if (none) out.push(['warn', t('{0} of the {1} users read from AD get no signature for new mail (Users → Only without a signature).', none, snapUsers.length) + (!plain.some((s) => s.defaultNew) ? ' ' + t('The default for new mail is limited to some (Who gets it).') : ''), null, t('{0} users without a signature', none)]);
    else if (!snapUsers && !plain.some((s) => s.defaultNew)) out.push(['warn', t('The default for new mail is limited to some users: everyone else gets none.'), null, t('The default for new mail is for some only')]);
    if (!c.signatures.some((s) => s.defaultReply)) out.push(['warn', t('No signature is the default for replies.')]);
    // Images the active signatures show that are not in the draft: left out of every mail. One of
    // the same name in another form (the GIF made a PNG, a -2) is offered in its place. Only the
    // server's copy endings (one or two digits) are left out: xmas-2025 is no xmas-2026.
    const named = (list) => [...new Set(list.flatMap((s) => [...(s.html || '').matchAll(/src="assets\/([^"]+)"/g)].map((m) => m[1])))].filter((n) => !S.assets.some((a) => a.name === n));
    const missing = named(active), stem = (n) => n.toLowerCase().replace(/(?:-\d{1,2})*\.[a-z0-9]+$/, '');
    const twin = (n) => (S.assets.find((a) => stem(a.name) === stem(n) && /\.(png|jpe?g|gif)$/i.test(a.name)) || {}).name || '';
    if (missing.length) {
      const swaps = missing.filter(twin);
      out.push(['warn', t('Images not in the draft, left out of every mail until added: {0}', missing.map((n) => twin(n) ? n + ' (→ ' + twin(n) + ')' : n).join(', ')),
        swaps.length ? { label: t("Use the draft's images instead"), fn: () => { swaps.forEach((n) => renameAsset(n, twin(n))); reopenChanged(); return swaps.length === 1 ? t('1 image put in place.') : t('{0} images put in place.', swaps.length); } } : null,
        t('Images missing in the draft: {0}', missing.length)]);
    }
    const big = classicHeavy().filter((x) => active.includes(x.s) || active.some((a) => a.id === x.s.id));
    if (big.length) out.push(['warn', t('Too large for classic Outlook to put in by itself (about 30 KB fit in the mailbox): {0}. There, images are left out, or Insert in the pane is needed. What helps: a smaller width, a JPEG instead of a PNG for photos, or fewer images.', big.map((x) => x.s.name + ' (~' + x.kb + ' KB)').join(', ')), { label: t('Prepare for classic Outlook'), fn: prepareClassic, leave: true }, t('Too large for classic Outlook: {0}', count(big.map((x) => x.s)))]);
    const lost = active.filter((s) => /<img[^>]+src="(?!assets\/|photos\/|https:)[^"]*"/i.test(s.html || '')).map((s) => s.name);
    if (lost.length) out.push(['warn', t('Images not stored with the signatures are left out in: {0}.', lost.join(', ')), null, t('Images left out: {0}', count(lost))]);
    const keys = fieldKeys();
    const used = [...new Set(active.flatMap((s) => core.placeholdersIn(s.html)))];
    const unknown = used.filter((k) => keys.indexOf(k) < 0);
    if (unknown.length) out.push(['warn', t('Unknown placeholders, left as typed: {0}', unknown.map((k) => '@@' + k).join(', ')), null, t('Unknown placeholders: {0}', unknown.length)]);
    // {{campaign}} is the campaigns' place, no building block.
    const noBlock = [...new Set(active.flatMap((s) => core.blocksIn(s.html)))].filter((k) => k.toLowerCase() !== 'campaign');
    if (noBlock.length) out.push(['warn', t('Unknown building blocks, left as typed: {0}', noBlock.map((k) => '{{' + k + '}}').join(', ')), null, t('Unknown building blocks: {0}', noBlock.length)]);
    const noGraph = used.filter((k) => { const f = c.fields.find((x) => x.key === k); return f && core.fieldSource(c, f) === 'none'; });
    if (noGraph.length) out.push(['warn', t('Empty in Outlook (no source): {0}', noGraph.map(fieldName).join(', ')), null, t('Fields empty in Outlook: {0}', noGraph.length)]);
    if (core.groupsUsed(c).length && core.valueSource(c) === 'graph') out.push(['warn', t('Signatures limited to AD groups: the app registration needs GroupMember.Read.All (delegated, with admin consent), and the groups have to be synced to Entra ID.'), null, t('Groups need GroupMember.Read.All')]);
    if (core.usesAdExport(c)) {
      const sn = S.snap;
      if (sn && S.snapMatches && !sn.written) later.push(['good', t('The users as read {0} are written: {1} new, {2} changed, {3} removed.', when(sn.readAt), sn.counts.new, sn.counts.changed, sn.counts.removed), null, t('Users: {0} new, {1} changed', sn.counts.new, sn.counts.changed)]);
      // As published, nothing to do: said, not a note.
      else if (sn && S.snapMatches && sn.written) later.push(['good', t('The users stay as published {0}. To update them, read and check them under Users first.', when(sn.readAt)), null, t('The users stay as published')]);
      else out.push(['warn', t('The users are not written: read and check them under Users first. Until then they stay as they are in the folder.'), null, t('The users are not written')]);
    }
    // How it looks: in the active signatures, then the banners.
    signatureChecks(ca).forEach((x) => out.push(x));
    campaignChecks(c).forEach((x) => out.push(x));
    const perPerson = active.filter((s) => /<img[^>]+src="photos\//i.test(s.html || '')).map((s) => s.name);
    if (perPerson.length) later.push(['good', t('Photos per person ({0}) come from photos\\<username>.jpg in the OSign folder; users without a photo get the signature without one.', perPerson.join(', ')), null, t('Photos per person')]);
    if (idle.length) {
      const gone = named(idle);
      later.push(['good', t('Nobody gets these, so they are not checked: {0} (Who gets it).', idle.map((s) => s.name).join(', ')) + (gone.length ? ' ' + t('Images they show that are not in the draft: {0}.', gone.join(', ')) : ''), null, t('Not checked (nobody gets them): {0}', count(idle))]);
    }
    const unused = unusedImages();
    if (unused.length) later.push(['good', t('Images in the draft that nothing uses: {0}. They go out all the same.', unused.join(', ')), { label: t('Remove them'), fn: removeUnused }, t('Unused images: {0}', unused.length)]);
    return out.concat(later);
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
    if (mod && !e.shiftKey && !e.altKey && e.key.toLowerCase() === 's') { e.preventDefault(); save().then((ok) => { if (ok) toast(t('Draft saved.')); }); return; }
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
    items.push({ kind: t('Go to'), label: t('Campaigns'), go: () => goView('campaigns') });
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
  // The quick search (Ctrl+K), under its field.
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
  // nothing red, the notes are folded under one line. A note flagged to stand alone (the fifth
  // entry: no signature left) is said in full, not folded.
  function renderChecks(checks) {
    const order = { bad: 0, warn: 1, good: 2 };
    const list = checks.slice().sort((a, b) => order[a[0]] - order[b[0]]);
    const bad = list.filter((c) => c[0] === 'bad'), warn = list.filter((c) => c[0] === 'warn'), good = list.filter((c) => c[0] === 'good');
    // leave: the dialog closed first, so that what the work says (and offers next) can be seen.
    // Else the checks made anew after it, the note typed kept, and what the work says (its
    // answer) shown in the dialog: a message under it would be hidden.
    const actBtn = (act) => el('button', { type: 'button', class: 'small', onclick: async (e) => {
      if (act.leave) { $('dlgPublish').close(); await act.fn(); return; }
      e.currentTarget.disabled = true;
      const said = await act.fn(), note = $('pubMsg').value;
      $('dlgPublish').close(); await openPublish();
      $('pubMsg').value = note;
      if (said) $('pubOk').textContent = said;
    } }, act.label);
    const item = ([k, text, act, short, alone]) => el('li', { class: k }, short && !alone
      ? el('details', { class: 'chk' }, el('summary', null, short), el('p', null, text), act ? actBtn(act) : null)
      : el('span', null, text, act ? actBtn(act) : null));
    const out = bad.map(item);
    if (warn.length) {
      const items = warn.map(item);
      if (bad.length || warn.some((c) => c[4])) out.push(...items);
      else out.push(el('li', { class: 'chk-head' }, el('details', { class: 'chk-all' },
        el('summary', null, el('b', null, t('Ready to publish.')), ' ', warn.length === 1 ? t('1 note') : t('{0} notes', warn.length)),
        el('ul', { class: 'checks chk-sub' }, ...items))));
    } else if (!bad.length) out.push(el('li', { class: 'good' }, el('b', null, t('Ready to publish.'))));
    if (good.length) out.push(el('li', { class: 'good chk-ok' }, good.map((g) => g[3] || g[1]).join(' · '),
      el('details', { class: 'chk' }, el('summary', null, t('details')), ...good.map((g) => el('p', null, g[1], g[2] ? ' ' : null, g[2] ? actBtn(g[2]) : null)))));
    $('pubChecks').replaceChildren(...out);
  }
  // Publish…: the checks (the PNGs used looked at for dark mode first, 3 seconds at most), where
  // it goes, what waits for approval.
  async function openPublish() {
    flushSource();
    if (core.usesAdExport(S.config)) { await saving; await loadSnapshot(); }
    const pngs = [...new Set(usedHtml().flatMap((h) => [...h.matchAll(/assets\/([^"'?#\s>]+\.png)/gi)].map((m) => m[1])))].filter((n) => (S.assets || []).some((a) => a.name === n));
    await Promise.race([Promise.all(pngs.map(darkCheck)), new Promise((r) => setTimeout(r, 3000))]);
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
  }
  $('btnPublish').addEventListener('click', () => openPublish());
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
    // With approval on, it is asked for, not published.
    const approval = !!(S.info && S.info.features && S.info.features.requireApproval);
    if (!prepared().signatures.length && !(await ask(t('Publish with no signature? Outlook then puts none in, for anyone.'), approval ? t('Ask for approval with none') : t('Publish with none'), true))) return;
    $('pubGo').disabled = true; $('pubErr').textContent = ''; $('pubOk').textContent = t('Publishing…');
    try {
      // A change still waiting for its save goes first, so the draft and what is published match.
      // Not saved (the server did not take it, or it was changed elsewhere): nothing goes out. A
      // save that failed before (since loaded anew, say) is tried once more first.
      let ok = await (saveTimer ? save() : saving);
      if (ok === false) ok = await save();
      if (ok === false) { $('pubOk').textContent = ''; $('pubErr').textContent = t('Not saved'); return; }
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
      // Published; only the images next to the add-in's files (for classic Outlook) did not go up.
      if (r.hostedError) $('pubErr').textContent = t('The images could not be put next to the add-in\'s files on Cloudflare Pages ({0}): classic Outlook takes the smaller copies in the mailbox until the next publish.', t(r.hostedError));
      // The users written now: their marks (new, changed) go on the Users page.
      if (core.usesAdExport(S.config)) loadSnapshot().then(() => { if (S.view === 'users') renderUsers(); });
      // Done: only Close is left.
      $('pubGo').hidden = true;
    } catch (err) { $('pubOk').textContent = ''; $('pubErr').textContent = t(err.message); }
    finally { $('pubGo').disabled = false; }
  });

  // The web certificate the proxy serves for this address: what it is, and whether PCs trust it.
  let tlsLast = null;
  // The web certificate of this address, for the status and the status bar.
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
  // The audit log as the filter asks, loaded anew.
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
  $('bakCheck').addEventListener('click', async () => {
    $('bakErr').textContent = ''; $('bakRestoreOk').textContent = '';
    const f = $('bakFile').files[0];
    if (!f) { $('bakErr').textContent = t('Choose the backup file.'); return; }
    $('bakCheck').disabled = true; $('bakRestoreOk').textContent = t('Opening it…');
    try {
      const r = await api('backup/check', { data: await b64Of(f), passphrase: $('bakRestorePass').value });
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
      const data = await b64Of(f);
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
  // A published version brought back into the draft (asked first).
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

