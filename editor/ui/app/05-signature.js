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
    if (on) S.config.signatures.forEach((x) => { if (x !== s && x[flag] && audienceText(x) === audienceText(s) && (x.language || '') === (s.language || '')) x[flag] = false; });
    s[flag] = on; changed();
  }
  // The panels of When and Who gets it: one at a time, closed by a click elsewhere or Esc.
  const propPanels = () => document.querySelectorAll('.sig-props > details');
  // Kept inside the page's area (not under the sidebar): as wide as fits, moved left as far as it
  // would stick out on the right. Again when the chips move (a longer summary, another row).
  function placePop(d) {
    const pop = d.querySelector('.pop'), box = d.closest('main').getBoundingClientRect();
    pop.style.left = '0px'; pop.style.maxWidth = Math.max(240, box.width - 32) + 'px';
    const over = pop.getBoundingClientRect().right - (box.right - 16);
    if (over > 0) pop.style.left = -Math.max(0, Math.min(over, d.getBoundingClientRect().left - box.left - 16)) + 'px';
  }
  const placeOpen = () => propPanels().forEach((d) => { if (d.open) placePop(d); });
  propPanels().forEach((d) => d.addEventListener('toggle', () => {
    if (!d.open) return;
    propPanels().forEach((o) => { if (o !== d) o.open = false; });
    placePop(d);
  }));
  if (window.ResizeObserver) {
    const ro = new ResizeObserver(placeOpen);
    ro.observe(document.querySelector('.sig-props'));
    propPanels().forEach((d) => ro.observe(d.querySelector('summary')));
  }
  addEventListener('resize', placeOpen);
  document.addEventListener('mousedown', (e) => {
    if (e.target.closest('.sig-props > details') || e.target.closest('.suggest')) return;
    propPanels().forEach((d) => { d.open = false; });
  });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') propPanels().forEach((d) => { if (d.open) { d.open = false; d.querySelector('summary').focus(); } }); });
  $('sigNew').addEventListener('change', () => setDefault('defaultNew', $('sigNew').checked));
  $('sigReply').addEventListener('change', () => setDefault('defaultReply', $('sigReply').checked));
  // A campaign: from / until.
  ['sigFrom', 'sigUntil'].forEach((id) => $(id).addEventListener('change', () => {
    const s = sig(); if (!s) return;
    s.validFrom = $('sigFrom').value; s.validUntil = $('sigUntil').value;
    $('campaignSummary').textContent = campaignText(s) || t('always'); placeOpen();
    changed({ list: true });
  }));
  // A language variant: the person's language comes from the field preferredLanguage, added when missing.
  $('sigLang').addEventListener('change', () => {
    const s = sig(); if (!s) return;
    s.language = $('sigLang').value;
    if (!s.language) delete s.language;
    $('sigLangChip').textContent = s.language ? ' · ' + s.language.toUpperCase() : '';
    if (s.language && !S.config.fields.some((f) => f.key === 'preferredLanguage')) {
      S.config.fields.push(core.fieldFor('preferredLanguage'));
      toast(t('Added the field «Language» (preferredLanguage): it tells which variant someone gets.'));
    }
    changed();
  });
  $('sigRemoveEmpty').addEventListener('change', () => { const s = sig(); if (s) { s.removeEmptyLines = $('sigRemoveEmpty').checked; changed({ list: false }); } });
  // Who gets it, as in the Microsoft 365 admin center: the entire organization, specific groups,
  // departments or sender domains (one field, each a token), or not assigned.
  const AUD_KINDS = { groups: { label: 'AD group', icon: '👥' }, departments: { label: 'Department', icon: '🏢' }, domains: { label: 'Sender domain', icon: '🌐' } };
  const audMode = (s) => {
    const a = s.audience || {};
    if ((a.groups || []).length || (a.departments || []).length || (a.domains || []).length || s.__audSome) return 'some';
    return a.everyone === false ? 'none' : 'all';
  };
  function audienceShow(s) {
    const a = s.audience || (s.audience = { groups: [], departments: [], domains: [] }), mode = audMode(s);
    $('audienceSummary').textContent = audienceText(s); placeOpen();
    $('audienceSummary').classList.toggle('warn', core.forNobody(s));
    document.querySelectorAll('input[name=audMode]').forEach((r) => { r.checked = r.value === mode; });
    $('audSome').hidden = mode !== 'some';
    $('audNone').hidden = mode !== 'none';
    const toks = [];
    for (const kind of Object.keys(AUD_KINDS)) {
      for (const v of a[kind] || []) {
        toks.push(el('span', { class: 'tok tok-' + kind, title: t(AUD_KINDS[kind].label) },
          el('span', { class: 'tok-i', 'aria-hidden': 'true' }, AUD_KINDS[kind].icon), kind === 'domains' ? '@' + v : v,
          el('button', { type: 'button', class: 'tok-x', 'aria-label': t('Remove {0}', v), onclick: () => audRemove(kind, v) }, '×')));
      }
    }
    $('audTokens').replaceChildren(...toks, $('audInput'));
  }
  function audAdd(kind, value) {
    const s = sig(); if (!s) return;
    let v = String(value || '').trim(); if (!v) return;
    if (kind === 'domains') v = v.replace(/^@/, '').toLowerCase();
    const list = s.audience[kind] || (s.audience[kind] = []);
    if (!list.some((x) => x.toLowerCase() === v.toLowerCase())) list.push(v);
    s.audience.everyone = false; delete s.__audSome;
    audienceShow(s); changed();
    $('audInput').value = ''; $('audInput').focus();
  }
  function audRemove(kind, v) {
    const s = sig(); if (!s) return;
    s.audience[kind] = (s.audience[kind] || []).filter((x) => x !== v);
    if (!s.audience.groups.length && !s.audience.departments.length && !s.audience.domains.length) s.__audSome = true;
    audienceShow(s); changed();
  }
  // What was there before "not assigned" or "entire organization": back when "specific" is chosen again.
  const audKept = new Map();
  document.querySelectorAll('input[name=audMode]').forEach((r) => r.addEventListener('change', () => {
    const s = sig(); if (!s || !r.checked) return;
    const a = s.audience, had = { groups: a.groups || [], departments: a.departments || [], domains: a.domains || [] };
    if (had.groups.length || had.departments.length || had.domains.length) audKept.set(s.id, had);
    if (r.value === 'some') {
      Object.assign(a, audKept.get(s.id) || { groups: [], departments: [], domains: [] }, { everyone: false });
      s.__audSome = true;
    } else {
      Object.assign(a, { groups: [], departments: [], domains: [], everyone: r.value === 'all' });
      delete s.__audSome;
    }
    audienceShow(s); changed();
    if (r.value === 'some') $('audInput').focus();
  }));
  // Typed: picked from AD's groups and departments, or as typed (@… a domain, else a group).
  suggest($('audInput'), { min: 2, fetch: async (q) => {
    if (/^@/.test(q)) return [{ value: q.replace(/^@/, ''), label: '@' + q.replace(/^@/, ''), sub: t('Sender domain'), kind: 'domains' }];
    if (!adOn()) return [];
    const [g, d] = await Promise.all([api('ad/groups', { query: q }).then((r) => r.groups).catch(() => []), api('ad/departments', { query: q }).then((r) => r.departments).catch(() => [])]);
    return g.map((x) => ({ value: x, label: x, sub: t('AD group'), kind: 'groups' })).concat(d.map((x) => ({ value: x, label: x, sub: t('Department'), kind: 'departments' })));
  }, pick: (it) => audAdd(it.kind, it.value) });
  $('audInput').addEventListener('keydown', (e) => {
    if (e.defaultPrevented) return;
    const v = $('audInput').value.trim();
    if ((e.key === 'Enter' || e.key === ',' || e.key === ';') && v) {
      e.preventDefault();
      audAdd(/^@/.test(v) || /^[a-z0-9-]+(\.[a-z0-9-]+)+$/i.test(v) ? 'domains' : 'groups', v);
    } else if (e.key === 'Backspace' && !v) {
      const s = sig(); if (!s) return;
      for (const kind of ['domains', 'departments', 'groups']) { const l = s.audience[kind] || []; if (l.length) { audRemove(kind, l[l.length - 1]); break; } }
    }
  });
  $('audTokens').addEventListener('click', (e) => { if (e.target === $('audTokens')) $('audInput').focus(); });
  // ---------------------------------------------------------------------------------------------
  // Right-click on a signature in the list (or the menu key, or Shift+F10): what the signature's
  // own page offers, without opening it first.
  // ---------------------------------------------------------------------------------------------
  function closeMenu() {
    const m = $('ctxMenu'); if (m.hidden) return;
    m.hidden = true; m.textContent = '';
    if (closeMenu.back && document.contains(closeMenu.back)) closeMenu.back.focus();
  }
  $('sigPickedDel').addEventListener('click', () => deletePicked());
  $('sigPickedNone').addEventListener('click', () => pickClear());
  function sigMenu(s, x, y) {
    closeMenu();
    // On one of several picked: what is done with all of them.
    if (picked.size > 1 && picked.has(s.id)) {
      const m0 = $('ctxMenu');
      const b = (label, run, danger) => { const x0 = el('button', { type: 'button', role: 'menuitem', class: danger ? 'danger' : '' }, el('span', { class: 'check', 'aria-hidden': 'true' }), label); x0.addEventListener('click', () => { closeMenu.back = null; closeMenu(); run(); }); return x0; };
      closeMenu.back = document.activeElement;
      m0.replaceChildren(b(t('Delete {0} signatures…', picked.size), deletePicked, true), el('hr'), b(t('Cancel the selection'), pickClear));
      m0.hidden = false;
      const r0 = m0.getBoundingClientRect();
      m0.style.left = Math.min(x, innerWidth - r0.width - 8) + 'px'; m0.style.top = Math.min(y, innerHeight - r0.height - 8) + 'px';
      m0.querySelector('button').focus();
      return;
    }
    closeMenu.back = document.activeElement;
    const g0 = sigGroups(), grp = core.forNobody(s) && g0.assigned.length ? g0.unassigned : core.forNobody(s) ? g0.unassigned : g0.assigned;
    const m = $('ctxMenu'), i = grp.indexOf(s), n = grp.length;
    const item = (label, run, opts) => {
      opts = opts || {};
      const b = el('button', { type: 'button', role: opts.check === undefined ? 'menuitem' : 'menuitemcheckbox', class: opts.danger ? 'danger' : '' },
        el('span', { class: 'check', 'aria-hidden': 'true' }, opts.check ? '✓' : ''), label, opts.key ? el('span', { class: 'key' }, opts.key) : null);
      if (opts.check !== undefined) b.setAttribute('aria-checked', String(!!opts.check));
      b.disabled = !!opts.disabled;
      b.addEventListener('click', () => { closeMenu.back = null; closeMenu(); run(); });
      return b;
    };
    const asDefault = (flag, id) => () => { openSig(s.id); $(id).checked = !s[flag]; $(id).dispatchEvent(new Event('change')); };
    // Up and down as the list shows them: within its group (assigned, or nobody's).
    const move = (by) => () => {
      const list = S.config.signatures, k = grp.indexOf(s), other = grp[k + by];
      if (!other) return;
      const a = list.indexOf(s), b = list.indexOf(other);
      list[a] = other; list[b] = s; changed({ list: true });
    };
    m.replaceChildren(
      item(t('Open'), () => openSig(s.id)),
      item(t('Rename'), () => { openSig(s.id); $('sigName').focus(); $('sigName').select(); }),
      item(t('Duplicate'), () => { openSig(s.id); flushSource(); addSig(s); }),
      el('hr'),
      item(t('Default for new mail'), asDefault('defaultNew', 'sigNew'), { check: !!s.defaultNew }),
      item(t('Default for replies & forwards'), asDefault('defaultReply', 'sigReply'), { check: !!s.defaultReply }),
      el('hr'),
      item(t('Move up'), move(-1), { disabled: i <= 0 }),
      item(t('Move down'), move(1), { disabled: i < 0 || i >= n - 1 }),
      el('hr'),
      item(t('Delete…'), () => { openSig(s.id); $('btnDel').click(); }, { danger: true }));
    m.hidden = false;
    // Inside the window, also near its right or bottom edge.
    const w = m.offsetWidth, h = m.offsetHeight;
    m.style.left = Math.max(4, Math.min(x, innerWidth - w - 4)) + 'px';
    m.style.top = Math.max(4, Math.min(y, innerHeight - h - 4)) + 'px';
    m.querySelector('button:not(:disabled)').focus();
  }
  $('ctxMenu').addEventListener('keydown', (e) => {
    const items = [...$('ctxMenu').querySelectorAll('button:not(:disabled)')], at = items.indexOf(document.activeElement);
    if (e.key === 'Escape' || e.key === 'Tab') { e.preventDefault(); closeMenu(); }
    else if (e.key === 'ArrowDown') { e.preventDefault(); items[(at + 1) % items.length].focus(); }
    else if (e.key === 'ArrowUp') { e.preventDefault(); items[(at - 1 + items.length) % items.length].focus(); }
    else if (e.key === 'Home') { e.preventDefault(); items[0].focus(); }
    else if (e.key === 'End') { e.preventDefault(); items[items.length - 1].focus(); }
  });
  document.addEventListener('mousedown', (e) => { if (!$('ctxMenu').contains(e.target)) { closeMenu.back = null; closeMenu(); } });
  addEventListener('blur', () => { closeMenu.back = null; closeMenu(); });
  addEventListener('resize', () => closeMenu());
  document.addEventListener('scroll', () => closeMenu(), true);

  // What is published (S.pub: id → signature as it went out), to see what the draft changed and to
  // go back to it, one signature at a time. Nothing published yet: no marks.
  S.pub = null;
  async function loadPublishedCopy() {
    let c = null;
    try { const r = await api('published/get'); c = r.configJson ? JSON.parse(r.configJson) : null; } catch (e) { c = null; }
    S.pub = c && Array.isArray(c.signatures) ? new Map(c.signatures.map((x) => [x.id, x])) : null;
    // As the editor reads it, for comparing the rest (fields, building blocks, options).
    S.pubCfg = c ? normalize(JSON.parse(JSON.stringify(c))) : null;
    renderList(); revertShow(); renderStatus();
  }
  // What a signature is made of, as one text: its HTML as written and the rules who gets it.
  const sigKey = (s, html) => JSON.stringify([s.name || '', (html || '').replace(/\s+/g, ' ').trim(), !!s.defaultNew, !!s.defaultReply, s.validFrom || '', s.validUntil || '',
    s.language || '', s.removeEmptyLines !== false, !!(s.audience && s.audience.everyone !== false),
    ...['groups', 'departments', 'domains'].map((k) => ((s.audience && s.audience[k]) || []).map((x) => String(x).trim().toLowerCase()).sort().join(','))]);
  // 'same', 'changed', 'new' (not published), or '' (nothing published at all).
  function sigState(s) {
    if (!S.pub || !s) return '';
    const p = S.pub.get(s.id);
    if (!p) return 'new';
    return sigKey(s, sanitize(s.html)) === sigKey(p, sanitize(typeof p.source === 'string' ? p.source : p.html)) ? 'same' : 'changed';
  }
  // What changed in a signature against what is published, in words: its parts that differ.
  function sigDiff(s, p) {
    const out = [], same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
    const aud = (x) => [!!(x.audience && x.audience.everyone !== false), ...['groups', 'departments', 'domains'].map((k) => ((x.audience && x.audience[k]) || []).map((v) => String(v).trim().toLowerCase()).sort().join(','))];
    if ((s.name || '') !== (p.name || '')) out.push(t('name'));
    if (sanitize(s.html).replace(/\s+/g, ' ').trim() !== sanitize(typeof p.source === 'string' ? p.source : p.html).replace(/\s+/g, ' ').trim()) out.push(t('text and layout'));
    if (!same(aud(s), aud(p))) out.push(t('who gets it'));
    if ((s.validFrom || '') !== (p.validFrom || '') || (s.validUntil || '') !== (p.validUntil || '')) out.push(t('when'));
    if (!!s.defaultNew !== !!p.defaultNew || !!s.defaultReply !== !!p.defaultReply) out.push(t('default for'));
    if ((s.language || '') !== (p.language || '')) out.push(t('language'));
    if ((s.removeEmptyLines !== false) !== (p.removeEmptyLines !== false)) out.push(t('empty lines'));
    return out;
  }
  // A part of the configuration other than the signatures, against what is published: both as the
  // editor reads them, objects whatever the order of their keys, HTML as it shows.
  const stableJson = (v) => JSON.stringify(v, (k, x) => (x && typeof x === 'object' && !Array.isArray(x) ? Object.keys(x).sort().reduce((o, key) => { if (x[key] !== undefined && x[key] !== '') o[key] = x[key]; return o; }, {}) : x));
  const CFG_VIEW = { blocks: (c) => (c.blocks || []).map((b) => [b.key, b.name, sanitize(b.html || '').replace(/\s+/g, ' ').trim()]), fields: (c) => c.fields || [], options: (c) => c.options || {} };
  function cfgPartChanged(k, pubN) { return !!pubN && stableJson(CFG_VIEW[k](S.config || {})) !== stableJson(CFG_VIEW[k](pubN)); }
  // The draft the same in content as what is published (whatever way it is saved): then nothing
  // waits to go out, also after undoing a change.
  function sameAsPublished() {
    if (!S.pub || !S.pubCfg || !S.config) return false;
    const sigs = S.config.signatures || [];
    if (sigs.length !== S.pub.size || sigs.map((s) => s.id).join() !== [...S.pub.keys()].join()) return false;
    if (sigs.some((s) => sigState(s) !== 'same')) return false;
    return !['fields', 'blocks', 'options'].some((k) => cfgPartChanged(k, S.pubCfg));
  }
  async function openChanges() {
    flushSource();
    let pubCfg = null;
    try { const r = await api('published/get'); pubCfg = r.configJson ? JSON.parse(r.configJson) : null; } catch (e) { pubCfg = null; }
    const pub = new Map(((pubCfg && pubCfg.signatures) || []).map((x) => [x.id, x]));
    const rows = [];
    const row = (cls, name, what, ...btns) => el('li', { class: cls }, el('div', { class: 'chg-text' }, el('b', null, name), el('small', null, what)), el('div', { class: 'chg-btns' }, ...btns));
    const btn = (label, fn, danger) => el('button', { type: 'button', class: 'small' + (danger ? ' ghost danger' : ' ghost'), onclick: fn }, label);
    const open = (id) => () => { $('dlgChanges').close(); openSig(id); };
    const revert = (id) => async () => { $('dlgChanges').close(); openSig(id); fromChanges = true; $('btnRevert').click(); };
    for (const s of S.config.signatures) {
      const p = pub.get(s.id);
      if (!p) rows.push(row('new', s.name, t('New, not published yet.'), btn(t('Open'), open(s.id)), btn(t('Delete…'), () => { $('dlgChanges').close(); openSig(s.id); $('btnDel').click(); }, true)));
      else { const d = sigDiff(s, p); if (d.length) rows.push(row('chg', s.name, t('Changed: {0}.', d.join(', ')), btn(t('Compare'), () => { $('dlgChanges').close(); openCompare(s.id); }), btn(t('Open'), open(s.id)), btn(t('Discard changes'), revert(s.id), true))); }
    }
    for (const [id, p] of pub) if (!S.config.signatures.some((x) => x.id === id)) {
      rows.push(row('gone', p.name, t('Removed: still in Outlook until you publish.'), btn(t('Bring back'), async () => {
        try {
          const r = await api('published/revert', { id });
          const back = Object.assign(core.newSignature(r.signature.id, r.signature.name), r.signature);
          if (typeof back.source === 'string') back.html = back.source;
          delete back.source; delete back.compiled;
          back.audience = Object.assign({ departments: [], domains: [], groups: [] }, back.audience || {});
          S.config.signatures.push(back); S.assets = r.assets || S.assets; changed(); openChanges();
        } catch (e) { toast(t(e.message), true); }
      })));
    }
    // The rest of the configuration: said, discarded only with all of it. Both sides as the editor
    // reads them (a version published by an older editor has other defaults, no names on its
    // building blocks, its options in another order), the HTML as it shows: only a real change counts.
    const pubN = pubCfg ? normalize(JSON.parse(JSON.stringify(pubCfg))) : null;
    const part = (k) => cfgPartChanged(k, pubN);
    const other = [part('fields') && t('Fields'), part('blocks') && t('Building blocks'), part('options') && t('Add-in')].filter(Boolean);
    if (other.length) rows.push(row('chg', other.join(', '), t('Changed too: they go out with the next publish.')));
    // Nothing to show while the status says otherwise: the same in content, saved another way.
    const none = S.dirty ? t('Nothing to see: the draft is the same as what is published, only saved another way (e.g. by a newer version). Publishing again makes them one.') : t('Nothing: the draft is as published.');
    $('chgList').replaceChildren(...(rows.length ? rows : [el('li', { class: 'none' }, none)]));
    $('dlgChanges').showModal();
  }
  $('chgAll').addEventListener('click', () => { $('dlgChanges').close(); $('btnLoad').click(); });
  $('chgPublish').addEventListener('click', () => { $('dlgChanges').close(); $('btnPublish').click(); });
  function revertShow() { const s = sig(); $('btnRevert').hidden = $('btnCompare').hidden = sigState(s) !== 'changed'; }
  // Before and after: the signature as published and as it is now, each with the preview's values,
  // and what changed in who gets it, when and as default.
  function openCompare(id) {
    flushSource();
    const s = S.config.signatures.find((x) => x.id === id), p = S.pub && S.pub.get(id);
    if (!s || !p) return;
    const values = previewValues();
    const page = (html) => '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:12px 14px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222;zoom:.9}p{margin:0}table{border-collapse:collapse}</style></head><body>' + html + '</body></html>';
    const show = (frame, html, removeEmpty) => {
      frame.srcdoc = page(core.render(core.compile(sanitize(html), scratchDoc), values, { assetBase: '/', removeEmptyLines: removeEmpty }));
      frame.onload = () => { const d = frame.contentDocument; if (d) d.querySelectorAll('img').forEach((i) => { if (i.complete && !i.naturalWidth) i.remove(); else i.addEventListener('error', () => i.remove()); }); };
    };
    show($('cmpBefore'), p.html || '', p.removeEmptyLines !== false);
    show($('cmpAfter'), sigFull(s), s.removeEmptyLines !== false);
    $('cmpTitle').textContent = t('Before and after: {0}', s.name);
    const src = $('pvSource');
    $('cmpLead').textContent = t('With the values: {0}. What changed: {1}.', src.options[src.selectedIndex].textContent, sigDiff(s, p).join(', ') || t('nothing'));
    const aud = (x) => { const a = x.audience || {}; return audienceText(x) || (a.everyone === false ? t('not assigned') : t('entire organization')); };
    const yes = (b) => (b ? t('yes') : t('no'));
    const rules = [
      [t('Name'), p.name || '', s.name || ''],
      [t('Who gets it'), aud(p), aud(s)],
      [t('When'), campaignText(p) || t('always'), campaignText(s) || t('always')],
      [t('Default for new mail'), yes(p.defaultNew), yes(s.defaultNew)],
      [t('Default for replies & forwards'), yes(p.defaultReply), yes(s.defaultReply)],
      [t('Language'), (p.language || '').toUpperCase() || t('All languages'), (s.language || '').toUpperCase() || t('All languages')],
      [t('Leave out empty lines'), yes(p.removeEmptyLines !== false), yes(s.removeEmptyLines !== false)]
    ].filter((r) => r[1] !== r[2]);
    $('cmpRules').closest('table').hidden = !rules.length;
    $('cmpRules').replaceChildren(...rules.map(([k, a, b]) => el('tr', null, el('th', null, k), el('td', { class: 'was' }, a), el('td', { class: 'now' }, b))));
    $('cmpRevert').onclick = () => { $('dlgCompare').close(); openSig(id); $('btnRevert').click(); };
    $('dlgCompare').showModal();
  }
  $('btnCompare').addEventListener('click', () => { const s = sig(); if (s) openCompare(s.id); });
  let fromChanges = false;
  $('btnRevert').addEventListener('click', async () => {
    const s = sig(); if (!s) return;
    if (!(await ask(t('Discard the changes to «{0}» since it was published? It goes back to what Outlook has.', s.name), t('Discard changes'), true))) { fromChanges = false; return; }
    flushSource();
    try {
      const r = await api('published/revert', { id: s.id });
      const before = JSON.parse(JSON.stringify(s)), i = S.config.signatures.indexOf(s);
      const back = Object.assign(core.newSignature(r.signature.id, r.signature.name), r.signature);
      if (typeof back.source === 'string') back.html = back.source;
      delete back.source; delete back.compiled;
      back.audience = Object.assign({ departments: [], domains: [], groups: [] }, back.audience || {});
      S.config.signatures[i] = back; S.assets = r.assets || S.assets;
      changed(); openSig(back.id);
      // Saved at once, so the status at the top and the marks say it right away.
      await save(); await loadPublishedCopy();
      if (fromChanges) { fromChanges = false; openChanges(); }
      toast(t('«{0}» is as published again.', back.name), false, { label: t('Undo'), run: () => { const j = S.config.signatures.findIndex((x) => x.id === before.id); if (j >= 0) { S.config.signatures[j] = before; changed(); openSig(before.id); } } });
    } catch (e) { toast(t(e.message), true); }
  });

  // New signature: first the layout, each shown with the sample values.
  function openTemplates() {
    const silhouette = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" width="80" height="100"><rect width="80" height="100" fill="#e3e7ee"/><circle cx="40" cy="38" r="17" fill="#b9c2d0"/><path d="M8 100c2-24 16-34 32-34s30 10 32 34z" fill="#b9c2d0"/></svg>');
    const values = {};
    S.config.fields.forEach((f) => { values[f.key] = f.sample || ''; });
    $('tplList').replaceChildren(...TEMPLATES.map((tp) => {
      const keys = core.placeholdersIn(tp.html);
      keys.forEach((k) => { if (!(k in values)) values[k] = core.fieldFor(k).sample || ''; });
      const html = core.render(core.compile(sanitize(tp.html.replace(/photos\/@@sAMAccountName\.jpg/g, silhouette)), scratchDoc), values, { assetBase: '/' });
      const frame = el('iframe', { tabindex: '-1', 'aria-hidden': 'true', sandbox: '' });
      frame.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:14px 16px;font-family:Arial,sans-serif;font-size:10pt;color:#222}p{margin:0}table{border-collapse:collapse}</style></head><body>' + html + '</body></html>';
      const b = el('button', { type: 'button', class: 'tpl-card', 'data-tpl': tp.key },
        frame, el('span', { class: 'tpl-name' }, t(tp.name)), el('span', { class: 'tpl-text' }, t(tp.text)));
      b.addEventListener('click', () => { $('dlgTpl').close(); addSig(null, tp.html); });
      return b;
    }), (() => {
      // Or what there is already: templates from CI-Sign or classic Outlook.
      const art = el('span', { class: 'tpl-art', 'aria-hidden': 'true' });
      art.innerHTML = '<svg viewBox="0 0 24 24"><path d="M12 4v11M7 10l5 5 5-5"/><path d="M5 19h14"/></svg>';
      const b = el('button', { type: 'button', class: 'tpl-card', 'data-tpl': 'import' }, art,
        el('span', { class: 'tpl-name' }, t('Import templates')), el('span', { class: 'tpl-text' }, t('From CI-Sign or classic Outlook (.htm with its images).')));
      b.addEventListener('click', () => { $('dlgTpl').close(); $('btnImportOpen').click(); });
      return b;
    })());
    $('dlgTpl').showModal();
    $('tplList').querySelector('.tpl-card').focus();
  }
  $('btnNew').addEventListener('click', openTemplates);
  $('btnNew2').addEventListener('click', openTemplates);
  $('btnImport2').addEventListener('click', () => $('btnImportOpen').click());
  $('btnDup').addEventListener('click', () => { flushSource(); const s = sig(); if (s) addSig(s); });
  $('btnDel').addEventListener('click', async () => {
    const s = sig(); if (!s || !(await ask(t('Delete the signature "{0}"? It goes from Outlook with the next publish.', s.name), t('Delete'), true))) return;
    S.config.signatures = S.config.signatures.filter((x) => x !== s);
    changed(); openSig(S.config.signatures[0] ? S.config.signatures[0].id : null);
  });
  document.querySelectorAll('.tab[data-mode]').forEach((t) => t.addEventListener('click', () => setMode(t.dataset.mode)));
  ['imgW', 'imgAlt', 'imgHref'].forEach((id) => $(id).addEventListener('change', applyImg));
  $('imgDone').addEventListener('click', () => { applyImg(); hideImgBar(); });


  // Campaigns on a timeline: from today (or the first one) to the last day, a month per tick; the
  // ones for the same people at the same time marked.
  function renderCampaigns() {
    const camps = S.config.signatures.filter((s) => core.isCampaign(s));
    const DAY = 86400000, day = (x) => Date.parse(x + 'T00:00:00');
    const now = new Date(); const today = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) + now.getTimezoneOffset() * 60000;
    const froms = camps.map((s) => (s.validFrom ? day(s.validFrom) : null)).filter((x) => x !== null);
    const untils = camps.map((s) => (s.validUntil ? day(s.validUntil) + DAY : null)).filter((x) => x !== null);
    const start = Math.min(today - 7 * DAY, ...froms), end = Math.max(today + 60 * DAY, ...untils);
    const span = end - start, pct = (x) => (100 * (x - start) / span).toFixed(2) + '%';
    const loc = I18N.lang === 'de' ? 'de-CH' : 'en-GB';
    const range = (s) => [s.validFrom ? day(s.validFrom) : start, s.validUntil ? day(s.validUntil) + DAY : end];
    const clash = (a, b) => a !== b && audienceText(a) === audienceText(b) && (a.language || '') === (b.language || '') && range(a)[0] < range(b)[1] && range(b)[0] < range(a)[1];
    const ticks = [];
    for (let d = new Date(start); d.getTime() <= end; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) {
      const m = new Date(d.getFullYear(), d.getMonth(), 1).getTime();
      if (m >= start) ticks.push(el('span', { class: 'camp-tick', style: 'left:' + pct(m) }, new Date(m).toLocaleDateString(loc, { month: 'short', year: '2-digit' })));
    }
    const fmt = (x) => (x ? new Date(day(x)).toLocaleDateString(loc) : '…');
    $('campTimeline').replaceChildren(
      el('div', { class: 'camp-axis' }, ...ticks, el('span', { class: 'camp-today', style: 'left:' + pct(today), title: t('Today') })),
      ...camps.map((s) => {
        const [a, b] = range(s), over = camps.some((o) => clash(s, o));
        const bar = el('button', { type: 'button', class: 'camp-bar' + (core.inPeriod(s) ? ' now' : b <= today ? ' past' : '') + (over ? ' clash' : ''),
          style: 'left:' + pct(a) + ';width:calc(' + pct(b) + ' - ' + pct(a) + ')',
          title: fmt(s.validFrom) + ' - ' + fmt(s.validUntil) + ' · ' + audienceText(s) + (over ? ' · ' + t('at the same time as another one for the same people') : ''),
          onclick: () => { $('dlgCampaigns').close(); openSig(s.id); } });
        return el('div', { class: 'camp-row' }, el('span', { class: 'camp-name' }, s.name || t('(no name)')), el('div', { class: 'camp-track' }, bar, el('span', { class: 'camp-today', style: 'left:' + pct(today) })));
      }),
      ...(camps.length ? [] : [el('p', { class: 'muted' }, t('No campaigns: a signature becomes one with its days under When.'))]));
  }
  $('btnCampaigns').addEventListener('click', () => { renderCampaigns(); $('dlgCampaigns').showModal(); });
