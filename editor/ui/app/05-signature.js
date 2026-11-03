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
    const m = $('ctxMenu'), i = S.config.signatures.indexOf(s), n = S.config.signatures.length;
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
    const move = (by) => () => {
      const list = S.config.signatures, j = list.indexOf(s);
      list.splice(j, 1); list.splice(j + by, 0, s); changed({ list: true });
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

