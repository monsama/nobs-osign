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
  // A campaign: from / until.
  ['sigFrom', 'sigUntil'].forEach((id) => $(id).addEventListener('change', () => {
    const s = sig(); if (!s) return;
    s.validFrom = $('sigFrom').value; s.validUntil = $('sigUntil').value;
    $('campaignSummary').textContent = campaignText(s) || t('always');
    changed({ list: true });
  }));
  $('sigRemoveEmpty').addEventListener('change', () => { const s = sig(); if (s) { s.removeEmptyLines = $('sigRemoveEmpty').checked; changed({ list: false }); } });
  $('sigGroups').addEventListener('input', () => { const s = sig(); if (s) { s.audience.groups = splitList($('sigGroups').value); $('audienceSummary').textContent = audienceText(s); changed({ preview: false }); } });
  // Group names and departments from AD to pick from, for what is typed after the last comma.
  suggest($('sigGroups'), { multi: true, fetch: async (q) => (adOn() ? (await api('ad/groups', { query: q })).groups : []).map((g) => ({ value: g, label: g })) });
  suggest($('sigDeps'), { multi: true, min: 0, fetch: async (q) => (adOn() ? (await api('ad/departments', { query: q })).departments : []).map((g) => ({ value: g, label: g })) });
  $('sigDeps').addEventListener('input', () => { const s = sig(); if (s) { s.audience.departments = splitList($('sigDeps').value); $('audienceSummary').textContent = audienceText(s); changed({ preview: false }); } });
  $('sigDoms').addEventListener('input', () => { const s = sig(); if (s) { s.audience.domains = splitList($('sigDoms').value).map((d) => d.replace(/^@/, '').toLowerCase()); $('audienceSummary').textContent = audienceText(s); changed({ preview: false }); } });
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

  $('btnNew').addEventListener('click', () => addSig());
  $('btnNew2').addEventListener('click', () => addSig());
  $('btnDup').addEventListener('click', () => { flushSource(); const s = sig(); if (s) addSig(s); });
  $('btnDel').addEventListener('click', async () => {
    const s = sig(); if (!s || !(await ask(t('Delete the signature "{0}"? It goes from Outlook with the next publish.', s.name), t('Delete'), true))) return;
    S.config.signatures = S.config.signatures.filter((x) => x !== s);
    changed(); openSig(S.config.signatures[0] ? S.config.signatures[0].id : null);
  });
  document.querySelectorAll('.tab[data-mode]').forEach((t) => t.addEventListener('click', () => setMode(t.dataset.mode)));
  ['imgW', 'imgAlt', 'imgHref'].forEach((id) => $(id).addEventListener('change', applyImg));
  $('imgDone').addEventListener('click', () => { applyImg(); hideImgBar(); });

