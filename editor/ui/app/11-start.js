  // ---------------------------------------------------------------------------------------------
  // Start
  // ---------------------------------------------------------------------------------------------
  // A configuration from an older editor or another hand gets whatever it lacks, so the rest of
  // the UI can rely on every part being there.
  function normalize(c) {
    const d = core.defaultConfig();
    c = c && typeof c === 'object' ? c : d;
    // Published with a pilot by 0.9.57–0.9.60 (loaded or brought back from History): its newer version.
    if (c.pilot && c.pilot.config && typeof c.pilot.config === 'object') c = c.pilot.config;
    delete c.pilot;
    // One made before the value source could be chosen took its values from Entra ID: it stays so
    // (new setups start with Entra ID and AD).
    if (c !== d && !(c.options && c.options.valueSource)) c.options = Object.assign({}, c.options || {}, { valueSource: 'graph' });
    c.schemaVersion = core.SCHEMA_VERSION;
    c.options = Object.assign({}, d.options, c.options || {});
    delete c.auth; delete c.assetBaseUrl;
    if (!Array.isArray(c.fields)) c.fields = d.fields;
    c.blocks = (Array.isArray(c.blocks) ? c.blocks : []).filter((b) => b && b.key).map((b) => ({ key: String(b.key), name: String(b.name || b.key), html: String(b.html || '') }));
    c.signatures = (Array.isArray(c.signatures) ? c.signatures : []).map((s) => {
      const n = Object.assign(core.newSignature(s.id, s.name), s);
      // As written, with its building blocks as {{key}} (html is them put in, for Outlook).
      if (typeof s.source === 'string') { n.html = s.source; delete n.source; }
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
    loadPublishedCopy();
    // The logos looked at for dark mode, for the checks before publishing.
    setTimeout(() => (S.assets || []).forEach((a) => { if (/\.png$/i.test(a.name)) darkCheck(a.name); }), 500);
    if (S.config.signatures.length) openSig(S.sigId && S.config.signatures.some((s) => s.id === S.sigId) ? S.sigId : S.config.signatures[0].id);
    else show('empty');
  }

  let toastTimer = null;
  // A question with its answer as a button (Delete, Restore …), not the browser's OK: true when
  // that was chosen. danger: the button in red, and Cancel the one ready for Enter.
  // With a value (askFor): a field to answer in; the answer, or null when cancelled.
  function askFor(text, value, okLabel) {
    $('askInput').hidden = false; $('askInput').value = value || '';
    const p = ask(text, okLabel || t('OK'), false, null, true).then((yes) => (yes ? $('askInput').value : null));
    return p;
  }
  function ask(text, okLabel, danger, cancelLabel, withInput) {
    const d = $('dlgAsk');
    if (!withInput) $('askInput').hidden = true;
    $('askText').textContent = text;
    $('askOk').textContent = okLabel || t('OK'); $('askOk').className = danger ? 'danger-btn' : 'primary';
    $('askCancel').textContent = cancelLabel || t('Cancel');
    d.returnValue = '';
    return new Promise((resolve) => {
      d.addEventListener('close', () => { $('askInput').hidden = true; resolve(d.returnValue === 'ok'); }, { once: true });
      d.showModal();
      if (withInput) { $('askInput').focus(); $('askInput').select(); }
      else (danger ? $('askCancel') : $('askOk')).focus();
    });
  }
  // Enter in the answer field: as OK.
  $('askInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); $('dlgAsk').close('ok'); } });
  // action: { label, run }, a button in it (e.g. Undo), and it stays a little longer.
  function toast(msg, isError, action) {
    const box = $('toast'); box.replaceChildren(msg); box.className = 'toast show' + (isError ? ' error' : '') + (action ? ' act' : '');
    if (action) {
      const b = el('button', { type: 'button', class: 'small' }, action.label);
      b.addEventListener('click', () => { clearTimeout(toastTimer); box.className = 'toast'; action.run(); });
      box.append(' ', b);
    }
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { box.className = 'toast'; }, action ? 9000 : isError ? 6000 : 3000);
  }

  // A file picked through a button-like label: its name next to it.
  function fileName(id) { const s = document.querySelector('.fname[data-for="' + id + '"]'); if (s) s.textContent = $(id).files && $(id).files[0] ? $(id).files[0].name : ''; }
  document.querySelectorAll('.fname[data-for]').forEach((s) => $(s.dataset.for).addEventListener('change', () => fileName(s.dataset.for)));

  // Accessibility: each dialog is named by its heading; tabs move with the arrow keys; the list
  // says which signature is open.
  function a11y() {
    document.querySelectorAll('dialog').forEach((d, i) => {
      const h = d.querySelector('h2');
      if (!h) return;
      if (!h.id) h.id = 'dlgTitle' + i;
      d.setAttribute('aria-labelledby', h.id);
    });
    document.querySelectorAll('[role=tablist]').forEach((list) => list.addEventListener('keydown', (e) => {
      if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(e.key)) return;
      const tabs = [...list.querySelectorAll('[role=tab]')].filter((x) => !x.hidden);
      const i = tabs.indexOf(document.activeElement);
      if (i < 0) return;
      e.preventDefault();
      const next = e.key === 'Home' ? 0 : e.key === 'End' ? tabs.length - 1 : (i + (e.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length;
      tabs[next].focus(); tabs[next].click();
    }));
  }

  // Dialogs as windows, as in NOBS Inventory: moved by the title bar (where each was left is kept),
  // maximised (or a double-click on the title), minimised to a chip at the bottom right, closed.
  // On a phone they stay where they are. A window opened from another (Settings → Audit log) is
  // minimised with it, as one chip: the page under an open dialog cannot be used.
  const POS_KEY = 'osign-dlg-pos';
  const WIN_ICON = {
    min: '<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M2 9.5h8" stroke="currentColor" stroke-width="1.4"/></svg>',
    max: '<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><rect x="2" y="2" width="8" height="8" fill="none" stroke="currentColor" stroke-width="1.3"/></svg>',
    unmax: '<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M4 2h6v6M2 4h6v6H2z" fill="none" stroke="currentColor" stroke-width="1.2"/></svg>',
    close: '<svg viewBox="0 0 12 12" width="12" height="12" aria-hidden="true"><path d="M2.5 2.5l7 7M9.5 2.5l-7 7" stroke="currentColor" stroke-width="1.4"/></svg>'
  };
  function windows() {
    let pos = {};
    try { pos = JSON.parse(localStorage.getItem(POS_KEY)) || {}; } catch (e) { pos = {}; }
    const keep = () => { try { localStorage.setItem(POS_KEY, JSON.stringify(pos)); } catch (e) { /* not kept */ } };
    const narrow = () => window.matchMedia('(max-width: 760px)').matches;
    const tray = el('div', { id: 'dlgTray', role: 'toolbar', 'aria-label': t('Minimised windows') });
    document.body.append(tray);
    const title = (d) => (d.querySelector('.dlg > h2:first-child, .wiz-head h2') || {}).textContent || '';
    const opened = []; // the open dialogs, the one on top last
    let groups = []; // minimised: each the dialogs that were open, the top one last
    function renderTray() {
      groups = groups.filter((g) => g.length);
      tray.replaceChildren(...groups.map((g) => el('span', { class: 'tray-chip' },
        el('button', { type: 'button', class: 'tray-open', title: t('Restore'), onclick: () => restore(g) }, title(g[g.length - 1])),
        el('button', { type: 'button', class: 'tray-x', title: t('Close'), 'aria-label': t('Close'), onclick: () => drop(g) }, '×'))));
      tray.hidden = !tray.children.length;
    }
    function minimize() {
      const g = opened.slice();
      g.forEach((d) => { d.dataset.min = '1'; });
      g.slice().reverse().forEach((d) => d.close());
      groups.push(g); renderTray();
    }
    function restore(g) {
      groups = groups.filter((x) => x !== g);
      g.forEach((d) => { delete d.dataset.min; if (!d.open) d.showModal(); });
      renderTray();
    }
    // Closed from the tray: as if closed (the assistant, for one, starts again next time).
    function drop(g) {
      groups = groups.filter((x) => x !== g);
      g.forEach((d) => { delete d.dataset.min; d.dispatchEvent(new Event('close')); });
      renderTray();
    }
    function place(d) {
      const p = pos[d.id];
      d.style.translate = p && !narrow() && !d.classList.contains('maxed') ? p.x + 'px ' + p.y + 'px' : '';
    }
    function maximize(d, on) {
      d.classList.toggle('maxed', on);
      const b = d.querySelector('.wc-max');
      b.innerHTML = on ? WIN_ICON.unmax : WIN_ICON.max;
      b.title = on ? t('Restore') : t('Maximise'); b.setAttribute('aria-label', b.title);
      place(d);
    }
    const BIG_DIALOGS = ['dlgWizard', 'dlgSettings', 'dlgHistory', 'dlgDiff', 'dlgAudit', 'dlgImport'];
    document.querySelectorAll('dialog').forEach((d) => {
      const bar = d.querySelector('.dlg > h2:first-child, .wiz-head');
      if (!bar) return;
      // What scrolls is the content only: the title (with these buttons), the tabs and the buttons
      // at the bottom stay where they are, and the scrollbar runs between them.
      const form = bar.matches('.dlg > h2') ? bar.parentNode : null;
      if (form) {
        const body = el('div', { class: 'dlg-body' });
        const fixed = (n) => n === bar || (n.nodeType === 1 && (n.classList.contains('dlg-foot') || n.classList.contains('dtabs') || (d.id === 'dlgAudit' && n.classList.contains('dlg-toolbar'))));
        const kids = [...form.childNodes];
        const firstFoot = kids.findIndex((n) => n.nodeType === 1 && n.classList.contains('dlg-foot'));
        const top = kids.filter((n, i) => fixed(n) && (firstFoot < 0 || i < firstFoot) && !(n.nodeType === 1 && n.classList.contains('dlg-foot')));
        body.append(...kids.filter((n) => !fixed(n)));
        const lastTop = top[top.length - 1] || bar;
        lastTop.after(body);
      }
      const btn = (cls, icon, label, fn) => { const b = el('button', { type: 'button', class: cls, title: label, 'aria-label': label, onclick: fn }); b.innerHTML = icon; return b; };
      // Minimise and maximise only where there is much to read or work through; the others just close.
      const big = BIG_DIALOGS.includes(d.id);
      d.append(el('div', { class: 'wctl' },
        big ? btn('wc-min', WIN_ICON.min, t('Minimise'), minimize) : null,
        big ? btn('wc-max', WIN_ICON.max, t('Maximise'), () => maximize(d, !d.classList.contains('maxed'))) : null,
        btn('wc-close', WIN_ICON.close, t('Close'), () => d.close())));
      bar.classList.add('movable');
      bar.addEventListener('dblclick', (e) => { if (big && !narrow() && !e.target.closest('button, input, select, a')) maximize(d, !d.classList.contains('maxed')); });
      bar.addEventListener('pointerdown', (e) => {
        if (e.button !== 0 || narrow() || d.classList.contains('maxed') || e.target.closest('button, input, select, a, label')) return;
        const start = pos[d.id] || { x: 0, y: 0 };
        const r = d.getBoundingClientRect();
        // Where it stands without the move: the title bar stays in reach, 120px of it at least.
        const left = r.left - start.x, top = r.top - start.y;
        const sx = e.clientX, sy = e.clientY;
        let moved = false;
        const move = (ev) => {
          const x = Math.min(window.innerWidth - 120 - left, Math.max(120 - left - r.width, start.x + ev.clientX - sx));
          const y = Math.min(window.innerHeight - 40 - top, Math.max(-top, start.y + ev.clientY - sy));
          pos[d.id] = { x: Math.round(x), y: Math.round(y) }; moved = true;
          d.classList.add('moving'); place(d);
        };
        const up = () => { bar.removeEventListener('pointermove', move); d.classList.remove('moving'); if (moved) keep(); };
        bar.setPointerCapture(e.pointerId);
        bar.addEventListener('pointermove', move);
        bar.addEventListener('pointerup', up, { once: true });
        bar.addEventListener('pointercancel', up, { once: true });
      });
      // Opened again by its button while minimised: no chip for it any more; in its place.
      new MutationObserver(() => {
        const i = opened.indexOf(d);
        if (i >= 0) opened.splice(i, 1);
        if (!d.open) return;
        opened.push(d);
        if (d.dataset.min) { delete d.dataset.min; groups.forEach((g) => { const j = g.indexOf(d); if (j >= 0) g.splice(j, 1); }); renderTray(); }
        place(d);
      }).observe(d, { attributes: true, attributeFilter: ['open'] });
    });
    window.addEventListener('resize', () => document.querySelectorAll('dialog[open]').forEach(place));
    renderTray();
  }

  async function start() {
    I18N.apply(document.body);
    a11y();
    windows();
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
      loadApproval();
      loadTls();
      document.querySelectorAll('button.pick-btn').forEach((b) => { b.hidden = !info.features.pickFolder; });
      if (onServer()) serverTexts();
      document.querySelector('#pvSource option[value=entra]').hidden = !(onServer() && info.features.publishTarget === 'sharepoint');
      if (info.features.signOut) {
        const out = $('btnLogout');
        out.hidden = false; $('whoSep').hidden = false;
        if (info.user) {
          out.title = t('Signed in as {0}', info.user.name || info.user.login);
          $('who').textContent = info.user.name || info.user.login; $('who').title = out.title; $('who').hidden = false;
        }
        out.addEventListener('click', async () => {
          if (saveTimer) await save();
          try { await api('logout'); } catch (e) { /* signed out either way */ }
          location.replace('/login');
        });
      }
      const wl = $('wizLang');
      Object.entries(I18N.langs).forEach(([k, v]) => wl.append(el('option', { value: k }, v)));
      wl.value = I18N.lang;
      // A step stored: the assistant was open when the page started again (another language).
      const resume = stored(WIZ_KEY);
      // The server version always has a target: the assistant opens until it was finished once, or
      // until what it sets up is there anyway (the app registration and the add-in's address).
      // On the server the setup is the local account admin's: the others neither get nor see it.
      const mayWizard = !onServer() || !!(info.user && info.user.local);
      const setUp = S.settings.setupDone === '1' || (onServer() ? !!(S.settings.clientId && S.settings.addinUrl) : !!(S.settings.folder && S.settings.clientId));
      // In the top bar only until it is set up; then in Settings → General.
      $('btnWizard').hidden = !mayWizard || setUp; $('wizSep').hidden = $('btnWizard').hidden;
      $('setWizRow').hidden = !mayWizard;
      if (mayWizard && (resume !== null || !setUp)) wizOpen(+resume || 0);
      if (!info.features.adLookup) $('pvSource').querySelector('option[value=ad]').textContent = t('User from Active Directory (not reachable)');
      // Loaded: what waits for the editor (the tests) can go on.
      document.documentElement.dataset.ready = '1';
    } catch (e) {
      document.body.classList.add('failed');
      $('main').replaceChildren(el('section', { class: 'view empty' }, el('h1', null, t('Not connected')), el('p', null, t(e.message))));
    }
  }
  start();
})();
