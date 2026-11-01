  // ---------------------------------------------------------------------------------------------
  // Fields
  // ---------------------------------------------------------------------------------------------
  const GRAPH_PROPS = ['givenName', 'surname', 'displayName', 'jobTitle', 'department', 'companyName', 'officeLocation', 'mobilePhone', 'businessPhones[0]', 'faxNumber', 'mail', 'userPrincipalName', 'streetAddress', 'postalCode', 'city', 'state', 'country', 'employeeId']
    .concat(Array.from({ length: 15 }, (_, i) => 'onPremisesExtensionAttributes.extensionAttribute' + (i + 1)));
  $('graphProps').append(...GRAPH_PROPS.map((p) => el('option', { value: p })));

  function editable(f) {
    const c = el('input', { type: 'checkbox', 'aria-label': t('Users may change this field in their own variants'), title: t('Users may change this field in their own variants') });
    c.checked = !!f.userEditable;
    c.addEventListener('change', () => { if (c.checked) f.userEditable = true; else delete f.userEditable; changed({ list: false }); });
    return c;
  }
  // The AD attributes the server lets fields read (the local admin can allow more): offered, and
  // one it does not allow marked, since it would stay empty.
  const adAllowed = () => ((S.info && S.info.features && S.info.features.adAttributes) || []);
  function adInput(f, x) {
    const mark = () => {
      const bad = !!f.ad && adAllowed().length > 0 && !adAllowed().some((a) => a.toLowerCase() === f.ad.toLowerCase());
      x.classList.toggle('bad', bad);
      x.title = bad ? t('Not read from AD: the local admin can allow it under Server connections → Active Directory.') : '';
    };
    x.addEventListener('input', mark); mark();
    return x;
  }
  function renderFields() {
    $('adAttrs').replaceChildren(...adAllowed().map((a) => el('option', { value: a })));
    const tb = $('fieldRows'); tb.textContent = '';
    S.config.fields.forEach((f, i) => {
      const inp = (prop, attrs) => {
        const x = el('input', Object.assign({ value: f[prop] || '' }, attrs || {}));
        x.addEventListener('input', () => { f[prop] = x.value.trim(); changed({ list: false }); if (prop === 'graph') x.closest('td').classList.toggle('nograph', !f.graph); });
        return x;
      };
      const key = el('input', { value: f.key, 'aria-label': t('Placeholder name') });
      let oldKey = f.key;
      key.addEventListener('change', async () => {
        const nk = key.value.trim().replace(/^@@/, '');
        if (!/^[A-Za-z][A-Za-z0-9_]*$/.test(nk)) { toast(t('A placeholder name is letters, digits and _, starting with a letter.'), true); key.value = oldKey; return; }
        if (S.config.fields.some((x) => x !== f && x.key === nk)) { toast(t('@@{0} is already a field.', nk), true); key.value = oldKey; return; }
        const uses = S.config.signatures.filter((s) => core.placeholdersIn(s.html).indexOf(oldKey) >= 0);
        if (uses.length && await ask(t('Change @@{0} to @@{1} in {2} signature(s) too?', oldKey, nk, uses.length), t('Change them too'), false, t('Only the field'))) {
          const re = new RegExp('@@' + oldKey + '(?![A-Za-z0-9_])', 'g');
          uses.forEach((s) => { s.html = s.html.replace(re, '@@' + nk); });
        }
        f.key = nk; oldKey = nk; changed({ list: false });
      });
      const del = el('button', { class: 'ghost small danger', title: t('Remove field'), 'aria-label': t('Remove field @@{0}', f.key), onclick: async () => {
        const uses = S.config.signatures.filter((s) => core.placeholdersIn(s.html).indexOf(f.key) >= 0).map((s) => s.name);
        if (uses.length && !(await ask(t('@@{0} is used in: {1}. Remove the field anyway? The placeholder then stays as typed.', f.key, uses.join(', ')), t('Remove'), true))) return;
        S.config.fields.splice(i, 1); changed({ list: false }); renderFields();
        // Back as it was, at its place, while the note shows.
        toast(t('Field @@{0} removed.', f.key), false, { label: t('Undo'), run: () => {
          if (S.config.fields.some((x) => x.key === f.key)) return;
          S.config.fields.splice(Math.min(i, S.config.fields.length), 0, f); changed({ list: false }); renderFields();
        } });
      } }, '✕');
      tb.append(el('tr', null,
        el('td', null, inp('label', { 'aria-label': t('Field') })),
        el('td', null, el('div', { class: 'ph' }, el('span', null, '@@'), key)),
        el('td', null, adInput(f, inp('ad', { list: 'adAttrs', 'aria-label': t('AD attribute') }))),
        el('td', { class: f.graph ? '' : 'nograph' }, inp('graph', { list: 'graphProps', placeholder: t('not available'), 'aria-label': t('Graph source') })),
        el('td', null, inp('sample', { 'aria-label': t('Sample value') })),
        el('td', { class: 'center' }, editable(f)),
        el('td', null, del)));
    });
  }
  // The starter fields as they were at first: missing ones come back, changed ones are reset
  // (label, AD attribute, Graph source, sample, editable in the add-in); fields of your own stay.
  $('btnResetFields').addEventListener('click', async () => {
    const defs = core.DEFAULT_FIELDS;
    const changedKeys = defs.filter((d) => { const f = S.config.fields.find((x) => x.key === d.key); return !f || ['label', 'ad', 'graph', 'sample'].some((k) => (f[k] || '') !== (d[k] || '')) || f.userEditable; }).map((d) => '@@' + d.key);
    if (!changedKeys.length) { toast(t('The default fields are as they were at first.')); return; }
    if (!(await ask(t('Restore the default fields? This resets: {0}. Fields of your own stay.', changedKeys.join(', ')), t('Restore'), true))) return;
    const own = S.config.fields.filter((f) => !defs.some((d) => d.key === f.key));
    S.config.fields = defs.map((d) => Object.assign({}, d)).concat(own);
    changed({ list: false }); renderFields();
    toast(t('Default fields restored.'));
  });
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
    // The manifest's address, as the add-in's files were made for it.
    const base = S.settings && S.settings.addinUrl ? S.settings.addinUrl.replace(/\/+$/, '') + '/' : '';
    $('installManifest').textContent = base ? base + 'manifest.xml' : '';
    $('installLink').hidden = !base; $('installNone').hidden = !!base;
    const mode = core.valueSource(S.config);
    document.querySelectorAll('input[name=vsrc]').forEach((r) => { r.checked = r.value === mode; });
    // The server version: it exports the users itself, every day at the time set on it.
    const sn = $('vsrcServerNote');
    sn.hidden = mode === 'graph' || !onServer();
    if (!sn.hidden) {
      const f = S.info.features, x = f.adExport;
      let text = t('With AD, the users are written into users/ in the OSign folder on Publish: one small file per person with their fields and groups. The add-in reads only the file of whoever writes.');
      if (!f.adLookup) text += ' ' + t('Active Directory is not set up on the server: {0}', t(f.adProblem));
      else if (!x) text += ' ' + t('The server does not export them by itself (no daily time in Settings → Server connections): the values change only when you publish.');
      else {
        text += ' ' + t('The server exports them every day at {0}.', x.at);
        const when = x.last && new Date(x.last.at).toLocaleString(I18N.lang);
        // Skipped (nothing published that takes values from AD, or AD answered nobody): said as that.
        if (x.last) text += ' ' + (!x.last.ok ? t('Last run {0} failed: {1}', when, t(x.last.error)) : x.last.result && x.last.result.skipped ? t('Last run {0}: skipped, {1}', when, t(x.last.result.skipped)) : t('Last run {0}: fine.', when));
      }
      sn.textContent = text;
      sn.className = 'lead' + (!f.adLookup || (x && x.last && !x.last.ok) ? ' warn' : '');
    }
  }
  document.querySelectorAll('input[name=vsrc]').forEach((r) => r.addEventListener('change', () => {
    if (!r.checked) return;
    S.config.options.valueSource = r.value; changed({ list: false }); renderOptions();
  }));
  Object.entries(OPTS).forEach(([id, k]) => $(id).addEventListener('change', () => { S.config.options[k] = $(id).checked; changed({ list: false }); }));
  $('lnkSetup').addEventListener('click', (e) => { e.preventDefault(); $('dlgHelp').showModal(); });

  document.querySelectorAll('.nav-list a[data-view]').forEach((a) => a.addEventListener('click', (e) => {
    e.preventDefault();
    // Not loaded yet (a slow server): nothing to show.
    if (!S.config) return;
    flushSource();
    goView(a.dataset.view);
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
    head.append(el('tr', null, el('th', null, ''), el('th', null, t('Name')), el('th', null, t('Mail')), fields.map((f) => el('th', null, (f.label || f.key))), showGroups ? el('th', null, t('Groups')) : null));
    const q = $('usersFilter').value.trim().toLowerCase();
    const only = $('usersChanges').checked;
    const rows = snap.users.filter((u) => (!only || u.status !== 'unchanged') &&
      (!q || u.mail.toLowerCase().includes(q) || String(u.name || '').toLowerCase().includes(q) || Object.values(u.values || {}).some((v) => String(v).toLowerCase().includes(q))));
    const LIMIT = 500;
    rows.slice(0, LIMIT).forEach((u) => {
      body.append(el('tr', null,
        el('td', null, el('span', { class: 'st ' + u.status }, t(u.status))),
        el('td', { class: 'nowrap' }, u.name || ''),
        el('td', null, u.mail),
        fields.map((f) => {
          const v = (u.values || {})[f.key] || '', was = u.old ? (u.old[f.key] || '') : null;
          const chg = u.status === 'changed' && was !== null && was !== v;
          return el('td', { class: chg ? 'chg' : '' }, v, chg ? el('span', { class: 'was' }, was || t('(empty)')) : null);
        }),
        showGroups ? el('td', null, (u.groups || []).join(', ')) : null));
    });
    snap.removed.filter((m) => !q || m.includes(q)).forEach((m) => {
      body.append(el('tr', { class: 'removed' }, el('td', null, el('span', { class: 'st removed' }, t('removed'))), el('td', { colspan: fields.length + (showGroups ? 3 : 2) }, m)));
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

