  // ---------------------------------------------------------------------------------------------
  // Settings
  // ---------------------------------------------------------------------------------------------
  function renderSettingsNotes(r) {
    if (onServer()) {
      const w = $('setServerWhere');
      const where = toSharePoint()
        ? t('Publishing goes to {0}: the SharePoint site and folder of the add-in below, through Microsoft Graph.', S.settings.folder)
        : t("Publishing goes into the server's own folder ({0}), not to SharePoint: Microsoft Graph is not set up on the server.", S.settings.folder);
      const state = !toSharePoint() ? '' : r.folderError ? t('SharePoint cannot be reached: {0}', t(r.folderError)) : r.folderFound ? t('Found.') : t('The folder is not there yet: it is made on the first publish.');
      w.textContent = where + (state ? ' ' + state : '');
      w.className = 'lead' + (r.folderError || !toSharePoint() ? ' warn' : '');
      return;
    }
    $('setFolderNote').textContent = !S.settings.folder ? t('Not set yet.') : r.folderFound ? t('Found.') : t('Not found on this PC: is the library synced?');
    $('setPhotosNote').textContent = r.photosFound ? t('Found.') : t('No photos folder there (yet): the preview shows no photos.');
  }
  let lastSettings = {};
  const ADDIN_FIELDS = [['setAddinUrl', 'addinUrl'], ['setAddinFolder', 'addinFolder'], ['setAddinHost', 'addinHost'], ['setClientId', 'clientId'], ['setTenantId', 'tenantId'], ['setSpSite', 'spSite'], ['setSpFolder', 'spFolder']];
  const f0 = () => (S.info && S.info.features) || {};
  function settingsFromDialog() {
    const s = { folder: $('setFolder').value, photosFolder: $('setPhotos').value };
    for (const [id, k] of ADDIN_FIELDS) s[k] = $(id).value;
    s.addinMobile = $('setAddinMobile').checked ? 'on' : '';
    return s;
  }
  $('btnSettings').addEventListener('click', () => {
    $('setFolder').value = S.settings.folder || '';
    $('setPhotos').value = S.settings.photosFolder || '';
    for (const [id, k] of ADDIN_FIELDS) $(id).value = S.settings[k] || '';
    $('setAddinMobile').checked = S.settings.addinMobile === 'on';
    $('setAddinBuild').textContent = onServer() && f0().pagesUpload ? t('Create and upload the add-in files') : t('Create the add-in files');
    $('setAddinHost').options[0].textContent = onServer() ? t('Automatic (Cloudflare Pages when set up)') : t('Choose…');
    // The server: what follows from the connections and the setup script, said in one line; the
    // fields by hand folded away, and the SharePoint folder only in Server connections.
    const st = S.settings, f = S.info.features || {};
    $('setAddinInfo').hidden = !onServer();
    $('setAddinInfo').textContent = t('Served at {0}', st.addinUrl || t('(no address yet: set up Cloudflare Pages in Server connections)')) + ' · ' +
      (st.clientId ? t('app registration {0} (from the setup script)', st.clientId) : S.info.user && S.info.user.local ? t('no app registration yet: run the setup script (Server connections → SharePoint)') : t('no app registration yet')) +
      (f.pagesProject ? ' · ' + t('Cloudflare Pages project {0}', f.pagesProject) : '');
    $('setAddinMore').open = !onServer();
    document.querySelector('#dlgSettings [data-th=addinMoreHint]').hidden = !onServer();
    for (const id of ['setSpSite', 'setSpFolder']) $(id).closest('label').hidden = onServer();
    $('setAddinBuild').disabled = !(S.info.features && S.info.features.addinFiles);
    // The server version: where publishing goes is the site and folder below; no folders on a PC.
    $('setServerWhere').hidden = !onServer();
    for (const id of ['setFolder', 'setPhotos', 'setAddinFolder']) $(id).closest('label').hidden = onServer();
    document.querySelector('#dlgSettings [data-th=settingsLead]').hidden = onServer();
    document.querySelector('#dlgSettings [data-th=addinLead]').hidden = onServer();
    document.querySelector('#dlgSettings [data-th=addinLeadServer]').hidden = !onServer();
    $('setAddinBuild').closest('p').hidden = onServer() && $('setAddinBuild').disabled;
    // The server's connections, and where it publishes: the local account admin's alone.
    $('setConnRow').hidden = true;
    $('setTabServer').hidden = !(onServer() && S.info.user && S.info.user.local);
    $('setApproval').checked = S.settings.requireApproval === 'on';
    $('setApprovalNote').textContent = '';
    const locked = onServer() && !(S.info.user && S.info.user.local);
    $('setLocked').hidden = !locked;
    for (const [id] of ADDIN_FIELDS) $(id).disabled = locked;
    $('setAddinMobile').disabled = locked;
    // Making and uploading the add-in's files is the local admin's too: not offered to others.
    $('setAddinBuild').hidden = locked;
    $('setAddinUpload').hidden = true; $('setAddinZip').hidden = true;
    // For "Add from file" in Outlook, or "Upload manifest file" in the admin center.
    $('setAddinManifest').hidden = !(onServer() && S.info.features.addinBuilt);
    $('setAddinOk').textContent = $('setAddinBuild').disabled && !onServer() ? t('The add-in files are not next to OSign.ps1: use the release zip.') : '';
    $('setErr').textContent = '';
    renderSettingsNotes(lastSettings);
    $('setWhere').textContent = '';
    $('aboutVersion').textContent = S.info.version || '';
    setTab('general');
    $('dlgSettings').showModal();
  });
  // Settings in tabs: general, the server (the local admin's), the Outlook add-in.
  function setTab(name) {
    document.querySelectorAll('#dlgSettings [data-stab]').forEach((b) => { b.classList.toggle('active', b.dataset.stab === name); b.setAttribute('aria-selected', b.dataset.stab === name); });
    document.querySelectorAll('#dlgSettings [data-spane]').forEach((p) => { p.hidden = p.dataset.spane !== name; });
    document.querySelector('#dlgSettings .dlg').scrollTop = 0;
    // Info saves nothing: only Close there.
    const view = name === 'about';
    $('setSave').hidden = view;
    document.querySelector('#dlgSettings .dlg-foot button[value=cancel]').textContent = view ? t('Close') : t('Cancel');
  }
  document.querySelectorAll('#dlgSettings [data-stab]').forEach((b) => b.addEventListener('click', () => { setTab(b.dataset.stab); if (b.dataset.stab === 'server') runHealth(); }));

  // Status: every part the server depends on, tried now, in one list; a part to set up opens
  // its tab in Server connections.
  // Asked again while it runs (say, after the add-in's files were made again): once more after it,
  // so the rows show what is there now.
  let healthRunning = null, healthAgain = false;
  function runHealth() {
    if (healthRunning) { healthAgain = true; return healthRunning; }
    healthRunning = health().finally(() => {
      healthRunning = null;
      if (healthAgain) { healthAgain = false; runHealth(); }
    });
    return healthRunning;
  }
  async function health() {
    $('healthRun').disabled = true;
    try { S.info = await api('info'); } catch (e) { /* the last one known */ }
    const f = S.info.features || {};
    const test = (what) => async () => {
      try { return { state: 'good', text: t((await api('connections/test', { what })).message) }; } catch (e) {
        return { state: /^(Not set up yet|Noch nicht eingerichtet)/.test(e.message) ? 'off' : 'bad', text: e.message };
      }
    };
    const checks = [
      { name: 'Active Directory', tab: 'ad', run: test('ad') },
      { name: 'SharePoint (Microsoft 365)', tab: 'graph', run: test('graph') },
      { name: 'Cloudflare Pages', tab: 'pages', run: test('pages') },
      { name: 'Web certificate', run: async () => {
        const c = await loadTls();
        const x = tlsText(c).replace(/^[^:]*:\s*/, '');
        return { state: !c ? 'off' : !c.error && c.trusted && c.nameOk !== false && c.daysLeft > 30 ? 'good' : c.error || c.daysLeft < 0 ? 'bad' : 'warn', text: x.charAt(0).toUpperCase() + x.slice(1) };
      } },
      { name: 'Sign-in to Microsoft 365', run: async () => {
        if (f.publishTarget !== 'sharepoint') return { state: 'off', text: t('Not set up yet: the setup script does it (Server connections → SharePoint).') };
        const left = f.secretDaysLeft, how = f.graphCredential === 'certificate' ? t('With the server\'s certificate') : t('With a client secret');
        if (left === null) return { state: 'good', text: how + '.' };
        return { state: left < 0 ? 'bad' : left <= 30 ? 'warn' : 'good', text: how + ', ' + (left < 0 ? t('ran out on {0}: run the setup script again.', day(f.secretExpires)) : t('valid until {0} ({1}).', day(f.secretExpires), t('{0} days', left))) };
      } },
      { name: 'Key for passwords', run: async () => (f.keyOk ? { state: 'good', text: t('There (secrets/app_key): passwords and the certificate\'s key are kept encrypted.') }
        : { state: 'bad', text: t('Missing (secrets/app_key): run sudo ./deploy-nobs-osign.sh again.') }) },
      { name: 'Users from AD, daily', run: async () => {
        const x = f.adExport;
        if (!x) return { state: 'off', text: t('Not scheduled: the users are written when publishing.') };
        if (!x.last) return { state: 'good', text: t('Every day at {0}; not run yet since the server started.', x.at) };
        return x.last.ok ? { state: 'good', text: t('Every day at {0}; last run {1}.', x.at, new Date(x.last.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB')) }
          : { state: 'bad', text: t('Every day at {0}; the last run failed: {1}', x.at, t(x.last.error)) };
      } },
      { name: 'Last publish', run: async () => {
        const v = (await api('history/list')).versions[0];
        if (!v) return { state: 'off', text: t('Nothing published yet.') };
        return { state: 'good', text: new Date(v.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB') + ' · ' + v.by + (f.requireApproval ? ' · ' + t('four eyes on') : '') };
      } },
      { name: 'Outlook add-in files', act: () => (!f.addinVersion || !f.addinFiles || f.addinVersion === S.info.version && (!f.pagesUpload || (f.addinUploaded && f.addinUploaded.version === f.addinVersion)) ? null
        : { label: f.pagesUpload ? t('Make and upload') : t('Make again'), fn: addinRemake }), run: async () => {
        const v = f.addinVersion, up = f.addinUploaded;
        if (!v) return { state: 'off', text: t('Not made yet (Settings → Outlook add-in).') };
        if (v !== S.info.version) return { state: 'warn', text: t('Made with {0}; this server is {1}: make them again, so Outlook gets the new version.', v, S.info.version) };
        if (f.pagesUpload && (!up || up.version !== v)) return { state: 'warn', text: t('Made, but not uploaded to Cloudflare Pages yet.') };
        return { state: 'good', text: f.pagesUpload ? t('Version {0}, on Cloudflare Pages since {1}.', v, new Date(up.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB')) : t('Version {0}.', v) };
      } },
      // Where the add-in is deployed in Microsoft 365: only a script with an admin's sign-in can
      // ask Microsoft; its last answer is shown, with when it came.
      { name: 'Add-in in Microsoft 365', act: () => (S.info.user && S.info.user.local ? { label: t('Check…'), fn: deployCheck } : null), run: async () => {
        const st = await api('deploy/status'), x = st.last;
        if (!x) return { state: 'off', text: S.info.user && S.info.user.local ? t('Not checked yet: "Check…" downloads a script that only reads where the add-in is deployed (run it as a Global or Exchange administrator).') : t('Not checked yet: the local account admin can check it here.') };
        const when = new Date(x.at).toLocaleString(I18N.lang === 'de' ? 'de-CH' : 'en-GB');
        if (x.error) return { state: 'bad', text: t('The check on {0} stopped: {1}', when, x.error) };
        const m = x.summary;
        if (!m.found) return { state: 'bad', text: t('Not deployed (checked {0}): Microsoft 365 admin center → Settings → Integrated apps → Upload custom apps, with the manifest address.', when) };
        const who = m.everyone ? t('everyone') : m.assigned || t('nobody named');
        const parts = [t('Deployed to: {0}', who)];
        if (m.version) parts.push(t('version {0}', m.version));
        parts.push(t('checked {0}', when));
        // The manifest's version is the add-in files' (0.9.26 → 0.9.26.0).
        const want = f.addinVersion ? f.addinVersion + '.0' : '';
        if (m.enabled === false) return { state: 'warn', text: t('Switched off in Integrated apps.') + ' ' + parts.join(' · ') };
        if (want && m.version && m.version !== want) return { state: 'warn', text: parts.join(' · ') + ' · ' + t('the current manifest is {0}: Microsoft takes it over from the manifest address by itself, within a day.', want) };
        return { state: 'good', text: parts.join(' · ') };
      } },
      { name: 'Version', run: async () => (f.update ? { state: 'warn', text: t('{0} runs here; {1} is out: sudo ./scripts/update.sh {1}', S.info.version, f.update.version) }
        : { state: 'good', text: S.info.version }) }
    ];
    const rows = checks.map((c) => {
      const dot = el('span', { class: 'hdot run' }), text = el('small', null, t('Checking…'));
      const act = c.act && c.act();
      const row = el('div', { class: 'setrow' }, dot, el('div', { class: 'setrow-text' }, el('b', null, t(c.name)), text),
        c.tab ? el('button', { type: 'button', id: c.tab === 'ad' ? 'setConn' : null, onclick: () => openConn(c.tab) }, t(c.name.replace(/ \(.*\)$/, '') + '…')) : null,
        act ? el('button', { type: 'button', class: 'primary', onclick: (e) => act.fn(e.currentTarget) }, act.label) : null);
      return { c, row, dot, text };
    });
    $('healthRows').replaceChildren(...rows.map((r) => r.row));
    $('healthSum').textContent = t('Checking…'); $('healthSum').className = 'health-sum';
    await Promise.all(rows.map(async (r) => {
      let x;
      try { x = await r.c.run(); } catch (e) { x = { state: 'bad', text: t(e.message) }; }
      r.dot.className = 'hdot ' + x.state; r.dot.title = t({ good: 'All right', warn: 'Needs attention', bad: 'Not working', off: 'Not set up' }[x.state]);
      r.text.textContent = x.text; r.row.dataset.state = x.state;
    }));
    const n = (s) => rows.filter((r) => r.row.dataset.state === s).length;
    const bad = n('bad'), warn = n('warn');
    $('healthSum').textContent = bad && warn ? t('{0} not working, {1} to look at.', bad, warn) : bad ? t('{0} not working.', bad) : warn ? t('{0} to look at.', warn) : t('Everything set up works.');
    $('healthSum').className = 'health-sum ' + (bad ? 'bad' : warn ? 'warn' : 'good');
    $('healthRun').disabled = false;
  }
  $('healthRun').addEventListener('click', runHealth);
  // The script that reads where the add-in is deployed; the row shows its answer when it comes.
  async function deployCheck(btn) {
    btn.disabled = true;
    try {
      const r = await api('deploy/start', { origin: location.origin });
      const a = el('a', { href: '/setup/nobs-osign-deployment.ps1?id=' + encodeURIComponent(r.id), download: 'nobs-osign-deployment.ps1' });
      document.body.append(a); a.click(); a.remove();
      toast(t('Downloaded. Run it in Windows PowerShell: powershell -ExecutionPolicy Bypass -File .\\nobs-osign-deployment.ps1 (valid until {0}). It only reads.', new Date(r.until).toLocaleTimeString()));
      // Its answer, when it comes: the status again.
      const until = Date.parse(r.until);
      const poll = async () => {
        if (Date.now() > until || $('dlgSettings').open === false) return;
        const st = await api('deploy/status').catch(() => null);
        if (st && !st.waiting) { runHealth(); return; }
        setTimeout(poll, 4000);
      };
      setTimeout(poll, 4000);
    } catch (e) { toast(t(e.message), true); } finally { btn.disabled = false; }
  }
  // After the add-in's files were made or uploaded: what the server now says about them, in the
  // status bar (no more "Add-in files from …") and the status list, without reloading the page.
  async function afterAddinChange() {
    try { S.info = await api('info'); renderStatusBar(); } catch (e) { /* as it was */ }
    if ($('dlgSettings').open) runHealth();
  }
  // The add-in's files made again for this version (and uploaded, with Cloudflare Pages).
  async function addinRemake(btn) {
    btn.disabled = true;
    try {
      const r = await api('addin/build');
      if (S.info.features.pagesUpload) await api('addin/upload');
      toast(t('Add-in files made for {0}.', r.version) + (S.info.features.pagesUpload ? ' ' + t('Uploaded to Cloudflare Pages.') : '') + (r.note ? ' ' + t(r.note) : ''), !!r.note);
    } catch (e) { toast(t(e.message), true); }
    await afterAddinChange();
  }
  $('setSave').addEventListener('click', async (e) => {
    e.preventDefault();
    try {
      const r = await api('settings/save', { settings: settingsFromDialog() });
      // Another language: the page starts again in it (the draft is saved, nothing is lost).
      if ($('setLang').value !== I18N.lang) { await saving; I18N.setLang($('setLang').value); return; }
      S.settings = r.settings; lastSettings = r;
      $('dlgSettings').close();
      renderStatusBar();
      if (onServer()) toast(r.folderError ? t('SharePoint cannot be reached: {0}', t(r.folderError)) : t('Settings saved.'), !!r.folderError);
      else toast(r.folderFound || !r.settings.folder ? t('Settings saved.') : t('Saved, but the folder is not there on this PC.'), !r.folderFound && !!r.settings.folder);
      schedulePreview();
    } catch (err) { $('setErr').textContent = t(err.message); }
  });

  $('setAddinBuild').addEventListener('click', async () => {
    $('setErr').textContent = ''; $('setAddinOk').textContent = '';
    try {
      const saved = await api('settings/save', { settings: settingsFromDialog() });
      S.settings = saved.settings; lastSettings = saved;
      const r = await api('addin/build');
      if (onServer()) { $('setAddinManifest').hidden = false; S.info.features.addinBuilt = true; }
      if (onServer() && S.info.features.pagesUpload) {
        // Made and uploaded in one go; the ZIP stays there for another web server.
        $('setAddinZip').hidden = false;
        await pagesUpload($('setAddinOk'), $('setAddinBuild'));
        $('setAddinOk').textContent += ' ' + t('Once: upload {0} in the Microsoft 365 admin center (Settings → Integrated apps).', r.manifestUrl);
      } else if (onServer()) {
        $('setAddinZip').hidden = false;
        $('setAddinOk').textContent = t('{0} files (version {1}) made. Upload them, add the redirect URI {2} to the app registration, and upload {3} in the admin center.', r.files, r.version, r.redirectUri, r.manifestUrl);
      } else $('setAddinOk').textContent = t('{0} files (version {1}) in {2}. Copy them to the web server, add the redirect URI {3} to the app registration, and upload {4} in the admin center.', r.files, r.version, r.folder, r.redirectUri, r.manifestUrl);
      // Made, but without the folder's ids: said, so they are made again once SharePoint answers.
      if (r.note) $('setErr').textContent = t(r.note);
    } catch (err) { $('setErr').textContent = t(err.message); }
    await afterAddinChange();
  });

  // The server's connections (AD, Graph, Cloudflare), set by the local account admin: what is
  // saved wins over .env, an empty field takes .env's value (its placeholder). Secrets are only
  // ever sent, never shown: the server says where each comes from.
  const CONN_FIELDS = ['ldapUrl', 'adDomain', 'ldapStartTls', 'ldapBind', 'ldapBase', 'ldapGroup', 'ldapLoginAttr', 'ldapServiceDn', 'ldapDirectoryBase', 'adExportAt', 'adAttributes', 'graphTenantId', 'graphClientId', 'cfAccountId', 'cfProject'];
  const CONN_SECRETS = ['ldapServicePassword', 'cfApiToken'];
  let connDrop = {}, connCa;
  function day(iso) { return iso ? new Date(iso).toLocaleDateString(I18N.lang === 'de' ? 'de-CH' : 'en-GB') : ''; }
  function renderConnections(r) {
    $('connKey').hidden = r.keyOk;
    for (const k of CONN_FIELDS) {
      const el = $('cn_' + k), f = r.fields[k];
      el.value = f.value;
      if (el.tagName === 'SELECT') el.options[0].textContent = f.env ? t('As in .env ({0})', t(f.env === 'on' ? 'On' : 'Off')) : t('As in .env');
      else el.placeholder = f.env || (r.derived && r.derived[k]) || t(el.dataset.ph);
    }
    for (const k of CONN_SECRETS) {
      const x = r.secrets[k];
      $('cs_' + k).value = '';
      $('cs_' + k).placeholder = x.saved ? '••••••••' : '';
      document.querySelector('#dlgConn [data-drop=' + k + ']').hidden = !x.saved;
      $('cx_' + k).textContent = x.saved ? (x.readable ? t('Set here on {0}.', day(x.saved)) : t('Set here, but the server cannot read it any more (another secrets/app_key): enter it again.'))
        : x.file ? t('From {0}.', x.fileName) : t('Not set.');
      $('cx_' + k).className = x.saved && !x.readable ? 'err' : '';
    }
    // The server's own certificate, when the setup script put it on the app: no secret then.
    $('connCert').textContent = r.cert ? t('Thumbprint {0}, valid until {1}.', r.cert.thumbprint, day(r.cert.expires)) : t('None yet: the setup script makes it, or Download makes one now.');
    $('connCaRemove').hidden = !r.ca.saved;
    cfAddress();
    $('connCa').textContent = r.ca.saved ? t('Set here on {0}: {1}.', day(r.ca.saved.at), r.ca.saved.certs.map((c) => c.name + ' (' + t('until {0}', day(c.expires)) + ')').join(', '))
      : r.ca.env ? t('From {0}.', r.ca.env) : t('None: only the public CAs are trusted.');
    $('connCaFile').value = ''; fileName('connCaFile');
    const line = (id, problem, okText) => { $(id).textContent = problem ? connText(problem) : t(okText); $(id).className = 'lead' + (problem ? ' warn' : ' ok'); };
    // SharePoint: the setup script first; "again" once it is set up (a new certificate).
    connSpRender();
    $('connScriptGet').textContent = r.problems.graph ? t('Download the setup script') : t('Download the setup script again');
    $('connScriptGet').className = r.problems.graph ? 'primary' : '';
    // Active Directory: the sign-in first, then the service account.
    r.problems.ad = r.problems.signIn || r.problems.directory;
    line('connAd', r.problems.ad, 'Ready: members of the sign-in group can sign in, and the editor reads the users\' values from AD.');
    line('connGraph', r.problems.graph, 'Ready: publishing goes into SharePoint.');
    line('connPages', r.problems.pages, 'Ready: the add-in\'s files can be uploaded from here.');
    // Each tab's dot: ready, or something to set.
    document.querySelectorAll('#dlgConn [data-cdot]').forEach((d) => { d.className = 'dot ' + (r.problems[d.dataset.cdot] ? 'warn' : 'good'); });
    connDrop = {}; connCa = undefined; connCaReading = null;
    $('connErr').textContent = '';
  }
  // In this dialog, "(Settings → Server connections)" would point to itself; the line at the top of
  // a tab and the answer of its test say a problem in the same words.
  // Inside Server connections, a pointer to it (or to one of its tabs) says nothing: left out.
  function connText(p) { return t(String(p).replace(/ \(secrets\/[a-z_]+, or Server connections\)/, '').replace(/ \(Settings → Server connections\)\./, '.').replace(/, or Server connections\)/, ')')).replace(/ \((?:Server connections|Server-Verbindungen) → [^)]+\)/, ''); }

  // The setup script, from the SharePoint tab: as in the setup assistant, its result comes back
  // here by itself.
  let connScriptTimer = null;
  function connScriptPoll() {
    clearTimeout(connScriptTimer);
    connScriptTimer = setTimeout(async () => {
      if (!$('dlgConn').open) return;
      let r;
      try { r = await api('setup/status'); } catch (e) { return connScriptPoll(); }
      const out = $('connScriptNote');
      if (r.state === 'done') {
        out.className = 'where good'; out.textContent = '✓ ' + setupDoneText(r.result);
        try { const [c, info, st] = await Promise.all([api('connections/get'), api('info'), api('settings/get')]); S.info = info; S.settings = st.settings; lastSettings = st; renderConnections(c); renderStatusBar(); } catch (e) { /* shown on the next open */ }
        out.className = 'where good'; out.textContent = '✓ ' + setupDoneText(r.result);
        return;
      }
      if (r.state === 'failed') { out.className = 'where warn'; out.textContent = t('The setup script stopped: {0}', r.problem) + ' ' + t('Fix it, then run the script again (it is valid until {0}).', new Date(r.until).toLocaleTimeString()); }
      if (r.state !== 'none') connScriptPoll();
    }, 3000);
  }
  function setupDoneText(x) {
    const until = x.secretExpires ? new Date(x.secretExpires).toLocaleDateString(I18N.lang === 'de' ? 'de-CH' : 'en-GB') : '';
    return x.credential === 'certificate'
      ? t('Microsoft 365 is set up: the add-in\'s app {0}, the server\'s app {1} with its certificate, write access to {2}.', x.addinClientId, x.serverClientId, x.siteUrl)
      : t('Microsoft 365 is set up: the add-in\'s app {0}, the server\'s app {1}, write access to {2}. The client secret is valid until {3}.', x.addinClientId, x.serverClientId, x.siteUrl, until || '?');
  }
  // The folder of the signatures: the site the setup script gives the server, shown and changed here.
  const personalNow = () => isPersonal((S.settings && S.settings.spSite) || '');
  function connSpRender() {
    const st = S.settings || {};
    const site = !st.spSite || st.spSite === 'root' ? t('the root site') : st.spSite.replace(/:\//, '/').replace(/:$/, '');
    $('connSpWhere').textContent = st.spFolder ? 'SharePoint: ' + site + ' › ' + st.spFolder : t('Not chosen yet.');
    $('connSpWarn').textContent = personalNow() ? t('A personal OneDrive: only its owner can give the server access, and it goes with that person. Choose a folder in a SharePoint site.') : '';
  }
  $('connSpSet').addEventListener('click', async () => {
    const out = document.querySelector('#dlgConn [data-ctestout="graph"]');
    out.className = 'where'; out.textContent = '';
    const p = parseSpAddress($('connSpUrl').value);
    if (!$('connSpUrl').value.trim() || !p.folder || !/^[A-Za-z0-9.-]+:\/(sites|teams|personal)\/[^:/]+:$|^root$/.test(p.site)) {
      out.className = 'where warn'; out.textContent = t('That address names no folder: open the folder in SharePoint and copy the address from the browser.'); return;
    }
    try {
      const r = await api('settings/save', { settings: { spSite: p.site, spFolder: p.folder } });
      S.settings = r.settings; lastSettings = r;
      $('connSpUrl').value = '';
      connSpRender(); renderStatusBar();
      out.className = 'where good'; out.textContent = '✓ ' + t('Saved: {0}. The setup script gives the server write access to this site.', $('connSpWhere').textContent);
    } catch (err) { out.className = 'where warn'; out.textContent = connText(err.message); }
  });
  // The script's result as a file (nobs-osign-setup-result.json), where the PC could not reach the
  // server (its certificate not trusted there): the browser hands it over.
  async function setupResultUpload(input, out) {
    const f = input.files[0]; input.value = '';
    if (!f) return false;
    out.className = 'where'; out.textContent = t('Handing over the result…');
    try {
      let body;
      try { body = JSON.parse(await f.text()); } catch (e) { throw new Error(t('This is not the setup script\'s result (nobs-osign-setup-result.json).')); }
      await api('setup/finish', body);
      const r = await api('setup/status');
      if (r.state !== 'done') throw new Error(r.problem || t('The server did not take the result.'));
      const [info, st] = await Promise.all([api('info'), api('settings/get')]);
      S.info = info; S.settings = st.settings; lastSettings = st; renderStatusBar();
      out.className = 'where good'; out.textContent = '✓ ' + setupDoneText(r.result) + ' ' + t('Delete the file now.');
      return true;
    } catch (err) { out.className = 'where warn'; out.textContent = t(err.message); return false; }
  }
  $('connScriptResult').addEventListener('change', async () => {
    if (await setupResultUpload($('connScriptResult'), $('connScriptNote'))) { const keep = $('connScriptNote').textContent; renderConnections(await api('connections/get')); $('connScriptNote').className = 'where good'; $('connScriptNote').textContent = keep; }
  });
  $('connScriptGet').addEventListener('click', async () => {
    const out = $('connScriptNote');
    out.className = 'where'; out.textContent = ''; $('connScriptWizard').hidden = true;
    if (personalNow() && !(await ask(t('The folder is in a personal OneDrive: the script only works when run by its owner. Download it anyway?'), t('Download')))) return;
    try {
      const r = await api('setup/start', { origin: location.origin, groups: core.valueSource(S.config) === 'graph', users: $('connScriptUsers').checked });
      const a = el('a', { href: '/setup/nobs-osign-setup.ps1?id=' + encodeURIComponent(r.id), download: 'nobs-osign-setup.ps1' });
      document.body.append(a); a.click(); a.remove();
      out.textContent = t('Downloaded. Run it now: this page goes on by itself when it is done (the script is valid until {0}).', new Date(r.until).toLocaleTimeString());
      connScriptPoll();
    } catch (err) {
      out.className = 'where warn'; out.textContent = t(err.message);
      // The folder or the add-in's address is missing: the assistant asks for them.
      $('connScriptWizard').hidden = !/first/i.test(err.message);
    }
  });
  $('connScriptWizard').addEventListener('click', () => { $('dlgConn').close(); wizOpen(Math.max(0, wizSteps().indexOf('folder'))); });

  // The add-in's address that follows from the Pages project.
  function cfAddress() {
    const v = $('cn_cfProject').value.trim() || $('cn_cfProject').placeholder;
    $('cfAddress').textContent = /^[a-z0-9][a-z0-9-]{0,57}$/.test(v) ? t('(the add-in: https://{0}.pages.dev/)', v) : t('(the add-in: https://<name>.pages.dev/)');
  }
  $('cn_cfProject').addEventListener('input', cfAddress);

  // Server connections in tabs: one part at a time.
  function connTestsClear() { document.querySelectorAll('#dlgConn [data-ctestout]').forEach((o) => { o.textContent = ''; o.className = 'where'; }); }
  function connTab(name) {
    document.querySelectorAll('#dlgConn [data-ctab]').forEach((b) => { b.classList.toggle('active', b.dataset.ctab === name); b.setAttribute('aria-selected', b.dataset.ctab === name); });
    document.querySelectorAll('#dlgConn [data-cpane]').forEach((p) => { p.hidden = p.dataset.cpane !== name; });
  }
  document.querySelectorAll('#dlgConn [data-ctab]').forEach((b) => b.addEventListener('click', () => connTab(b.dataset.ctab)));
  async function openConn(tab) {
    $('setErr').textContent = '';
    try { renderConnections(await api('connections/get')); connTab(tab || 'ad'); connTestsClear(); $('dlgConn').showModal(); } catch (err) { $('setErr').textContent = t(err.message); }
  }
  // Back from Server connections to Status: checked again, with what was just changed.
  $('dlgConn').addEventListener('close', () => { if ($('dlgSettings').open && !document.querySelector('#dlgSettings [data-spane=server]').hidden) runHealth(); });
  $('dlgConn').addEventListener('click', (e) => {
    const b = e.target.closest('[data-drop]');
    if (!b) return;
    connDrop[b.dataset.drop] = true; b.hidden = true;
    $('cx_' + b.dataset.drop).textContent = t('Removed when you save.');
  });
  // By hand: the server's certificate (made now if there is none) to upload to the app in Entra.
  $('connCertGet').addEventListener('click', async () => {
    const out = document.querySelector('#dlgConn [data-ctestout="graph"]');
    out.className = 'where'; out.textContent = '';
    try {
      const r = await api('connections/cert');
      const bytes = Uint8Array.from(atob(r.der), (c) => c.charCodeAt(0));
      const a = el('a', { href: URL.createObjectURL(new Blob([bytes], { type: 'application/pkix-cert' })), download: 'nobs-osign-server.cer' });
      document.body.append(a); a.click(); a.remove();
      $('connCert').textContent = t('Thumbprint {0}, valid until {1}.', r.thumbprint, day(r.expires));
      out.textContent = t('Downloaded. Upload it to the server\'s app in Entra, then Save and test the connection.');
    } catch (err) { out.className = 'where warn'; out.textContent = connText(err.message); }
  });
  $('connCaRemove').addEventListener('click', () => { connCa = null; connCaReading = null; $('connCaRemove').hidden = true; $('connCa').textContent = t('Removed when you save.'); });
  let connCaReading = null;
  $('connCaFile').addEventListener('change', async () => {
    const f = $('connCaFile').files[0];
    if (!f) return;
    connCaReading = new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = rej; r.readAsDataURL(f); });
  });
  // Saves what the dialog holds; the dialog stays open (Test) or closes (Save).
  async function connSaveNow() {
    const values = {}, secrets = {};
    for (const k of CONN_FIELDS) values[k] = $('cn_' + k).value;
    for (const k of CONN_SECRETS) { if ($('cs_' + k).value) secrets[k] = $('cs_' + k).value; else if (connDrop[k]) secrets[k] = null; }
    const body = { values, secrets };
    // A CA file chosen just now may still be being read.
    if (connCaReading) { connCa = await connCaReading; connCaReading = null; }
    if (connCa !== undefined) body.ca = connCa;
    const r = await api('connections/save', body);
    // What the server can do now (publishing, AD, the upload) follows from what was saved.
    const [info, st] = await Promise.all([api('info'), api('settings/get')]);
    S.info = info; S.settings = st.settings; lastSettings = st;
    wizConn = r; wizConnRender();
    renderStatusBar();
    if ($('dlgSettings').open) renderSettingsNotes(st);
    return r;
  }
  $('connSave').addEventListener('click', async (e) => {
    e.preventDefault();
    $('connErr').textContent = '';
    try { await connSaveNow(); $('dlgConn').close(); toast(t('Server connections saved.')); }
    catch (err) { $('connErr').textContent = t(err.message); }
  });
  // Detect: the domain from the domain controller (saved first, with its CA), into the field.
  $('connDetect').addEventListener('click', async () => {
    const out = document.querySelector('#dlgConn [data-ctestout="ad"]');
    out.className = 'where'; out.textContent = t('Asking the domain controller…'); $('connDetect').disabled = true; $('connErr').textContent = '';
    try {
      renderConnections(await connSaveNow());
      const r = await api('connections/detect');
      $('cn_adDomain').value = r.domain;
      renderConnections(await connSaveNow());
      out.className = 'where good'; out.textContent = '✓ ' + t('Domain {0} ({1}). Now the sign-in group and the service account.', r.domain, r.host || r.base);
    } catch (err) { out.className = 'where warn'; out.textContent = t(err.message); }
    finally { $('connDetect').disabled = false; }
  });
  // Test: saved first, so what is tested is what the dialog shows; the answer next to the button.
  document.querySelectorAll('#dlgConn [data-ctest]').forEach((b) => b.addEventListener('click', async () => {
    const out = document.querySelector('#dlgConn [data-ctestout="' + b.dataset.ctest + '"]');
    out.className = 'where'; out.textContent = t('Testing…'); b.disabled = true; $('connErr').textContent = '';
    try {
      renderConnections(await connSaveNow());
      const r = await api('connections/test', { what: b.dataset.ctest });
      out.className = 'where good'; out.textContent = '✓ ' + connText(r.message);
    } catch (err) { out.className = 'where warn'; out.textContent = connText(err.message); }
    finally { b.disabled = false; }
  }));

  // The server version: the files made go to Cloudflare Pages from the server.
  async function pagesUpload(out, btn) {
    btn.disabled = true; out.textContent = t('Uploading…');
    try {
      const r = await api('addin/upload');
      out.textContent = t('Uploaded to Cloudflare Pages ({0} of {1} files new): {2}', r.uploaded, r.files, r.url);
      afterAddinChange();
      return true;
    } catch (err) { out.textContent = ''; throw err; } finally { btn.disabled = false; }
  }
  $('setAddinUpload').addEventListener('click', async () => {
    $('setErr').textContent = '';
    try { await pagesUpload($('setAddinOk'), $('setAddinUpload')); } catch (err) { $('setErr').textContent = t(err.message); }
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
    const id = u.searchParams.get('id') || u.searchParams.get('RootFolder');
    // OneDrive's own pages (…-my.sharepoint.com/my?id=/personal/<name>/Documents/…) name the site
    // only in ?id=.
    const idParts = id ? id.split('/').filter(Boolean) : [];
    if (site === 'root' && idParts.length >= 2 && /^(sites|teams|personal)$/i.test(idParts[0])) {
      parts = idParts.slice(0, 2);
      site = host + ':/' + parts[0].toLowerCase() + '/' + parts[1] + ':';
    }
    const sitePath = site === 'root' ? [] : parts.slice(0, 2).map((x) => x.toLowerCase());
    let path = id ? idParts : (shared ? [] : rest);
    if (id && sitePath.length && path.slice(0, 2).map((x) => x.toLowerCase()).join('/') === sitePath.join('/')) path = path.slice(2);
    let folder = null, url = null;
    const lib = path.findIndex((x) => LIBRARIES.includes(x.toLowerCase()));
    if (lib === 0) {
      const inLib = path.slice(1);
      const forms = inLib.findIndex((x) => /^forms$/i.test(x) || /\.aspx$/i.test(x));
      const f = (forms < 0 ? inLib : inLib.slice(0, forms));
      // The folder's own address: what OneDrive notes for the library, and the folder in it.
      if (f.length) { folder = f.join('/'); url = 'https://' + host + (site === 'root' ? '' : '/' + parts.slice(0, 2).join('/')) + '/' + [path[0]].concat(f).join('/'); }
    }
    return { site, folder, url };
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
    if (!r.found || r.location.uncertain) {
      out.textContent = !r.found
        ? t('NOBS OSign cannot tell where this folder is in SharePoint: OneDrive does not sync it. Choose «SharePoint link» instead, or a folder OneDrive syncs.')
        : t('This is a shortcut in your OneDrive («Add shortcut to My files»): NOBS OSign cannot tell which SharePoint folder it is. Choose «SharePoint link» and paste the folder\'s address.');
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
  $('wizFolder').addEventListener('change', () => {
    $('wizSpSite').value = ''; $('wizSpFolder').value = '';
    wiz.checking = detectWhere($('wizFolder').value, 'wizSpSite', 'wizSpFolder', $('wizWhereFound'), null).then(wizPermNote);
  });

  // The texts that speak of "this PC" (OSign.ps1 on Windows): on the server, what it does instead.
  const SERVER_TEXTS = {
    wizValuesNote: ['You can change this at any time under <b>Add-in</b>. The server reads Active Directory itself, with its service account (<b>Settings → Server connections</b>).',
      'Sie können das jederzeit unter <b>Add-in</b> ändern. Active Directory liest der Server selbst, mit seinem Dienstkonto (<b>Einstellungen → Server-Verbindungen</b>).'],
    wizFolderLead: ['NOBS OSign keeps the signatures, their images and the photos in a folder in your company\'s SharePoint, where Outlook reads them with each user\'s own sign-in. Make a folder <b>OSign</b> there (with a folder <b>photos</b> in it), and paste its address here: the server writes into it through Microsoft Graph.',
      'NOBS OSign bewahrt die Signaturen, ihre Bilder und die Fotos in einem Ordner im SharePoint Ihrer Firma auf, wo Outlook sie mit der Anmeldung der jeweiligen Person liest. Legen Sie dort einen Ordner <b>OSign</b> an (mit einem Ordner <b>photos</b> darin) und fügen Sie seine Adresse hier ein: der Server schreibt über Microsoft Graph hinein.'],
    hSpSite: ['Where the add-in finds the OSign folder in SharePoint, and where the server publishes. <code>root</code> is the home site (<code>https://&lt;tenant&gt;.sharepoint.com</code>). Paste the folder\'s address from the browser or a sharing link to it; NOBS OSign converts it, e.g. to <code>contoso.sharepoint.com:/sites/IT:</code>.',
      'Wo das Add-in den OSign-Ordner in SharePoint findet und wohin der Server veröffentlicht. <code>root</code> ist die Startseite (<code>https://&lt;tenant&gt;.sharepoint.com</code>). Fügen Sie die Adresse des Ordners aus dem Browser oder einen Freigabelink ein; NOBS OSign wandelt sie um, z. B. in <code>contoso.sharepoint.com:/sites/IT:</code>.'],
    hPreview: ['<b>Sample values</b>: each field\'s sample value (see <i>Fields</i>). <b>All fields empty</b>: shows which lines disappear. <b>User from Active Directory</b>: a real user\'s values and groups, read from AD by the server; shows which signature they would get.',
      '<b>Beispielwerte</b>: der Beispielwert jedes Feldes (siehe <i>Felder</i>). <b>Alle Felder leer</b>: zeigt, welche Zeilen wegfallen. <b>Benutzer aus Active Directory</b>: die Werte und Gruppen einer echten Person, vom Server aus AD gelesen; zeigt, welche Signatur sie bekäme.']
  };
  function serverTexts() {
    for (const [k, [en, de]] of Object.entries(SERVER_TEXTS)) document.querySelectorAll('[data-th="' + k + '"]').forEach((e) => { e.innerHTML = I18N.lang === 'de' ? de : en; });
  }

