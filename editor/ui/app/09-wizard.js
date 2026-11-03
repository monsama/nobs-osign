  // ---------------------------------------------------------------------------------------------
  // Setup assistant: opens by itself until it was finished once, while NOBS OSign is not set up
  // (no publish folder or no app registration). Each step saves what it asks for.
  // ---------------------------------------------------------------------------------------------
  const ALL_WIZ = Array.from(document.querySelectorAll('#dlgWizard .wiz-step')).map((sec) => sec.dataset.step);
  // The steps for this backend: the server version cannot make the add-in's files yet.
  let WIZ = ALL_WIZ;
  const wizSteps = () => ALL_WIZ.filter((n) => !(onServer() && (n === 'files' || n === 'deploy') && !S.info.features.addinFiles)
    // The server's connections: only on the server, and only for who may set them.
    && !(n === 'connect' && !(onServer() && S.info.user && S.info.user.local)));
  const WIZ_KEY = 'osign.wizStep';
  const wiz = { step: 0, reached: 0, built: false, warned: '', checking: null, spOk: false, checkSeq: 0, scriptWaiting: false };
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
    $('wizFolder').value = st.folder || ''; $('wizLink').value = '';
    $('wizSpSite').value = st.spSite || 'root'; $('wizSpFolder').value = st.folder ? st.spFolder || 'OSign' : '';
    // Nothing made yet: the recommended one; otherwise what is set.
    setRadio('wvsrc', S.config.signatures.length || S.settings.setupDone === '1' ? core.valueSource(S.config) : 'mixed');
    setRadio('whost', st.addinHost || 'cloudflare');
    $('wizUrl').value = st.addinUrl || '';
    $('wizClientId').value = st.clientId || ''; $('wizTenantId').value = st.tenantId || '';
    if (onServer()) { $('wizSpSite').value = st.spSite || 'root'; $('wizSpFolder').value = st.spFolder || 'OSign'; }
    $('wizAddinFolder').value = st.addinFolder || (S.info.backend === 'powershell' && navigator.platform.startsWith('Win') ? 'C:\\OSign-addin' : '');
  }
  function wizRender() {
    const name = WIZ[wiz.step];
    document.querySelectorAll('#dlgWizard .wiz-step').forEach((sec) => { sec.hidden = sec.dataset.step !== name; });
    const sec = document.querySelector('#dlgWizard .wiz-step[data-step=' + name + ']');
    $('wizTitle').textContent = t(sec.dataset.title);
    $('wizCount').textContent = t('Step {0} of {1}', wiz.step + 1, WIZ.length);
    $('wizBar').style.width = Math.round(100 * wiz.step / (WIZ.length - 1)) + '%';
    $('wizBack').hidden = wiz.step === 0;
    wizNavRender();
    if (name === 'done') wizSummary();
    $('wizNext').textContent = name === 'welcome' ? t('Start') : name === 'done' ? t('Finish') : t('Next');
    $('wizLater').hidden = name === 'done';
    $('wizErr').textContent = '';
    $('wizScript').hidden = !(name === 'app' && wizHasScript());
    // Done already (here or in Server connections): said so; the script only to run again, by hand folded away.
    if (name === 'app') {
      const f = (S.info && S.info.features) || {}, done = wizHasScript() && !!S.settings.clientId && f.graphCredential === 'certificate';
      $('wizAppDone').hidden = !done;
      $('wizAppDone').textContent = done ? t('✓ Set up already: the add-in\'s app {0}, the server\'s app {1} with its certificate. Nothing to do here; run the script again only to change something.', S.settings.clientId, f.graphClientId) : '';
      $('wizScriptGet').textContent = done ? t('Download the setup script again') : t('Download the setup script');
      $('wizScriptGet').className = done ? '' : 'primary';
      $('wizAppManual').open = !wizHasScript();
    }
    // Back on the step while the script is out: ask again (the answer may be there already).
    if (name === 'app' && wiz.scriptWaiting && !wizScriptTimer) wizScriptPoll();
    const host = radio('whost');
    // The server uploads to its own Pages project: its address is known before the first upload.
    const project = onServer() && host === 'cloudflare' && S.info.features && S.info.features.pagesProject;
    if (project && !$('wizUrl').value.trim()) $('wizUrl').value = 'https://' + project + '.pages.dev/';
    const base = addinBase();
    $('wizUrlHint').textContent = project ? t('(the server\'s Pages project {0}, made on the first upload; or your own domain set up on it)', project) : t(URL_HINTS[host] || URL_HINTS.other);
    $('wizRedirect').textContent = base ? 'brk-multihub://' + new URL(base).host : t('(enter the address in the step before)');
    $('wizManifest').textContent = base ? base + 'manifest.xml' : '';
    const perms = ['User.Read', 'Sites.Read.All'];
    if (radio('wvsrc') === 'graph') perms.push('GroupMember.Read.All');
    $('wizPerms').textContent = perms.join(', ');
    // The server version makes the files there: uploaded from there, or downloaded as a ZIP. The
    // steps to upload by hand only when the server cannot.
    const srv = onServer();
    $('wizAddinFolder').closest('label').hidden = srv;
    document.querySelector('#dlgWizard [data-th=wizFilesLead]').hidden = srv;
    $('wizFilesServer').hidden = !srv;
    $('wizServerFiles').hidden = !srv || !wiz.built;
    $('wizCfUpload').hidden = !srv || !S.info.features.pagesUpload;
    $('wizUpload').hidden = !wiz.built || (srv && S.info.features.pagesUpload && host === 'cloudflare');
    document.querySelectorAll('#wizUpload [data-host]').forEach((p) => { p.hidden = p.dataset.host !== (host || 'other'); });
  }
  // Step 1: where the folder is, given one way or the other - its SharePoint address (NOBS OSign finds
  // the synced folder on this PC), or the folder on this PC (NOBS OSign finds where it is in SharePoint).
  const isPersonal = (site) => /:\/personal\//.test(site || '');
  // The steps as a row to jump to: once set up, any of them; before, those reached already.
  const WIZ_SHORT = { welcome: 'Welcome', connect: 'Connections', folder: 'Signatures', values: 'Values', host: 'Add-in address', app: 'App registration', files: 'Add-in files', deploy: 'Deployment', done: 'Ready' };
  const wizAnyStep = () => S.settings.setupDone === '1';
  // How far the assistant got, kept in this browser: opened again, the steps up to there are a
  // click away, as they are once it is set up.
  function wizFar() { try { return WIZ.indexOf(localStorage.getItem('osign.wizReached')); } catch (e) { return -1; } }
  function wizKeepFar() { try { if (wiz.reached > wizFar()) localStorage.setItem('osign.wizReached', WIZ[wiz.reached]); } catch (e) { /* this time only */ } }
  // As breadcrumbs: each step numbered, the ones gone through ticked, the current one marked; once
  // set up, any of them a click away.
  // The last step: what is set up, and a step still open a click away (it can be left for later).
  function wizSummary() {
    const f = (S.info && S.info.features) || {}, st = S.settings || {};
    const rows = [
      ['connect', t('Active Directory: signing in with Windows accounts'), onServer() ? !f.adProblem : true],
      ['folder', t('SharePoint: where the signatures are kept'), onServer() ? f.publishTarget === 'sharepoint' && !f.graphProblem : !!st.folder],
      ['host', t('The add-in\'s address'), !!st.addinUrl],
      ['app', t('App registration in Entra ID'), !!st.clientId],
      ['files', t('The add-in\'s files'), !!f.addinVersion && f.addinVersion === S.info.version]
    ].filter(([step]) => WIZ.includes(step));
    $('wizSummary').replaceChildren(...rows.map(([step, what, ok]) => el('li', { class: ok ? 'good' : 'warn' }, what + (ok ? '' : ' – ' + t('not done yet')),
      ok ? null : el('button', { type: 'button', class: 'small ghost', onclick: () => wizGo(WIZ.indexOf(step)) }, t('To the step')))));
  }
  function wizNavRender() {
    const nav = $('wizNav'), any = wizAnyStep();
    wizKeepFar();
    nav.replaceChildren(...WIZ.map((n, i) => {
      const done = i !== wiz.step && i < wiz.reached;
      // Its full title is the heading above once there: named so for a screen reader, no tooltip.
      const b = el('button', { type: 'button', 'data-step': n, class: done ? 'done' : null,
        'aria-label': (i + 1) + '. ' + t(document.querySelector('#dlgWizard .wiz-step[data-step=' + n + ']').dataset.title) + (done ? ' ✓' : '') },
        el('span', { class: 'wn-n', 'aria-hidden': 'true' }, done ? '✓' : String(i + 1)), t(WIZ_SHORT[n] || n));
      if (i === wiz.step) b.setAttribute('aria-current', 'step');
      b.disabled = i !== wiz.step && !any && i > wiz.reached;
      b.addEventListener('click', () => wizGo(i));
      return b;
    }));
    const cur = nav.querySelector('[aria-current]'); if (cur && cur.scrollIntoView) cur.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }
  // What stops a step, where it is seen (on a small screen it is below what shows).
  function wizError(err) {
    $('wizErr').textContent = err && err.message ? t(err.message) : '';
    if ($('wizErr').textContent) $('wizErr').scrollIntoView({ block: 'nearest' });
  }
  // To another step: forward keeps what this step asks for, as Next does (it stays here if that
  // fails); back, as Back, keeps nothing. The add-in files need not be made again to pass by.
  async function wizGo(i) {
    if (i === wiz.step || $('wizNext').disabled) return;
    $('wizErr').textContent = '';
    const name = WIZ[wiz.step];
    $('wizNext').disabled = true; $('wizBack').disabled = true;
    try {
      if (i > wiz.step && wizLeave[name] && name !== 'files' && name !== 'done') {
        try { await wizLeave[name](); }
        catch (err) {
          // Set up already: the step is left as it is saved, and says so; before, it is finished first.
          if (!wizAnyStep()) throw err;
          if (err.message) toast(t('Not saved: {0}', t(err.message)));
        }
      }
      wiz.step = i; wiz.reached = Math.max(wiz.reached, i); wiz.warned = ''; wizRender();
    } catch (err) { wizError(err); }
    finally { $('wizNext').disabled = false; $('wizBack').disabled = false; }
  }
  function wizPermNote() {
    const site = $('wizSpFolder').value ? $('wizSpSite').value : '';
    $('wizPerm').hidden = !site;
    $('wizPerm').innerHTML = !site ? '' : isPersonal(site)
      ? t('<b>Only you</b> can read your OneDrive, so only you get signatures: good for testing. For everyone, choose a SharePoint site later.')
      : t('<b>Everyone who gets a signature</b> needs read access to the folder OSign: in SharePoint, on the folder, <b>Share</b> → <i>Everyone except external users</i> → <b>Can view</b>. While testing, leave it to yourself.');
  }
  function wizLibChanged() {
    $('wizLinkBox').hidden = radio('wlib') !== 'link';
    $('wizOwnBox').hidden = radio('wlib') !== 'own';
    $('wizErr').textContent = '';
    wizPermNote();
  }
  document.querySelectorAll('input[name=wlib]').forEach((r) => r.addEventListener('change', wizLibChanged));
  async function wizLinkCheck() {
    const out = $('wizLinkNote'), v = $('wizLink').value.trim();
    out.textContent = ''; out.className = 'where';
    $('wizSpSite').value = ''; $('wizSpFolder').value = ''; $('wizFolder').value = '';
    if (!v) { wizPermNote(); return; }
    let p = parseSpAddress(v);
    const warn = (msg) => { out.textContent = msg; out.classList.add('warn'); wizPermNote(); };
    if (!/^https:\/\/[^/]+\.sharepoint\.com/i.test(v)) return warn(t('That is not an address in SharePoint (https://….sharepoint.com/…).'));
    // A sharing link names no folder: the server asks Graph where it leads. Elsewhere, or when it
    // may not, the link is opened here, to copy the address from there.
    const sharing = /^https:\/\/[^/]+\/:[a-z]:\//i.test(v);
    if (!p.folder && sharing && onServer()) {
      out.textContent = t('Opening the sharing link…');
      try {
        const r = await api('sp/share', { url: v });
        p = parseSpAddress(r.webUrl);
        if (p.folder) $('wizLink').value = r.webUrl;
      } catch (e) { return shareHelp(out, v, t(e.message)); }
    }
    if (!p.folder && sharing) return shareHelp(out, v, t('A sharing link does not say which folder it is.'));
    if (!p.folder) return warn(t('This address does not say which folder it is. Open the folder in the browser and copy the address from the address bar.'));
    $('wizSpSite').value = p.site; $('wizSpFolder').value = p.folder;
    if (onServer()) { await wizServerCheck(p.site, p.folder, true); return; }
    let r = { found: false };
    try { r = await api('sp/local', { url: p.url }); } catch (e) { /* said below */ }
    if (r.found && r.exists) {
      $('wizFolder').value = r.path;
      out.textContent = '✓ ' + t('On this PC: {0}', r.path); out.classList.add('good');
    } else {
      warn(t('This folder is not on this PC yet: NOBS OSign publishes through OneDrive. In SharePoint, open the library and choose Sync, wait until the folder is there, then paste the link again.'));
    }
    wizPermNote();
  }
  // What to do with a sharing link that cannot be read here: open it, and copy the address the
  // browser shows then.
  function shareHelp(out, link, why) {
    out.replaceChildren(document.createTextNode(why + ' '));
    const a = document.createElement('a');
    a.href = link; a.target = '_blank'; a.rel = 'noopener noreferrer';
    a.textContent = t('Open the link');
    out.append(a, document.createTextNode(', ' + t('then copy the address from the address bar and paste it here.')));
    out.className = 'where warn';
    wizPermNote();
  }
  // Next may come right after the change (a click on it ends the field): it waits for the check.
  $('wizLink').addEventListener('change', () => { wiz.checking = wizLinkCheck(); });
  // Test: the same check again, e.g. once the server was given access to the site.
  $('wizLinkTest').addEventListener('click', () => { wiz.checking = wizLinkCheck(); });

  // The server version writes there itself: whether it reaches the site and folder, said under the
  // link, with how to give it access when it does not.
  // given: the admin entered or tested this address (not only what is saved, checked on opening).
  async function wizServerCheck(site, folder, given) {
    const out = $('wizLinkNote'), my = ++wiz.checkSeq;
    out.textContent = ''; out.className = 'where';
    wiz.spOk = false;
    wizGrantFill(site);
    $('wizSpSite').value = site; $('wizSpFolder').value = folder;
    try {
      const c = await api('sp/check', { spSite: site, spFolder: folder });
      // A newer check (another address, Test) came in meanwhile: that one says.
      if (my !== wiz.checkSeq) return;
      wiz.spOk = true;
      $('wizGrant').hidden = true;
      out.textContent = c.found ? t('✓ Found in SharePoint: {0}', c.where) : t('The server reaches the site; the folder is not there yet and is made on the first publish: {0}', c.where);
      out.classList.add('good');
      if (/:\/personal\//.test(site)) { out.textContent += ' ' + t('A personal OneDrive goes with its owner: a SharePoint site is better for signatures.'); }
    } catch (e) {
      if (my !== wiz.checkSeq) return;
      // Microsoft 365 not set up yet: the setup script does it (step "Register the app"), and the
      // folder is checked then.
      if (!toSharePoint() && wizHasScript() && given) {
        wiz.spOk = true;
        out.textContent = t('Noted. The server checks this folder once Microsoft 365 is set up, with the setup script a few steps on.');
        if (/:\/personal\//.test(site)) { out.textContent += ' ' + t('But this is a personal OneDrive: only its owner can give the server access (no administrator can), and it goes with that person. A SharePoint site is better.'); out.classList.add('warn'); }
      } else {
        $('wizGrant').hidden = !toSharePoint();
        out.textContent = t(e.message); out.classList.add('warn');
      }
    }
    wizPermNote();
  }
  const wizHasScript = () => onServer() && !!(S.info.user && S.info.user.local);
  // The server version: how to give its app registration write access to the site (Sites.Selected).
  function wizGrantFill(site) {
    const sitePath = !site || site === 'root' ? 'root' : site;
    $('wizGrantGet').textContent = 'GET https://graph.microsoft.com/v1.0/sites/' + sitePath + '?$select=id,webUrl';
    $('wizGrantBody').textContent = JSON.stringify({ roles: ['write'], grantedToIdentities: [{ application: { id: (S.info.features && S.info.features.graphClientId) || '<the server app\'s client ID>', displayName: 'NOBS OSign server' } }] });
    if (!site) $('wizGrant').hidden = true;
    $('wizGrantRedirect').textContent = location.origin + '/ui/signin.html';
    $('wizSetupClient').value = stored(SETUP_KEY) || '';
  }

  // The quick way to give the server the site: the admin signs in with Microsoft here (a popup,
  // the authorization code flow with PKCE, for the app registration "NOBS OSign setup": a
  // single-page application without a secret), and this page makes the two requests to Graph
  // itself. The token stays in this page, for these requests only; the server never gets it.
  const SETUP_KEY = 'osign.setupClientId';
  const b64url = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  async function msToken(clientId, tenant, scope) {
    const verifier = b64url(crypto.getRandomValues(new Uint8Array(32)));
    const challenge = b64url(await crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)));
    const state = b64url(crypto.getRandomValues(new Uint8Array(16)));
    const redirect = location.origin + '/ui/signin.html';
    const auth = 'https://login.microsoftonline.com/' + encodeURIComponent(tenant || 'organizations') + '/oauth2/v2.0';
    const popup = window.open(auth + '/authorize?' + new URLSearchParams({
      client_id: clientId, response_type: 'code', redirect_uri: redirect, response_mode: 'fragment', scope: scope + ' openid',
      code_challenge: challenge, code_challenge_method: 'S256', state, prompt: 'select_account'
    }), 'osign-signin', 'width=520,height=700');
    if (!popup) throw new Error(t('The browser blocked the sign-in window. Allow pop-ups for this page and try again.'));
    // The answer comes back to /ui/signin.html: read from there (same site again), or sent from there.
    const answer = await new Promise((resolve, reject) => {
      let timer = null;
      const done = (a) => { clearInterval(timer); window.removeEventListener('message', onMsg); try { popup.close(); } catch (e) { /* closed */ } if (a) resolve(a); else reject(new Error(t('The sign-in window was closed.'))); };
      const onMsg = (e) => { if (e.origin === location.origin && e.data && e.data.osignSignIn && e.data.state === state) done(e.data); };
      window.addEventListener('message', onMsg);
      timer = setInterval(() => {
        let h = '';
        try { if (popup.location.origin === location.origin) h = popup.location.hash; } catch (e) { /* still at Microsoft */ }
        const p = new URLSearchParams(h.slice(1));
        if (h && p.get('state') === state) return done({ code: p.get('code') || '', error: p.get('error') || '', errorText: p.get('error_description') || '' });
        if (popup.closed) done(null);
      }, 250);
    });
    const said = (x) => String(x || '').split(/\r?\n/)[0];
    if (answer.error) throw new Error(t('Microsoft did not sign you in: {0}', said(answer.errorText || answer.error)));
    let r, j;
    try {
      r = await fetch(auth + '/token', { method: 'POST', body: new URLSearchParams({ client_id: clientId, grant_type: 'authorization_code', code: answer.code, redirect_uri: redirect, code_verifier: verifier, scope: scope + ' openid' }) });
      j = await r.json().catch(() => ({}));
    } catch (e) { throw new Error(t('Microsoft cannot be reached from this browser: {0}', e.message)); }
    if (!r.ok || !j.access_token) throw new Error(t('Microsoft did not sign you in: {0}', said(j.error_description || j.error || r.status)));
    return j.access_token;
  }
  async function grantSite(site) {
    const clientId = $('wizSetupClient').value.trim();
    const serverApp = S.info.features && S.info.features.graphClientId;
    if (!serverApp) throw new Error(t('Set up publishing to SharePoint first (Server connections): the server\'s client ID is what gets the site.'));
    if (!/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(clientId)) throw new Error(t('Enter the client ID of NOBS OSign setup, a GUID.'));
    stored(SETUP_KEY, clientId);
    const tok = await msToken(clientId, S.info.features.graphTenantId, 'https://graph.microsoft.com/Sites.FullControl.All');
    async function graph(method, path, body) {
      const r = await fetch('https://graph.microsoft.com/v1.0' + path, { method, headers: Object.assign({ Authorization: 'Bearer ' + tok }, body ? { 'Content-Type': 'application/json' } : {}), body: body ? JSON.stringify(body) : undefined });
      const j = await r.json().catch(() => ({}));
      if (r.status === 401 || r.status === 403) throw new Error(t('Your account may not do this: sign in as a SharePoint or Global administrator.'));
      if (r.status === 404) throw new Error(t('There is no such SharePoint site. Check the address.'));
      if (!r.ok) throw new Error(t('Microsoft Graph answered {0}: {1}', r.status, (j.error && j.error.message) || ''));
      return j;
    }
    const m = /^([A-Za-z0-9.-]+):\/(sites|teams|personal)\/([^:/]+):$/.exec(site || '');
    const s = await graph('GET', (m ? '/sites/' + m[1] + ':/' + m[2] + '/' + encodeURIComponent(m[3]) + ':' : '/sites/root') + '?$select=id,webUrl');
    // Granted already (write or more): nothing to add.
    const perms = await graph('GET', '/sites/' + encodeURIComponent(s.id) + '/permissions');
    const has = (perms.value || []).some((p) => (p.roles || []).some((x) => ['write', 'owner', 'fullcontrol'].includes(String(x).toLowerCase()))
      && (p.grantedToIdentitiesV2 || p.grantedToIdentities || []).some((x) => x.application && String(x.application.id).toLowerCase() === serverApp.toLowerCase()));
    if (!has) await graph('POST', '/sites/' + encodeURIComponent(s.id) + '/permissions', { roles: ['write'], grantedToIdentities: [{ application: { id: serverApp, displayName: 'NOBS OSign server' } }] });
    return { already: has, webUrl: s.webUrl };
  }
  $('wizGrantGo').addEventListener('click', async () => {
    const out = $('wizGrantMsNote');
    out.className = 'where'; out.textContent = t('Signing in with Microsoft…');
    $('wizGrantGo').disabled = true;
    try {
      const r = await grantSite($('wizSpSite').value || 'root');
      out.textContent = r.already ? t('✓ The server had write access to {0} already.', r.webUrl) : t('✓ The server may now write to {0}.', r.webUrl);
      out.classList.add('good');
      toast(out.textContent);
      // And whether it can, now: as the Test button does (the box goes once it can).
      wiz.checking = wizServerCheck($('wizSpSite').value || 'root', $('wizSpFolder').value);
    } catch (err) { out.textContent = t(err.message); out.classList.add('warn'); }
    finally { $('wizGrantGo').disabled = false; }
  });

  function wizOpen(step) {
    WIZ = wizSteps();
    wizFill();
    wiz.step = Math.max(0, Math.min(WIZ.length - 1, step || 0)); wiz.reached = Math.max(wiz.step, wizFar()); wiz.built = false; wiz.warned = '';
    $('wizMakeOk').textContent = ''; $('wizChecks').replaceChildren();
    $('wizWhereFound').textContent = ''; $('wizLinkNote').textContent = ''; $('wizGrantMsNote').textContent = '';
    wiz.spOk = false; $('wizFolderNote').textContent = ''; $('wizCfOk').textContent = '';
    if (!wiz.scriptWaiting) $('wizScriptNote').textContent = '';
    // A folder set already: that one, with where it is in SharePoint as set. The server version
    // has no folder on a PC: only the SharePoint link.
    setRadio('wlib', onServer() ? 'link' : S.settings.folder ? 'own' : '');
    document.querySelector('#dlgWizard input[name=wlib][value=own]').closest('label').hidden = onServer();
    document.querySelector('#dlgWizard input[name=wlib][value=link]').closest('label').hidden = onServer();
    wizGrantFill(null);
    // What is set already: shown, and checked right away, so the step says whether it works; the
    // field is then only for another folder.
    const saved = onServer() && S.settings.spFolder && S.settings.spSite && S.settings.spSite !== 'root';
    $('wizSpSaved').hidden = !saved;
    $('wizSpSaved').textContent = saved ? t('Chosen: SharePoint {0} › {1}. For another folder, paste its address below.', S.settings.spSite.replace(/:\//, '/').replace(/:$/, ''), S.settings.spFolder) : '';
    $('wizLink').placeholder = saved ? t('Only for another folder') : 'https://contoso.sharepoint.com/sites/IT/Shared Documents/Forms/AllItems.aspx?id=…';
    if (onServer() && S.settings.spFolder) wiz.checking = wizServerCheck(S.settings.spSite || 'root', S.settings.spFolder);
    if ($('wizFolder').value) detectWhere($('wizFolder').value, 'wizSpSite', 'wizSpFolder', $('wizWhereFound'), null);
    wizLibChanged();
    if (WIZ.includes('connect')) wizConnLoad();
    wizRender();
    if (!$('dlgWizard').open) $('dlgWizard').showModal();
  }
  // The server's connections, as the step shows them: ready, or what is missing.
  let wizConn = null;
  function wizConnRender() {
    const p = wizConn && wizConn.problems;
    if (!p) { $('wizConnList').replaceChildren(); return; }
    const rows = [
      ['graph', t('SharePoint (Microsoft Graph): publishing'), true],
      ['signIn', t('Active Directory: signing in with Windows accounts'), false],
      ['directory', t('Active Directory: the users\' values, with the service account'), false],
      ['pages', t('Cloudflare Pages: uploading the add-in\'s files'), false]
    ];
    // SharePoint is set up by the setup script, a few steps on: nothing to do here for it.
    $('wizConnList').replaceChildren(...rows.map(([k, what, needed]) => el('li', { class: p[k] ? 'warn' : 'good' },
      what + (!p[k] ? ' – ' + t('ready') : needed ? ' – ' + t('set up by the setup script in the step "Register the app in Entra ID"') : ' – ' + t(p[k]) + ' ' + t('(optional)')))));
  }
  async function wizConnLoad() {
    try { wizConn = await api('connections/get'); } catch (e) { wizConn = null; }
    wizConnRender();
  }
  // Back from the connections (Settings closed): the step says how they are now.
  $('dlgSettings').addEventListener('close', () => { if ($('dlgWizard').open && WIZ.includes('connect')) wizConnLoad(); });
  $('wizConnOpen').addEventListener('click', async () => {
    $('wizErr').textContent = '';
    try { await openConn('ad'); } catch (err) { $('wizErr').textContent = t(err.message); }
  });
  // The setup script (server/setup.js): made for the folder and the add-in's address as saved,
  // downloaded, run on a Windows PC; its result comes back to the server, and here by asking.
  let wizScriptTimer = null;
  function wizScriptPoll(stop) {
    clearTimeout(wizScriptTimer); wizScriptTimer = null;
    if (stop) return;
    wizScriptTimer = setTimeout(async () => {
      let r;
      try { r = await api('setup/status'); } catch (e) { return wizScriptPoll(); }
      const out = $('wizScriptNote');
      if (r.state === 'none') { wiz.scriptWaiting = false; return wizScriptPoll(true); }
      if (r.state === 'done') {
        wiz.scriptWaiting = false;
        let info, st;
        try { [info, st] = await Promise.all([api('info'), api('settings/get')]); } catch (e) { wiz.scriptWaiting = true; return wizScriptPoll(); }
        S.info = info; S.settings = st.settings; lastSettings = st;
        $('wizClientId').value = S.settings.clientId || ''; $('wizTenantId').value = S.settings.tenantId || '';
        out.className = 'where good';
        out.textContent = '✓ ' + setupDoneText(r.result);
        return wizScriptPoll(true);
      }
      if (r.state === 'failed') { out.className = 'where warn'; out.textContent = t('The setup script stopped: {0}', r.problem) + ' ' + t('Fix it, then run the script again (it is valid until {0}).', new Date(r.until).toLocaleTimeString()); }
      wizScriptPoll();
    }, 3000);
  }
  $('wizScriptResult').addEventListener('change', async () => {
    if (await setupResultUpload($('wizScriptResult'), $('wizScriptNote'))) {
      wiz.scriptWaiting = false; wizScriptPoll(true);
      $('wizClientId').value = S.settings.clientId || ''; $('wizTenantId').value = S.settings.tenantId || '';
    }
  });
  $('wizScriptGet').addEventListener('click', async () => {
    const out = $('wizScriptNote');
    out.className = 'where'; $('wizErr').textContent = '';
    try {
      const r = await api('setup/start', { origin: location.origin, groups: radio('wvsrc') === 'graph', users: $('wizScriptUsers').checked });
      const a = el('a', { href: '/setup/nobs-osign-setup.ps1?id=' + encodeURIComponent(r.id), download: 'nobs-osign-setup.ps1' });
      document.body.append(a); a.click(); a.remove();
      out.textContent = t('Downloaded. Run it now: this page goes on by itself when it is done (the script is valid until {0}).', new Date(r.until).toLocaleTimeString());
      wiz.scriptWaiting = true;
      wizScriptPoll();
    } catch (err) { out.className = 'where warn'; out.textContent = t(err.message); }
  });
  async function wizSave(settings) {
    const r = await api('settings/save', { settings });
    S.settings = r.settings; lastSettings = r;
    return r;
  }
  // What a step does on Next; throws to stay on it.
  const wizLeave = {
    async folder() {
      await wiz.checking;
      if (onServer()) {
        const chosen = S.settings.spSite && S.settings.spSite !== 'root' && S.settings.spFolder;
        if (!$('wizSpFolder').value || (!wiz.spOk && !toSharePoint() && !$('wizLink').value.trim() && !chosen)) throw new Error(t('Paste the address of the folder in SharePoint.'));
        // Not checked when SharePoint is not set up yet (the setup script, a few steps on, does it).
        if (!wiz.spOk && wiz.warned !== 'folder' && S.info.features.publishTarget === 'sharepoint') {
          wiz.warned = 'folder';
          throw new Error(t('The server cannot reach this folder yet. Give it access as described below, then paste the link again. Click Next again to continue anyway.'));
        }
        await wizSave({ spSite: $('wizSpSite').value || 'root', spFolder: $('wizSpFolder').value });
        schedulePreview();
        return;
      }
      const choice = radio('wlib');
      if (!choice) throw new Error(t('Choose where the signatures are kept.'));
      if (choice === 'link' && !$('wizSpFolder').value) throw new Error(t('Paste the address of the folder in SharePoint.'));
      if (!$('wizFolder').value.trim()) throw new Error(choice === 'link' ? t('This folder is not on this PC yet: NOBS OSign publishes through OneDrive. In SharePoint, open the library and choose Sync, wait until the folder is there, then paste the link again.') : t('Choose the folder on this PC.'));
      if (!$('wizSpFolder').value) throw new Error(t('NOBS OSign cannot tell where this folder is in SharePoint. Choose «SharePoint link» and paste the folder\'s address.'));
      const r = await wizSave({ folder: $('wizFolder').value, spSite: $('wizSpSite').value || 'root', spFolder: $('wizSpFolder').value });
      if (!r.folderFound && wiz.warned !== 'folder') {
        wiz.warned = 'folder';
        $('wizFolderNote').textContent = t('Not found on this PC: is the library synced? Click Next again to continue anyway.');
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
      // Without the setup script's result, the field by hand is where the ID goes: opened.
      if (!$('wizClientId').value.trim()) { $('wizAppManual').open = true; throw new Error(t('Run the setup script first, or enter the client ID under «Or by hand».')); }
      await wizSave({ clientId: $('wizClientId').value, tenantId: $('wizTenantId').value });
    },
    async files() {
      if (!wiz.built) throw new Error(t('Create the add-in files first.'));
    },
    async done() {
      await wizSave({ setupDone: '1' });
    }
  };
  $('wizNext').addEventListener('click', async () => {
    $('wizErr').textContent = '';
    const name = WIZ[wiz.step];
    $('wizNext').disabled = true; $('wizBack').disabled = true;
    try {
      if (wizLeave[name]) await wizLeave[name]();
      if (name === 'done') { $('dlgWizard').close(); toast(t('NOBS OSign is set up.')); return; }
      wiz.step++; wiz.reached = Math.max(wiz.reached, wiz.step); wiz.warned = ''; wizRender();
    } catch (err) { wizError(err); }
    finally { $('wizNext').disabled = false; $('wizBack').disabled = false; }
  });
  $('wizBack').addEventListener('click', () => { if (wiz.step) { wiz.step--; wizRender(); } });
  $('wizLater').addEventListener('click', () => $('dlgWizard').close());
  // Closed in any way (Later, Esc, finished; not minimised): it opens by itself again at the next
  // start only if NOBS OSign is still not set up.
  $('dlgWizard').addEventListener('close', () => { if (!$('dlgWizard').dataset.min) stored(WIZ_KEY, null); });
  document.querySelectorAll('input[name=whost], input[name=wvsrc]').forEach((r) => r.addEventListener('change', wizRender));
  $('wizUrl').addEventListener('input', wizRender);
  $('wizFolder').addEventListener('input', () => { $('wizFolderNote').textContent = ''; wiz.warned = ''; });
  $('wizCfUpload').addEventListener('click', async () => {
    $('wizErr').textContent = '';
    try { await pagesUpload($('wizCfOk'), $('wizCfUpload')); } catch (err) { $('wizErr').textContent = t(err.message); }
  });
  $('wizMake').addEventListener('click', async () => {
    $('wizErr').textContent = ''; $('wizMakeOk').textContent = '';
    try {
      if (!onServer()) await wizSave({ addinFolder: $('wizAddinFolder').value });
      const r = await api('addin/build');
      wiz.built = true;
      afterAddinChange();
      $('wizMakeOk').textContent = onServer() ? t('{0} files (version {1}) made.', r.files, r.version) : t('{0} files (version {1}) in {2}.', r.files, r.version, r.folder);
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
  document.querySelectorAll('button.copy').forEach((b) => b.addEventListener('click', async () => {
    try { await navigator.clipboard.writeText($(b.dataset.copy).textContent); b.textContent = t('Copied'); setTimeout(() => { b.textContent = t('Copy'); }, 1500); } catch (e) { /* select it by hand */ }
  }));
  // Another language: the page starts again in it, and the assistant with it, on this step.
  $('wizLang').addEventListener('change', async () => { await saving; stored(WIZ_KEY, String(wiz.step)); I18N.setLang($('wizLang').value); });
  const wizFinishThen = (btn) => async () => {
    try { await wizLeave.done(); } catch (err) { $('wizErr').textContent = t(err.message); return; }
    $('dlgWizard').close(); $(btn).click();
  };
  $('wizImport').addEventListener('click', wizFinishThen('btnImportOpen'));
  $('wizNewSig').addEventListener('click', wizFinishThen('btnNew'));
  $('btnWizard').addEventListener('click', () => wizOpen(0));
  $('setWizard').addEventListener('click', () => { $('dlgSettings').close(); wizOpen(0); });

