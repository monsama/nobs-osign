/*
 * The sign-in page of the server version: your Windows sign-in name and password (members of the
 * NOBS OSign group in Active Directory), or the local account admin.
 */
'use strict';
(function () {
  const $ = (id) => document.getElementById(id);
  I18N.apply(document.body);

  function error(text) { const e = $('loginError'); e.textContent = text || ''; e.hidden = !text; }

  fetch('/api/login-methods', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' })
    .then((r) => r.json())
    .then((m) => {
      $('loginHint').textContent = m.ad
        ? t('Your Windows sign-in name and password. Only members of the NOBS OSign group may sign in.')
        : t('Active Directory is not set up on the server yet: sign in as admin.');
    })
    .catch(() => {});

  $('loginForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    error('');
    $('loginGo').disabled = true;
    try {
      const r = await fetch('/api/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ login: $('loginName').value, password: $('loginPassword').value })
      });
      const j = await r.json().catch(() => ({ ok: false, error: 'The server answered ' + r.status + '.' }));
      if (!j.ok) { error(t(j.error || 'Failed.')); $('loginPassword').select(); return; }
      location.replace('/');
    } catch (err) {
      error(t('The NOBS OSign server cannot be reached.'));
    } finally {
      $('loginGo').disabled = false;
    }
  });
  $('loginName').focus();
})();
