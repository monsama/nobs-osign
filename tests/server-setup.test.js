// node --test tests/server-setup.test.js   (the script itself needs pwsh, or Windows PowerShell)
//
// The setup script for Microsoft 365 (server/setup.js, server/setup-template.ps1): made for one
// setup by the local account admin, run as it is against a stand-in for Microsoft's sign-in and
// Graph, it makes both app registrations, gives the server's the site, sends its result back with
// its one-time code, takes back the rights it used and deletes itself. Run again, it uses the apps
// that are there.
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync, execFile } = require('node:child_process');
const { createApp } = require('../server/app');
const { fakeMicrosoft, TENANT } = require('./support/fake-microsoft');

const PS = process.env.OSIGN_TEST_SHELL || 'pwsh';
let psOk = true;
try { execFileSync(PS, ['-NoProfile', '-Command', '1'], { stdio: 'ignore' }); } catch (e) { psOk = false; }
const skip = !psOk && PS + ' is not installed';

const SITE = 'contoso.sharepoint.com:/sites/osign:';

let ms, msUrl, srv, base, dataDir, admin;
const log = { info: () => {}, warn: () => {}, error: () => {} };
test.before(async () => {
  ms = fakeMicrosoft();
  msUrl = await ms.listen();
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-setup-'));
  srv = http.createServer(createApp({ log, env: { OSIGN_DATA: dataDir, OSIGN_ADMIN_PASSWORD: 'pw', OSIGN_SECURE_COOKIE: '0', OSIGN_APP_KEY: 'a-key-for-the-setup-test' } }));
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  base = 'http://127.0.0.1:' + srv.address().port;
  const r = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'admin', password: 'pw' }) });
  admin = r.headers.get('set-cookie').split(';')[0];
});
test.after(() => { srv.close(); ms.close(); });
const api = (name, data, cookie) => fetch(base + '/api/' + name, { method: 'POST', headers: Object.assign({ 'Content-Type': 'application/json' }, cookie === null ? {} : { Cookie: cookie || admin }), body: JSON.stringify(data || {}) }).then((r) => r.json());

async function makeScript(groups, users) {
  await api('settings/save', { settings: { spSite: SITE, spFolder: 'OSign', addinUrl: 'https://osign.contoso.com/' } });
  const s = await api('setup/start', { origin: base, groups, users });
  assert.equal(s.ok, true, s.error);
  const r = await fetch(base + '/setup/nobs-osign-setup.ps1?id=' + s.id, { headers: { Cookie: admin } });
  assert.equal(r.status, 200);
  assert.match(r.headers.get('content-disposition'), /nobs-osign-setup\.ps1/);
  const file = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'osign-script-')), 'nobs-osign-setup.ps1');
  fs.writeFileSync(file, Buffer.from(await r.arrayBuffer()));
  return file;
}
function run(file) {
  return new Promise((resolve) => execFile(PS, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', file], {
    env: Object.assign({}, process.env, { OSIGN_SETUP_LOGIN_URL: msUrl, OSIGN_SETUP_GRAPH_URL: msUrl + '/v1.0', OSIGN_SETUP_TEST_FETCH: '1' }), timeout: 120000
  }, (err, stdout, stderr) => resolve({ code: err ? err.code : 0, out: stdout + stderr })));
}

test('the setup script: both apps, the site, the result to the server, the rights taken back, and gone', { skip }, async () => {
  assert.equal((await api('setup/status')).state, 'none');
  // Someone else registered apps with these names first: they are never used.
  ms.state.apps.push({ id: 'obj-planted', appId: '99999999-9999-9999-9999-999999999999', displayName: 'NOBS OSign server', owners: ['someone-else'], requiredResourceAccess: [] });
  ms.state.apps.push({ id: 'obj-planted2', appId: '88888888-8888-8888-8888-888888888888', displayName: 'NOBS OSign', owners: ['someone-else'], requiredResourceAccess: [] });
  const file = await makeScript(true, true);
  const text = fs.readFileSync(file, 'utf8');
  assert.ok(text.startsWith('﻿'), 'with a BOM for Windows PowerShell');
  assert.match(text, /\$Site = 'contoso\.sharepoint\.com:\/sites\/osign:'/);
  assert.match(text, /\$AddinHost = 'osign\.contoso\.com'/);
  assert.equal((await api('setup/status')).state, 'waiting');

  const r = await run(file);
  assert.equal(r.code, 0, r.out);
  assert.equal(fs.existsSync(file), false, 'it deleted itself');

  const st = ms.state;
  const addin = st.apps.find((a) => a.displayName === 'NOBS OSign' && a.id !== 'obj-planted2');
  const server = st.apps.find((a) => a.displayName === 'NOBS OSign server' && a.id !== 'obj-planted');
  assert.deepEqual(st.apps.filter((a) => a.id.startsWith('obj-planted')).map((a) => a.requiredResourceAccess.length), [0, 0], 'the planted apps got nothing');
  assert.ok(!st.sitePerms.some((x) => x.grantedToIdentitiesV2[0].application.id.startsWith('9999')), 'the planted app has no site access');
  assert.deepEqual(addin.spa.redirectUris, ['brk-multihub://osign.contoso.com']);
  assert.deepEqual(addin.requiredResourceAccess[0].resourceAccess.map((x) => x.id).sort(), ['scope-GroupMember.Read.All', 'scope-Sites.Read.All', 'scope-User.Read']);
  const addinSp = st.sps.find((s) => s.appId === addin.appId);
  assert.equal(st.grants.find((g) => g.clientId === addinSp.id).scope, 'User.Read Sites.Read.All GroupMember.Read.All');
  assert.equal(st.grants.find((g) => g.clientId === addinSp.id).consentType, 'AllPrincipals');
  assert.deepEqual(server.requiredResourceAccess[0].resourceAccess.map((x) => x.id), ['role-sites-selected', 'role-users', 'role-groups']);
  // Sites.Selected, and (asked for) reading users and groups for the preview.
  assert.deepEqual(st.roleAssignments.map((r) => r.appRoleId).sort(), ['role-groups', 'role-sites-selected', 'role-users']);
  assert.deepEqual(st.sitePerms.map((x) => [x.roles, x.grantedToIdentitiesV2[0].application.id]), [[['write'], server.appId]]);
  // The sign-in's consent is gone again.
  assert.equal(st.grants.find((g) => g.id === 'grant-cli-me'), undefined);

  // The server signs in with its own certificate: only its public half went to the app, and no
  // secret was made.
  assert.equal(server.keyCredentials.length, 1);
  assert.equal(server.keyCredentials[0].type, 'AsymmetricX509Cert');
  assert.equal(st.passwords.length, 0, 'no client secret');
  const certFile = fs.readFileSync(path.join(dataDir, 'graph-cert.json'), 'utf8');
  assert.ok(!certFile.includes('PRIVATE KEY'), 'the key is kept sealed');
  assert.ok(!text.includes('PRIVATE KEY'), 'the key never goes into the script');
  const cert = new crypto.X509Certificate(Buffer.from(server.keyCredentials[0].key, 'base64'));

  // The server has what it needs, and the assistant sees it.
  const s = await api('setup/status');
  assert.equal(s.state, 'done');
  assert.equal(s.result.addinClientId, addin.appId);
  assert.equal(s.result.serverClientId, server.appId);
  assert.equal(s.result.credential, 'certificate');
  let info = await api('info');
  assert.equal(info.features.publishTarget, 'sharepoint');
  assert.equal(info.features.graphClientId, server.appId);
  assert.equal(info.features.graphCredential, 'certificate');
  // The certificate's end, so the editor warns in time.
  assert.equal(info.features.secretExpires, new Date(cert.validTo).toISOString().slice(0, 10));
  const settings = (await api('settings/get')).settings;
  assert.equal(settings.clientId, addin.appId);
  assert.equal(settings.tenantId, TENANT);
  let conn = await api('connections/get');
  assert.equal(conn.cert.thumbprint, cert.fingerprint.replace(/:/g, ''));

  // Again (another site, say), where the tenant refuses certificates on apps: the apps of the first
  // run are used again (the server knows them, and they are yours), nothing twice, and the server
  // gets a client secret instead.
  st.refuseCerts = true;
  const again = await run(await makeScript(false));
  assert.equal(again.code, 0, again.out);
  assert.equal(st.apps.length, 4);
  assert.equal(st.roleAssignments.length, 3);
  assert.equal(st.sitePerms.length, 1);
  assert.equal(st.passwords.length, 1, 'a secret instead');
  assert.equal(addin.requiredResourceAccess[0].resourceAccess.length, 3, 'what it asked for stays');
  info = await api('info');
  assert.equal(info.features.graphCredential, 'secret');
  assert.equal(info.features.secretExpires, st.passwords[0].endDateTime.slice(0, 10));
  assert.ok(!JSON.stringify(await api('setup/status')).includes(st.passwords[0].secretText), 'the secret is never sent to the page');
  conn = await api('connections/get');
  assert.equal(conn.secrets.graphClientSecret.readable, true);
  assert.equal(conn.cert, null);
  assert.ok(!fs.readFileSync(path.join(dataDir, 'connections.json'), 'utf8').includes(st.passwords[0].secretText), 'kept sealed');

  // And with certificates allowed again: the certificate, and the old secret goes from the app.
  st.refuseCerts = false;
  const third = await run(await makeScript(false));
  assert.equal(third.code, 0, third.out);
  assert.equal((await api('info')).features.graphCredential, 'certificate');
  assert.equal(st.passwords.filter((x) => !x.removed).length, 0, 'the old secret is removed');
  assert.equal((await api('connections/get')).secrets.graphClientSecret.saved, '');
});

test('the one-time code: only with it, only once, and the script only for the admin', async () => {
  await api('settings/save', { settings: { spSite: SITE, spFolder: 'OSign', addinUrl: 'https://osign.contoso.com/' } });
  const s = await api('setup/start', { origin: base });
  assert.equal(s.ok, true, s.error);
  // A member of the sign-in group cannot change where the server publishes (it goes into the script).
  // (No AD here: the rule is checked in server-ad.test.js.)
  // Without a session: not the script.
  assert.equal((await fetch(base + '/setup/nobs-osign-setup.ps1?id=' + s.id)).status, 401);
  assert.equal((await fetch(base + '/setup/nobs-osign-setup.ps1?id=nope', { headers: { Cookie: admin } })).status, 404);
  const text = await (await fetch(base + '/setup/nobs-osign-setup.ps1?id=' + s.id, { headers: { Cookie: admin } })).text();
  const token = /\$Token = '([^']+)'/.exec(text)[1];
  const result = { tenantId: TENANT, siteUrl: 'https://contoso.sharepoint.com/sites/osign', addin: { clientId: '22222222-2222-2222-2222-222222222222' }, server: { clientId: '33333333-3333-3333-3333-333333333333', secret: 'the-secret' } };
  // A wrong code: refused, and nothing changes.
  let r = await fetch(base + '/api/setup/finish', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(Object.assign({ token: 'wrong' }, result)) });
  assert.equal(r.status, 403);
  // What went wrong, said by the script: the assistant shows it, and the code still holds.
  r = await api('setup/finish', { token, error: 'Microsoft Graph: Insufficient privileges' }, null);
  assert.equal(r.ok, true);
  assert.deepEqual([(await api('setup/status')).state, (await api('setup/status')).problem], ['failed', 'Microsoft Graph: Insufficient privileges']);
  r = await api('setup/finish', Object.assign({ token }, result), null);
  assert.equal(r.ok, true, r.error);
  assert.equal((await api('info')).features.graphClientId, '33333333-3333-3333-3333-333333333333');
  // Used up.
  r = await fetch(base + '/api/setup/finish', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(Object.assign({ token }, result)) });
  assert.equal(r.status, 403);
  // Not from another site's page.
  r = await fetch(base + '/api/setup/finish', { method: 'POST', headers: { 'Content-Type': 'application/json', Origin: 'https://evil.example' }, body: JSON.stringify(Object.assign({ token }, result)) });
  assert.equal(r.status, 403);
});

test('the setup script needs the folder and the add-in\'s address first', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-setup-'));
  const other = http.createServer(createApp({ log, env: { OSIGN_DATA: dir, OSIGN_ADMIN_PASSWORD: 'pw', OSIGN_SECURE_COOKIE: '0', OSIGN_APP_KEY: 'a-key-for-the-setup-test' } }));
  await new Promise((r) => other.listen(0, '127.0.0.1', r));
  const at = 'http://127.0.0.1:' + other.address().port;
  const c = (await fetch(at + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'admin', password: 'pw' }) })).headers.get('set-cookie').split(';')[0];
  const call = (n, d) => fetch(at + '/api/' + n, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: c }, body: JSON.stringify(d || {}) }).then((r) => r.json());
  await call('settings/save', { settings: { spFolder: '' } });
  assert.match((await call('setup/start', { origin: at })).error, /SharePoint folder first/);
  await call('settings/save', { settings: { spSite: SITE, spFolder: 'OSign' } });
  assert.match((await call('setup/start', { origin: at })).error, /add-in's address first/);
  assert.match((await call('setup/start', { origin: 'javascript:alert(1)' })).error, /editor's address/);
  // Nothing that could end a string in PowerShell gets into the settings or a script.
  for (const site of ["contoso.sharepoint.com:/sites/IT\u2019;New-Item(\u2019x\u2019);\u2019:", "contoso.sharepoint.com:/sites/IT';x;':"]) {
    assert.match((await call('settings/save', { settings: { spSite: site } })).error, /SharePoint site/);
  }
  other.close();
});
