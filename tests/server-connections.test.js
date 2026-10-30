// node --test tests/server-connections.test.js
//
// The server's connections set in the editor (server/connections.js): only the local account admin
// may see or change them, what is saved wins over .env and takes effect without a restart, the
// passwords are kept sealed and never sent back. Against a stand-in AD and a stand-in SharePoint.
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');
const { createApp } = require('../server/app');
const { createConnections, parseCa } = require('../server/connections');
const { fakeAd } = require('./support/fake-ad');
const { fakeSharePoint } = require('./support/fake-sharepoint');

const BASE = 'DC=example,DC=ch';
const SERVICE = { dn: 'CN=svc-osign,OU=Service,DC=example,DC=ch', pw: 'svc-pw' };
const TENANT = '11111111-2222-3333-4444-555555555555';
const CLIENT = '66666666-7777-8888-9999-000000000000';
const SITE = 'contoso.sharepoint.com:/sites/IT:';
const PEOPLE = { anna: { pw: 'anna-pw', name: 'Anna Muster', mail: 'anna@example.ch', attrs: {}, groups: ['OSign-Admins'] } };
const GROUPS = { 'OSign-Admins': [] };

let ad, sp, srv, base, dataDir, ldapUrl, spUrl;
const log = { info: () => {}, warn: () => {}, error: () => {} };

test.before(async () => {
  ad = fakeAd({ base: BASE, people: PEOPLE, groups: GROUPS, service: SERVICE });
  ldapUrl = await ad.listen();
  sp = fakeSharePoint({ tenantId: TENANT, clientId: CLIENT, secret: 'graph-secret', granted: [SITE] });
  spUrl = await sp.listen();
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-conn-'));
  fs.writeFileSync(path.join(dataDir, 'app_key'), 'a-key-of-at-least-sixteen-characters\n');
  // Nothing set up in .env but what a test needs: the stand-ins' addresses, and plain LDAP on
  // the stand-in's port.
  const env = {
    OSIGN_DATA: dataDir, OSIGN_SECURE_COOKIE: '0', OSIGN_ADMIN_PASSWORD: 'admin-pw', OSIGN_APP_KEY_FILE: path.join(dataDir, 'app_key'),
    OSIGN_LDAP_ALLOW_PLAIN: '1', OSIGN_LDAP_PORTS: new URL(ldapUrl).port,
    OSIGN_GRAPH_URL: spUrl + '/v1.0', OSIGN_LOGIN_URL: spUrl
  };
  srv = http.createServer(createApp({ log, env }));
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  base = 'http://127.0.0.1:' + srv.address().port;
});
test.after(() => { srv.close(); ad.close(); sp.close(); });

async function signIn(login, password) {
  const r = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login, password }) });
  const body = await r.json();
  return { status: r.status, body, cookie: r.status === 200 ? r.headers.get('set-cookie').split(';')[0] : null };
}
const api = (cookie, name, data) => fetch(base + '/api/' + name, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: cookie }, body: JSON.stringify(data || {}) }).then((r) => r.json());

const LDAP_VALUES = () => ({
  ldapUrl, ldapBind: 'cn={login}@example.ch,' + BASE, ldapBase: 'OU=Users,' + BASE, ldapGroup: 'CN=OSign-Admins,OU=Groups,' + BASE,
  ldapServiceDn: SERVICE.dn
});

test('set in the editor: AD sign-in and the service account work at once, without a restart', async () => {
  const admin = (await signIn('admin', 'admin-pw')).cookie;
  let r = await api(admin, 'connections/get');
  assert.equal(r.ok, true);
  assert.equal(r.keyOk, true);
  assert.match(r.problems.signIn, /not set up/i);
  assert.equal((await api(admin, 'info')).user.local, true);
  // Nobody from AD yet.
  assert.equal((await signIn('anna', 'anna-pw')).status, 401);

  r = await api(admin, 'connections/save', { values: LDAP_VALUES(), secrets: { ldapServicePassword: SERVICE.pw } });
  assert.equal(r.ok, true, r.error);
  assert.equal(r.problems.signIn, '');
  assert.equal(r.problems.directory, '');
  assert.equal(r.fields.ldapUrl.value, ldapUrl);
  // The password: where it is from, never what it is.
  assert.ok(r.secrets.ldapServicePassword.saved);
  assert.equal(r.secrets.ldapServicePassword.readable, true);
  assert.ok(!JSON.stringify(r).includes(SERVICE.pw));
  assert.ok(!fs.readFileSync(path.join(dataDir, 'connections.json'), 'utf8').includes(SERVICE.pw), 'kept sealed');

  const anna = await signIn('anna', 'anna-pw');
  assert.equal(anna.status, 200, JSON.stringify(anna.body));
  const info = await api(anna.cookie, 'info');
  assert.equal(info.features.adLookup, true);
  assert.equal(info.user.local, false);
  const who = await api(anna.cookie, 'ad/lookup', { query: 'anna' });
  assert.equal(who.ok, true, who.error);
  // Test, per part: for real, with what is saved.
  let tst = await api(admin, 'connections/test', { what: 'signIn' });
  assert.equal(tst.ok, true, tst.error);
  assert.match(tst.message, /sign-in group is there/);
  tst = await api(admin, 'connections/test', { what: 'directory' });
  assert.equal(tst.ok, true, tst.error);
  assert.match(tst.message, /reads users/);
  assert.match((await api(admin, 'connections/test', { what: 'graph' })).error, /not set up/i);
  assert.match((await api(anna.cookie, 'connections/test', { what: 'signIn' })).error, /Only the local account admin/);
  // A group that is not there is said as that.
  await api(admin, 'connections/save', { values: { ldapGroup: 'CN=Nobody,OU=Groups,' + BASE } });
  assert.match((await api(admin, 'connections/test', { what: 'signIn' })).error, /no sign-in group by that name/);
  await api(admin, 'connections/save', { values: { ldapGroup: LDAP_VALUES().ldapGroup } });
});

test('only the local account admin may see or change the connections', async () => {
  const anna = (await signIn('anna', 'anna-pw')).cookie;
  for (const name of ['connections/get', 'connections/save']) {
    const r = await api(anna, name, { values: { ldapUrl: 'ldaps://evil.example.com' } });
    assert.equal(r.ok, false);
    assert.match(r.error, /Only the local account admin/);
  }
  const admin = (await signIn('admin', 'admin-pw')).cookie;
  assert.equal((await api(admin, 'connections/get')).fields.ldapUrl.value, ldapUrl);
});

test('Graph set in the editor: publishing goes into SharePoint at once', async () => {
  const admin = (await signIn('admin', 'admin-pw')).cookie;
  assert.equal((await api(admin, 'info')).features.publishTarget, 'folder');
  const r = await api(admin, 'connections/save', { values: { graphTenantId: TENANT, graphClientId: CLIENT }, secrets: { graphClientSecret: 'graph-secret' } });
  assert.equal(r.ok, true, r.error);
  assert.equal(r.problems.graph, '');
  assert.equal((await api(admin, 'info')).features.publishTarget, 'sharepoint');
  const found = await api(admin, 'sp/check', { spSite: SITE, spFolder: 'OSign' });
  assert.equal(found.ok, true, found.error);
  // The site not saved yet: the test says the server may not reach it.
  assert.match((await api(admin, 'connections/test', { what: 'graph' })).error, /may not reach the site yet/);
  await api(admin, 'settings/save', { settings: { spSite: SITE, spFolder: 'OSign' } });
  let tg = await api(admin, 'connections/test', { what: 'graph' });
  assert.equal(tg.ok, true, tg.error);
  assert.match(tg.message, /reaches the site of the signatures/);
  // When the secret runs out: said from 30 days before (the setup script sets the date itself).
  const soon = new Date(Date.now() + 10 * 86400000).toISOString().slice(0, 10);
  assert.equal((await api(admin, 'connections/save', { values: { graphSecretExpires: soon } })).ok, true);
  let f = (await api(admin, 'info')).features;
  assert.equal(f.secretExpires, soon);
  assert.ok(f.secretDaysLeft >= 9 && f.secretDaysLeft <= 10, String(f.secretDaysLeft));
  assert.match((await api(admin, 'connections/save', { values: { graphSecretExpires: '31.10.2027' } })).error, /is a date/);
  // Removed again: back to the server's folder.
  await api(admin, 'connections/save', { values: { graphTenantId: '', graphClientId: '' }, secrets: { graphClientSecret: null } });
  assert.equal((await api(admin, 'info')).features.publishTarget, 'folder');
});

test('what is not a directory address, a GUID or one line is refused, and nothing changes', async () => {
  const admin = (await signIn('admin', 'admin-pw')).cookie;
  const bad = [
    [{ values: { ldapUrl: 'ldap://dc01.example.ch:8080' } }, /Port 8080 is not a directory port/],
    [{ values: { ldapUrl: 'https://dc01.example.ch' } }, /like ldaps:\/\//],
    [{ values: { ldapBind: 'example\\user' } }, /\{login\}/],
    [{ values: { graphTenantId: 'contoso' } }, /GUID/],
    [{ values: { cfAccountId: 'abc' } }, /32 letters/],
    [{ values: { adExportAt: '25:00' } }, /time of day/],
    [{ secrets: { cfApiToken: 'two\nlines' } }, /one line/],
    [{ ca: Buffer.from('not a certificate').toString('base64') }, /not a certificate/]
  ];
  for (const [body, err] of bad) {
    const r = await api(admin, 'connections/save', body);
    assert.equal(r.ok, false, JSON.stringify(body));
    assert.match(r.error, err);
  }
  assert.equal((await api(admin, 'connections/get')).fields.ldapUrl.value, ldapUrl);
});

test('an empty field takes .env\'s value; the password in ./secrets only goes to the domain controller .env names', () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-conn-'));
  fs.writeFileSync(path.join(dir, 'svc'), 'file-pw\n');
  const env = { OSIGN_APP_KEY: 'another-key-of-sixteen-chars', OSIGN_LDAP_URL: 'ldaps://dc01.example.ch', OSIGN_LDAP_SERVICE_PASSWORD_FILE: path.join(dir, 'svc'), OSIGN_CF_PROJECT: 'from-env' };
  const c = createConnections({ dataDir: dir, env });
  let e = c.effectiveEnv();
  assert.equal(e.OSIGN_LDAP_URL, 'ldaps://dc01.example.ch');
  assert.equal(e.OSIGN_CF_PROJECT, 'from-env');
  assert.equal(c.answer().fields.cfProject.env, 'from-env');
  assert.equal(c.answer().secrets.ldapServicePassword.file, true);
  // Another domain controller set here: the file's password does not go there.
  assert.deepEqual(c.save({ values: { ldapUrl: 'ldaps://dc02.example.ch', cfProject: 'from-here' } }), { saved: true });
  e = c.effectiveEnv();
  assert.equal(e.OSIGN_LDAP_URL, 'ldaps://dc02.example.ch');
  assert.equal(e.OSIGN_CF_PROJECT, 'from-here');
  assert.equal(e.OSIGN_LDAP_SERVICE_PASSWORD, '');
  assert.equal(e.OSIGN_LDAP_SERVICE_PASSWORD_FILE, undefined);
  assert.equal(c.answer().passwordPinned, 'dc01.example.ch');
  // Entered here, it is used.
  c.save({ secrets: { ldapServicePassword: 'here-pw' } });
  assert.equal(c.effectiveEnv().OSIGN_LDAP_SERVICE_PASSWORD, 'here-pw');
  // With another key it cannot be read, and says so.
  const other = createConnections({ dataDir: dir, env: Object.assign({}, env, { OSIGN_APP_KEY: 'yet-another-key-of-sixteen' }) });
  assert.equal(other.answer().secrets.ldapServicePassword.readable, false);
  assert.equal(other.effectiveEnv().OSIGN_LDAP_SERVICE_PASSWORD, '');
  // Without a key no password is kept at all.
  const none = createConnections({ dataDir: dir, env: { OSIGN_LDAP_URL: 'ldaps://dc01.example.ch' } });
  assert.match(none.save({ secrets: { cfApiToken: 'tok' } }).error, /no key/);
  assert.equal(none.answer().keyOk, false);
});

test('a CA certificate: only a CA\'s, as Windows exports it or as PEM', () => {
  assert.match(parseCa('').error, /Choose/);
  assert.match(parseCa(Buffer.from('-----BEGIN CERTIFICATE-----\nAAAA\n-----END CERTIFICATE-----\n').toString('base64')).error, /not a certificate/);
  // A CA and a server certificate it issued, made with openssl where it is there.
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-ca-'));
  const ssl = (args) => execFileSync('openssl', args, { cwd: dir, stdio: 'ignore' });
  try {
    ssl(['req', '-x509', '-newkey', 'rsa:2048', '-nodes', '-keyout', 'ca.key', '-out', 'ca.pem', '-days', '30', '-subj', '/CN=Example Root CA', '-addext', 'basicConstraints=critical,CA:TRUE']);
    ssl(['req', '-newkey', 'rsa:2048', '-nodes', '-keyout', 'dc.key', '-out', 'dc.csr', '-subj', '/CN=dc01.example.ch']);
    ssl(['x509', '-req', '-in', 'dc.csr', '-CA', 'ca.pem', '-CAkey', 'ca.key', '-CAcreateserial', '-out', 'dc.pem', '-days', '30']);
    ssl(['x509', '-in', 'ca.pem', '-outform', 'DER', '-out', 'ca.cer']);
  } catch (e) { return; }
  const ca = parseCa(fs.readFileSync(path.join(dir, 'ca.cer')).toString('base64'));
  assert.equal(ca.error, undefined);
  assert.equal(ca.certs[0].name, 'Example Root CA');
  assert.match(ca.pem, /^-----BEGIN CERTIFICATE-----/);
  assert.match(parseCa(fs.readFileSync(path.join(dir, 'dc.pem')).toString('base64')).error, /dc01\.example\.ch is a server's own certificate.*Example Root CA/);
});
