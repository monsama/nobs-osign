// node --test tests/server.test.js
//
// The editor's backend for the container (server/), started here as it runs there, with its data
// in a temporary folder and a small LDAP server standing in for Active Directory. It is talked to
// over HTTP the way the web UI does, with the session cookie the sign-in page gets.
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const ldap = require('ldapjs');
const { createApp } = require('../server/app');

const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');
const ADMIN_PW = 'local-admin-secret';
const GROUP = 'CN=OSign-Admins,OU=Groups,DC=example,DC=ch';
// The directory: anna is in the sign-in group (through a nested group), beat is not.
const PEOPLE = {
  anna: { pw: 'anna-pw', name: 'Anna Muster', member: true },
  beat: { pw: 'beat-pw', name: 'Beat Beispiel', member: false }
};

let server, base, dataDir, publishDir, ldapServer, ldapUrl;
const logs = [];
const log = { info: (m) => logs.push(m), warn: (m) => logs.push(m), error: (m) => logs.push(m) };

function startLdap() {
  const s = ldap.createServer();
  s.bind('DC=example,DC=ch', (req, res, next) => {
    const dn = req.dn.toString().toLowerCase();
    const who = Object.keys(PEOPLE).find((k) => dn.startsWith('cn=' + k + '@'));
    if (!who || req.credentials !== PEOPLE[who].pw) return next(new ldap.InvalidCredentialsError());
    res.end(); return next();
  });
  s.search('OU=Users,DC=example,DC=ch', (req, res, next) => {
    const f = req.filter.toString();
    for (const [login, p] of Object.entries(PEOPLE)) {
      if (!f.toLowerCase().includes('samaccountname=' + login)) continue;
      // The "in chain" rule: answered as AD does, by membership through nested groups.
      if (/1\.2\.840\.113556\.1\.4\.1941/.test(f) && !p.member) continue;
      res.send({ dn: 'CN=' + p.name + ',OU=Users,DC=example,DC=ch', attributes: { displayName: p.name, sAMAccountName: login, memberOf: p.member ? ['CN=Nested,OU=Groups,DC=example,DC=ch'] : [] } });
    }
    res.end(); return next();
  });
  return new Promise((resolve) => s.listen(0, '127.0.0.1', () => resolve(s)));
}

function startApp(env) {
  const handle = createApp({ env, log });
  const srv = http.createServer(handle);
  return new Promise((resolve) => srv.listen(0, '127.0.0.1', () => resolve(srv)));
}

test.before(async () => {
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-server-'));
  publishDir = path.join(dataDir, 'publish');
  fs.writeFileSync(path.join(dataDir, 'admin_password'), ADMIN_PW + '\n');
  ldapServer = await startLdap();
  ldapUrl = 'ldap://127.0.0.1:' + ldapServer.address().port;
  server = await startApp({
    OSIGN_DATA: dataDir, OSIGN_PUBLISH_DIR: publishDir,
    OSIGN_ADMIN_PASSWORD_FILE: path.join(dataDir, 'admin_password'), OSIGN_SECURE_COOKIE: '0',
    OSIGN_LDAP_URL: ldapUrl, OSIGN_LDAP_BIND: 'cn={login}@example.ch,DC=example,DC=ch', OSIGN_LDAP_BASE: 'OU=Users,DC=example,DC=ch',
    OSIGN_LDAP_GROUP: GROUP, OSIGN_LDAP_ALLOW_PLAIN: '1'
  });
  base = 'http://127.0.0.1:' + server.address().port;
});
test.after(() => { server.close(); ldapServer.close(); });

// One HTTP request, with everything the tests need to see.
function request(p, { method = 'GET', headers = {}, body } = {}) {
  return new Promise((resolve, reject) => {
    const req = http.request(base + p, { method, headers }, (res) => {
      const parts = [];
      res.on('data', (d) => parts.push(d));
      res.on('end', () => {
        const buf = Buffer.concat(parts);
        let json = null; try { json = JSON.parse(buf.toString('utf8')); } catch (e) { /* not JSON */ }
        resolve({ status: res.statusCode, headers: res.headers, body: buf, json });
      });
    });
    req.on('error', reject);
    if (body !== undefined) req.write(body);
    req.end();
  });
}
const post = (name, data, cookie, headers) => request('/api/' + name, {
  method: 'POST', body: JSON.stringify(data || {}),
  headers: Object.assign({ 'Content-Type': 'application/json', Origin: base }, cookie ? { Cookie: cookie } : {}, headers || {})
});
async function signIn(login, password) {
  const r = await post('login', { login, password });
  assert.equal(r.status, 200, JSON.stringify(r.json));
  const c = String(r.headers['set-cookie']);
  assert.match(c, /HttpOnly/); assert.match(c, /SameSite=Strict/);
  return c.split(';')[0];
}
let admin;
const api = (name, data) => post(name, data, admin).then((r) => r.json);
const cfg = (sigs) => JSON.stringify({ schemaVersion: 1, options: {}, fields: [], signatures: sigs || [] });

test('without a sign-in: the sign-in page and the health check, nothing else', async () => {
  assert.equal((await request('/health')).status, 200);
  const root = await request('/');
  assert.equal(root.status, 302);
  assert.equal(root.headers.location, '/login');
  const page = await request('/login');
  assert.equal(page.status, 200);
  assert.match(page.headers['content-security-policy'], /script-src 'self'/);
  assert.match(page.body.toString(), /loginForm/);
  for (const p of ['/ui/app.js', '/shared/signature-core.js', '/assets/x.png', '/photos/x.jpg']) assert.equal((await request(p)).status, 401, p);
  // Where Microsoft's sign-in comes back to (a cross-site navigation, so no session cookie): it holds nothing.
  const back = await request('/ui/signin.html');
  assert.equal(back.status, 200);
  assert.match(back.body.toString(), /signin\.js/);
  assert.equal((await request('/ui/signin.js')).status, 200);
  // The page may talk to Microsoft's sign-in and Graph, and nothing else besides its server.
  assert.match(page.headers['content-security-policy'], /connect-src 'self' https:\/\/login\.microsoftonline\.com https:\/\/graph\.microsoft\.com;/);
  const info = await post('info', {});
  assert.equal(info.status, 401);
  assert.equal(info.json.ok, false);
});

test('the local admin signs in with the password from the secret file', async () => {
  assert.equal((await post('login', { login: 'admin', password: 'wrong' })).status, 401);
  admin = await signIn('admin', ADMIN_PW);
  const info = await api('info');
  assert.equal(info.backend, 'node');
  assert.equal(info.features.signOut, true);
  assert.equal(info.features.adLookup, false);
  assert.equal(info.user.login, 'admin');
  assert.equal((await request('/', { headers: { Cookie: admin } })).status, 200);
});

test('Active Directory: members of the group (nested ones too) get in; others and wrong passwords not', async () => {
  const anna = await signIn('anna', 'anna-pw');
  assert.equal((await post('info', {}, anna)).json.user.name, 'Anna Muster');
  const beat = await post('login', { login: 'beat', password: 'beat-pw' });
  assert.equal(beat.status, 403);
  assert.match(beat.json.error, /not in the group/);
  assert.equal((await post('login', { login: 'anna', password: 'nope' })).status, 401);
  // An LDAP filter in the name is a name, not a filter.
  assert.equal((await post('login', { login: 'anna)(cn=*', password: 'anna-pw' })).status, 401);
});

test('signing out ends the session', async () => {
  const anna = await signIn('anna', 'anna-pw');
  const out = await post('logout', {}, anna);
  assert.match(String(out.headers['set-cookie']), /Max-Age=0/);
  assert.equal((await post('info', {}, anna)).status, 401);
});

test('after five wrong passwords that name waits, even with the right one', async () => {
  for (let i = 0; i < 5; i++) assert.equal((await post('login', { login: 'Anna', password: 'wrong' + i })).status, 401);
  const r = await post('login', { login: 'anna', password: 'anna-pw' });
  assert.equal(r.status, 429);
});

test('the directory not answering is said as that, and the local admin still works', async () => {
  const other = await startApp({
    OSIGN_DATA: fs.mkdtempSync(path.join(os.tmpdir(), 'osign-server-')), OSIGN_ADMIN_PASSWORD: ADMIN_PW, OSIGN_SECURE_COOKIE: '0',
    OSIGN_LDAP_URL: 'ldap://127.0.0.1:1', OSIGN_LDAP_BIND: '{login}@example.ch', OSIGN_LDAP_BASE: 'DC=example,DC=ch', OSIGN_LDAP_GROUP: GROUP, OSIGN_LDAP_ALLOW_PLAIN: '1'
  });
  const at = 'http://127.0.0.1:' + other.address().port;
  const login = (l, p) => fetch(at + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: l, password: p }) });
  const r = await login('anna', 'anna-pw');
  assert.equal(r.status, 503);
  assert.match((await r.json()).error, /cannot be reached/);
  assert.equal((await login('admin', ADMIN_PW)).status, 200);
  other.close();
});

test('plain ldap:// without StartTLS is refused unless allowed', async () => {
  const other = await startApp({
    OSIGN_DATA: fs.mkdtempSync(path.join(os.tmpdir(), 'osign-server-')), OSIGN_SECURE_COOKIE: '0',
    OSIGN_LDAP_URL: ldapUrl, OSIGN_LDAP_BIND: 'cn={login}@example.ch,DC=example,DC=ch', OSIGN_LDAP_BASE: 'OU=Users,DC=example,DC=ch', OSIGN_LDAP_GROUP: GROUP
  });
  const r = await fetch('http://127.0.0.1:' + other.address().port + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'anna', password: 'anna-pw' }) });
  assert.equal(r.status, 503);
  assert.match((await r.json()).error, /unencrypted/);
  other.close();
});

test('the API takes only JSON POSTs from this site', async () => {
  assert.equal((await post('info', {}, admin, { Origin: 'https://evil.example' })).status, 403);
  assert.equal((await post('info', {}, admin, { 'Sec-Fetch-Site': 'cross-site' })).status, 403);
  assert.equal((await request('/api/info', { method: 'POST', body: '{}', headers: { 'Content-Type': 'text/plain', Cookie: admin } })).status, 403);
  assert.equal((await request('/api/info', { headers: { Cookie: admin } })).status, 403);
  // A sign-in from another site is refused too: nobody can sign a browser in to their own account.
  assert.equal((await post('login', { login: 'admin', password: ADMIN_PW }, null, { Origin: 'https://evil.example' })).status, 403);
});

test('serves the UI and the shared engine, nothing outside them', async () => {
  const ui = await request('/', { headers: { Cookie: admin } });
  assert.match(ui.body.toString(), /<title>NOBS OSign<\/title>/);
  assert.equal(ui.headers['x-frame-options'], 'DENY');
  assert.match((await request('/shared/signature-core.js', { headers: { Cookie: admin } })).body.toString(), /SignatureCore/);
  for (const p of ['/ui/../../package.json', '/ui/..%2f..%2fpackage.json', '/assets/..%2fdraft.json', '/ui/%2e%2e/%2e%2e/server/app.js', '/photos/..%2f..%2fadmin_password']) {
    assert.equal((await request(p, { headers: { Cookie: admin } })).status, 404, p);
  }
});

test('draft: saved as sent, and checked', async () => {
  assert.equal((await api('draft/save', { configJson: '{"schemaVersion":2}' })).ok, false);
  assert.equal((await api('draft/save', { configJson: 'nope' })).ok, false);
  const text = cfg([{ id: 'a', name: 'Ä "quoted" <b>', html: '<p>@@mail</p>', compiled: '', audience: { departments: ['X'], domains: [] } }]);
  assert.equal((await api('draft/save', { configJson: text })).ok, true);
  const g = await api('draft/get');
  assert.equal(g.configJson, text);
  assert.equal(g.dirty, true);
});

test('images: only real PNG/JPG/GIF, safe names, never overwritten', async () => {
  assert.equal((await api('assets/upload', { name: 'x.svg', data: Buffer.from('<svg/>').toString('base64') })).ok, false);
  assert.equal((await api('assets/upload', { name: 'fake.png', data: Buffer.from('GIF89a....').toString('base64') })).ok, false);
  assert.equal((await api('assets/upload', { name: 'cut.png', data: '!!not base64!!' })).ok, false);
  const a = await api('assets/upload', { name: '../../Our Logo!.png', data: PNG.toString('base64') });
  assert.equal(a.path, 'assets/Our-Logo.png');
  assert.equal((await api('assets/upload', { name: 'Our Logo.png', data: PNG.toString('base64') })).name, 'Our-Logo.png');
  const b = await api('assets/upload', { name: 'Our Logo.png', data: Buffer.concat([PNG, Buffer.from([0])]).toString('base64') });
  assert.equal(b.name, 'Our-Logo-2.png');
  const img = await request('/assets/Our-Logo.png', { headers: { Cookie: admin } });
  assert.equal(img.headers['content-type'], 'image/png');
  assert.ok(img.body.equals(PNG));
  assert.equal((await api('assets/delete', { name: 'Our-Logo-2.png' })).ok, true);
  assert.equal((await api('assets/delete', { name: '../draft.json' })).ok, false);
  assert.deepEqual((await api('draft/get')).assets.map((x) => x.name), ['Our-Logo.png']);
});

test('settings: the publish folder is the server\'s; the add-in\'s settings are checked', async () => {
  const s = await api('settings/save', { settings: { folder: '/etc', photosFolder: '/root', clientId: '00000000-0000-0000-0000-000000000001', addinUrl: 'https://osign.example.pages.dev/' } });
  assert.equal(s.ok, true);
  assert.equal(s.settings.folder, publishDir);
  assert.equal(s.settings.photosFolder, '');
  assert.equal(s.settings.clientId, '00000000-0000-0000-0000-000000000001');
  assert.equal(s.settings.addinUrl, 'https://osign.example.pages.dev');
  assert.equal((await api('settings/save', { settings: { clientId: 'nope' } })).ok, false);
  assert.equal((await api('settings/save', { settings: { addinUrl: 'http://plain.example/' } })).ok, false);
  assert.equal((await api('settings/save', { settings: { spFolder: 'a/../b' } })).ok, false);
  assert.equal((await api('settings/get')).settings.clientId, '00000000-0000-0000-0000-000000000001');
});

test('load published before anything was published says so', async () => {
  const r = await api('load-published');
  assert.equal(r.ok, false);
  assert.match(r.error, /Nothing has been published/);
});

test('publish writes signatures.json and assets, only what changed, and never touches photos', async () => {
  fs.mkdirSync(path.join(publishDir, 'photos'), { recursive: true });
  fs.writeFileSync(path.join(publishDir, 'photos', 'mustera.jpg'), PNG);
  fs.mkdirSync(path.join(publishDir, 'assets'), { recursive: true });
  fs.writeFileSync(path.join(publishDir, 'assets', 'old.png'), PNG);
  fs.writeFileSync(path.join(publishDir, 'assets', 'notes.txt'), 'not an image: left alone');
  const text = cfg([{ id: 'std', name: 'Standard', html: '<p><img src="assets/Our-Logo.png"></p>' }]);
  const p = await api('publish', { configJson: text, message: 'First' });
  assert.equal(p.ok, true, JSON.stringify(p));
  assert.equal(p.changed, true);
  assert.equal(fs.readFileSync(path.join(publishDir, 'signatures.json'), 'utf8'), text);
  assert.deepEqual(fs.readdirSync(path.join(publishDir, 'assets')).sort(), ['Our-Logo.png', 'notes.txt']);
  assert.deepEqual(fs.readdirSync(path.join(publishDir, 'photos')), ['mustera.jpg']);
  assert.equal((await api('draft/get')).dirty, false);
  const before = fs.statSync(path.join(publishDir, 'signatures.json')).mtimeMs;
  assert.equal((await api('publish', { configJson: text })).changed, false);
  assert.equal(fs.statSync(path.join(publishDir, 'signatures.json')).mtimeMs, before);
  // The preview's photos come from the published folder.
  const photo = await request('/photos/mustera.jpg', { headers: { Cookie: admin } });
  assert.equal(photo.status, 200);
  assert.equal(photo.headers['content-type'], 'image/jpeg');
  assert.equal((await api('settings/get')).photosFound, true);
});

test('load published replaces the draft and its images, and keeps the old draft once', async () => {
  await api('assets/upload', { name: 'stray.png', data: PNG.toString('base64') });
  await api('draft/save', { configJson: cfg([{ id: 'local', name: 'Local only', html: '' }]) });
  const r = await api('load-published');
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.match(r.configJson, /"Standard"/);
  assert.deepEqual(r.assets.map((a) => a.name), ['Our-Logo.png']);
  assert.equal(r.dirty, false);
  assert.match(fs.readFileSync(path.join(dataDir, 'draft.json.bak'), 'utf8'), /Local only/);
});

test('what the server version cannot do is said plainly', async () => {
  for (const name of ['pick-folder', 'sp/resolve', 'ad/lookup', 'ad/read', 'quit']) {
    const r = await api(name);
    assert.equal(r.ok, false, name);
    assert.ok(r.error, name);
  }
  assert.deepEqual(await api('ad/snapshot'), { ok: true, snapshot: null, matches: false });
  assert.equal((await post('nope', {}, admin)).status, 404);
});
