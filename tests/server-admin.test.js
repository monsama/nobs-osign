// node --test tests/server-admin.test.js
//
// What the server keeps for its admins: the audit log (server/audit.js), the backup and bringing
// it back (server/backup.js), the notice of a newer version (server/updates.js), four eyes before
// publishing, and what changed between versions (server/diff.js). Two people from a stand-in AD
// (tests/support/fake-ad.js) and the local admin.
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../server/app');
const { fakeAd } = require('./support/fake-ad');
const { diffConfigs } = require('../server/diff');
const { newer } = require('../server/updates');

const BASE = 'DC=example,DC=ch';
const SERVICE = { dn: 'CN=svc-osign,OU=Service,DC=example,DC=ch', pw: 'svc-pw' };
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');
const log = { info: () => {}, warn: () => {}, error: () => {} };
const cfg = (sigs, options) => JSON.stringify({ schemaVersion: 1, options: options || {}, fields: [], signatures: sigs || [] });

let ad, ldapUrl, release;
test.before(async () => {
  ad = fakeAd({ base: BASE, service: SERVICE, groups: { 'NOBS-OSign-Admins': [] }, people: {
    anna: { pw: 'anna-pw', name: 'Anna Muster', mail: 'anna@example.ch', groups: ['NOBS-OSign-Admins'] },
    beat: { pw: 'beat-pw', name: 'Beat Beispiel', mail: 'beat@example.ch', groups: ['NOBS-OSign-Admins'] }
  } });
  ldapUrl = await ad.listen();
  // GitHub's latest release, as the update check asks for it.
  release = http.createServer((req, res) => { res.writeHead(200, { 'Content-Type': 'application/json' }); res.end(JSON.stringify(release.body)); });
  await new Promise((r) => release.listen(0, '127.0.0.1', r));
  release.body = { tag_name: 'v99.0.0', html_url: 'https://github.com/monsama/nobs-osign/releases/tag/v99.0.0', published_at: '2030-01-01T00:00:00Z' };
});
test.after(() => { ad.close(); release.close(); });

async function server(extra) {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-admin-'));
  const env = Object.assign({
    OSIGN_DATA: dataDir, OSIGN_SECURE_COOKIE: '0', OSIGN_ADMIN_PASSWORD: 'admin-pw', OSIGN_APP_KEY: 'the-first-servers-key',
    OSIGN_LDAP_URL: ldapUrl, OSIGN_LDAP_ALLOW_PLAIN: '1', OSIGN_LDAP_BIND: 'cn={login}@example.ch,' + BASE, OSIGN_LDAP_BASE: 'OU=Users,' + BASE,
    OSIGN_LDAP_GROUP: 'CN=NOBS-OSign-Admins,OU=Groups,' + BASE
  }, extra || {});
  const srv = http.createServer(createApp({ log, env }));
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  const at = 'http://127.0.0.1:' + srv.address().port;
  const signIn = async (login, password, headers) => {
    const r = await fetch(at + '/api/login', { method: 'POST', headers: Object.assign({ 'Content-Type': 'application/json' }, headers || {}), body: JSON.stringify({ login, password }) });
    return (r.headers.get('set-cookie') || '').split(';')[0];
  };
  const as = (cookie) => (name, data) => fetch(at + '/api/' + name, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: cookie }, body: JSON.stringify(data || {}) }).then((r) => r.json());
  return { srv, at, dataDir, signIn, as };
}

test('the audit log: sign-ins, failed ones, publishing and the connections, for the local admin only', async () => {
  const s = await server();
  try {
    const admin = s.as(await s.signIn('admin', 'admin-pw', { 'X-Forwarded-For': '10.1.2.3' }));
    await s.signIn('anna', 'Pa55-not-logged');
    const anna = s.as(await s.signIn('anna', 'anna-pw'));
    assert.equal((await anna('publish', { configJson: cfg([{ id: 'a', name: 'A', html: '<p>a</p>' }]), message: 'first' })).ok, true);
    await admin('connections/save', { values: {}, secrets: { cfApiToken: 'tok-123' } });
    assert.match((await anna('audit/list')).error, /Only the local account admin/);
    const e = (await admin('audit/list')).entries;
    assert.deepEqual(e.map((x) => [x.who, x.what, x.ok]), [
      ['admin', 'connections/save', true], ['anna', 'publish', true], ['anna', 'signin', true], ['anna', 'signin', false], ['admin', 'signin', true]
    ]);
    assert.equal(e[4].ip, '10.1.2.3');
    assert.equal(e[1].detail, 'changed: first');
    assert.match(e[0].detail, /cfApiToken set/);
    assert.ok(!fs.readFileSync(path.join(s.dataDir, 'audit.log'), 'utf8').includes('tok-123'), 'never a secret');
    assert.ok(!fs.readFileSync(path.join(s.dataDir, 'audit.log'), 'utf8').includes('Pa55-not-logged'), 'never a password');
    assert.deepEqual((await admin('audit/list', { failed: true })).entries.map((x) => x.who), ['anna']);
    assert.deepEqual((await admin('audit/list', { what: 'publish' })).entries.length, 1);
  } finally { s.srv.close(); }
});

test('four eyes: asked for by one, published by another, never by the same; a changed draft withdraws it', async () => {
  const s = await server();
  try {
    const admin = s.as(await s.signIn('admin', 'admin-pw'));
    const anna = s.as(await s.signIn('anna', 'anna-pw'));
    const beat = s.as(await s.signIn('beat', 'beat-pw'));
    assert.equal((await anna('publish', { configJson: cfg([{ id: 'a', name: 'A', html: '<p>one</p>' }]) })).ok, true);
    // Only the local admin turns it on.
    assert.match((await anna('settings/save', { settings: { requireApproval: 'on' } })).error, /Only the local account admin/);
    assert.equal((await admin('settings/save', { settings: { requireApproval: 'on' } })).ok, true);
    assert.equal((await anna('info')).features.requireApproval, true);
    assert.match((await anna('publish', { configJson: cfg([]) })).error, /second person's approval/);

    const next = cfg([{ id: 'a', name: 'A', html: '<p>two</p>' }, { id: 'b', name: 'B', html: '<p>b</p>' }]);
    let r = await anna('approval/request', { configJson: next, message: 'B for sales' });
    assert.equal(r.ok, true, r.error);
    assert.equal(r.pending.mine, true);
    assert.equal(r.pending.mayApprove, false);
    // What would go out, against what is published.
    assert.deepEqual(r.pending.changes.map((c) => c.kind + ' ' + c.name), ['html A', 'added B']);
    assert.match((await anna('approval/approve')).error, /Someone else/);
    r = await beat('approval/get');
    assert.equal(r.pending.by, 'Anna Muster');
    assert.equal(r.pending.mayApprove, true);
    r = await beat('approval/approve');
    assert.equal(r.ok, true, r.error);
    assert.equal(r.changed, true);
    assert.equal((await beat('approval/get')).pending, null);
    const h = (await beat('history/list')).versions;
    assert.equal(h[0].by, 'Anna Muster (Beat Beispiel approved)');
    assert.equal(h[0].message, 'B for sales');

    // Asked for, then the draft changes: it has to be asked for again.
    await anna('approval/request', { configJson: cfg([{ id: 'a', name: 'A', html: '<p>three</p>' }]) });
    await anna('draft/save', { configJson: cfg([{ id: 'a', name: 'A', html: '<p>four</p>' }]) });
    assert.equal((await beat('approval/get')).pending.current, false);
    assert.match((await beat('approval/approve')).error, /draft changed/);
    assert.equal((await beat('approval/cancel')).ok, true);
    const log = (await admin('audit/list', { what: 'approval/approve' })).entries;
    assert.deepEqual(log.map((x) => [x.who, x.ok]), [['beat', false], ['beat', true], ['anna', false]]);
  } finally { s.srv.close(); }
});

test('history: what changed in a version, and against the draft', async () => {
  const s = await server();
  try {
    const admin = s.as(await s.signIn('admin', 'admin-pw'));
    await admin('publish', { configJson: cfg([{ id: 'a', name: 'A', html: '<table><tr><td>Anna</td></tr><tr><td>Mobile</td></tr></table>', isDefaultNew: true }]) });
    await admin('publish', { configJson: cfg([{ id: 'a', name: 'A sales', html: '<table><tr><td>Anna</td></tr><tr><td>Phone</td></tr></table>', isDefaultNew: true, groups: ['sg-sales'] }], { disableClientSignature: true }) });
    const [latest, first] = (await admin('history/list')).versions;
    const d = (await admin('history/diff', { id: latest.id })).changes;
    assert.deepEqual(d.map((c) => c.kind), ['renamed', 'html', 'rules', 'options']);
    assert.equal(d[0].detail, 'A');
    assert.deepEqual(d[1].lines, [{ op: '-', text: '<td>Mobile</td>' }, { op: '+', text: '<td>Phone</td>' }]);
    assert.equal(d[2].detail, 'groups');
    assert.equal(d[3].detail, 'disableClientSignature');
    assert.deepEqual((await admin('history/diff', { id: first.id })).changes.map((c) => c.kind), ['added']);
    await admin('draft/save', { configJson: cfg([], { disableClientSignature: true }) });
    assert.deepEqual((await admin('history/diff', { id: latest.id, against: 'draft' })).changes.map((c) => c.kind + ' ' + c.name), ['removed A sales']);
    assert.match((await admin('history/diff', { id: '../x' })).error, /no such version/);
  } finally { s.srv.close(); }
  assert.deepEqual(diffConfigs('', cfg([])), []);
});

test('a backup: locked with a passphrase, brought back on a new server with another key', async () => {
  const s = await server();
  const t = await server({ OSIGN_APP_KEY: 'the-second-servers-own-key' });
  try {
    const admin = s.as(await s.signIn('admin', 'admin-pw'));
    const anna = s.as(await s.signIn('anna', 'anna-pw'));
    await admin('assets/upload', { name: 'logo.png', data: PNG.toString('base64') });
    await admin('publish', { configJson: cfg([{ id: 'a', name: 'A', html: '<img src="assets/logo.png">' }]), message: 'with logo' });
    await admin('connections/save', { values: { cfProject: 'osign' }, secrets: { cfApiToken: 'tok-secret-1' } });
    assert.match((await anna('backup/export', { passphrase: 'correct horse battery' })).error, /Only the local account admin/);
    assert.match((await admin('backup/export', { passphrase: 'short' })).error, /at least 12/);
    const b = await admin('backup/export', { passphrase: 'correct horse battery' });
    assert.equal(b.ok, true, b.error);
    assert.match(b.name, /^nobs-osign-backup-\d{4}-\d{2}-\d{2}\.osignbak$/);
    const raw = Buffer.from(b.data, 'base64').toString('utf8');
    assert.ok(!raw.includes('tok-secret-1') && !raw.includes('with logo') && !raw.includes('the-first-servers-key'), 'nothing readable without the passphrase');

    const admin2 = t.as(await t.signIn('admin', 'admin-pw'));
    await admin2('draft/save', { configJson: cfg([{ id: 'z', name: 'Z', html: '<p>z</p>' }]) });
    assert.match((await admin2('backup/restore', { data: b.data, passphrase: 'wrong passphrase!' })).error, /passphrase is not the one/);
    assert.match((await admin2('backup/restore', { data: Buffer.from('{}').toString('base64'), passphrase: 'x' })).error, /not a backup/);
    const r = await admin2('backup/restore', { data: b.data, passphrase: 'correct horse battery' });
    assert.equal(r.ok, true, r.error);
    // The draft, its image, the history and the connections, with the token readable again.
    const draft = await admin2('draft/get');
    assert.match(draft.configJson, /assets\/logo\.png/);
    assert.deepEqual(draft.assets.map((a) => a.name), ['logo.png']);
    assert.equal((await admin2('history/list')).versions[0].message, 'with logo');
    const c = await admin2('connections/get');
    assert.equal(c.fields.cfProject.value, 'osign');
    assert.equal(c.secrets.cfApiToken.readable, true);
    // This server's own audit log stays, with the restore in it.
    assert.equal((await admin2('audit/list', { what: 'backup/restore' })).entries[0].ok, true);
  } finally { s.srv.close(); t.srv.close(); }
});

test('a newer version: said in info when the server may look, never otherwise', async () => {
  assert.equal(newer('0.10.0', '0.9.9'), true);
  assert.equal(newer('0.9.0', '0.9.0'), false);
  assert.equal(newer('1.0.0-rc.1', '0.9.0'), false);
  const off = await server();
  const on = await server({ OSIGN_UPDATE_CHECK: '1', OSIGN_UPDATE_URL: 'http://127.0.0.1:' + release.address().port + '/latest' });
  try {
    const a = off.as(await off.signIn('admin', 'admin-pw'));
    assert.equal((await a('info')).features.update, null);
    const b = on.as(await on.signIn('admin', 'admin-pw'));
    await b('info');
    await new Promise((r) => setTimeout(r, 200));
    assert.deepEqual((await b('info')).features.update, { version: '99.0.0', url: 'https://github.com/monsama/nobs-osign/releases/tag/v99.0.0', at: '2030-01-01T00:00:00Z' });
  } finally { off.srv.close(); on.srv.close(); }
});

test('Active Directory made simple: the domain detected, the group by its name (nested too), the rest from the domain', async () => {
  const { ldapConfig } = require('../server/ldap');
  const { directoryConfig } = require('../server/directory');
  // What follows from the domain alone.
  const c = ldapConfig({ OSIGN_LDAP_URL: 'ldaps://dc.argo.intern', OSIGN_AD_DOMAIN: 'Argo.Intern' });
  assert.equal(c.bindTemplate, '{login}@argo.intern');
  assert.equal(c.searchBase, 'DC=argo,DC=intern');
  assert.equal(directoryConfig({ OSIGN_LDAP_SERVICE_DN: 'osign' }, c).serviceDn, 'osign@argo.intern');
  assert.equal(directoryConfig({ OSIGN_LDAP_SERVICE_DN: 'CN=osign,OU=x,DC=argo,DC=intern' }, c).serviceDn, 'CN=osign,OU=x,DC=argo,DC=intern');
  assert.equal(directoryConfig({}, c).base, 'DC=argo,DC=intern');
  // Set by hand, they win; an OU still finds groups in the whole domain.
  const d = ldapConfig({ OSIGN_AD_DOMAIN: 'argo.intern', OSIGN_LDAP_BIND: 'ARGO\\{login}', OSIGN_LDAP_BASE: 'OU=Users,DC=argo,DC=intern' });
  assert.deepEqual([d.bindTemplate, d.searchBase, d.domainBase], ['ARGO\\{login}', 'OU=Users,DC=argo,DC=intern', 'DC=argo,DC=intern']);

  // The group by name, against the stand-in AD: Carla is in it through a nested group.
  const nested = fakeAd({ base: BASE, service: SERVICE, groups: { 'NOBS-OSign-Admins': [], 'Marketing': ['NOBS-OSign-Admins'] }, people: {
    carla: { pw: 'carla-pw', name: 'Carla Nested', mail: 'carla@example.ch', groups: ['Marketing'] },
    dani: { pw: 'dani-pw', name: 'Dani Outside', mail: 'dani@example.ch', groups: [] }
  } });
  const url = await nested.listen();
  const s = await server({ OSIGN_LDAP_URL: url, OSIGN_LDAP_GROUP: 'NOBS-OSign-Admins' });
  try {
    assert.ok(await s.signIn('carla', 'carla-pw'), 'nested membership counts');
    assert.equal(await s.signIn('dani', 'dani-pw'), '', 'not in the group');
    const admin = s.as(await s.signIn('admin', 'admin-pw'));
    // Detect: the domain from the domain controller's RootDSE.
    const r = await admin('connections/detect');
    assert.equal(r.ok, true, r.error);
    assert.deepEqual([r.domain, r.base, r.host], ['example.ch', BASE, 'dc01.example.ch']);
    // The test as one: the group found by name, the service account reading users.
    await admin('connections/save', { values: { ldapServiceDn: SERVICE.dn }, secrets: { ldapServicePassword: SERVICE.pw } });
    const t = await admin('connections/test', { what: 'ad' });
    assert.equal(t.ok, true, t.error);
    assert.match(t.message, /All set.*sign-in group is there/);
    await admin('connections/save', { values: { ldapGroup: 'No-Such-Group' } });
    assert.match((await admin('connections/test', { what: 'ad' })).error, /no sign-in group by that name/);
    assert.match((await admin('connections/save', { values: { adDomain: 'not a domain' } })).error, /DNS name/);
    const conn = await admin('connections/get');
    assert.equal(conn.derived.ldapBase, 'OU=Users,' + BASE);
  } finally { s.srv.close(); nested.close(); }
});

test('the server\'s certificate is valid DER whatever its random serial starts with', () => {
  const crypto = require('node:crypto');
  const { makeCert } = require('../server/cert');
  for (const first of [0x00, 0x01, 0x7f, 0x80, 0xff]) {
    const c = makeCert('NOBS OSign server', new Date(), Buffer.concat([Buffer.from([first]), crypto.randomBytes(15)]));
    const x = new crypto.X509Certificate(Buffer.from(c.certDer, 'base64'));
    assert.match(x.subject, /NOBS OSign server/);
    assert.equal(x.verify(x.publicKey), true);
  }
});
