// node --test tests/server-ad.test.js
//
// The server version and Active Directory, read with a service account (server/directory.js,
// server/users.js): the preview's users and groups, and the users' values for signatures that
// take them from AD, read, checked, and written to users/<mail>.json exactly once. Against a
// stand-in AD (tests/support/fake-ad.js) and a stand-in SharePoint (tests/support/fake-sharepoint.js).
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../server/app');
const { fakeAd } = require('./support/fake-ad');
const { fakeSharePoint } = require('./support/fake-sharepoint');
const { nextAt } = require('../server/users');

const BASE = 'DC=example,DC=ch';
const SERVICE = { dn: 'CN=svc-osign,OU=Service,DC=example,DC=ch', pw: 'svc-pw' };
const TENANT = '11111111-2222-3333-4444-555555555555';
const CLIENT = '66666666-7777-8888-9999-000000000000';
const SITE = 'contoso.sharepoint.com:/sites/IT:';
// anna is in sg-field through a nested group; carl is disabled; dora has no mail.
const PEOPLE = {
  anna: { pw: 'anna-pw', name: 'Anna Muster', mail: 'Anna.Muster@example.ch', attrs: { description: 'Dipl. Ing.', homePhone: '+41 41 000 00 01' }, groups: ['sg-field-north', 'OSign-Admins'] },
  beat: { pw: 'beat-pw', name: 'Beat Beispiel', mail: 'beat@example.ch', attrs: { description: '', homePhone: '+41 41 000 00 02' }, groups: [] },
  carl: { pw: 'carl-pw', name: 'Carl Gone', mail: 'carl@example.ch', attrs: {}, disabled: true, groups: ['sg-field'] },
  dora: { pw: 'dora-pw', name: 'Dora Nomail', attrs: {}, groups: [] }
};
const GROUPS = { 'sg-field': [], 'sg-field-north': ['sg-field'], 'sg-sales': [], 'OSign-Admins': [] };
const FIELDS = [{ key: 'description', ad: 'description' }, { key: 'homePhone', ad: 'homePhone' }];

let ad, sp, srv, base, cookie, dataDir, handle;
const logs = [];
const log = { info: (m) => logs.push(m), warn: (m) => logs.push(m), error: (m) => logs.push(m) };

test.before(async () => {
  ad = fakeAd({ base: BASE, people: PEOPLE, groups: GROUPS, service: SERVICE });
  const ldapUrl = await ad.listen();
  sp = fakeSharePoint({ tenantId: TENANT, clientId: CLIENT, secret: 's', granted: [SITE] });
  const spUrl = await sp.listen();
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-ad-'));
  fs.writeFileSync(path.join(dataDir, 'svc'), SERVICE.pw + '\n');
  handle = createApp({ log, env: {
    OSIGN_DATA: dataDir, OSIGN_SECURE_COOKIE: '0', OSIGN_ADMIN_PASSWORD: 'admin-pw',
    OSIGN_LDAP_URL: ldapUrl, OSIGN_LDAP_ALLOW_PLAIN: '1', OSIGN_LDAP_BIND: 'cn={login}@example.ch,' + BASE, OSIGN_LDAP_BASE: 'OU=Users,' + BASE,
    OSIGN_LDAP_GROUP: 'CN=OSign-Admins,OU=Groups,' + BASE, OSIGN_LDAP_SERVICE_DN: SERVICE.dn, OSIGN_LDAP_SERVICE_PASSWORD_FILE: path.join(dataDir, 'svc'),
    OSIGN_GRAPH_TENANT_ID: TENANT, OSIGN_GRAPH_CLIENT_ID: CLIENT, OSIGN_GRAPH_CLIENT_SECRET: 's', OSIGN_GRAPH_URL: spUrl + '/v1.0', OSIGN_LOGIN_URL: spUrl
  } });
  srv = http.createServer(handle);
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  base = 'http://127.0.0.1:' + srv.address().port;
  const r = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'anna', password: 'anna-pw' }) });
  assert.equal(r.status, 200);
  cookie = r.headers.get('set-cookie').split(';')[0];
  // Where publishing goes is the local admin's to set; a member of the sign-in group may not.
  const denied = await api('settings/save', { settings: { spSite: SITE, spFolder: 'OSign' } });
  assert.match(denied.error, /Only the local account admin/);
  const admin = (await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'admin', password: 'admin-pw' }) })).headers.get('set-cookie').split(';')[0];
  const saved = await fetch(base + '/api/settings/save', { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: admin }, body: JSON.stringify({ settings: { spSite: SITE, spFolder: 'OSign' } }) }).then((r) => r.json());
  assert.equal(saved.ok, true, saved.error);
  // What is set already may be sent again unchanged (the Settings dialog sends everything).
  assert.equal((await api('settings/save', { settings: { spSite: SITE, spFolder: 'OSign', setupDone: '1' } })).ok, true);
});
test.after(() => { srv.close(); ad.close(); sp.close(); });

function api(name, data) {
  return fetch(base + '/api/' + name, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: cookie }, body: JSON.stringify(data || {}) }).then((r) => r.json());
}
const files = () => sp.sites[SITE].files;
const userFile = (n) => { const b = files().get('OSign/users/' + n); return b ? JSON.parse(b.toString()) : null; };

test('info: AD can be asked, with the service account', async () => {
  const info = await api('info');
  assert.equal(info.features.adLookup, true);
  assert.equal(info.features.adProblem, '');
  assert.equal(info.features.adExport, null);
});

test('lookup: one user\'s values and groups, nested ones included; nothing typed is the one signed in', async () => {
  const r = await api('ad/lookup', { query: 'Anna.Muster@example.ch', fields: [...FIELDS, { key: 'nope', ad: 'not valid!' }] });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.displayName, 'Anna Muster');
  assert.deepEqual(r.values, { description: 'Dipl. Ing.', homePhone: '+41 41 000 00 01', nope: '' });
  assert.deepEqual(r.groups, ['OSign-Admins', 'sg-field', 'sg-field-north']);
  const me = await api('ad/lookup', { query: '', fields: FIELDS });
  assert.equal(me.displayName, 'Anna Muster');
  assert.match((await api('ad/lookup', { query: 'nobody', fields: FIELDS })).error, /No user nobody/);
  // The LAPS password and its kind are never asked for, whatever a field names.
  ad.set('beat', 'ms-Mcs-AdmPwd', 'secret!');
  const laps = await api('ad/lookup', { query: 'beat', fields: [{ key: 'pw', ad: 'ms-Mcs-AdmPwd' }] });
  assert.deepEqual(laps.values, { pw: '' });
  assert.ok(!ad.searches.some((f) => /admpwd/i.test(f)));
  ad.set('beat', 'ms-Mcs-AdmPwd', null);
  // A filter in what was typed is text: it matches nobody, not everybody.
  assert.match((await api('ad/lookup', { query: '*', fields: FIELDS })).error, /No user \*/);
});

test('groups to pick from: by the start of their name, from two letters', async () => {
  assert.deepEqual((await api('ad/groups', { query: 'sg-f' })).groups, ['sg-field', 'sg-field-north']);
  assert.deepEqual((await api('ad/groups', { query: 's' })).groups, []);
  assert.deepEqual((await api('ad/groups', { query: 'sg-*' })).groups, []);
});

test('read: every active user with a mail address, compared with what is there; nothing written', async () => {
  const r = await api('ad/read', { fields: FIELDS, groups: ['SG-Field'] });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.users.map((u) => u.mail).sort(), ['Anna.Muster@example.ch', 'beat@example.ch']);
  assert.deepEqual(r.counts, { new: 2, changed: 0, unchanged: 0, removed: 0 });
  const anna = r.users.find((u) => u.mail === 'Anna.Muster@example.ch');
  assert.deepEqual(anna.groups, ['sg-field']);
  assert.deepEqual(anna.values, { description: 'Dipl. Ing.', homePhone: '+41 41 000 00 01' });
  assert.ok(![...files().keys()].some((k) => k.startsWith('OSign/users/')));
  const snap = await api('ad/snapshot', { fields: FIELDS, groups: ['sg-field'] });
  assert.equal(snap.matches, true);
  assert.equal((await api('ad/snapshot', { fields: FIELDS, groups: [] })).matches, false);
});

test('export: exactly what was read, once, as users/<mail>.json', async () => {
  assert.match((await api('ad/export', { fields: FIELDS, groups: [] })).error, /fields or groups changed/);
  const x = await api('ad/export', { fields: FIELDS, groups: ['sg-field'] });
  assert.deepEqual(x, { ok: true, users: 2, written: 2, unchanged: 0, removed: 0 });
  assert.deepEqual(userFile('anna.muster@example.ch.json'), { values: { description: 'Dipl. Ing.', homePhone: '+41 41 000 00 01' }, groups: ['sg-field'] });
  assert.deepEqual(userFile('beat@example.ch.json'), { values: { description: '', homePhone: '+41 41 000 00 02' }, groups: [] });
  assert.match((await api('ad/export', { fields: FIELDS, groups: ['sg-field'] })).error, /written already/);
});

test('read again: what changed in AD, with the old value; who left; only changed files are fetched', async () => {
  ad.set('beat', 'homePhone', '+41 41 999 99 99');
  ad.remove('anna');
  sp.requests.length = 0;
  const r = await api('ad/read', { fields: FIELDS, groups: ['sg-field'] });
  assert.deepEqual(r.counts, { new: 0, changed: 1, unchanged: 0, removed: 1 });
  assert.deepEqual(r.users[0].old, { description: '', homePhone: '+41 41 000 00 02' });
  assert.deepEqual(r.removed, ['anna.muster@example.ch']);
  // Written by this server: known from its mirror, not read from SharePoint again.
  assert.ok(!sp.requests.some((q) => /users\/.*:\/content/.test(q) && q.startsWith('GET')), sp.requests.join('\n'));
  const x = await api('ad/export', { fields: FIELDS, groups: ['sg-field'] });
  assert.deepEqual(x, { ok: true, users: 1, written: 1, unchanged: 0, removed: 1 });
  assert.equal(userFile('anna.muster@example.ch.json'), null);
  assert.equal(userFile('beat@example.ch.json').values.homePhone, '+41 41 999 99 99');
});

test('a file changed in SharePoint by someone else is read again, not taken from the mirror', async () => {
  files().set('OSign/users/beat@example.ch.json', Buffer.from(JSON.stringify({ values: { description: 'by hand', homePhone: '' }, groups: [] })));
  const r = await api('ad/read', { fields: FIELDS, groups: ['sg-field'] });
  assert.deepEqual(r.counts, { new: 0, changed: 1, unchanged: 0, removed: 0 });
  assert.equal(r.users[0].old.description, 'by hand');
});

test('a read that finds nobody writes nothing', async () => {
  const keep = ad.entries.slice();
  ad.hideAll();
  await api('ad/read', { fields: FIELDS, groups: ['sg-field'] });
  const x = await api('ad/export', { fields: FIELDS, groups: ['sg-field'] });
  assert.equal(x.ok, false);
  assert.match(x.error, /returned no users/);
  assert.ok(userFile('beat@example.ch.json'));
  ad.entries.length = 0; ad.entries.push(...keep);
});

test('the scheduled export: what the published signatures ask for, read and written in one go', async () => {
  const cfg = { schemaVersion: 1, options: { valueSource: 'mixed' }, fields: [{ key: 'givenName', ad: 'givenName', graph: 'givenName' }, { key: 'homePhone', ad: 'homePhone', graph: '' }],
    signatures: [{ id: 's', name: 'S', html: '', audience: { groups: ['sg-field'] } }] };
  assert.equal((await api('publish', { configJson: JSON.stringify(cfg) })).ok, true);
  const r = await handle.runExport();
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.result, { users: 1, written: 1, unchanged: 0, removed: 0 });
  // Only the fields without a Graph source come from AD.
  assert.deepEqual(userFile('beat@example.ch.json'), { values: { homePhone: '+41 41 999 99 99' }, groups: [] });
  // The editor's read is older than that now: it does not go out.
  assert.equal((await api('ad/snapshot', { fields: [{ key: 'homePhone', ad: 'homePhone' }], groups: ['sg-field'] })).snapshot.written, true);
  // Values from Entra ID: nothing to export.
  cfg.options.valueSource = 'graph';
  await api('publish', { configJson: JSON.stringify(cfg) });
  assert.match((await handle.runExport()).result.skipped, /Entra ID/);
});

test('the time of the scheduled export', () => {
  const at = new Date(2026, 9, 29, 5, 0).getTime();
  assert.equal(new Date(nextAt('06:00', at)).getHours(), 6);
  assert.equal(new Date(nextAt('06:00', at)).getDate(), 29);
  assert.equal(new Date(nextAt('04:30', at)).getDate(), 30);
  assert.equal(nextAt('6 Uhr', at), null);
  assert.equal(nextAt('24:00', at), null);
});

test('without a service account AD says what is missing', async () => {
  const other = http.createServer(createApp({ log, env: { OSIGN_DATA: fs.mkdtempSync(path.join(os.tmpdir(), 'osign-ad-')), OSIGN_ADMIN_PASSWORD: 'pw', OSIGN_SECURE_COOKIE: '0', OSIGN_LDAP_URL: 'ldaps://dc.example.ch', OSIGN_LDAP_BIND: '{login}@example.ch', OSIGN_LDAP_BASE: BASE, OSIGN_LDAP_GROUP: 'CN=X,' + BASE } }));
  await new Promise((r) => other.listen(0, '127.0.0.1', r));
  const at = 'http://127.0.0.1:' + other.address().port;
  const c = (await fetch(at + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'admin', password: 'pw' }) })).headers.get('set-cookie').split(';')[0];
  const call = (n, d) => fetch(at + '/api/' + n, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: c }, body: JSON.stringify(d || {}) }).then((r) => r.json());
  const info = await call('info');
  assert.equal(info.features.adLookup, false);
  assert.match(info.features.adProblem, /no service account/);
  assert.match((await call('ad/lookup', { query: 'x' })).error, /no service account/);
  assert.match((await call('ad/read', { fields: FIELDS })).error, /no service account/);
  other.close();
});
