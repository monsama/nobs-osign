'use strict';
// Who may call what: every route of the API, without a session, as a member of the sign-in group,
// as one who also publishes, and as the local account admin. Read from the code, so a route added
// without a place in ACCESS below fails here until it is given one.
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../server/app');
const { fakeAd } = require('./support/fake-ad');

// admin: the local account admin only; publish: whoever publishes (and admin); member: everyone signed in.
const ADMIN = ['addin/build', 'addin/upload', 'audit/list', 'backup/auto/download', 'backup/auto/get', 'backup/auto/run', 'backup/auto/save',
  'backup/check', 'backup/export', 'backup/restore', 'connections/cert', 'connections/detect', 'connections/get', 'connections/save',
  'connections/test', 'deploy/start', 'notify/get', 'notify/save', 'notify/test', 'sessions/end', 'sessions/list', 'setup/start',
  'setup/status', 'sp/check', 'sp/share'];
const PUBLISH = ['ad/export', 'photos/delete', 'photos/upload', 'schedule/cancel', 'schedule/set', 'publish'];
const MEMBER = ['ad/departments', 'ad/groups', 'ad/lookup', 'ad/read', 'ad/snapshot', 'ad/users', 'addin/check', 'approval/approve',
  'approval/cancel', 'approval/get', 'approval/request', 'assets/delete', 'assets/upload', 'deploy/status', 'draft/get', 'draft/save',
  'entra/lookup', 'entra/users', 'history/diff', 'history/list', 'history/restore', 'load-published', 'photos/fetch', 'photos/list',
  'published/get', 'published/revert', 'schedule/get', 'server/tls', 'settings/get', 'settings/save', 'info', 'ping'];
const DENIED = /Only the local|local account admin|publishing group|Only whoever|may not|not allowed|campaigns only/i;
// The campaign role (OSIGN_LDAP_CAMPAIGN_GROUP): these and nothing else.
const CAMPAIGN = ['ad/lookup', 'ad/users', 'approval/get', 'approval/request', 'assets/delete', 'assets/upload', 'draft/get', 'draft/save', 'entra/lookup', 'entra/users',
  'history/diff', 'history/list', 'info', 'photos/list', 'ping', 'publish', 'published/get', 'schedule/get', 'settings/get'];

function routesInCode() {
  const dir = path.join(__dirname, '..', 'server');
  const src = ['app.js', ...fs.readdirSync(dir).filter((f) => /^routes-.*\.js$/.test(f))].map((f) => fs.readFileSync(path.join(dir, f), 'utf8')).join('\n');
  const quoted = [...src.matchAll(/^\s+'([a-z][a-z-]*(?:\/[a-z-]+)*)'\s*:\s*(?:async\s*)?\(\s*(?:d|data)?\s*[,)]/gm)].map((m) => m[1]);
  const bare = [...src.matchAll(/^\s+(publish|info|ping)\s*:\s*(?:async\s*)?\(/gm)].map((m) => m[1]);
  return [...new Set(quoted.concat(bare))].sort();
}

test('every route of the API has its place, and each role gets exactly its own', async () => {
  const known = new Set([...ADMIN, ...PUBLISH, ...MEMBER]);
  const inCode = routesInCode();
  assert.deepEqual(inCode.filter((r) => !known.has(r)), [], 'routes without a place in ACCESS');
  assert.deepEqual([...known].filter((r) => !inCode.includes(r)).sort(), [], 'ACCESS names routes the code no longer has');

  const BASE = 'DC=example,DC=ch', SERVICE = { dn: 'CN=svc-osign,OU=Service,DC=example,DC=ch', pw: 'svc-pw' };
  const ad = fakeAd({ base: BASE, service: SERVICE, groups: { 'NOBS-OSign-Admins': [], 'NOBS-OSign-Pub': [], 'NOBS-OSign-Marketing': [] }, people: {
    mar: { pw: 'mar-pw', name: 'Mar', mail: 'mar@example.ch', groups: ['NOBS-OSign-Marketing'] },
    mem: { pw: 'mem-pw', name: 'Mem', mail: 'mem@example.ch', groups: ['NOBS-OSign-Admins'] },
    pub: { pw: 'pub-pw', name: 'Pub', mail: 'pub@example.ch', groups: ['NOBS-OSign-Admins', 'NOBS-OSign-Pub'] } } });
  const ldapUrl = await ad.listen();
  const env = { OSIGN_DATA: fs.mkdtempSync(path.join(os.tmpdir(), 'osign-authz-')), OSIGN_PUBLISH_DIR: fs.mkdtempSync(path.join(os.tmpdir(), 'osign-authz-pub-')),
    OSIGN_SECURE_COOKIE: '0', OSIGN_ADMIN_PASSWORD: 'admin-pw', OSIGN_APP_KEY: 'authz-key',
    OSIGN_LDAP_URL: ldapUrl, OSIGN_LDAP_ALLOW_PLAIN: '1', OSIGN_LDAP_BIND: 'cn={login}@example.ch,' + BASE, OSIGN_LDAP_BASE: 'OU=Users,' + BASE,
    OSIGN_LDAP_GROUP: 'CN=NOBS-OSign-Admins,OU=Groups,' + BASE, OSIGN_LDAP_PUBLISH_GROUP: 'CN=NOBS-OSign-Pub,OU=Groups,' + BASE, OSIGN_LDAP_CAMPAIGN_GROUP: 'NOBS-OSign-Marketing',
    OSIGN_LDAP_SERVICE_DN: SERVICE.dn, OSIGN_LDAP_SERVICE_PASSWORD: SERVICE.pw };
  const srv = http.createServer(createApp({ log: { info() {}, warn() {}, error() {} }, env }));
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  const at = 'http://127.0.0.1:' + srv.address().port;
  try {
    const signIn = async (l, p) => (await fetch(at + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: l, password: p }) })).headers.get('set-cookie').split(';')[0];
    const who = { mem: await signIn('mem', 'mem-pw'), pub: await signIn('pub', 'pub-pw'), admin: await signIn('admin', 'admin-pw') };
    const call = async (route, cookie) => {
      const r = await fetch(at + '/api/' + route, { method: 'POST', headers: Object.assign({ 'Content-Type': 'application/json' }, cookie ? { Cookie: cookie } : {}), body: '{}' });
      const j = await r.json().catch(() => ({}));
      return r.status === 401 ? 'none' : !j.ok && DENIED.test(j.error || '') ? 'denied' : 'through';
    };
    const wrong = [];
    for (const route of inCode) {
      if ((await call(route, '')) !== 'none') wrong.push(route + ': answered without a session');
      const may = ADMIN.includes(route) ? ['admin'] : PUBLISH.includes(route) ? ['pub', 'admin'] : ['mem', 'pub', 'admin'];
      for (const [role, cookie] of Object.entries(who)) {
        const got = await call(route, cookie);
        if (may.includes(role) && got !== 'through') wrong.push(route + ': refused to ' + role);
        if (!may.includes(role) && got !== 'denied') wrong.push(route + ': let ' + role + ' through');
      }
    }
    // The campaign role: its routes, and every other one refused.
    const mar = await signIn('mar', 'mar-pw');
    for (const route of inCode) {
      const got = await call(route, mar);
      if (CAMPAIGN.includes(route) && got !== 'through') wrong.push(route + ': refused to the campaign role');
      if (!CAMPAIGN.includes(route) && got !== 'denied') wrong.push(route + ': let the campaign role through');
    }
    assert.deepEqual(CAMPAIGN.filter((r) => !inCode.includes(r)), [], 'the campaign role names routes the code no longer has');
    assert.deepEqual(wrong, []);
  } finally { srv.close(); ad.close(); }
});
