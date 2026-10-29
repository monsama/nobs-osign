// node --test tests/server-graph.test.js
//
// The server version publishing into SharePoint through Microsoft Graph (server/graph.js), against
// a stand-in Entra ID and Graph (tests/support/fake-sharepoint.js) that, like an app with
// Sites.Selected, reaches only the one site it was granted.
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../server/app');
const { fakeSharePoint } = require('./support/fake-sharepoint');

const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');
const TENANT = '11111111-2222-3333-4444-555555555555';
const CLIENT = '66666666-7777-8888-9999-000000000000';
const SITE = 'contoso.sharepoint.com:/sites/IT:';

let sp, base, srv, cookie, dataDir;
const logs = [];
const log = { info: (m) => logs.push(m), warn: (m) => logs.push(m), error: (m) => logs.push(m) };

async function start(env) {
  const s = http.createServer(createApp({ env, log }));
  await new Promise((r) => s.listen(0, '127.0.0.1', r));
  return s;
}
function envFor(dir, extra) {
  fs.writeFileSync(path.join(dir, 'graph_secret'), 'the-secret\n');
  return Object.assign({
    OSIGN_DATA: dir, OSIGN_ADMIN_PASSWORD: 'pw', OSIGN_SECURE_COOKIE: '0',
    OSIGN_GRAPH_TENANT_ID: TENANT, OSIGN_GRAPH_CLIENT_ID: CLIENT, OSIGN_GRAPH_CLIENT_SECRET_FILE: path.join(dir, 'graph_secret'),
    OSIGN_GRAPH_URL: sp.url + '/v1.0', OSIGN_LOGIN_URL: sp.url
  }, extra || {});
}
async function signIn(at) {
  const r = await fetch(at + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'admin', password: 'pw' }) });
  assert.equal(r.status, 200);
  return r.headers.get('set-cookie').split(';')[0];
}
const call = (at, c, name, data) => fetch(at + '/api/' + name, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: c }, body: JSON.stringify(data || {}) }).then((r) => r.json());
const api = (name, data) => call(base, cookie, name, data);
const cfg = (sigs) => JSON.stringify({ schemaVersion: 1, options: {}, fields: [], signatures: sigs || [] });

test.before(async () => {
  sp = fakeSharePoint({ tenantId: TENANT, clientId: CLIENT, secret: 'the-secret', granted: [SITE] });
  sp.url = await sp.listen();
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-graph-'));
  srv = await start(envFor(dataDir));
  base = 'http://127.0.0.1:' + srv.address().port;
  cookie = await signIn(base);
});
test.after(() => { srv.close(); sp.close(); });

test('info says publishing goes to SharePoint, with the server\'s app registration', async () => {
  const info = await api('info');
  assert.equal(info.features.publishTarget, 'sharepoint');
  assert.equal(info.features.graphClientId, CLIENT);
  assert.equal(info.features.graphProblem, '');
});

test('sp/check: a site the app may not reach says what to grant; the granted one is found', async () => {
  const denied = await api('sp/check', { spSite: 'contoso.sharepoint.com:/sites/Sales:', spFolder: 'OSign' });
  assert.equal(denied.ok, false);
  assert.match(denied.error, /may not reach this site yet/);
  assert.match(denied.error, /Sites\.Selected/);
  const missing = await api('sp/check', { spSite: SITE, spFolder: 'OSign' });
  assert.equal(missing.ok, true, JSON.stringify(missing));
  assert.equal(missing.found, false);
  assert.match(missing.where, /sites\/IT/);
  assert.equal((await api('sp/check', { spSite: 'nope', spFolder: 'OSign' })).ok, false);
  assert.equal((await api('sp/check', { spSite: SITE, spFolder: '../x' })).ok, false);
});

test('publish makes the folder and writes into SharePoint: images, then signatures.json', async () => {
  assert.equal((await api('settings/save', { settings: { spSite: SITE, spFolder: 'OSign' } })).ok, true);
  assert.equal((await api('assets/upload', { name: 'logo.png', data: PNG.toString('base64') })).name, 'logo.png');
  const text = cfg([{ id: 'std', name: 'Standard', html: '<p><img src="assets/logo.png"></p>' }]);
  sp.requests.length = 0;
  const p = await api('publish', { configJson: text });
  assert.equal(p.ok, true, JSON.stringify(p));
  assert.equal(p.changed, true);
  assert.equal(p.folder, 'SharePoint contoso.sharepoint.com/sites/IT › OSign');
  const files = sp.sites[SITE].files;
  assert.ok(files.get('OSign/assets/logo.png').equals(PNG));
  assert.equal(files.get('OSign/signatures.json').toString(), text);
  const puts = sp.requests.filter((r) => r.startsWith('PUT '));
  assert.deepEqual(puts.map((r) => r.split('root:/')[1]), ['OSign/assets/logo.png:/content', 'OSign/signatures.json:/content']);
  const s = await api('settings/get');
  assert.equal(s.folderFound, true);
  assert.equal(s.photosFound, false);
  assert.equal((await api('draft/get')).dirty, false);
});

test('publishing again writes nothing that did not change; an image no longer used goes after the configuration', async () => {
  sp.requests.length = 0;
  const text = cfg([{ id: 'std', name: 'Standard', html: '<p><img src="assets/logo.png"></p>' }]);
  assert.equal((await api('publish', { configJson: text })).changed, false);
  assert.ok(!sp.requests.some((r) => /^(PUT|DELETE) /.test(r)), sp.requests.join('\n'));
  // A new logo, the old one deleted from the draft.
  const other = Buffer.concat([PNG, Buffer.from([1])]);
  assert.equal((await api('assets/upload', { name: 'new.png', data: other.toString('base64') })).name, 'new.png');
  assert.equal((await api('assets/delete', { name: 'logo.png' })).ok, true);
  const next = cfg([{ id: 'std', name: 'Standard', html: '<p><img src="assets/new.png"></p>' }]);
  sp.requests.length = 0;
  assert.equal((await api('publish', { configJson: next })).changed, true);
  const writes = sp.requests.filter((r) => /^(PUT|DELETE) /.test(r)).map((r) => r.split(' ')[0] + ' ' + r.split('root:/')[1]);
  assert.deepEqual(writes, ['PUT OSign/assets/new.png:/content', 'PUT OSign/signatures.json:/content', 'DELETE OSign/assets/logo.png:']);
  assert.deepEqual([...sp.sites[SITE].files.keys()].sort(), ['OSign/assets/new.png', 'OSign/signatures.json']);
});

test('photos for the preview come from SharePoint; load published reads it back', async () => {
  sp.sites[SITE].folders.add('OSign/photos');
  sp.sites[SITE].files.set('OSign/photos/mustera.jpg', PNG);
  const r = await fetch(base + '/photos/mustera.jpg', { headers: { Cookie: cookie } });
  assert.equal(r.status, 200);
  assert.equal(r.headers.get('content-type'), 'image/jpeg');
  assert.ok(Buffer.from(await r.arrayBuffer()).equals(PNG));
  assert.equal((await fetch(base + '/photos/nobody.jpg', { headers: { Cookie: cookie } })).status, 404);
  assert.equal((await fetch(base + '/photos/..%2fsignatures.json', { headers: { Cookie: cookie } })).status, 404);
  await api('draft/save', { configJson: cfg([{ id: 'x', name: 'Not published', html: '' }]) });
  const l = await api('load-published');
  assert.equal(l.ok, true, JSON.stringify(l));
  assert.match(l.configJson, /assets\/new\.png/);
  assert.deepEqual(l.assets.map((a) => a.name), ['new.png']);
});

test('a busy SharePoint (429) is asked once more', async () => {
  sp.busyOnce();
  const r = await api('sp/check', { spSite: SITE, spFolder: 'OSign' });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.found, true);
});

test('a wrong client secret is said as that, not as a missing folder', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-graph-'));
  const env = envFor(dir);
  fs.writeFileSync(path.join(dir, 'graph_secret'), 'wrong');
  const other = await start(env);
  const at = 'http://127.0.0.1:' + other.address().port;
  const c = await signIn(at);
  const r = await call(at, c, 'publish', { configJson: cfg([]) });
  assert.equal(r.ok, false);
  assert.match(r.error, /could not sign in/);
  assert.match(r.error, /AADSTS7000215/);
  assert.ok(!/Trace ID/.test(r.error));
  const s = await call(at, c, 'settings/get');
  assert.equal(s.folderFound, false);
  assert.match(s.folderError, /client secret/);
  other.close();
});

test('Graph set up only half: said in info, and publishing stays in the server\'s folder', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-graph-'));
  const other = await start({ OSIGN_DATA: dir, OSIGN_ADMIN_PASSWORD: 'pw', OSIGN_SECURE_COOKIE: '0', OSIGN_GRAPH_TENANT_ID: TENANT });
  const at = 'http://127.0.0.1:' + other.address().port;
  const c = await signIn(at);
  const info = await call(at, c, 'info');
  assert.equal(info.features.publishTarget, 'folder');
  assert.match(info.features.graphProblem, /OSIGN_GRAPH_CLIENT_ID/);
  assert.match((await call(at, c, 'sp/check', { spSite: SITE, spFolder: 'OSign' })).error, /OSIGN_GRAPH_CLIENT_ID/);
  other.close();
});

test('sp/share: a sharing link becomes the folder\'s own address; what it cannot open, it says', async () => {
  const folderLink = 'https://contoso-my.sharepoint.com/:f:/g/personal/anna_contoso_com/IgDpY5m2CxbbRYgU2eNTaxZP?e=GOPGuO';
  const fileLink = 'https://contoso.sharepoint.com/:w:/s/IT/EaBcDeF?e=x';
  sp.shares[folderLink] = { name: 'OSign', webUrl: 'https://contoso-my.sharepoint.com/personal/anna_contoso_com/Documents/OSign', folder: true };
  sp.shares[fileLink] = { name: 'Notes.docx', webUrl: 'https://contoso.sharepoint.com/sites/IT/Shared%20Documents/Notes.docx' };
  const r = await api('sp/share', { url: folderLink });
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.webUrl, 'https://contoso-my.sharepoint.com/personal/anna_contoso_com/Documents/OSign');
  assert.match((await api('sp/share', { url: fileLink })).error, /a file \(Notes\.docx\), not to a folder/);
  assert.match((await api('sp/share', { url: 'https://contoso.sharepoint.com/:f:/s/IT/Gone' })).error, /removed or has expired/);
  assert.match((await api('sp/share', { url: 'https://example.com/:f:/g/x' })).error, /not a sharing link/);
  sp.shares.denied = true;
  try { assert.match((await api('sp/share', { url: folderLink })).error, /may not open sharing links/); } finally { delete sp.shares.denied; }
});
