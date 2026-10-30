// node --test tests/server-addin.test.js
//
// The server version and the add-in's files (server/addin.js): made from the template exactly as
// addin/fill.mjs makes them, offered as a ZIP, checked online, and uploaded to Cloudflare Pages
// the way wrangler does (against a stand-in Cloudflare API).
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const zlib = require('node:zlib');
const { blake3 } = require('@noble/hashes/blake3');
const { bytesToHex } = require('@noble/hashes/utils');
const { createApp } = require('../server/app');

const ROOT = path.join(__dirname, '..');
const ACCOUNT = '0123456789abcdef0123456789abcdef';
const SETTINGS = { addinUrl: 'https://company-osign.pages.dev/', clientId: '00000000-0000-0000-0000-000000000001', addinHost: 'cloudflare', spSite: 'contoso.sharepoint.com:/sites/IT:', spFolder: 'OSign' };

// The Cloudflare API, as far as Direct Upload uses it.
function fakeCloudflare() {
  const projects = new Set(), assets = new Map(), requests = [], deployments = [];
  const server = http.createServer(async (req, res) => {
    const body = await new Promise((r) => { const p = []; req.on('data', (d) => p.push(d)); req.on('end', () => r(Buffer.concat(p))); });
    const u = new URL(req.url, 'http://x'), auth = String(req.headers.authorization || '');
    requests.push(req.method + ' ' + u.pathname);
    const json = (status, result, errors) => { res.writeHead(status, { 'Content-Type': 'application/json' }); res.end(JSON.stringify({ success: status < 400, result, errors: errors || [] })); };
    const api = auth === 'Bearer cf-token', jwt = auth === 'Bearer the-jwt';
    let m;
    if ((m = /^\/accounts\/([^/]+)\/pages\/projects\/([^/]+)$/.exec(u.pathname)) && req.method === 'GET') {
      if (!api) return json(403, null, [{ code: 10000, message: 'Authentication error' }]);
      return projects.has(m[2]) ? json(200, { name: m[2] }) : json(404, null, [{ code: 8000007, message: 'Project not found' }]);
    }
    if (/^\/accounts\/[^/]+\/pages\/projects$/.test(u.pathname) && req.method === 'POST' && api) { projects.add(JSON.parse(body).name); return json(200, { name: JSON.parse(body).name }); }
    if (/\/upload-token$/.test(u.pathname) && api) return json(200, { jwt: 'the-jwt' });
    if (u.pathname === '/pages/assets/check-missing' && jwt) return json(200, JSON.parse(body).hashes.filter((h) => !assets.has(h)));
    if (u.pathname === '/pages/assets/upload' && jwt) { for (const f of JSON.parse(body)) assets.set(f.key, { data: Buffer.from(f.value, 'base64'), type: f.metadata.contentType }); return json(200, {}); }
    if (u.pathname === '/pages/assets/upsert-hashes' && jwt) return json(200, {});
    if (/\/deployments$/.test(u.pathname) && req.method === 'POST' && api) {
      const form = await new Request('http://x', { method: 'POST', headers: { 'content-type': req.headers['content-type'] }, body }).formData();
      const manifest = JSON.parse(form.get('manifest'));
      const missing = Object.values(manifest).filter((h) => !assets.has(h));
      if (missing.length) return json(400, null, [{ code: 8000000, message: 'assets missing' }]);
      const headers = form.get('_headers') ? Buffer.from(await form.get('_headers').arrayBuffer()).toString() : null;
      deployments.push({ manifest, headers, branch: form.get('branch') });
      return json(200, { id: 'dep-' + deployments.length, url: 'https://abc' + deployments.length + '.company-osign.pages.dev' });
    }
    return json(400, null, [{ code: 1, message: 'not understood: ' + req.method + ' ' + u.pathname }]);
  });
  return { projects, assets, requests, deployments, listen: () => new Promise((r) => server.listen(0, '127.0.0.1', () => r('http://127.0.0.1:' + server.address().port))), close: () => server.close() };
}

let cf, srv, base, cookie, dataDir;
test.before(async () => {
  cf = fakeCloudflare();
  const cfUrl = await cf.listen();
  cf.url = cfUrl;
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-addin-'));
  // The add-in's address, served from the files made here, as Cloudflare would.
  const online = (url, init) => {
    const u = new URL(url);
    if (u.hostname !== 'company-osign.pages.dev') return fetch(url, init);
    const f = path.join(dataDir, 'addin', u.pathname.slice(1));
    if (!fs.existsSync(f)) return Promise.resolve(new Response('not found', { status: 404 }));
    return Promise.resolve(new Response(fs.readFileSync(f), { status: 200, headers: { 'cache-control': 'no-cache' } }));
  };
  srv = http.createServer(createApp({ fetch: online, log: { info() {}, warn() {}, error() {} }, env: {
    OSIGN_DATA: dataDir, OSIGN_ADMIN_PASSWORD: 'pw', OSIGN_SECURE_COOKIE: '0',
    OSIGN_CF_ACCOUNT_ID: ACCOUNT, OSIGN_CF_PROJECT: 'company-osign', OSIGN_CF_API_TOKEN: 'cf-token', OSIGN_CF_API_URL: cfUrl
  } }));
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  base = 'http://127.0.0.1:' + srv.address().port;
  const r = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'admin', password: 'pw' }) });
  cookie = r.headers.get('set-cookie').split(';')[0];
});
test.after(() => { srv.close(); cf.close(); });
const api = (name, data) => fetch(base + '/api/' + name, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: cookie }, body: JSON.stringify(data || {}) }).then((r) => r.json());

test('info: the add-in\'s files can be made and uploaded', async () => {
  const info = await api('info');
  assert.equal(info.features.addinFiles, true);
  assert.equal(info.features.pagesUpload, true);
  assert.equal(info.features.pagesProject, 'company-osign');
});

test('build: the same files addin/fill.mjs makes, for the saved settings', async () => {
  assert.match((await api('addin/build')).error, /address/);
  assert.equal((await api('settings/save', { settings: SETTINGS })).ok, true);
  const r = await api('addin/build');
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.manifestUrl, 'https://company-osign.pages.dev/manifest.xml');
  assert.equal(r.redirectUri, 'brk-multihub://company-osign.pages.dev');
  const fill = await import(path.join(ROOT, 'addin', 'fill.mjs'));
  const ref = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-ref-'));
  fill.fillTemplate(path.join(ROOT, 'dist', 'template'), ref, { base: SETTINGS.addinUrl, clientId: SETTINGS.clientId, site: SETTINGS.spSite, folder: SETTINGS.spFolder, host: 'cloudflare' });
  const made = path.join(dataDir, 'addin');
  const list = (d) => require('../server/addin').walk(d);
  assert.deepEqual(list(made), list(ref));
  for (const f of list(ref)) assert.ok(fs.readFileSync(path.join(made, f)).equals(fs.readFileSync(path.join(ref, f))), f);
  // Outlook on iOS and Android only when asked for: then Mailbox 1.5 and the mobile launch event.
  let xml = fs.readFileSync(path.join(made, 'manifest.xml'), 'utf8');
  assert.match(xml, /DefaultMinVersion="1\.10"/);
  assert.doesNotMatch(xml, /MobileFormFactor/);
  assert.equal((await api('settings/save', { settings: { addinMobile: 'on' } })).ok, true);
  assert.equal((await api('addin/build')).ok, true);
  xml = fs.readFileSync(path.join(made, 'manifest.xml'), 'utf8');
  assert.match(xml, /DefaultMinVersion="1\.5"/);
  assert.match(xml, /<MobileFormFactor>[\s\S]*<LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"\/>[\s\S]*<\/MobileFormFactor>/);
  assert.equal((await api('settings/save', { settings: { addinMobile: '' } })).ok, true);
  assert.equal((await api('addin/build')).ok, true);
  assert.ok(list(made).includes('_headers'));
  assert.equal(r.files, list(ref).length);
});

test('the files as a ZIP, for an upload by hand', async () => {
  const z = Buffer.from(await (await fetch(base + '/addin.zip', { headers: { Cookie: cookie } })).arrayBuffer());
  // Read back entry by entry: name, stored data, CRC.
  const got = {};
  for (let at = 0; z.readUInt32LE(at) === 0x04034b50;) {
    const size = z.readUInt32LE(at + 18), nameLen = z.readUInt16LE(at + 26), name = z.subarray(at + 30, at + 30 + nameLen).toString();
    const data = z.subarray(at + 30 + nameLen, at + 30 + nameLen + size);
    assert.equal(zlib.crc32(data), z.readUInt32LE(at + 14), name);
    got[name] = data; at += 30 + nameLen + size;
  }
  assert.ok(got['manifest.xml'].equals(fs.readFileSync(path.join(dataDir, 'addin', 'manifest.xml'))));
  assert.ok(got['assets/icon-16.png']);
  assert.equal(z.readUInt32LE(z.length - 22), 0x06054b50);
  assert.equal((await fetch(base + '/addin.zip')).status, 401);
});

test('upload to Cloudflare Pages: the project made, only missing files sent, _headers with the deployment', async () => {
  const r = await api('addin/upload');
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.url, 'https://company-osign.pages.dev', 'the address people use, not the one deployment\'s');
  assert.equal(r.deployment, 'https://abc1.company-osign.pages.dev');
  assert.ok(cf.projects.has('company-osign'));
  const d = cf.deployments[0];
  assert.match(d.headers, /Cache-Control: no-cache/);
  assert.equal(d.branch, 'main');
  assert.ok(!('/_headers' in d.manifest));
  // Each file under its BLAKE3 name, as wrangler names them.
  const m = fs.readFileSync(path.join(dataDir, 'addin', 'manifest.xml'));
  const hash = bytesToHex(blake3(Buffer.from(m.toString('base64') + 'xml'))).slice(0, 32);
  assert.equal(d.manifest['/manifest.xml'], hash);
  assert.ok(cf.assets.get(hash).data.equals(m));
  assert.equal(cf.assets.get(hash).type, 'application/xml');
  // Again: nothing new to send.
  const again = await api('addin/upload');
  assert.equal(again.uploaded, 0);
  assert.equal(cf.deployments.length, 2);
});

test('check: the address serves exactly the files made here', async () => {
  const r = await api('addin/check');
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.equal(r.allOk, true, JSON.stringify(r.checks));
  assert.deepEqual(r.checks.map((c) => c.what), ['manifest.xml', 'Cache-Control', 'launchevent.js', 'taskpane.js', 'taskpane.html']);
});

test('a wrong token is said as that, with the permission it needs', async () => {
  const { pagesDeploy, pagesConfig } = require('../server/addin');
  const cfUrl = 'http://127.0.0.1:' + new URL(cf.url).port;
  const cfg = pagesConfig({ OSIGN_CF_ACCOUNT_ID: ACCOUNT, OSIGN_CF_PROJECT: 'company-osign', OSIGN_CF_API_TOKEN: 'wrong', OSIGN_CF_API_URL: cfUrl });
  await assert.rejects(pagesDeploy(cfg, path.join(dataDir, 'addin'), fetch), /refused the API token.*Cloudflare Pages: Edit/);
  assert.match(require('../server/addin').pagesProblem(pagesConfig({ OSIGN_CF_ACCOUNT_ID: 'x', OSIGN_CF_PROJECT: 'p' })), /account ID/);
});

test('Test in Server connections: the token works, and whether the project is there', async () => {
  const { pagesTest, pagesConfig } = require('../server/addin');
  const cfUrl = 'http://127.0.0.1:' + new URL(cf.url).port;
  const cfg = (project, token) => pagesConfig({ OSIGN_CF_ACCOUNT_ID: ACCOUNT, OSIGN_CF_PROJECT: project, OSIGN_CF_API_TOKEN: token, OSIGN_CF_API_URL: cfUrl });
  assert.deepEqual(await pagesTest(cfg('not-there-yet', 'cf-token'), fetch), { project: false });
  await assert.rejects(pagesTest(cfg('company-osign', 'wrong'), fetch), /refused the API token/);
});
