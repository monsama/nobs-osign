// node --test tests/editor.test.js   (needs pwsh)
//
// Starts editor/OSign.ps1 as it runs for real, with its data in a temporary folder and a temporary
// folder standing in for the synced SharePoint folder, and talks to it over HTTP the way the web UI
// does.
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn, execFileSync } = require('node:child_process');

const SCRIPT = path.join(__dirname, '..', 'editor', 'OSign.ps1');
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');

// ---------------------------------------------------------------------------------------------
// OSIGN_TEST_SHELL=powershell runs it under Windows PowerShell 5.1 instead of PowerShell 7.
const PS = process.env.OSIGN_TEST_SHELL || 'pwsh';
let pwshOk = true;
try { execFileSync(PS, ['-NoProfile', '-Command', '1'], { stdio: 'ignore' }); } catch (e) { pwshOk = false; }

const PORT = 17781;
let proc, token, dataDir, folder;
async function api(p, body, opts) {
  const r = await fetch('http://127.0.0.1:' + PORT + '/api/' + p, {
    method: 'POST',
    headers: Object.assign({ 'Content-Type': 'application/json', Authorization: 'Bearer ' + token }, (opts && opts.headers) || {}),
    body: JSON.stringify(body || {})
  });
  return { status: r.status, json: await r.json().catch(() => null) };
}

test.before(async () => {
  if (!pwshOk) return;
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-test-'));
  folder = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-sp-'));
  proc = spawn(PS, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT, '-NoBrowser', '-Port', String(PORT)], {
    env: Object.assign({}, process.env, { OSIGN_DATA: dataDir })
  });
  token = await new Promise((resolve, reject) => {
    let out = '';
    const t = setTimeout(() => reject(new Error('OSign did not start: ' + out)), 30000);
    proc.stdout.on('data', (d) => { out += d; const m = /#t=([0-9a-f]{64})/.exec(out); if (m) { clearTimeout(t); resolve(m[1]); } });
    proc.stderr.on('data', (d) => { out += d; });
  });
});
test.after(async () => {
  if (!pwshOk) return;
  try { await api('quit'); } catch (e) { /* gone */ }
  setTimeout(() => { try { proc.kill(); } catch (e) { /* gone */ } }, 1500).unref();
});

const opts = { skip: !pwshOk && PS + ' is not installed' };
const cfg = (sigs) => JSON.stringify({ schemaVersion: 1, options: {}, fields: [], signatures: sigs || [] });

test('the API wants the token, a JSON POST, and this machine as the host', opts, async () => {
  assert.equal((await api('info', {}, { headers: { Authorization: 'Bearer ' + '0'.repeat(64) } })).status, 401);
  assert.equal((await api('info', {}, { headers: { Authorization: '' } })).status, 401);
  const evilOrigin = await api('info', {}, { headers: { Origin: 'https://evil.example' } });
  assert.equal(evilOrigin.status, 403);
  const r = await new Promise((resolve) => http.get({ host: '127.0.0.1', port: PORT, path: '/', headers: { Host: 'evil.example' } }, (res) => resolve(res.statusCode)));
  assert.equal(r, 403);
  const get = await fetch('http://127.0.0.1:' + PORT + '/api/info');
  assert.equal(get.status, 403);
  await get.text();
  const info = await api('info');
  assert.equal(info.json.backend, 'powershell');
  assert.deepEqual(Object.keys(info.json.features), ['adLookup']);
});

test('serves the UI and the shared engine, nothing outside them', opts, async () => {
  // Every body is read: an answer left unread while the server closes the connection trips Node's
  // fetch on Windows.
  const ui = await fetch('http://127.0.0.1:' + PORT + '/');
  assert.equal(ui.status, 200);
  assert.match(ui.headers.get('content-security-policy'), /script-src 'self'/);
  assert.match(await ui.text(), /<title>OSign<\/title>/);
  const coreJs = await fetch('http://127.0.0.1:' + PORT + '/shared/signature-core.js');
  assert.equal(coreJs.status, 200);
  assert.match(await coreJs.text(), /SignatureCore/);
  for (const p of ['/ui/../OSign.ps1', '/ui/..%2fOSign.ps1', '/assets/../draft.json', '/ui/%2e%2e/%2e%2e/package.json']) {
    const r = await new Promise((resolve) => http.get({ host: '127.0.0.1', port: PORT, path: p }, (res) => { res.resume(); resolve(res.statusCode); }));
    assert.equal(r, 404, p);
  }
});

test('draft: saved as sent, and checked', opts, async () => {
  assert.equal((await api('draft/save', { configJson: '{"schemaVersion":2}' })).json.ok, false);
  assert.equal((await api('draft/save', { configJson: 'nope' })).json.ok, false);
  const text = cfg([{ id: 'a', name: 'Ä "quoted" <b>', html: '<p>@@mail</p>', compiled: '', audience: { departments: ['X'], domains: [] } }]);
  assert.equal((await api('draft/save', { configJson: text })).json.ok, true);
  const g = await api('draft/get');
  assert.equal(g.json.configJson, text);
  assert.equal(g.json.dirty, true);
});

test('images: only real PNG/JPG/GIF, safe names, never overwritten', opts, async () => {
  assert.equal((await api('assets/upload', { name: 'x.svg', data: Buffer.from('<svg/>').toString('base64') })).json.ok, false);
  assert.equal((await api('assets/upload', { name: 'fake.png', data: Buffer.from('GIF89a....').toString('base64') })).json.ok, false);
  const a = await api('assets/upload', { name: '../../Our Logo!.png', data: PNG.toString('base64') });
  assert.equal(a.json.path, 'assets/Our-Logo.png');
  const again = await api('assets/upload', { name: 'Our Logo.png', data: PNG.toString('base64') });
  assert.equal(again.json.name, 'Our-Logo.png');
  const other = Buffer.concat([PNG, Buffer.from([0])]);
  const b = await api('assets/upload', { name: 'Our Logo.png', data: other.toString('base64') });
  assert.equal(b.json.name, 'Our-Logo-2.png');
  const img = await fetch('http://127.0.0.1:' + PORT + '/assets/Our-Logo.png');
  assert.equal(img.headers.get('content-type'), 'image/png');
  assert.ok(Buffer.from(await img.arrayBuffer()).equals(PNG));
  assert.equal((await api('assets/delete', { name: 'Our-Logo-2.png' })).json.ok, true);
  assert.deepEqual((await api('draft/get')).json.assets.map((x) => x.name), ['Our-Logo.png']);
});

test('settings: full paths only; publishing needs a folder that is there', opts, async () => {
  assert.equal((await api('settings/save', { settings: { folder: 'relative\\path' } })).json.ok, false);
  assert.equal((await api('publish', { configJson: cfg([]), message: 'x' })).json.error.includes('Pick the folder'), true);
  const gone = path.join(folder, 'not-there');
  assert.equal((await api('settings/save', { settings: { folder: gone } })).json.folderFound, false);
  assert.match((await api('publish', { configJson: cfg([]) })).json.error, /not there/);
  assert.match((await api('load-published')).json.error, /not there/);
  const s = await api('settings/save', { settings: { folder: folder + path.sep, photosFolder: '' } });
  assert.equal(s.json.ok, true);
  assert.equal(s.json.settings.folder, folder);
  assert.equal(s.json.folderFound, true);
  assert.equal(s.json.photosFound, false);
});

test('load published before anything was published says so', opts, async () => {
  const r = await api('load-published');
  assert.equal(r.json.ok, false);
  assert.match(r.json.error, /Nothing has been published/);
});

test('publish writes signatures.json and assets, only what changed, and never touches photos', opts, async () => {
  fs.mkdirSync(path.join(folder, 'photos'));
  fs.writeFileSync(path.join(folder, 'photos', 'mustera.jpg'), PNG);
  fs.mkdirSync(path.join(folder, 'assets'));
  fs.writeFileSync(path.join(folder, 'assets', 'old.png'), PNG);
  fs.writeFileSync(path.join(folder, 'assets', 'notes.txt'), 'not an image: left alone');
  const text = cfg([{ id: 'std', name: 'Standard', html: '<p><img src="assets/Our-Logo.png"></p>' }]);
  const p = await api('publish', { configJson: text, message: 'First' });
  assert.equal(p.json.ok, true, JSON.stringify(p.json));
  assert.equal(p.json.changed, true);
  assert.equal(fs.readFileSync(path.join(folder, 'signatures.json'), 'utf8'), text);
  assert.deepEqual(fs.readdirSync(path.join(folder, 'assets')).sort(), ['Our-Logo.png', 'notes.txt']);
  assert.ok(fs.readFileSync(path.join(folder, 'assets', 'Our-Logo.png')).equals(PNG));
  assert.deepEqual(fs.readdirSync(path.join(folder, 'photos')), ['mustera.jpg']);
  assert.equal((await api('draft/get')).json.dirty, false);

  // Same again: nothing written (OneDrive then uploads nothing).
  const before = fs.statSync(path.join(folder, 'signatures.json')).mtimeMs;
  const same = await api('publish', { configJson: text });
  assert.equal(same.json.changed, false);
  assert.equal(fs.statSync(path.join(folder, 'signatures.json')).mtimeMs, before);
});

test('the preview gets the photos: from the photos folder in the folder, or from the one set', opts, async () => {
  const get = (p) => new Promise((resolve) => http.get({ host: '127.0.0.1', port: PORT, path: p }, (res) => { const c = []; res.on('data', (d) => c.push(d)); res.on('end', () => resolve({ status: res.statusCode, type: res.headers['content-type'], body: Buffer.concat(c) })); }));
  const a = await get('/photos/mustera.jpg');
  assert.equal(a.status, 200);
  assert.equal(a.type, 'image/jpeg');
  assert.ok(a.body.equals(PNG));
  assert.equal((await get('/photos/nobody.jpg')).status, 404);
  assert.equal((await get('/photos/..%2fsignatures.json')).status, 404);
  const share = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-photos-'));
  fs.writeFileSync(path.join(share, 'other.jpg'), PNG);
  const s = await api('settings/save', { settings: { folder, photosFolder: share } });
  assert.equal(s.json.photosFound, true);
  assert.equal((await get('/photos/other.jpg')).status, 200);
  assert.equal((await get('/photos/mustera.jpg')).status, 404);
});

test('load published replaces the draft and its images', opts, async () => {
  await api('assets/upload', { name: 'stray.png', data: PNG.toString('base64') });
  await api('draft/save', { configJson: cfg([{ id: 'local', name: 'Local only', html: '' }]) });
  const r = await api('load-published');
  assert.equal(r.json.ok, true, JSON.stringify(r.json));
  assert.match(r.json.configJson, /"Standard"/);
  assert.deepEqual(r.json.assets.map((a) => a.name), ['Our-Logo.png']);
  assert.equal(r.json.dirty, false);
  assert.ok(fs.existsSync(path.join(dataDir, 'draft.json.bak')));
});

test('AD lookup says plainly where it cannot work', opts, async () => {
  const r = await api('ad/lookup', { query: 'x', fields: [] });
  if (process.platform !== 'win32') { assert.equal(r.json.ok, false); assert.match(r.json.error, /Windows/); }
});

test('helpers: LDAP escaping and asset names', opts, () => {
  const out = execFileSync(PS, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
    `$env:OSIGN_DATA='${dataDir}'; . '${SCRIPT}'; ConvertTo-LdapFilterValue 'a*)(|(x=1)\\'; Get-SafeAssetName 'C:\\x\\Logo Neu (2).JPG'; Get-SafeAssetName 'x.webp'; Get-SafeAssetName '....png'`], { encoding: 'utf8' });
  assert.deepEqual(out.trim().split(/\r?\n/), ['a\\2a\\29\\28|\\28x=1\\29\\5c', 'Logo-Neu-2.jpg', 'image.png']);
});
