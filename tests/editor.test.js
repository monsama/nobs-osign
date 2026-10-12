// node --test tests/editor.test.js   (needs pwsh)
//
// Starts editor/OSign.ps1 as it runs for real, with its data in a temporary folder, and a stand-in
// for GitHub's Git Data API, and talks to it over HTTP the way the web UI does.
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn, execFileSync } = require('node:child_process');

const SCRIPT = path.join(__dirname, '..', 'editor', 'OSign.ps1');
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');

// ---------------------------------------------------------------------------------------------
// A small in-memory GitHub: one repository, one branch, Git Data API only.
// ---------------------------------------------------------------------------------------------
function gitSha(buf) { return crypto.createHash('sha1').update(Buffer.concat([Buffer.from('blob ' + buf.length + '\0'), buf])).digest('hex'); }
function fakeGitHub() {
  const blobs = {}, trees = {}, commits = {};
  let n = 0; const id = () => crypto.createHash('sha1').update(String(++n)).digest('hex');
  trees.t0 = { 'README.md': 'b0' }; blobs.b0 = Buffer.from('hello');
  commits.c0 = { tree: 't0', parents: [] };
  // main holds the code; the editor's branch (osign-data) is made by the first publish.
  const state = { refs: { main: 'c0' }, branch: 'osign-data', requests: [], dispatches: [], failNextRef: false };
  const server = http.createServer((req, res) => {
    let body = '';
    req.on('data', (d) => { body += d; });
    req.on('end', () => {
      const send = (code, obj) => { res.writeHead(code, { 'Content-Type': 'application/json' }); res.end(obj === undefined ? '' : JSON.stringify(obj)); };
      state.requests.push(req.method + ' ' + req.url);
      if (req.headers.authorization !== 'Bearer test-token') return send(401, { message: 'Bad credentials' });
      const u = new URL(req.url, 'http://x');
      if (req.method === 'GET' && u.pathname === '/repos/monsama/osign') return send(200, { default_branch: 'main' });
      if (req.method === 'POST' && u.pathname === '/repos/monsama/osign/actions/workflows/pages.yml/dispatches') { state.dispatches.push(JSON.parse(body)); return send(204); }
      const p = u.pathname.replace(/^\/repos\/monsama\/osign\/git/, '');
      const data = body ? JSON.parse(body) : null;
      let m;
      if (req.method === 'GET' && (m = /^\/ref\/heads\/(.+)$/.exec(p))) return state.refs[m[1]] ? send(200, { object: { sha: state.refs[m[1]] } }) : send(404, { message: 'Not Found' });
      if (req.method === 'GET' && (m = /^\/commits\/(\w+)$/.exec(p))) return commits[m[1]] ? send(200, { sha: m[1], tree: { sha: commits[m[1]].tree } }) : send(404, {});
      if (req.method === 'GET' && (m = /^\/trees\/(\w+)$/.exec(p))) return send(200, { sha: m[1], truncated: false, tree: Object.entries(trees[m[1]]).map(([path, sha]) => ({ path, sha, type: 'blob', mode: '100644' })) });
      if (req.method === 'GET' && (m = /^\/blobs\/(\w+)$/.exec(p))) return send(200, { content: blobs[m[1]].toString('base64').replace(/(.{60})/g, '$1\n'), encoding: 'base64' });
      if (req.method === 'POST' && p === '/blobs') { const b = Buffer.from(data.content, 'base64'); const s = gitSha(b); blobs[s] = b; return send(201, { sha: s }); }
      if (req.method === 'POST' && p === '/trees') {
        const t = data.base_tree ? Object.assign({}, trees[data.base_tree]) : {};
        data.tree.forEach((e) => { if (e.sha === null) delete t[e.path]; else t[e.path] = e.sha; });
        const s = id(); trees[s] = t; return send(201, { sha: s });
      }
      if (req.method === 'POST' && p === '/commits') { const s = id(); commits[s] = { tree: data.tree, parents: data.parents, message: data.message }; return send(201, { sha: s }); }
      if (req.method === 'POST' && p === '/refs') {
        const name = data.ref.replace(/^refs\/heads\//, '');
        if (state.refs[name]) return send(422, { message: 'Reference already exists' });
        state.refs[name] = data.sha; return send(201, { ref: data.ref, object: { sha: data.sha } });
      }
      if (req.method === 'PATCH' && (m = /^\/refs\/heads\/(.+)$/.exec(p))) {
        if (state.failNextRef || commits[data.sha].parents[0] !== state.refs[m[1]]) { state.failNextRef = false; return send(422, { message: 'Update is not a fast forward' }); }
        state.refs[m[1]] = data.sha; return send(200, { object: { sha: data.sha } });
      }
      send(404, { message: 'Not Found' });
    });
  });
  const head = () => state.refs[state.branch];
  const files = (branch) => { const t = trees[commits[state.refs[branch || state.branch]].tree]; const o = {}; Object.entries(t).forEach(([k, v]) => { o[k] = blobs[v]; }); return o; };
  return { server, state, files, commits, head };
}

// ---------------------------------------------------------------------------------------------
// OSIGN_TEST_SHELL=powershell runs it under Windows PowerShell 5.1 instead of PowerShell 7.
const PS = process.env.OSIGN_TEST_SHELL || 'pwsh';
let pwshOk = true;
try { execFileSync(PS, ['-NoProfile', '-Command', '1'], { stdio: 'ignore' }); } catch (e) { pwshOk = false; }

const PORT = 17781;
let proc, token, gh, dataDir;
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
  gh = fakeGitHub();
  await new Promise((r) => gh.server.listen(0, '127.0.0.1', r));
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-test-'));
  proc = spawn(PS, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT, '-NoBrowser', '-Port', String(PORT)], {
    env: Object.assign({}, process.env, { OSIGN_DATA: dataDir, GITHUB_TOKEN: 'test-token', OSIGN_GITHUB_API: 'http://127.0.0.1:' + gh.server.address().port })
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
  gh.server.close();
});

const opts = { skip: !pwshOk && PS + ' is not installed' };
const cfg = (sigs) => JSON.stringify({ schemaVersion: 1, assetBaseUrl: 'https://monsama.github.io/osign/config/', options: {}, auth: {}, fields: [], signatures: sigs || [] });

test('the API wants the token, a JSON POST, and this machine as the host', opts, async () => {
  assert.equal((await api('info', {}, { headers: { Authorization: 'Bearer ' + '0'.repeat(64) } })).status, 401);
  assert.equal((await api('info', {}, { headers: { Authorization: '' } })).status, 401);
  const evilOrigin = await api('info', {}, { headers: { Origin: 'https://evil.example' } });
  assert.equal(evilOrigin.status, 403);
  const r = await new Promise((resolve) => http.get({ host: '127.0.0.1', port: PORT, path: '/', headers: { Host: 'evil.example' } }, (res) => resolve(res.statusCode)));
  assert.equal(r, 403);
  const get = await fetch('http://127.0.0.1:' + PORT + '/api/info');
  assert.equal(get.status, 403);
  const info = await api('info');
  assert.equal(info.json.backend, 'powershell');
});

test('serves the UI and the shared engine, nothing outside them', opts, async () => {
  const ui = await fetch('http://127.0.0.1:' + PORT + '/');
  assert.equal(ui.status, 200);
  assert.match(ui.headers.get('content-security-policy'), /script-src 'self'/);
  assert.equal((await fetch('http://127.0.0.1:' + PORT + '/shared/signature-core.js')).status, 200);
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
  assert.equal((await api('assets/delete', { name: 'Our-Logo-2.png' })).json.ok, true);
  assert.deepEqual((await api('draft/get')).json.assets.map((x) => x.name), ['Our-Logo.png']);
});

test('settings: checked, and the token never comes back', opts, async () => {
  assert.equal((await api('settings/save', { settings: { target: { kind: 'github', owner: 'bad owner!' } } })).json.ok, false);
  assert.equal((await api('settings/save', { settings: { target: { kind: 'github', path: '../x' } } })).json.ok, false);
  assert.equal((await api('settings/save', { settings: { publicBaseUrl: 'http://insecure/' } })).json.ok, false);
  const s = await api('settings/save', { settings: { target: { kind: 'github', owner: 'monsama', repo: 'osign', branch: 'osign-data', path: 'config' }, publicBaseUrl: 'https://monsama.github.io/osign' } });
  assert.equal(s.json.ok, true);
  assert.equal(s.json.settings.publicBaseUrl, 'https://monsama.github.io/osign/');
  assert.equal(s.json.tokenSet, true);
  assert.ok(!JSON.stringify(s.json).includes('test-token'));
});

test('load published before anything was published says so', opts, async () => {
  const r = await api('load-published');
  assert.equal(r.json.ok, false);
  assert.match(r.json.error, /Nothing has been published/);
});

test('publish to GitHub: its own branch, one commit, unchanged files not sent, removed images deleted', opts, async () => {
  const text = cfg([{ id: 'std', name: 'Standard', html: '<p><img src="assets/Our-Logo.png"></p>' }]);
  const p = await api('publish', { configJson: text, message: 'First' });
  assert.equal(p.json.ok, true, JSON.stringify(p.json));
  assert.equal(p.json.changed, true);
  assert.equal(p.json.note, '');
  assert.match(p.json.commitUrl, /^https:\/\/github\.com\/monsama\/osign\/commit\/\w+$/);
  // A branch of its own, with no history shared with the code; main untouched.
  assert.deepEqual(gh.commits[gh.head()].parents, []);
  assert.equal(gh.state.refs.main, 'c0');
  let files = gh.files();
  assert.deepEqual(Object.keys(files).sort(), ['config/assets/Our-Logo.png', 'config/signatures.json']);
  assert.equal(files['config/signatures.json'].toString('utf8'), text);
  assert.ok(files['config/assets/Our-Logo.png'].equals(PNG));
  assert.equal((await api('draft/get')).json.dirty, false);
  // The Pages build was started, on the code's branch.
  assert.deepEqual(gh.state.dispatches, [{ ref: 'main' }]);

  // Same again: nothing to commit.
  const head = gh.head();
  const same = await api('publish', { configJson: text, message: 'Again' });
  assert.equal(same.json.changed, false);
  assert.equal(gh.head(), head);
  assert.equal(gh.state.dispatches.length, 1);

  // The image goes from the draft, so from the branch.
  await api('assets/delete', { name: 'Our-Logo.png' });
  gh.state.requests.length = 0;
  const p2 = await api('publish', { configJson: cfg([]), message: 'Second' });
  assert.equal(p2.json.changed, true);
  files = gh.files();
  assert.deepEqual(Object.keys(files).sort(), ['config/signatures.json']);
  assert.equal(gh.state.requests.filter((r) => r.startsWith('POST') && r.endsWith('/git/blobs')).length, 1);
  assert.equal(gh.commits[gh.head()].message, 'Second');
  assert.deepEqual(gh.commits[gh.head()].parents, [head]);
});

test('publish refuses to overwrite a newer commit', opts, async () => {
  gh.state.failNextRef = true;
  const p = await api('publish', { configJson: cfg([{ id: 'x', name: 'X', html: '<p>x</p>' }]), message: 'Race' });
  assert.equal(p.json.ok, false);
  assert.match(p.json.error, /published in the meantime/);
});

test('load published replaces the draft and its images', opts, async () => {
  await api('assets/upload', { name: 'stray.png', data: PNG.toString('base64') });
  await api('draft/save', { configJson: cfg([{ id: 'local', name: 'Local only', html: '' }]) });
  const r = await api('load-published');
  assert.equal(r.json.ok, true, JSON.stringify(r.json));
  assert.equal(r.json.configJson, cfg([]));
  assert.deepEqual(r.json.assets, []);
  assert.equal(r.json.dirty, false);
  assert.ok(fs.existsSync(path.join(dataDir, 'draft.json.bak')));
});

test('a bad token is reported as such', opts, async () => {
  // A second server would be needed to change the environment; the fake GitHub's own check is used.
  const saved = gh.server.listeners('request')[0];
  gh.server.removeAllListeners('request');
  gh.server.on('request', (req, res) => { req.resume(); res.writeHead(401, { 'Content-Type': 'application/json' }); res.end('{"message":"Bad credentials"}'); });
  const p = await api('publish', { configJson: cfg([]), message: 'x' });
  gh.server.removeAllListeners('request'); gh.server.on('request', saved);
  assert.equal(p.json.ok, false);
  assert.match(p.json.error, /did not accept the token/);
});

test('publish to a folder', opts, async () => {
  const folder = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-pub-'));
  assert.equal((await api('settings/save', { settings: { target: { kind: 'folder', folder, path: 'config' }, publicBaseUrl: 'https://intranet.example/osign/' } })).json.ok, true);
  await api('assets/upload', { name: 'logo.png', data: PNG.toString('base64') });
  fs.mkdirSync(path.join(folder, 'config', 'assets'), { recursive: true });
  fs.writeFileSync(path.join(folder, 'config', 'assets', 'old.png'), PNG);
  const p = await api('publish', { configJson: cfg([]), message: 'x' });
  assert.equal(p.json.ok, true, JSON.stringify(p.json));
  assert.equal(fs.readFileSync(path.join(folder, 'config', 'signatures.json'), 'utf8'), cfg([]));
  assert.deepEqual(fs.readdirSync(path.join(folder, 'config', 'assets')), ['logo.png']);
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
