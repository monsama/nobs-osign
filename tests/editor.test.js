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
const { spawn, execFile, execFileSync } = require('node:child_process');

const SCRIPT = path.join(__dirname, '..', 'editor', 'OSign.ps1');
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');

// ---------------------------------------------------------------------------------------------
// OSIGN_TEST_SHELL=powershell runs it under Windows PowerShell 5.1 instead of PowerShell 7.
const PS = process.env.OSIGN_TEST_SHELL || 'pwsh';
let pwshOk = true;
try { execFileSync(PS, ['-NoProfile', '-Command', '1'], { stdio: 'ignore' }); } catch (e) { pwshOk = false; }

const PORT = 17781;
let proc, token, dataDir, folder, adJson, syncJson;
// The directory, for the AD export: what Get-AdUserRecords would find.
const AD_USERS = [
  { mail: 'Anna.Muster@example.com', attrs: { description: 'Dipl. Ing.', homePhone: '+41 41 000 00 01', info: '' }, groups: ['sg-field'] },
  { mail: 'beat@example.com', attrs: { description: '', homePhone: '+41 41 000 00 02', info: 'Mo-Do' }, groups: [] }
];
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
  adJson = path.join(dataDir, 'ad-users.json');
  fs.writeFileSync(adJson, JSON.stringify(AD_USERS));
  // What OneDrive notes about the libraries it syncs to this PC.
  syncJson = path.join(dataDir, 'sync.json');
  fs.writeFileSync(syncJson, JSON.stringify([
    { MountPoint: path.join(dataDir, 'Contoso', 'IT - Documents'), UrlNamespace: 'https://contoso.sharepoint.com/sites/IT/Shared Documents/' },
    { MountPoint: path.join(dataDir, 'Contoso', 'IT - Documents', 'General'), UrlNamespace: 'https://contoso.sharepoint.com/sites/IT/Shared Documents/General/' },
    { MountPoint: path.join(dataDir, 'OneDrive - Contoso'), UrlNamespace: 'https://contoso-my.sharepoint.com/personal/anna_contoso_com/Documents/' },
    { MountPoint: path.join(dataDir, 'Contoso', 'Home - Site Assets'), UrlNamespace: 'https://contoso.sharepoint.com/SiteAssets/' },
    // OneDrive notes some places twice; and a shortcut to another site's library, inside the OneDrive.
    { MountPoint: path.join(dataDir, 'OneDrive - Contoso'), UrlNamespace: 'https://contoso-my.sharepoint.com/personal/anna_contoso_com/Documents/' },
    { MountPoint: path.join(dataDir, 'OneDrive - Contoso', 'Sales - Team'), UrlNamespace: 'https://contoso.sharepoint.com/sites/Sales/Shared Documents/' }
  ]));
  folder = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-sp-'));
  proc = spawn(PS, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT, '-NoBrowser', '-Port', String(PORT)], {
    env: Object.assign({}, process.env, { OSIGN_DATA: dataDir, OSIGN_TEST_AD_JSON: adJson, OSIGN_TEST_SYNC_JSON: syncJson })
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
  assert.deepEqual(Object.keys(info.json.features).sort(), ['adLookup', 'addinFiles', 'pickFolder']);
  assert.equal(info.json.features.addinFiles, true);
});

test('serves the UI and the shared engine, nothing outside them', opts, async () => {
  // Every body is read: an answer left unread while the server closes the connection trips Node's
  // fetch on Windows.
  const ui = await fetch('http://127.0.0.1:' + PORT + '/');
  assert.equal(ui.status, 200);
  assert.match(ui.headers.get('content-security-policy'), /script-src 'self'/);
  assert.match(await ui.text(), /<title>NOBS OSign<\/title>/);
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

test('AD users: read first (nothing written), check, then exactly that is written - once', opts, async () => {
  const req = { fields: [{ key: 'description', ad: 'description' }, { key: 'homePhone', ad: 'homePhone' }, { key: 'info', ad: 'info' }], groups: ['SG-Field'] };
  fs.mkdirSync(path.join(folder, 'users'), { recursive: true });
  fs.writeFileSync(path.join(folder, 'users', 'left@example.com.json'), '{}');
  fs.writeFileSync(path.join(folder, 'users', 'beat@example.com.json'), JSON.stringify({ values: { description: '', homePhone: '+41 41 000 00 99', info: 'Mo-Do' }, groups: [] }));
  const files = () => fs.readdirSync(path.join(folder, 'users')).sort();

  // Nothing read yet: nothing is written.
  const none = await api('ad/export', req);
  assert.equal(none.json.ok, false);
  assert.match(none.json.error, /Read the users from AD and check them first/);

  // Read: what would change, with the old values; the folder is untouched.
  const r = await api('ad/read', req);
  assert.equal(r.json.ok, true, JSON.stringify(r.json));
  assert.deepEqual(r.json.counts, { new: 1, changed: 1, unchanged: 0, removed: 1 });
  const beat = r.json.users.find((u) => u.mail === 'beat@example.com');
  assert.equal(beat.status, 'changed');
  assert.equal(beat.values.homePhone, '+41 41 000 00 02');
  assert.equal(beat.old.homePhone, '+41 41 000 00 99');
  assert.deepEqual(r.json.removed, ['left@example.com']);
  assert.deepEqual(files(), ['beat@example.com.json', 'left@example.com.json']);
  assert.equal((await api('ad/snapshot', req)).json.matches, true);
  assert.equal((await api('ad/snapshot', { fields: req.fields, groups: [] })).json.matches, false);

  // Another set of fields than was read: refused.
  const other = await api('ad/export', { fields: req.fields.slice(0, 1), groups: req.groups });
  assert.match(other.json.error, /changed since the users were read/);

  // Write: exactly what was read, even if AD changed in between.
  const ad = JSON.parse(fs.readFileSync(adJson, 'utf8'));
  fs.writeFileSync(adJson, JSON.stringify(ad.map((u) => Object.assign({}, u, { attrs: Object.assign({}, u.attrs, { homePhone: 'CHANGED IN AD LATER' }) }))));
  const w = await api('ad/export', req);
  fs.writeFileSync(adJson, JSON.stringify(ad));
  assert.equal(w.json.ok, true, JSON.stringify(w.json));
  assert.deepEqual([w.json.users, w.json.written, w.json.removed], [2, 2, 1]);
  assert.deepEqual(files(), ['anna.muster@example.com.json', 'beat@example.com.json']);
  assert.deepEqual(JSON.parse(fs.readFileSync(path.join(folder, 'users', 'anna.muster@example.com.json'), 'utf8')),
    { values: { description: 'Dipl. Ing.', homePhone: '+41 41 000 00 01', info: '' }, groups: ['sg-field'] });

  // Once: an old read is not written again over what may be newer.
  const again = await api('ad/export', req);
  assert.match(again.json.error, /written already/);
  const re = await api('ad/read', req);
  assert.deepEqual(re.json.counts, { new: 0, changed: 0, unchanged: 2, removed: 0 });
});

test('-ExportUsers: the export alone, as the published signatures ask for', opts, async () => {
  const cfgPath = path.join(folder, 'signatures.json');
  const saved = fs.readFileSync(cfgPath, 'utf8');
  const run = () => { try { return { code: 0, out: execFileSync(PS, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', SCRIPT, '-ExportUsers'], { env: Object.assign({}, process.env, { OSIGN_DATA: dataDir, OSIGN_TEST_AD_JSON: adJson, OSIGN_TEST_SYNC_JSON: syncJson }), encoding: 'utf8' }) }; } catch (e) { return { code: e.status, out: String(e.stdout) + String(e.stderr) }; } };
  try {
    fs.writeFileSync(cfgPath, JSON.stringify({ schemaVersion: 1, options: { valueSource: 'graph' }, fields: [], signatures: [] }));
    assert.match(run().out, /Entra ID: nothing to export/);
    fs.writeFileSync(cfgPath, JSON.stringify({ schemaVersion: 1, options: { valueSource: 'mixed' },
      fields: [{ key: 'givenName', ad: 'givenName', graph: 'givenName' }, { key: 'homePhone', ad: 'homePhone', graph: '' }],
      signatures: [{ id: 'a', name: 'A', html: '', audience: { groups: ['SG-Field'] } }] }));
    const r = run();
    assert.equal(r.code, 0, r.out);
    assert.match(r.out, /Exported 2 users/);
    // A read the editor still held is older now: Publish must not write it.
    assert.equal((await api('ad/snapshot', { fields: [], groups: [] })).json.snapshot.written, true);
    // Mixed: only the field without a Graph source comes from AD.
    assert.deepEqual(JSON.parse(fs.readFileSync(path.join(folder, 'users', 'beat@example.com.json'), 'utf8')), { values: { homePhone: '+41 41 000 00 02' }, groups: [] });
  } finally { fs.writeFileSync(cfgPath, saved); }
});

test('the add-in files: the template filled in, the same as addin/fill.mjs does it', opts, async () => {
  const out = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-addin-out-'));
  const base = { folder, photosFolder: '' };
  const bad = await api('settings/save', { settings: { ...base, clientId: 'not-a-guid' } });
  assert.equal(bad.json.ok, false);
  assert.match(bad.json.error, /GUID/);
  assert.match((await api('settings/save', { settings: { ...base, addinUrl: 'http://osign.example.com' } })).json.error, /https/);
  assert.match((await api('settings/save', { settings: { ...base, spSite: 'contoso.sharepoint.com/sites/x' } })).json.error, /SharePoint site/);
  // A OneDrive, and /teams/ sites, are addressed the same way.
  for (const spSite of ['contoso-my.sharepoint.com:/personal/anna_contoso_com:', 'contoso.sharepoint.com:/teams/IT:']) assert.equal((await api('settings/save', { settings: { ...base, spSite } })).json.ok, true, spSite);
  assert.match((await api('settings/save', { settings: { ...base, spSite: 'contoso-my.sharepoint.com:/personal/a/b:' } })).json.error, /OneDrive/);
  await api('settings/save', { settings: { spSite: 'root' } });
  await api('settings/save', { settings: { ...base, addinUrl: 'https://osign.example.com/' } });
  assert.match((await api('addin/build')).json.error, /client ID/);
  await api('settings/save', { settings: { ...base, addinUrl: 'https://osign.example.com/', clientId: '11111111-2222-3333-4444-555555555555', addinFolder: folder } });
  assert.match((await api('addin/build')).json.error, /folder of their own/);

  const setup = { addinUrl: 'https://osign.example.com/sign/', clientId: '11111111-2222-3333-4444-555555555555', tenantId: 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee', spSite: 'contoso.sharepoint.com:/sites/Intranet:', spFolder: 'Vorlagen/Öffentlich OSign', addinHost: 'cloudflare' };
  assert.equal((await api('settings/save', { settings: { ...base, ...setup, addinFolder: out } })).json.ok, true);
  const r = (await api('addin/build')).json;
  assert.equal(r.ok, true, r.error);
  assert.equal(r.manifestUrl, 'https://osign.example.com/sign/manifest.xml');
  assert.equal(r.redirectUri, 'brk-multihub://osign.example.com');
  assert.equal(r.version, require('../package.json').version);

  const { fillTemplate } = await import('../addin/fill.mjs');
  const ref = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-addin-ref-'));
  const template = path.join(__dirname, '..', 'dist', 'template');
  fillTemplate(template, ref, { base: setup.addinUrl, clientId: setup.clientId, tenantId: setup.tenantId, site: setup.spSite, folder: setup.spFolder, host: setup.addinHost });
  const list = (d) => fs.readdirSync(d, { recursive: true }).filter((f) => fs.statSync(path.join(d, f)).isFile()).map((f) => f.split(path.sep).join('/')).sort();
  assert.deepEqual(list(out), list(ref));
  assert.equal(r.files, list(ref).length);
  for (const f of list(ref)) assert.ok(fs.readFileSync(path.join(out, f)).equals(fs.readFileSync(path.join(ref, f))), f + ' differs');
  for (const f of ['launchevent.js', 'taskpane.js', 'manifest.xml']) assert.doesNotMatch(fs.readFileSync(path.join(out, f), 'utf8'), /@@OSIGN_|\{\{[A-Z]+\}\}/, f);
  assert.ok(fs.readFileSync(path.join(out, 'launchevent.js'), 'utf8').startsWith('var OSIGN_SETUP={clientId:"11111111-2222-3333-4444-555555555555",tenantId:"aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",site:"contoso.sharepoint.com:/sites/Intranet:",folder:"Vorlagen/Öffentlich OSign"};'));
  // Cloudflare's caching file goes with them; the other web servers' files do not.
  assert.equal(fs.readFileSync(path.join(out, '_headers'), 'utf8'), '/*\n  Cache-Control: no-cache\n');
  assert.ok(!list(out).some((f) => /^hosting\/|web\.config|\.htaccess|staticwebapp/.test(f)), list(out).join(', '));
  assert.match((await api('settings/save', { settings: { addinHost: 'netlify' } })).json.error, /web server/);
  assert.match(fs.readFileSync(path.join(out, 'manifest.xml'), 'utf8'), /<SourceLocation DefaultValue="https:\/\/osign\.example\.com\/sign\/taskpane\.html"\/>/);
  // Saving sends only what changes; the rest stays.
  const kept = await api('settings/save', { settings: { photosFolder: '' } });
  assert.equal(kept.json.settings.clientId, setup.clientId);
  assert.equal(kept.json.settings.addinHost, 'cloudflare');

  // Checking the address: the same files are there, or not, and the caching header.
  const served = {};
  const srv = http.createServer((req, res) => {
    const f = served[req.url];
    if (!f) { res.writeHead(404); res.end(); return; }
    res.writeHead(200, { 'Content-Type': f.type, 'Cache-Control': f.cc || 'max-age=3600' }); res.end(f.body);
  });
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  try {
    for (const f of ['manifest.xml', 'launchevent.js', 'taskpane.js', 'taskpane.html']) served['/' + f] = { type: f.endsWith('.xml') ? 'text/xml' : 'text/plain', body: fs.readFileSync(path.join(out, f)), cc: 'no-cache' };
    served['/taskpane.js'].body = Buffer.from('old');
    // Not execFileSync: the server above has to answer while PowerShell asks.
    const check = (base) => new Promise((resolve, reject) => execFile(PS, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command',
      `$env:OSIGN_DATA='${dataDir}'; . '${SCRIPT}'; ConvertTo-Json -Compress -Depth 5 -InputObject (Test-AddinOnline '${base}' '${out}')`], { encoding: 'utf8' },
    (err, stdout) => (err ? reject(err) : resolve(JSON.parse(stdout)))));
    const r1 = await check('http://127.0.0.1:' + srv.address().port + '/');
    const by = Object.fromEntries(r1.map((c) => [c.what, c]));
    assert.equal(by['manifest.xml'].ok, true, by['manifest.xml'].detail);
    assert.equal(by['launchevent.js'].ok, true, by['launchevent.js'].detail);
    assert.equal(by['taskpane.js'].ok, false);
    assert.match(by['taskpane.js'].detail, /upload them again/);
    assert.equal(by['Cache-Control'].ok, true);
    delete served['/taskpane.html'];
    served['/manifest.xml'].cc = 'max-age=3600';
    const by2 = Object.fromEntries((await check('http://127.0.0.1:' + srv.address().port)).map((c) => [c.what, c]));
    assert.match(by2['taskpane.html'].detail, /cannot be loaded/);
    assert.equal(by2['Cache-Control'].ok, false);
  } finally { srv.close(); }
  await api('settings/save', { settings: { addinUrl: '', clientId: '', tenantId: '', spSite: 'root', spFolder: 'OSign', addinFolder: '', addinHost: '' } });
});

test('where a synced folder is in SharePoint, from what OneDrive notes', opts, async () => {
  const at = async (...p) => (await api('sp/resolve', { path: path.join(dataDir, ...p) })).json;
  let r = await at('Contoso', 'IT - Documents', 'General', 'OSign');
  assert.equal(r.found, true);
  // The longest match: the General library sync, not the whole library's.
  assert.deepEqual([r.location.site, r.location.siteName, r.location.library, r.location.folder, r.location.defaultLibrary], ['contoso.sharepoint.com:/sites/IT:', 'IT', 'Shared Documents', 'General/OSign', true]);
  assert.equal(r.location.url, 'https://contoso.sharepoint.com/sites/IT/Shared Documents/General/OSign');
  r = await at('OneDrive - Contoso', 'Tests', 'OSign');
  assert.deepEqual([r.location.site, r.location.folder], ['contoso-my.sharepoint.com:/personal/anna_contoso_com:', 'Tests/OSign']);
  r = await at('Contoso', 'Home - Site Assets', 'OSign');
  assert.deepEqual([r.location.site, r.location.library, r.location.folder, r.location.defaultLibrary], ['root', 'SiteAssets', 'OSign', false]);
  assert.equal((await at('Elsewhere', 'OSign')).found, false);
  assert.equal((await api('sp/resolve', { path: '' })).json.found, false);
});

test('a SharePoint folder\'s address: where OneDrive syncs it to on this PC; shortcuts are not guessed', opts, async () => {
  const local = async (url) => (await api('sp/local', { url })).json;
  let r = await local('https://contoso.sharepoint.com/sites/IT/Shared%20Documents/General/OSign');
  // The longest match: the General folder's own sync.
  assert.deepEqual([r.found, r.path, r.exists], [true, path.join(dataDir, 'Contoso', 'IT - Documents', 'General', 'OSign'), false]);
  r = await local('https://contoso-my.sharepoint.com/personal/anna_contoso_com/Documents/Tests/OSign');
  assert.equal(r.path, path.join(dataDir, 'OneDrive - Contoso', 'Tests', 'OSign'));
  // A shortcut in the OneDrive to another site's library: which folder of it is not known, so it
  // is never taken, neither for a link nor for a folder in it.
  r = await local('https://contoso.sharepoint.com/sites/Sales/Shared%20Documents/Team/OSign');
  assert.equal(r.found, false);
  const where = (await api('sp/resolve', { path: path.join(dataDir, 'OneDrive - Contoso', 'Sales - Team', 'OSign') })).json;
  assert.deepEqual([where.found, where.location.uncertain, where.location.site, where.location.folder], [true, true, 'contoso.sharepoint.com:/sites/Sales:', '']);
  assert.equal((await local('https://contoso.sharepoint.com/sites/Other/Shared%20Documents/OSign')).found, false);
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
