// node --test tests/ui.test.js   (needs pwsh and Playwright's Chromium)
//
// The editor in a real browser against the real backend: make a signature, give it an image and
// a field, see the preview, and publish it to a folder.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawn, execFileSync } = require('node:child_process');

let chromium = null;
try { ({ chromium } = require('playwright')); } catch (e) { /* skipped below */ }
// OSIGN_TEST_SHELL=powershell runs it under Windows PowerShell 5.1 instead of PowerShell 7.
const PS = process.env.OSIGN_TEST_SHELL || 'pwsh';
let pwshOk = true;
try { execFileSync(PS, ['-NoProfile', '-Command', '1'], { stdio: 'ignore' }); } catch (e) { pwshOk = false; }
const skip = (!pwshOk && PS + ' is not installed') || (!chromium && 'playwright is not installed');

const PORT = 17782;
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');
let proc, url, browser, page, folder, dataDir;
const errors = [], notFound = [];

test.before(async () => {
  if (skip) return;
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-ui-'));
  folder = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-ui-pub-'));
  // Publishing goes to a folder: set before the start, as an admin would in Settings.
  fs.writeFileSync(path.join(dataDir, 'settings.json'), JSON.stringify({ target: { kind: 'folder', folder, path: 'config' }, publicBaseUrl: 'https://monsama.github.io/osign/' }));
  proc = spawn(PS, ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', path.join(__dirname, '..', 'editor', 'OSign.ps1'), '-NoBrowser', '-Port', String(PORT)], {
    env: Object.assign({}, process.env, { OSIGN_DATA: dataDir, GITHUB_TOKEN: '' })
  });
  url = await new Promise((resolve, reject) => {
    let out = '';
    const t = setTimeout(() => reject(new Error('OSign did not start: ' + out)), 30000);
    proc.stdout.on('data', (d) => { out += d; const m = /(http:\/\/127\.0\.0\.1:\d+\/#t=[0-9a-f]{64})/.exec(out); if (m) { clearTimeout(t); resolve(m[1]); } });
  });
  browser = await chromium.launch(fs.existsSync('/opt/pw-browsers/chromium') ? {} : {});
  page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) errors.push(m.text()); });
  page.on('response', (r) => { if (r.status() >= 400) notFound.push(new URL(r.url()).pathname); });
  page.on('dialog', (d) => d.accept());
});
test.after(async () => {
  if (skip) return;
  if (browser) await browser.close();
  try { await fetch(url.replace(/\/#t=.*/, '/api/quit'), { method: 'POST', headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + url.split('#t=')[1] }, body: '{}' }); } catch (e) { /* gone */ }
  setTimeout(() => { try { proc.kill(); } catch (e) { /* gone */ } }, 1500).unref();
});

const preview = () => page.frameLocator('#preview').locator('body');
// Each test leaves the next one where it starts: in the signature, design tab.
const pvText = () => page.evaluate(() => document.querySelector('#preview').contentDocument.body.innerText);

test('make, preview and publish a signature', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('#viewEmpty:not([hidden])');
  await page.click('#btnNew2');
  await page.waitForSelector('#viewSig:not([hidden])');
  assert.equal(await page.inputValue('#sigName'), 'Standard');

  // The starter template, filled with the sample values.
  await page.waitForFunction(() => (((document.querySelector('#preview').contentDocument || {}).body || {}).textContent || '').includes('Anna Muster'));
  assert.match(await pvText(), /Mobile \+41 79 123 45 67/);

  // Empty fields: their lines go.
  await page.selectOption('#pvSource', 'empty');
  await page.waitForFunction(() => !(((document.querySelector('#preview').contentDocument || {}).body || {}).textContent || '').includes('Mobile'));
  assert.match(await pvText(), /Freundliche Grüsse/);
  await page.selectOption('#pvSource', 'sample');

  // An image, uploaded and inserted at the end.
  const design = page.frameLocator('#design').locator('body');
  await design.click();
  await page.keyboard.press('Control+End');
  await page.click('#toolbar button[title=Image]');
  await page.setInputFiles('#imgFile', { name: 'Logo.png', mimeType: 'image/png', buffer: PNG });
  await page.waitForFunction(() => (document.querySelector('#design').contentDocument || document).querySelector('img[src="assets/Logo.png"]'));

  // A field typed through the HTML tab.
  await page.click('.tab[data-mode=html]');
  const src = await page.inputValue('#source');
  assert.match(src, /<img src="assets\/Logo\.png"/);
  await page.fill('#source', src + '\n<p>Abteilung @@department</p><p onclick="x()">@@nope</p><script>alert(1)</script>');
  await page.click('.tab[data-mode=design]');
  await page.waitForFunction(() => (((document.querySelector('#preview').contentDocument || {}).body || {}).textContent || '').includes('Abteilung Informatik'));
  assert.match(await page.textContent('#pvNote'), /Not a field: @@nope/);

  // The Entra app registration, so the publish check is clean of that warning.
  await page.click('a[data-view=options]');
  await page.fill('#optClient', '11111111-2222-3333-4444-555555555555');
  await page.fill('#optTenant', '66666666-7777-8888-9999-000000000000');
  assert.equal(await page.textContent('#manifestUrl'), 'https://monsama.github.io/osign/manifest.xml');

  await page.click('#btnPublish');
  await page.waitForSelector('#dlgPublish[open]');
  assert.match(await page.textContent('#pubChecks'), /Not fields, left as typed: @@nope/);
  await page.fill('#pubMsg', 'First signatures');
  await page.click('#pubGo');
  await page.waitForFunction(() => /Published/.test(document.querySelector('#pubOk').textContent));

  const published = JSON.parse(fs.readFileSync(path.join(folder, 'config', 'signatures.json'), 'utf8'));
  assert.equal(published.assetBaseUrl, 'https://monsama.github.io/osign/config/');
  assert.equal(published.auth.clientId, '11111111-2222-3333-4444-555555555555');
  assert.equal(published.signatures.length, 1);
  const s = published.signatures[0];
  assert.equal(s.defaultNew, true);
  assert.equal(s.defaultReply, true);
  assert.doesNotMatch(s.html, /<script|onclick=/);
  assert.match(s.html, /<img src="assets\/Logo\.png" width="1" alt="" style="border:0">/);
  assert.match(s.compiled, /<!--sig-if:department-->/);
  assert.ok(fs.readFileSync(path.join(folder, 'config', 'assets', 'Logo.png')).equals(PNG));
  assert.match(await page.textContent('#status'), /Published/);
  assert.deepEqual(errors, []);
});

test('fields: renaming a placeholder renames it in the signatures', { skip }, async () => {
  await page.keyboard.press('Escape');
  await page.click('a[data-view=fields]');
  const key = page.locator('#fieldRows tr').filter({ has: page.locator('input[value="department"]') }).locator('.ph input');
  await key.fill('abteilung');
  await key.press('Tab');
  await page.click('#sigList a');
  await page.click('.tab[data-mode=html]');
  assert.match(await page.inputValue('#source'), /Abteilung @@abteilung/);
  assert.match(await page.textContent('#status'), /Unpublished|Saving/);
});

test('import: UTF-16 templates with their images, text-block files left out, groups set', { skip }, async () => {
  const utf16 = (s) => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(s, 'utf16le')]);
  const tpl = '<!DOCTYPE html><html><head><title></title></head><body><table><tbody><tr>' +
    '<td><img src="images/logo.png" width="90" /><br /><img src="images/@@sAMAccountName.jpg" width="100" /></td>' +
    '<td><strong>@@givenName @@sn</strong><br />@@title<br />Mobile @@mobile</td></tr></tbody></table>@@info</body></html>';
  const blocks = '<html><body>{##notice}{dats 01.01.2020}{date 31.01.2020}Closed in January{end ##notice}</body></html>';
  await page.click('#btnImportOpen');
  await page.waitForSelector('#dlgImport[open]');
  await page.setInputFiles('#impFiles', [
    { name: 'Vertrieb.htm', mimeType: 'text/html', buffer: utf16(tpl) },
    { name: 'textblocks.htm', mimeType: 'text/html', buffer: utf16(blocks) },
    { name: 'logo.png', mimeType: 'image/png', buffer: PNG }
  ]);
  await page.waitForSelector('#impList li');
  const items = await page.$$eval('#impList li', (lis) => lis.map((li) => ({ name: li.querySelector('.imp-name').textContent, on: li.querySelector('input').checked, meta: li.querySelector('.imp-meta').textContent })));
  assert.deepEqual(items.map((i) => [i.name, i.on]), [['Vertrieb', true], ['textblocks', false]]);
  assert.match(items[0].meta, /per-person image images\/@@sAMAccountName\.jpg/);
  assert.match(items[1].meta, /text blocks/);
  await page.click('#impGo');
  await page.waitForSelector('#dlgImport:not([open])', { state: 'attached' });
  assert.equal(await page.inputValue('#sigName'), 'Vertrieb');

  // The logo is now one of the signature's images; the per-person image is not in the preview.
  await page.click('.tab[data-mode=html]');
  const src = await page.inputValue('#source');
  assert.match(src, /<img src="assets\/logo\.png" width="90">/);
  await page.click('.tab[data-mode=design]');
  // (The preview of the signature shown before may still be there for a moment.)
  await page.waitForFunction(() => { const d = document.querySelector('#preview').contentDocument; return d && d.body && /Anna Muster/.test(d.body.textContent) && !/Freundliche/.test(d.body.textContent); });
  const imgs = await page.evaluate(() => [...document.querySelector('#preview').contentDocument.querySelectorAll('img')].map((i) => i.getAttribute('src')));
  assert.deepEqual(imgs, ['/assets/logo.png']);

  // Who gets it: an AD group.
  await page.click('.audience summary');
  await page.fill('#sigGroups', 'SG-Vertrieb, SG-Marketing');
  assert.equal(await page.textContent('#audienceSummary'), 'SG-Vertrieb, SG-Marketing');
  await page.click('#btnPublish');
  await page.waitForSelector('#dlgPublish[open]');
  const checks = await page.textContent('#pubChecks');
  assert.match(checks, /GroupMember\.Read\.All/);
  assert.match(checks, /Per-person images are left out for now in: Vertrieb/);
  await page.click('#dlgPublish button[value=cancel]');
  assert.deepEqual(errors, []);
  // Only the editor's own view tried the per-person image (it shows as missing there).
  assert.deepEqual([...new Set(notFound)], ['/images/@@sAMAccountName.jpg']);
});
