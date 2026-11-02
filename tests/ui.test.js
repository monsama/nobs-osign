// node --test tests/ui.test.js   (needs Playwright's Chromium)
//
// The editor in a real browser against the real server (server/app.js, in this process): make a
// signature, give it an image and a field, see the preview, publish it to the server's folder,
// read users from a stand-in AD, and go through the setup assistant.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const http = require('node:http');
const path = require('node:path');
const { createApp } = require('../server/app');
const { fakeAd } = require('./support/fake-ad');

let chromium = null;
try { ({ chromium } = require('playwright')); } catch (e) { /* skipped below */ }
const skip = !chromium && 'playwright is not installed';

const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');
const BASE = 'DC=example,DC=ch';
const SERVICE = { dn: 'CN=svc-osign,OU=Service,DC=example,DC=ch', pw: 'svc-pw' };
let srv, ad, base, url, cookie, browser, context, page, folder, dataDir;
const errors = [], notFound = [];

// A browser context signed in as the local account admin.
async function signedIn(opts) {
  const ctx = await browser.newContext(Object.assign({ viewport: { width: 1400, height: 900 } }, opts || {}));
  await ctx.addCookies([{ name: 'osign_session', value: cookie, url: base }]);
  return ctx;
}

test.before(async () => {
  if (skip) return;
  dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-ui-'));
  folder = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-ui-pub-'));
  ad = fakeAd({ base: BASE, service: SERVICE, groups: { 'NOBS-OSign-Admins': [] }, people: {
    anna: { pw: 'anna-pw', name: 'Anna Muster', mail: 'anna@example.com', attrs: { description: 'Dipl. Ing.', homePhone: '+41 41 000 00 01', info: '' }, groups: ['NOBS-OSign-Admins'] },
    beat: { pw: 'beat-pw', name: 'Beat Beispiel', mail: 'beat@example.com', attrs: { description: '', homePhone: '+41 41 000 00 02', info: '' }, groups: [] }
  } });
  const ldapUrl = await ad.listen();
  // Set up already (the assistant's own test starts it again), publishing into the server's folder.
  fs.writeFileSync(path.join(dataDir, 'settings.json'), JSON.stringify({ setupDone: '1' }));
  // A photo for the sample user (mustera), as an admin would have copied it there.
  fs.mkdirSync(path.join(folder, 'photos'));
  fs.writeFileSync(path.join(folder, 'photos', 'mustera.jpg'), PNG);
  const env = {
    OSIGN_DATA: dataDir, OSIGN_PUBLISH_DIR: folder, OSIGN_ADMIN_PASSWORD: 'admin-pw', OSIGN_SECURE_COOKIE: '0', OSIGN_APP_KEY: 'a-key-for-the-ui-test',
    OSIGN_LDAP_URL: ldapUrl, OSIGN_LDAP_ALLOW_PLAIN: '1', OSIGN_LDAP_BIND: 'cn={login}@example.ch,' + BASE, OSIGN_LDAP_BASE: 'OU=Users,' + BASE,
    OSIGN_LDAP_GROUP: 'CN=NOBS-OSign-Admins,OU=Groups,' + BASE, OSIGN_LDAP_SERVICE_DN: SERVICE.dn, OSIGN_LDAP_SERVICE_PASSWORD: SERVICE.pw
  };
  srv = http.createServer(createApp({ env, log: { info() {}, warn() {}, error() {} } }));
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  base = 'http://127.0.0.1:' + srv.address().port;
  url = base + '/';
  const r = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'admin', password: 'admin-pw' }) });
  cookie = /osign_session=([^;]+)/.exec(r.headers.get('set-cookie'))[1];
  browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  context = await signedIn();
  page = await context.newPage();
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) errors.push(m.text() + ' @ ' + JSON.stringify(m.location())); });
  page.on('response', (r) => { if (r.status() >= 400) notFound.push(new URL(r.url()).pathname); });
  page.on('dialog', (d) => d.accept());
});
test.after(async () => {
  if (skip) return;
  if (browser) await browser.close();
  srv.close(); ad.close();
});
const api = (p, body) => fetch(base + '/api/' + p, { method: 'POST', headers: { 'Content-Type': 'application/json', Cookie: 'osign_session=' + cookie }, body: JSON.stringify(body || {}) }).then((r) => r.json());

const preview = () => page.frameLocator('#preview').locator('body');
// Each test leaves the next one where it starts: in the signature, design tab.
const pvText = () => page.evaluate(() => document.querySelector('#preview').contentDocument.body.innerText);

test('make, preview and publish a signature', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('#viewEmpty:not([hidden])');
  await page.click('#btnNew2');
  await page.waitForSelector('#viewSig:not([hidden])');
  assert.equal(await page.inputValue('#sigName'), 'Main signature');

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


  await page.click('#btnPublish');
  await page.waitForSelector('#dlgPublish[open]');
  assert.match(await page.textContent('#pubChecks'), /Unknown placeholders, left as typed: @@nope/);
  await page.fill('#pubMsg', 'First signatures');
  await page.click('#pubGo');
  await page.waitForFunction(() => /Published/.test(document.querySelector('#pubOk').textContent));

  const published = JSON.parse(fs.readFileSync(path.join(folder, 'signatures.json'), 'utf8'));
  // No sign-in details, no addresses: only the signatures.
  assert.equal(published.auth, undefined);
  assert.equal(published.assetBaseUrl, undefined);
  assert.equal(published.signatures.length, 1);
  const s = published.signatures[0];
  assert.equal(s.defaultNew, true);
  assert.equal(s.defaultReply, true);
  assert.doesNotMatch(s.html, /<script|onclick=/);
  assert.match(s.html, /<img src="assets\/Logo\.png" width="1" alt="" style="border:0">/);
  assert.match(s.compiled, /<!--sig-if:department-->/);
  assert.ok(fs.readFileSync(path.join(folder, 'assets', 'Logo.png')).equals(PNG));
  await page.waitForFunction(() => /Published/.test(document.querySelector('#status').textContent));
  assert.deepEqual(errors, []);
});

test('fields: renaming a placeholder renames it in the signatures', { skip }, async () => {
  await page.keyboard.press('Escape');
  await page.click('a[data-view=fields]');
  const key = page.locator('#fieldRows tr').filter({ has: page.locator('input[value="department"]') }).locator('.ph input');
  await key.fill('abteilung');
  await key.press('Tab');
  // Asked in the editor's own dialog, its answer a button that says what it does.
  await page.waitForSelector('#dlgAsk[open]');
  assert.equal(await page.textContent('#askOk'), 'Change them too');
  await page.click('#askOk');
  await page.click('#sigList a');
  await page.click('.tab[data-mode=html]');
  assert.match(await page.inputValue('#source'), /Abteilung @@abteilung/);
  assert.match(await page.textContent('#status'), /Unpublished|Saving/);
});

test('import: UTF-16 templates with their images, text-block files left out, groups set', { skip }, async () => {
  const utf16 = (s) => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(s, 'utf16le')]);
  const tpl = '<!DOCTYPE html><html><head><title></title></head><body><table><tbody><tr>' +
    '<td><img src="images/logo.png" width="90" /><br /><img src="images/@@sAMAccountName.jpg" width="100" /></td>' +
    '<td><strong>@@givenName @@sn</strong><br />@@title<br />Mobile @@mobile</td></tr></tbody></table>@@info{##notice}{##legal}{##old}</body></html>';
  const blocks = '<html><body>{##notice}{dats 01.01.2099}{date 31.01.2099}<p>Closed in January</p>{end ##notice}{##legal}<p>Contoso AG</p>{end ##legal}{##old}{dats 01.01.2020}{date 31.01.2020}<p>Fair 2020</p>{end ##old}</body></html>';
  await page.click('#btnImportOpen');
  await page.waitForSelector('#dlgImport[open]');
  await page.setInputFiles('#impFiles', [
    { name: 'Vertrieb.htm', mimeType: 'text/html', buffer: utf16(tpl) },
    { name: 'textblocks.htm', mimeType: 'text/html', buffer: utf16(blocks) },
    { name: 'logo.png', mimeType: 'image/png', buffer: PNG },
    { name: 'mustera.jpg', mimeType: 'image/jpeg', buffer: Buffer.from('ffd8ffe000104a464946', 'hex') },
    { name: 'Newsletter-1.jpg', mimeType: 'image/jpeg', buffer: Buffer.from('ffd8ffe000104a464946', 'hex') }
  ]);
  await page.waitForSelector('#impList li');
  const items = await page.$$eval('#impList li', (lis) => lis.map((li) => ({ name: li.querySelector('.imp-name').textContent, on: li.querySelector('input').checked, meta: li.querySelector('.imp-meta').textContent })));
  assert.deepEqual(items.map((i) => [i.name, i.on]), [['Vertrieb', true], ['textblocks', false], ['Photos per person: 1 into photos\\ in the OSign folder', false]]);
  // Named like a user: a photo; a newsletter image is not one. Off unless ticked (photos count at once).
  assert.match(items[2].meta, /^mustera\.jpg · used at once/);
  assert.match(items[0].meta, /photo per person images\/@@sAMAccountName\.jpg → photos\\@@sAMAccountName\.jpg/);
  assert.match(items[1].meta, /text blocks/);
  assert.match(items[0].meta, /past text blocks left out: old · text blocks notice, legal, old \(with days: a campaign each\)/);
  await page.click('#impGo');
  await page.waitForSelector('#dlgImport:not([open])', { state: 'attached' });
  assert.equal(await page.inputValue('#sigName'), 'Vertrieb');
  // The text block without days in the template; the one with days a campaign of its own.
  let imported = '';
  for (let i = 0; i < 50 && !/Vertrieb – notice/.test(imported); i++) { await new Promise((r) => setTimeout(r, 100)); imported = (await api('draft/get')).configJson || ''; }
  const sigs = JSON.parse(imported).signatures;
  const vt = sigs.find((x) => x.name === 'Vertrieb'), camp = sigs.find((x) => x.name === 'Vertrieb – notice');
  assert.match(vt.html, /Contoso AG/);
  assert.doesNotMatch(vt.html, /Closed in January|Fair 2020|\{##/);
  assert.ok(!sigs.some((x) => /old$/.test(x.name)), 'a past block makes no campaign');
  assert.match(camp.html, /Closed in January.*Contoso AG|Contoso AG.*Closed in January/s);
  assert.deepEqual([camp.validFrom, camp.validUntil], ['2099-01-01', '2099-01-31']);
  await page.click('#sigList a:has-text("Vertrieb – notice")');
  await page.click('#btnDel');
  // Cancel keeps it; Delete deletes.
  await page.click('#askCancel');
  assert.ok(await page.isVisible('#sigList a:has-text("Vertrieb – notice")'));
  await page.click('#btnDel');
  assert.equal(await page.textContent('#askOk'), 'Delete');
  await page.click('#askOk');
  await page.click('#sigList a:has-text("Vertrieb")');

  // The logo is now one of the signature's images; the photo comes from the photos folder.
  await page.click('.tab[data-mode=html]');
  const src = await page.inputValue('#source');
  assert.match(src, /<img src="assets\/logo\.png" width="90">/);
  assert.match(src, /<img src="photos\/@@sAMAccountName\.jpg" width="100">/);
  await page.click('.tab[data-mode=design]');
  // (The preview of the signature shown before may still be there for a moment.)
  await page.waitForFunction(() => { const d = document.querySelector('#preview').contentDocument; return d && d.body && /Anna Muster/.test(d.body.textContent) && !/Freundliche/.test(d.body.textContent); });
  const imgs = await page.evaluate(() => [...document.querySelector('#preview').contentDocument.querySelectorAll('img')].map((i) => i.getAttribute('src')));
  assert.deepEqual(imgs, ['/assets/logo.png', '/photos/mustera.jpg']);
  // Both load: the logo from the draft, the photo from the photos folder.
  await page.waitForFunction(() => [...document.querySelector('#preview').contentDocument.querySelectorAll('img')].every((i) => i.complete && i.naturalWidth > 0), null, { timeout: 5000 });

  // Who gets it: an AD group.
  await page.click('.audience:not(.campaign) summary');
  await page.fill('#sigGroups', 'SG-Vertrieb, SG-Marketing');
  assert.equal(await page.textContent('#audienceSummary'), 'SG-Vertrieb, SG-Marketing');
  await page.click('#btnPublish');
  await page.waitForSelector('#dlgPublish[open]');
  const checks = await page.textContent('#pubChecks');
  // The groups come from AD (new setups take Entra ID and AD): no Graph group permission asked for.
  assert.doesNotMatch(checks, /GroupMember\.Read\.All/);
  assert.match(checks, /Photos per person \(Vertrieb\) come from photos\\<username>\.jpg/);
  await page.click('#dlgPublish button[value=cancel]');
  assert.deepEqual(errors, []);
  // The design view's photo with its placeholder gets the server's grey figure: nothing missing.
  assert.deepEqual([...new Set(notFound)], []);
  const ph = await page.evaluate(() => fetch('/photos/@@sAMAccountName.jpg').then((r) => r.status + ' ' + r.headers.get('content-type')));
  assert.equal(ph, '200 image/svg+xml');
});

test('German: in a German browser the editor speaks German, and offers no translation', { skip }, async () => {
  const deCtx = await signedIn({ locale: 'de-DE' });
  const de = await deCtx.newPage();
  const errs = [];
  de.on('pageerror', (e) => errs.push(e.message));
  await de.goto(url);
  await de.waitForSelector('#viewSig:not([hidden]), #viewEmpty:not([hidden])');
  assert.equal(await de.getAttribute('html', 'lang'), 'de');
  assert.equal(await de.getAttribute('html', 'translate'), 'no');
  assert.equal(await de.getAttribute('meta[name=google]', 'content'), 'notranslate');
  assert.equal((await de.textContent('#btnPublish')).trim(), 'Veröffentlichen…');
  assert.equal((await de.textContent('#btnSettings')).trim(), 'Einstellungen');
  assert.match(await de.textContent('.side'), /Signaturen.*Verwaltung.*Felder.*Veröffentlicht.*Verlauf/s);
  assert.match(await de.textContent('#toolbar'), /\+ Feld einfügen/);
  await de.click('a[data-view=fields]');
  assert.match(await de.textContent('#viewFields'), /Ein Feld ist ein Platzhalter.*Ohne Graph-Quelle/s);
  // English can be chosen, and stays chosen.
  await de.click('#btnSettings');
  assert.equal(await de.inputValue('#setLang'), 'de');
  await de.selectOption('#setLang', 'en');
  await Promise.all([de.waitForNavigation(), de.click('#setSave')]);
  await de.waitForSelector('html[data-ready]');
  assert.equal((await de.textContent('#btnPublish')).trim(), 'Publish…');
  assert.equal(await de.getAttribute('html', 'lang'), 'en');
  await de.evaluate(() => localStorage.removeItem('osign.lang'));
  assert.deepEqual(errs, []);
  await deCtx.close();
});

test('users from AD: read, look, and publish exactly that', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('#sigList a');
  await page.click('a[data-view=options]');
  await page.check('input[name=vsrc][value=mixed]');
  // Publishing before reading: the users are not written, and it says so.
  await page.click('#btnPublish');
  await page.waitForSelector('#dlgPublish[open]');
  assert.match(await page.textContent('#pubChecks'), /The users are not written: read and check them under Users first/);
  await page.click('#dlgPublish button[value=cancel]');

  await page.click('a[data-view=users]');
  assert.match(await page.textContent('#usersInfo'), /Not read yet/);
  await page.click('#btnAdRead');
  await page.waitForFunction(() => /Read .*2 users - 2 new/.test(document.querySelector('#usersInfo').textContent));
  const rows = await page.$$eval('#usersRows tr', (trs) => trs.map((tr) => [...tr.cells].map((c) => c.textContent)));
  // The last column: groups (an earlier test made a signature for groups).
  assert.deepEqual(rows[0], ['new', 'Anna Muster', 'anna@example.com', 'Dipl. Ing.', '+41 41 000 00 01', '', '']);
  assert.ok(!fs.existsSync(path.join(folder, 'users')), 'reading writes nothing');

  await page.click('#btnPublish');
  await page.waitForSelector('#dlgPublish[open]');
  assert.match(await page.textContent('#pubChecks'), /The users as read .* are written: 2 new, 0 changed, 0 removed/);
  await page.click('#pubGo');
  await page.waitForFunction(() => /users written/.test(document.querySelector('#pubOk').textContent));
  assert.deepEqual(fs.readdirSync(path.join(folder, 'users')).sort(), ['anna@example.com.json', 'beat@example.com.json']);
  await page.click('#dlgPublish button[value=cancel]');
  assert.deepEqual(errors, []);
});

test('setup assistant: opens by itself until NOBS OSign is set up, and walks through every step', { skip }, async () => {
  // Each step saves on Next: wait for where it leads, or for what it says.
  const next = async (title) => { await page.click('#wizNext'); await page.waitForFunction((x) => document.querySelector('#wizTitle').textContent === x, title); };
  const nextFails = async (rx) => { await page.click('#wizNext'); await page.waitForFunction((x) => new RegExp(x).test(document.querySelector('#wizErr').textContent), rx.source); };
  // Never finished: it opens by itself; Later closes it.
  // Never finished, but set up by hand (app registration and address there): it stays closed.
  await api('settings/save', { settings: { setupDone: '', clientId: '11111111-2222-3333-4444-555555555555', addinUrl: 'https://osign-addin.pages.dev/' } });
  await page.goto(url); await page.waitForSelector('html[data-ready="1"]');
  assert.equal(await page.isVisible('#dlgWizard'), false);
  await api('settings/save', { settings: { setupDone: '', clientId: '', addinUrl: '' } });
  await page.goto(url); await page.reload();
  await page.waitForSelector('#dlgWizard[open]');
  assert.equal(await page.textContent('#wizTitle'), 'Welcome to NOBS OSign');
  await page.click('#wizLater');
  assert.equal(await page.isVisible('#dlgWizard'), false);

  // From the top bar, step by step; each step keeps what it asks for.
  await page.click('#btnWizard');
  // Not set up yet: only the steps reached so far can be jumped to.
  assert.equal(await page.isDisabled('#wizNav button[data-step=host]'), true);
  await next("The server's connections");
  // AD is set up here; SharePoint is the setup script's, a few steps on.
  await page.waitForFunction(() => /SharePoint .* set up by the setup script/.test(document.querySelector('#wizConnList').textContent));
  assert.match(await page.textContent('#wizConnList'), /signing in with Windows accounts – ready/);
  await next('Where the signatures are kept');
  // Nothing given yet (only the saved default): the address is asked for.
  await nextFails(/Paste the address of the folder in SharePoint/);
  const link = async (address, rx) => {
    await page.fill('#wizLink', address); await page.dispatchEvent('#wizLink', 'change');
    await page.waitForFunction((x) => new RegExp(x).test(document.querySelector('#wizLinkNote').textContent), rx.source);
    return [await page.inputValue('#wizSpSite'), await page.inputValue('#wizSpFolder')];
  };
  // A sharing link: without Graph the server cannot open it; the note says what to do.
  await link('https://contoso-my.sharepoint.com/:f:/g/personal/anna_contoso_com/IgDpY5m2CxbbRYgU2eNTaxZP?e=GOPGuO', /Open the link/);
  assert.equal(await page.getAttribute('#wizLinkNote a', 'href'), 'https://contoso-my.sharepoint.com/:f:/g/personal/anna_contoso_com/IgDpY5m2CxbbRYgU2eNTaxZP?e=GOPGuO');
  assert.deepEqual(await link('https://example.com/x', /not an address in SharePoint/), ['', '']);
  // The OneDrive's folder, by its address from the browser: noted, checked once Microsoft 365 is set up.
  assert.deepEqual(await link('https://contoso-my.sharepoint.com/personal/anna_contoso_com/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fanna_contoso_com%2FDocuments%2FOSign', /Noted/), ['contoso-my.sharepoint.com:/personal/anna_contoso_com:', 'OSign']);
  assert.match(await page.textContent('#wizPerm'), /Only you/);
  await next("Where the users' values come from");
  assert.equal((await api('settings/get')).settings.spSite, 'contoso-my.sharepoint.com:/personal/anna_contoso_com:');
  await page.check('input[name=wvsrc][value=graph]');
  await next('Where the add-in is served');
  await page.check('input[name=whost][value=cloudflare]');
  await nextFails(/Enter the address/);
  await page.fill('#wizUrl', 'https://osign.example.com');
  await next('Register the app in Entra ID');
  // The setup script, for the local account admin; and the way by hand.
  assert.equal(await page.isVisible('#wizScriptGet'), true);
  // By hand: folded away under the script.
  assert.equal(await page.isVisible('#wizClientId'), false);
  await page.click('#wizAppManual summary');
  assert.equal(await page.textContent('#wizRedirect'), 'brk-multihub://osign.example.com');
  assert.equal(await page.textContent('#wizPerms'), 'User.Read, Sites.Read.All, GroupMember.Read.All');
  await page.fill('#wizClientId', 'nope');
  await nextFails(/GUID/);
  await page.fill('#wizClientId', '11111111-2222-3333-4444-555555555555');
  await next('Create the add-in files');
  await nextFails(/Create the add-in files first/);
  await page.click('#wizMake');
  await page.waitForFunction(() => /files \(version/.test(document.querySelector('#wizMakeOk').textContent));
  assert.equal(await page.isVisible('#wizServerFiles'), true);
  await next("Into everyone's Outlook");
  assert.equal(await page.textContent('#wizManifest'), 'https://osign.example.com/manifest.xml');
  await next('Ready');
  assert.equal(await page.textContent('#wizNext'), 'Finish');
  await page.click('#wizNext');
  await page.waitForSelector('#dlgWizard', { state: 'hidden' });
  const s = (await api('settings/get')).settings;
  assert.deepEqual([s.setupDone, s.addinHost, s.addinUrl, s.clientId, s.spSite, s.spFolder], ['1', 'cloudflare', 'https://osign.example.com', '11111111-2222-3333-4444-555555555555', 'contoso-my.sharepoint.com:/personal/anna_contoso_com:', 'OSign']);
  // The add-in's files, as downloaded: made for that site.
  const zip = await fetch(base + '/addin.zip', { headers: { Cookie: 'osign_session=' + cookie } });
  assert.equal(zip.status, 200);
  // Finished: it does not open by itself any more.
  await page.reload();
  await page.waitForSelector('html[data-ready]');
  assert.equal(await page.isVisible('#dlgWizard'), false);
  // Set up: any step can be jumped to; going forward keeps what the step asks for.
  await page.click('#btnWizard');
  await page.click('#wizNav button[data-step=host]');
  await page.waitForFunction(() => document.querySelector('#wizTitle').textContent === 'Where the add-in is served');
  assert.equal(await page.getAttribute('#wizNav button[data-step=host]', 'aria-current'), 'step');
  await page.fill('#wizUrl', 'https://osign2.example.com');
  await page.click('#wizNav button[data-step=deploy]');
  await page.waitForFunction(() => document.querySelector('#wizTitle').textContent === "Into everyone's Outlook");
  assert.equal((await api('settings/get')).settings.addinUrl, 'https://osign2.example.com');
  assert.equal(await page.textContent('#wizCount'), 'Step 8 of 9');
  // Set up already: a step that cannot be saved is left as saved, and says so (Next stays on it).
  await page.click('#wizNav button[data-step=host]');
  await page.fill('#wizUrl', '');
  await page.click('#wizNav button[data-step=app]');
  await page.waitForFunction(() => document.querySelector('#wizTitle').textContent === 'Register the app in Entra ID');
  await page.waitForFunction(() => /Not saved: Enter the address/.test(document.body.textContent));
  assert.equal((await api('settings/get')).settings.addinUrl, 'https://osign2.example.com');
  await page.click('#wizNav button[data-step=welcome]');
  await page.click('#wizLater');
  await api('settings/save', { settings: { addinUrl: 'https://osign.example.com' } });
  // Back to the folder the other tests publish to.
  await api('settings/save', { settings: { spSite: 'root', spFolder: 'OSign' } });
  assert.deepEqual(errors, []);
});

test('help: a ? opens a short explanation, above dialogs too, and closes again', { skip }, async () => {
  await page.goto(url); await page.reload();
  // Loaded from the server, not only the page.
  await page.waitForSelector('html[data-ready]');
  await page.click('a[data-view=fields]');
  await page.click('button.help[data-help=hGraph]');
  await page.waitForSelector('#helpPop:popover-open');
  assert.match(await page.textContent('#helpPop'), /as Microsoft Graph names it/);
  // A second click on the same ? closes it.
  await page.click('button.help[data-help=hGraph]');
  await page.waitForFunction(() => !document.querySelector('#helpPop').matches(':popover-open'));
  // In a dialog: shown over it, Esc closes the help, not the dialog.
  await page.click('#btnSettings');
  await page.click('#dlgSettings [data-stab=addin]');
  await page.click('#dlgSettings button.help[data-help=hMobile]');
  await page.waitForSelector('#helpPop:popover-open');
  assert.match(await page.textContent('#helpPop'), /iPhone/);
  assert.equal(await page.evaluate(() => document.elementFromPoint(...(() => { const r = document.querySelector('#helpPop').getBoundingClientRect(); return [r.left + 5, r.top + 5]; })()).closest('#helpPop') !== null), true);
  await page.keyboard.press('Escape');
  await page.waitForFunction(() => !document.querySelector('#helpPop').matches(':popover-open'));
  assert.equal(await page.isVisible('#dlgSettings'), true);
  await page.click('#dlgSettings button[value=cancel]');
  // Every ? has its text, in both languages.
  const missing = await page.evaluate(() => [...document.querySelectorAll('button.help')].map((b) => b.dataset.help).filter((k) => !document.querySelector('[data-help-text="' + k + '"]')));
  assert.deepEqual(missing, []);
  const keys = await page.evaluate(() => [...document.querySelectorAll('[data-help-text]')].map((d) => d.dataset.helpText));
  const i18n = fs.readFileSync(path.join(__dirname, '..', 'editor', 'ui', 'i18n.js'), 'utf8');
  assert.deepEqual(keys.filter((k) => !new RegExp('\\n    ' + k + ': ').test(i18n)), [], 'German help texts');
  assert.deepEqual(errors, []);
});

test('fields: chips with their names in the design view, @@key in the signature', { skip }, async () => {
  await page.goto(url); await page.reload();
  await page.waitForSelector('#sigList a');
  await page.click('#sigList a');
  await page.waitForSelector('#viewSig:not([hidden])');
  const dd = () => page.frameLocator('#design');
  // The starter template's placeholders, shown by name; a link's mailto:@@mail stays as it is.
  await page.waitForFunction(() => document.querySelector('#design').contentDocument.querySelectorAll('span.osf').length > 3);
  const chips = await page.evaluate(() => [...document.querySelector('#design').contentDocument.querySelectorAll('span.osf')].map((c) => c.dataset.f + '=' + c.textContent));
  assert.ok(chips.includes('givenName=Vorname'), chips.join(', '));
  assert.equal(await page.evaluate(() => document.querySelector('#design').contentDocument.body.textContent.includes('@@givenName')), false);
  assert.equal(await page.evaluate(() => document.querySelector('#design').contentDocument.querySelector('a[href^="mailto:"]').getAttribute('href')), 'mailto:@@mail');
  // Inserted from the list: a chip in the design, @@key in the signature.
  await dd().locator('body').click();
  await page.keyboard.press('Control+End');
  await page.selectOption('#toolbar select.field-pick', 'mobile');
  await page.waitForFunction(() => [...document.querySelector('#design').contentDocument.querySelectorAll('span.osf')].pop().dataset.f === 'mobile');
  // Typed by hand: a chip once it is complete.
  await page.keyboard.type(' @@title ');
  await page.waitForFunction(() => [...document.querySelector('#design').contentDocument.querySelectorAll('span.osf')].pop().dataset.f === 'title');
  await page.keyboard.type('@@nope ');
  await page.waitForFunction(() => document.querySelector('#design').contentDocument.querySelector('span.osf.unknown[data-f=nope]'));
  await page.click('.tab[data-mode=html]');
  const src = await page.inputValue('#source');
  // (a space beside a chip is kept as &nbsp;, as browsers do when editing)
  assert.match(src, /@@mobile(&nbsp;| )@@title(&nbsp;| )@@nope/);
  assert.doesNotMatch(src, /osf|contenteditable|data-f/);
  // Removed again, through the HTML tab, and back in the design.
  await page.fill('#source', src.replace(/@@mobile(&nbsp;| )@@title(&nbsp;| )@@nope(&nbsp;| )?/, ''));
  await page.click('.tab[data-mode=design]');
  assert.deepEqual(errors, []);
});


test('campaigns, dark mode, comparing versions, the audit log, the backup and four eyes', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('#sigList a');
  await page.click('#sigList a');
  await page.waitForSelector('#viewSig:not([hidden])');
  // A campaign that is over: said in the list and under the preview.
  await page.click('#sigCampaign summary');
  await page.fill('#sigFrom', '2020-01-01');
  await page.fill('#sigUntil', '2020-01-31');
  await page.waitForFunction(() => /Campaign, not now/.test(document.querySelector('#sigList').textContent));
  await page.waitForFunction(() => /A campaign, not running today/.test(document.querySelector('#pvNote').textContent));
  assert.match(await page.textContent('#campaignSummary'), /2020/);
  await page.fill('#sigFrom', ''); await page.fill('#sigUntil', '');
  await page.waitForFunction(() => !/Campaign/.test(document.querySelector('#sigList').textContent));
  // Dark mode in the preview.
  await page.check('#pvDark');
  await page.waitForFunction(() => /invert\(1\)/.test(document.querySelector('#preview').srcdoc));
  await page.uncheck('#pvDark');

  // A version against the one before it.
  await page.click('#navHistory');
  await page.waitForSelector('#historyRows button');
  await page.click('#historyRows tr:first-child button:has-text("Compare")');
  await page.waitForSelector('#dlgDiff[open] .change');
  assert.match(await page.textContent('#diffChanges'), /Added|Changed/);
  await page.check('#dlgDiff input[value=draft]');
  await page.waitForSelector('#dlgDiff[open] .change, #dlgDiff[open] .muted');
  await page.keyboard.press('Escape'); await page.keyboard.press('Escape');

  // The audit log and the backup, for the local admin.
  await page.click('#btnSettings');
  assert.equal(await page.isVisible('#setAudit'), false, 'on the Server tab');
  await page.click('#dlgSettings [data-stab=server]');
  await page.click('#setAudit');
  await page.waitForFunction(() => /Sign-in/.test(document.querySelector('#auditRows').textContent));
  await page.selectOption('#auditWhat', 'publish');
  await page.waitForFunction(() => /First signatures/.test(document.querySelector('#auditRows').textContent));
  // Searched in what is loaded.
  await page.fill('#auditFind', 'no such thing');
  await page.waitForFunction(() => /Nothing matches/.test(document.querySelector('#auditRows').textContent));
  await page.fill('#auditFind', 'first');
  await page.waitForFunction(() => /First signatures/.test(document.querySelector('#auditRows').textContent));
  assert.match(await page.textContent('#auditCount'), /^1 of \d+$/);
  // A window: minimised to the tray, brought back from it, moved by its title bar, maximised.
  await page.click('#dlgAudit .wc-min');
  await page.waitForSelector('#dlgTray .tray-chip');
  assert.equal(await page.isVisible('#dlgAudit'), false);
  assert.equal(await page.isVisible('#dlgSettings'), false, 'with the settings it was opened from');
  assert.match(await page.textContent('#dlgTray'), /Audit log/);
  await page.click('#dlgTray .tray-open');
  await page.waitForSelector('#dlgAudit[open]');
  assert.equal(await page.isVisible('#dlgSettings'), true);
  assert.equal(await page.isVisible('#dlgTray'), false);
  const before = await page.evaluate(() => document.querySelector('#dlgAudit').getBoundingClientRect().left);
  const bar = await page.$('#dlgAudit h2');
  const bb = await bar.boundingBox();
  await page.mouse.move(bb.x + 40, bb.y + 10); await page.mouse.down();
  await page.mouse.move(bb.x - 60, bb.y + 40, { steps: 4 }); await page.mouse.up();
  const after = await page.evaluate(() => document.querySelector('#dlgAudit').getBoundingClientRect().left);
  assert.ok(Math.abs(after - before + 100) < 3, 'moved 100px left: ' + before + ' → ' + after);
  assert.ok(await page.evaluate(() => JSON.parse(localStorage.getItem('osign-dlg-pos')).dlgAudit.x === -100));
  await page.click('#dlgAudit .wc-max');
  assert.equal(await page.evaluate(() => Math.round(document.querySelector('#dlgAudit').getBoundingClientRect().width)), 1400);
  await page.click('#dlgAudit .wc-max');
  await page.click('#dlgAudit .wc-close');
  assert.equal(await page.isVisible('#dlgAudit'), false);
  await page.click('#setBackup');
  await page.fill('#bakPass', 'a long passphrase'); await page.fill('#bakPass2', 'a long passphrase');
  const [dl] = await Promise.all([page.waitForEvent('download'), page.click('#bakExport')]);
  assert.match(dl.suggestedFilename(), /^nobs-osign-backup-.*\.osignbak$/);
  await page.keyboard.press('Escape');

  // Four eyes: on, asked for here (admin), approved by Anna in her own browser.
  await page.check('#setApproval');
  await page.waitForFunction(() => /On: publishing now waits/.test(document.querySelector('#setApprovalNote').textContent));
  await page.keyboard.press('Escape');
  await page.click('#btnPublish');
  await page.waitForSelector('#dlgPublish[open]');
  assert.equal(await page.textContent('#pubGo'), 'Ask for approval');
  await page.fill('#pubMsg', 'Please check');
  await page.click('#pubGo');
  await page.waitForFunction(() => /Asked for approval/.test(document.querySelector('#pubOk').textContent));
  assert.equal(await page.isVisible('#pubApprove'), false, 'not by the one who asked');
  await page.keyboard.press('Escape');
  assert.match(await page.textContent('#stApproval'), /Waiting for approval/);

  const r = await fetch(base + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'anna', password: 'anna-pw' }) });
  const annaCtx = await browser.newContext({ viewport: { width: 1400, height: 900 } });
  await annaCtx.addCookies([{ name: 'osign_session', value: /osign_session=([^;]+)/.exec(r.headers.get('set-cookie'))[1], url: base }]);
  const anna = await annaCtx.newPage();
  try {
    await anna.goto(url);
    await anna.waitForFunction(() => /asks for approval/.test(document.querySelector('#stApproval').textContent));
    await anna.click('#stApproval');
    await anna.waitForSelector('#dlgPublish[open] #pubApprove:not([hidden])');
    assert.match(await anna.textContent('#pubApprovalWho'), /admin asks for approval.*Please check/);
    await anna.click('#pubApprove');
    await anna.waitForFunction(() => /Approved and published|Nothing had changed/.test(document.querySelector('#pubOk').textContent));
  } finally { await annaCtx.close(); }
  // And off again, for the tests after this one.
  await api('settings/save', { settings: { requireApproval: '' } });
  assert.deepEqual(errors, []);
});

test('Server connections: SharePoint offers the setup script, and says what is missing in the same words as its test', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('#btnSettings');
  await page.click('#dlgSettings [data-stab=server]');
  await page.click('#setConn');
  await page.waitForSelector('#dlgConn[open]');
  await page.click('#dlgConn [data-ctab=graph]');
  assert.equal(await page.isVisible('#connScriptGet'), true);
  assert.equal(await page.isVisible('#cn_graphTenantId'), false, 'by hand only when opened');
  const lead = await page.textContent('#connGraph');
  assert.match(lead, /setup script does it/);
  await page.click('#dlgConn [data-ctest=graph]');
  await page.waitForFunction(() => document.querySelector('#dlgConn [data-ctestout=graph]').textContent && !/Testing/.test(document.querySelector('#dlgConn [data-ctestout=graph]').textContent));
  assert.equal(await page.textContent('#dlgConn [data-ctestout=graph]'), await page.textContent('#connGraph'));
  // The folder the script gives access to: shown, and changed by pasting its address.
  assert.match(await page.textContent('#connSpWhere'), /OSign/);
  await page.fill('#connSpUrl', 'https://contoso.sharepoint.com/sites/Signaturen/Shared%20Documents/Forms/AllItems.aspx');
  await page.click('#connSpSet');
  await page.waitForFunction(() => /names no folder/.test(document.querySelector('#dlgConn [data-ctestout=graph]').textContent));
  await page.fill('#connSpUrl', 'https://contoso-my.sharepoint.com/personal/anna_contoso_com/Documents/OSign');
  await page.click('#connSpSet');
  await page.waitForFunction(() => /personal OneDrive/.test(document.querySelector('#connSpWarn').textContent));
  await page.fill('#connSpUrl', 'https://contoso.sharepoint.com/sites/Signaturen/Shared%20Documents/OSign');
  await page.click('#connSpSet');
  await page.waitForFunction(() => /Saved/.test(document.querySelector('#dlgConn [data-ctestout=graph]').textContent));
  assert.equal(await page.textContent('#connSpWarn'), '');
  const st = (await api('settings/get')).settings;
  assert.deepEqual([st.spSite, st.spFolder], ['contoso.sharepoint.com:/sites/Signaturen:', 'OSign']);
  await api('settings/save', { settings: { spSite: 'root', spFolder: 'OSign' } });
  // A result file: handed over by the browser; one that is not valid (any more) is said so.
  await page.setInputFiles('#connScriptResult', { name: 'nobs-osign-setup-result.json', mimeType: 'application/json', buffer: Buffer.from('{"token":"old"}') });
  await page.waitForFunction(() => /not valid/.test(document.querySelector('#connScriptNote').textContent));
  await page.setInputFiles('#connScriptResult', { name: 'x.json', mimeType: 'application/json', buffer: Buffer.from('not json') });
  await page.waitForFunction(() => /not the setup script's result/.test(document.querySelector('#connScriptNote').textContent));
  // Without the folder and the add-in's address the script cannot be made: the assistant asks.
  await api('settings/save', { settings: { addinUrl: '' } });
  await page.click('#connScriptGet');
  await page.waitForSelector('#connScriptWizard:not([hidden])');
  assert.match(await page.textContent('#connScriptNote'), /first/);
  await page.keyboard.press('Escape'); await page.keyboard.press('Escape');
  assert.deepEqual(errors, []);
});

test('Status: every part the server depends on, tried now; a part opens its tab in Server connections', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('#btnSettings');
  await page.click('#dlgSettings [data-stab=server]');
  await page.waitForFunction(() => document.querySelectorAll('#healthRows .setrow').length >= 8 && !document.querySelector('#healthRows .hdot.run'));
  const rows = await page.$$eval('#healthRows .setrow', (rs) => rs.map((r) => [r.querySelector('b').textContent, r.dataset.state]));
  assert.deepEqual(rows.find((r) => r[0] === 'Active Directory'), ['Active Directory', 'good']);
  assert.deepEqual(rows.find((r) => r[0] === 'Key for passwords'), ['Key for passwords', 'good']);
  assert.equal(rows.find((r) => r[0] === 'Cloudflare Pages')[1], 'off');
  assert.match(await page.textContent('#healthSum'), /not working|to look at|works/);
  // Its button: Server connections on that tab; closed, Status is checked again.
  await page.click('#healthRows .setrow:has(b:text-is("SharePoint (Microsoft 365)")) button');
  await page.waitForSelector('#dlgConn[open]');
  assert.equal(await page.isVisible('#connScriptGet'), true);
  await page.click('#dlgConn .wc-close');
  await page.waitForFunction(() => !document.querySelector('#healthRows .hdot.run'));
  // The Server tab names Active Directory itself.
  await page.click('#dlgSettings [data-stab=server]');
  assert.equal((await page.textContent('#setConn')).trim(), 'Active Directory…');
  await page.click('#dlgSettings button[value=cancel]');
  assert.deepEqual(errors, []);
});

test('suggestions while typing: AD groups anywhere in their name, people for the preview; Enter takes one', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('#sigList a');
  await page.click('#sigList a');
  await page.waitForSelector('#viewSig:not([hidden])');
  // A group by a part of its name; ↓ and Enter take it, with a comma for the next one.
  await page.click('.audience:not(.campaign) summary');
  await page.fill('#sigGroups', 'osign');
  await page.waitForSelector('#sigGroupsSuggest:not([hidden]) .sg-item');
  assert.match(await page.textContent('#sigGroupsSuggest'), /NOBS-OSign-Admins/);
  await page.keyboard.press('ArrowDown'); await page.keyboard.press('Enter');
  assert.equal(await page.inputValue('#sigGroups'), 'NOBS-OSign-Admins, ');
  assert.equal(await page.isVisible('#sigGroupsSuggest'), false);
  await page.fill('#sigGroups', '');
  // A person by a part of the name: picked, looked up, in the preview.
  await page.selectOption('#pvSource', 'ad');
  await page.fill('#pvUser', 'Mus');
  await page.waitForSelector('#pvUserSuggest:not([hidden]) .sg-item');
  assert.match(await page.textContent('#pvUserSuggest'), /Anna Muster.*anna/);
  await page.keyboard.press('ArrowDown'); await page.keyboard.press('Enter');
  assert.equal(await page.inputValue('#pvUser'), 'anna');
  // Anna's own values (her description from AD), and no error for the part typed.
  await page.waitForFunction(() => /Dipl\. Ing\.<\/p>/.test(document.querySelector('#preview').srcdoc));
  assert.doesNotMatch(await page.textContent('#toast'), /No user/);
  await page.selectOption('#pvSource', 'sample');
  // Settings → Info: the version, the licence, and the coffee.
  await page.click('#btnSettings');
  await page.click('#dlgSettings [data-stab=about]');
  assert.match(await page.textContent('#dlgSettings [data-spane=about]'), /NOBS OSign \d+\.\d+\.\d+.*GNU General Public License.*Buy me a coffee/s);
  assert.equal(await page.isVisible('#coffeeBtn'), true, 'the coffee at the top right stays');
  await page.click('#dlgSettings button[value=cancel]');
  assert.deepEqual(errors, []);
});

test('the add-in\'s files older than the server: said in the status bar and under Server, made again with one click', { skip }, async () => {
  // Made "with an older version": their version file says so.
  await api('settings/save', { settings: { addinUrl: 'https://osign.example.com/', clientId: '11111111-2222-3333-4444-555555555555' } });
  const b = await api('addin/build');
  assert.equal(b.ok, true, b.error);
  const ver = path.join(dataDir, 'addin', 'osign-addin.json');
  const j = JSON.parse(fs.readFileSync(ver, 'utf8'));
  fs.writeFileSync(ver, JSON.stringify(Object.assign({}, j, { version: '0.0.1' })));
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.waitForSelector('#stAddin:not([hidden])');
  assert.match(await page.textContent('#stAddin'), /Add-in files from 0\.0\.1/);
  await page.click('#stAddin');
  await page.waitForSelector('#dlgSettings [data-spane=server]:not([hidden])');
  await page.waitForFunction(() => /Made with 0\.0\.1/.test(document.querySelector('#healthRows').textContent));
  await page.click('#healthRows button:has-text("Make again")');
  await page.waitForFunction(() => !/Made with 0\.0\.1/.test(document.querySelector('#healthRows').textContent) && !document.querySelector('#healthRows .hdot.run'));
  assert.equal(await page.isVisible('#stAddin'), false);
  await page.click('#dlgSettings button[value=cancel]');
  // Made again under Settings → Outlook add-in: the status bar says so at once, without a reload.
  const j2 = JSON.parse(fs.readFileSync(ver, 'utf8'));
  fs.writeFileSync(ver, JSON.stringify(Object.assign({}, j2, { version: '0.0.2' })));
  await page.goto(url);
  await page.waitForSelector('#stAddin:not([hidden])');
  await page.click('#btnSettings');
  await page.click('#dlgSettings .tab[data-stab=addin]');
  await page.click('#setAddinBuild');
  await page.waitForSelector('#stAddin', { state: 'hidden' });
  await page.click('#dlgSettings button[value=cancel]');
});

test('right-click on a signature: a menu with what its page offers, by mouse and by keyboard', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  const names = () => page.$$eval('#sigList li a', (as) => as.map((a) => a.childNodes[0].textContent));
  const before = await names();
  assert.ok(before.length >= 1, 'a signature to try it on');
  await page.click('#sigList li:first-child a', { button: 'right' });
  await page.waitForSelector('#ctxMenu:not([hidden])');
  const items = await page.$$eval('#ctxMenu button', (bs) => bs.map((b) => [b.textContent.replace('✓', '').trim(), b.disabled]));
  assert.deepEqual(items.map((x) => x[0]), ['Open', 'Rename', 'Duplicate', 'Default for new mail', 'Default for replies & forwards', 'Move up', 'Move down', 'Delete…']);
  assert.equal(items[5][1], true, 'the first cannot move up');
  // Escape closes it; nothing changes.
  await page.keyboard.press('Escape');
  assert.equal(await page.isHidden('#ctxMenu'), true);
  // Duplicate: one more; the copy goes last, and can then move up.
  await page.click('#sigList li:first-child a', { button: 'right' });
  await page.click('#ctxMenu button:has-text("Duplicate")');
  await page.waitForFunction((n) => document.querySelectorAll('#sigList li a').length === n + 1, before.length);
  const after = await names();
  await page.click('#sigList li:last-child a', { button: 'right' });
  await page.click('#ctxMenu button:has-text("Move up")');
  const moved = await names();
  assert.equal(moved[moved.length - 2], after[after.length - 1]);
  // By keyboard: Shift+F10 on the focused entry, arrows, Enter.
  await page.focus('#sigList li:last-child a');
  await page.keyboard.press('Shift+F10');
  await page.waitForSelector('#ctxMenu:not([hidden])');
  assert.equal(await page.evaluate(() => document.activeElement.textContent.trim()), 'Open');
  await page.keyboard.press('ArrowUp');
  assert.equal(await page.evaluate(() => document.activeElement.textContent.trim()), 'Delete…');
  await page.keyboard.press('Escape');
  // Delete the copy again (the dialog asks first; these tests answer yes).
  const copyName = after[after.length - 1];
  await page.click(`#sigList li a:has-text("${copyName}") >> nth=-1`, { button: 'right' });
  await page.click('#ctxMenu button:has-text("Delete…")');
  await page.click('#askOk');
  await page.waitForFunction((n) => document.querySelectorAll('#sigList li a').length === n, before.length);
});

test('back and forward (mouse side buttons, Alt+arrows): through the pages and signatures opened', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  const first = await page.$eval('#sigList li:first-child a', (a) => a.childNodes[0].textContent);
  await page.click('#sigList li:first-child a');
  await page.click('.nav-list a[data-view=fields]');
  await page.waitForSelector('#viewFields:not([hidden])');
  await page.click('.nav-list a[data-view=options]');
  await page.waitForSelector('#viewOptions:not([hidden])');
  const href = page.url();
  await page.goBack();
  await page.waitForSelector('#viewFields:not([hidden])');
  await page.goBack();
  await page.waitForSelector('#viewSig:not([hidden])');
  assert.equal(await page.inputValue('#sigName'), first);
  await page.goForward();
  await page.waitForSelector('#viewFields:not([hidden])');
  // The address (with its sign-in) stays the same throughout.
  assert.equal(page.url(), href);
});

test('an image uploaded is made as small as a signature needs, in its own type', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('#sigList li:first-child a');
  await page.click('#toolbar button[title=Image]');
  // A photo 2400 pixels wide, with detail (noise), as a JPEG.
  const b64 = await page.evaluate(() => {
    const c = document.createElement('canvas'); c.width = 2400; c.height = 1600;
    const g = c.getContext('2d'), d = g.createImageData(2400, 1600);
    for (let i = 0; i < d.data.length; i += 4) { d.data[i] = (i * 7) % 256; d.data[i + 1] = (i * 13) % 256; d.data[i + 2] = (i * 3) % 256; d.data[i + 3] = 255; }
    g.putImageData(d, 0, 0);
    return c.toDataURL('image/jpeg', 0.95).split(',')[1];
  });
  const big = Buffer.from(b64, 'base64');
  await page.setInputFiles('#imgFile', { name: 'Banner.jpg', mimeType: 'image/jpeg', buffer: big });
  await page.waitForFunction(() => /Image made smaller: .* \(2400 → 1200 pixels wide\)/.test(document.querySelector('#toast').textContent));
  // Uploaded and put in: then it is there to fetch.
  await page.waitForFunction(() => (document.querySelector('#design').contentDocument || document).querySelector('img[src="assets/Banner.jpg"]'));
  const r = await fetch(url.replace(/\/(#.*)?$/, '') + '/assets/Banner.jpg', { headers: { Cookie: (await page.context().cookies()).map((c) => c.name + '=' + c.value).join('; ') } });
  const got = Buffer.from(await r.arrayBuffer());
  assert.equal(got.subarray(0, 3).toString('hex'), 'ffd8ff', 'still a JPEG');
  assert.ok(got.length < big.length, got.length + ' < ' + big.length);
});

test('fields: the default fields restored, fields of your own kept', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('.nav-list a[data-view=fields]');
  await page.waitForSelector('#viewFields:not([hidden])');
  const keys = () => page.$$eval('#fieldRows tr .ph input', (xs) => xs.map((x) => x.value));
  // One of your own, a default one taken away and another changed.
  await page.click('#btnAddField');
  const before = await keys();
  const own = before[before.length - 1];
  // Removing it: asked first when a signature uses it; then it is gone.
  const removeMobile = async () => {
    await page.click('#fieldRows tr:has(input[value="mobile"]) button[aria-label="Remove field @@mobile"]');
    await page.waitForFunction(() => document.querySelector('#dlgAsk[open]') || !document.querySelector('#fieldRows button[aria-label="Remove field @@mobile"]'));
    if (await page.isVisible('#dlgAsk[open]')) await page.click('#askOk');
    await page.waitForFunction(() => !document.querySelector('#fieldRows button[aria-label="Remove field @@mobile"]'));
  };
  await removeMobile();
  // Removed: it can come back at once.
  await page.waitForFunction(() => /Field @@mobile removed/.test(document.querySelector('#toast').textContent));
  await page.click('#toast button');
  assert.ok((await keys()).includes('mobile'));
  await removeMobile();
  await page.fill('#fieldRows tr:first-child td:first-child input', 'Given name (changed)');
  assert.ok(!(await keys()).includes('mobile'));
  await page.click('#btnResetFields');
  await page.click('#askOk');
  await page.waitForFunction(() => /Default fields restored/.test(document.querySelector('#toast').textContent));
  const after = await keys();
  assert.ok(after.includes('mobile'), 'the default field is back');
  assert.ok(after.includes(own), 'your own field stays');
  assert.equal(await page.inputValue('#fieldRows tr:first-child td:first-child input'), 'Vorname');
  // As they were at first now: said as that.
  await page.click('#btnResetFields');
  await page.waitForFunction(() => /as they were at first/.test(document.querySelector('#toast').textContent));
});

test('photos: one per person, uploaded and made small, listed, removed; the preview says when one is missing', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('.nav-list a[data-view=photos]');
  await page.waitForSelector('#viewPhotos:not([hidden])');
  // A file for a person: under the name the signatures take (no rule yet: .jpg).
  await page.fill('#photoWho', 'neu1');
  assert.equal(await page.textContent('#photoName'), '→ photos/neu1.jpg');
  // An image larger than a signature needs.
  const big = await page.evaluate(async () => {
    const c = document.createElement('canvas'); c.width = 1200; c.height = 900;
    const g = c.getContext('2d'); g.fillStyle = '#c33'; g.fillRect(0, 0, 1200, 900);
    const b = await new Promise((r) => c.toBlob(r, 'image/png'));
    return Array.from(new Uint8Array(await b.arrayBuffer()));
  });
  await page.setInputFiles('#photoFile', { name: 'anna.png', mimeType: 'image/png', buffer: Buffer.from(big) });
  await page.waitForFunction(() => /neu1\.jpg/.test(document.querySelector('#photoList').textContent));
  const saved = fs.readFileSync(path.join(folder, 'photos', 'neu1.jpg'));
  assert.equal(saved.subarray(0, 3).toString('hex'), 'ffd8ff');
  // 400 pixels wide at most (the JPEG's SOF0 marker holds height and width).
  const sof = saved.indexOf(Buffer.from([0xff, 0xc0]));
  assert.equal(saved.readUInt16BE(sof + 7), 400);
  // One there already: asked before it is replaced.
  await page.fill('#photoWho', 'mustera');
  await page.setInputFiles('#photoFile', { name: 'anna.png', mimeType: 'image/png', buffer: Buffer.from(big) });
  await page.waitForSelector('#dlgAsk[open]');
  assert.match(await page.textContent('#askText'), /mustera\.jpg is there already/);
  await page.click('#askCancel');
  assert.ok(fs.readFileSync(path.join(folder, 'photos', 'mustera.jpg')).equals(PNG));
  // A web address in the company's network is refused, and said why.
  await page.fill('#photoWho', 'beispielb');
  await page.fill('#photoUrl', 'https://10.0.0.1/b.jpg');
  await page.click('#photoUrlGo');
  await page.waitForFunction(() => /own network/.test(document.querySelector('#photoErr').textContent));
  // A signature with a photo per person: the preview of someone without one says so.
  await page.click('#btnNew');
  await page.waitForSelector('#viewSig:not([hidden])');
  await page.click('.tab[data-mode=html]');
  await page.fill('#source', '<p>@@givenName</p><p><img src="photos/@@sn.jpg" width="100"></p>');
  await page.dispatchEvent('#source', 'input');
  await page.waitForFunction(() => /No photo .*\.jpg in the photos folder/.test(document.querySelector('#pvNote').textContent));
  // Removed again (asked first).
  await page.click('#pvNote a:has-text("Add the photo")');
  await page.waitForSelector('#viewPhotos:not([hidden])');
  assert.match(await page.textContent("#photosRule"), /takes photos\/@@\w+\.jpg/);
  assert.equal(await page.inputValue('#photoWho'), 'Muster');
  await page.click('#photoList button[aria-label="Remove photo neu1.jpg"]');
  await page.click('#askOk');
  await page.waitForFunction(() => !/neu1\.jpg/.test(document.querySelector('#photoList').textContent));
  assert.equal(fs.existsSync(path.join(folder, 'photos', 'neu1.jpg')), false);
  // A neutral placeholder for whoever has none: kept as photos/_default.jpg, not among the people.
  await page.click('#photoDefaultMake');
  await page.waitForFunction(() => !document.querySelector('#photoDefaultImg').hidden);
  const ph = fs.readFileSync(path.join(folder, 'photos', '_default.jpg'));
  assert.equal(ph.subarray(0, 3).toString('hex'), 'ffd8ff');
  assert.doesNotMatch(await page.textContent('#photoList'), /_default/);
  // The server gives it for someone without a photo, as the add-in takes it.
  const viaServer = await fetch(base + '/photos/Muster.jpg', { headers: { Cookie: 'osign_session=' + cookie } });
  assert.equal(viaServer.status, 200);
  assert.ok(Buffer.from(await viaServer.arrayBuffer()).equals(ph));
  // Removed again: nobody gets one.
  await page.click('#photoDefaultOff');
  await page.click('#askOk');
  await page.waitForFunction(() => document.querySelector('#photoDefaultImg').hidden);
  assert.equal(fs.existsSync(path.join(folder, 'photos', '_default.jpg')), false);
  // The signature made here goes again.
  await page.click('#sigList li:last-child a');
  await page.click('#btnDel');
  await page.click('#askOk');
});

test('zoom: the design view and the preview each their own, kept; the HTML tab indented, the signature as it was', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('#btnNew');
  await page.waitForSelector('#viewSig:not([hidden])');
  const html = '<table><tr><td>@@givenName<br>@@sn</td><td></td></tr></table><p>Text</p>';
  await page.click('.tab[data-mode=html]');
  await page.fill('#source', html);
  await page.dispatchEvent('#source', 'input');
  await page.click('.tab[data-mode=design]');
  await page.click('.tab[data-mode=html]');
  assert.equal(await page.inputValue('#source'), '<table>\n  <tbody>\n    <tr>\n      <td>\n        @@givenName<br>\n        @@sn\n      </td>\n      <td></td>\n    </tr>\n  </tbody>\n</table>\n<p>\n  Text\n</p>');
  // Kept as it was written: the indent is only for reading.
  await page.click('.tab[data-mode=design]');
  await page.waitForTimeout(800);
  const draft = JSON.parse((await page.evaluate(() => fetch('/api/draft/get', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: '{}' }).then((r) => r.json()))).configJson);
  assert.ok(draft.signatures.some((s) => s.html === '<table><tbody><tr><td>@@givenName<br>@@sn</td><td></td></tr></tbody></table><p>Text</p>'), 'the signature is stored without the indent');
  // Zoom: + twice in the design view, − once in the preview; Ctrl + wheel too; kept on reload.
  const z = (k) => page.textContent('.zoom[data-zoom=' + k + '] .zpct');
  await page.click('.zoom[data-zoom=design] [data-z="1"]');
  await page.click('.zoom[data-zoom=design] [data-z="1"]');
  // At first 80 %.
  assert.equal(await z('design'), '100 %');
  assert.equal(await page.evaluate(() => document.querySelector('#design').contentDocument.documentElement.style.zoom), '1');
  await page.click('.zoom[data-zoom=preview] [data-z="-1"]');
  assert.equal(await z('preview'), '70 %');
  await page.reload();
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('#sigList li:last-child a');
  assert.equal(await z('design'), '100 %');
  await page.waitForFunction(() => document.querySelector('#preview').contentDocument.documentElement.style.zoom === '0.7');
  await page.hover('#design');
  await page.keyboard.down('Control'); await page.mouse.wheel(0, -100); await page.keyboard.up('Control');
  await page.waitForFunction(() => document.querySelector('.zoom[data-zoom=design] .zpct').textContent === '110 %');
  // Back to 80 % with the percentage.
  await page.click('.zoom[data-zoom=design] .zpct');
  await page.click('.zoom[data-zoom=preview] .zpct');
  assert.equal(await z('design'), '80 %');
  await page.click('#btnDel');
  await page.click('#askOk');
});

test('undo and redo: typing and a field put in, with Ctrl+Z / Ctrl+Y, also when the focus left the editor', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('#btnNew');
  await page.waitForSelector('#viewSig:not([hidden])');
  await page.click('.tab[data-mode=html]');
  await page.fill('#source', '<p>Hello</p>');
  await page.dispatchEvent('#source', 'input');
  await page.click('.tab[data-mode=design]');
  const text = () => page.evaluate(() => document.querySelector('#design').contentDocument.body.innerText.trim());
  await page.frameLocator('#design').locator('p').click();
  await page.keyboard.press('End');
  await page.keyboard.type(' World');
  await page.waitForTimeout(700);
  await page.selectOption('#toolbar select.field-pick', 'mail');
  await page.waitForFunction(() => /Mail/.test(document.querySelector('#design').contentDocument.body.innerText));
  // The focus elsewhere on the page (not in a field): Ctrl+Z still undoes the signature.
  await page.click('.pane-title');
  await page.keyboard.press('Control+z');
  await page.waitForFunction(() => document.querySelector('#design').contentDocument.body.innerText.trim() === 'Hello World');
  await page.keyboard.press('Control+z');
  await page.waitForFunction(() => document.querySelector('#design').contentDocument.body.innerText.trim() === 'Hello');
  await page.keyboard.press('Control+y');
  assert.equal(await text(), 'Hello World');
  // The toolbar's buttons the same.
  await page.click('#toolbar [data-cmd=undo]');
  assert.equal(await text(), 'Hello');
  await page.click('#btnDel');
  await page.click('#askOk');
});

test('placeholders no field has: added as fields from the preview\'s note; the HTML tab formats again', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('#btnNew');
  await page.waitForSelector('#viewSig:not([hidden])');
  await page.click('.tab[data-mode=html]');
  assert.equal(await page.isVisible('#toolbar .html-only'), true);
  await page.fill('#source', '<p>@@company<br>@@streetAdress, @@postalCode @@l</p><table><tr><td>@@telephoneNumber</td></tr></table>');
  await page.dispatchEvent('#source', 'input');
  await page.click('#toolbar .html-only');
  assert.match(await page.inputValue('#source'), /^<p>\n  @@company<br>\n  @@streetAdress, @@postalCode @@l\n<\/p>\n<table>/);
  await page.waitForFunction(() => /Not a field: .*@@company/.test(document.querySelector('#pvNote').textContent));
  await page.click('#pvNote a:has-text("Add them as fields")');
  await page.waitForFunction(() => !/Not a field/.test(document.querySelector('#pvNote').textContent));
  await page.click('.tab[data-mode=design]');
  assert.equal(await page.isVisible('#toolbar .html-only'), false);
  await page.click('.nav-list a[data-view=fields]');
  const rows = await page.$$eval('#fieldRows tr', (trs) => trs.map((tr) => [...tr.querySelectorAll('input')].slice(1, 4).map((i) => i.value)));
  assert.deepEqual(rows.find((r) => r[0] === 'streetAdress'), ['streetAdress', 'streetAddress', 'streetAddress']);
  assert.deepEqual(rows.find((r) => r[0] === 'company'), ['company', 'company', 'companyName']);
  await page.click('#sigList li:last-child a');
  await page.click('#btnDel');
  await page.click('#askOk');
});

test('a link: asked in the editor\'s own dialog (no browser prompt), Enter takes it', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  await page.click('#btnNew');
  await page.waitForSelector('#viewSig:not([hidden])');
  await page.frameLocator('#design').locator('body').click({ position: { x: 60, y: 20 } });
  await page.keyboard.press('Shift+End');
  await page.click('#toolbar button[aria-label="Link"]');
  await page.waitForSelector('#dlgAsk[open] #askInput:not([hidden])');
  await page.fill('#askInput', 'https://www.example.com');
  await page.keyboard.press('Enter');
  await page.waitForFunction(() => !!document.querySelector('#design').contentDocument.querySelector('a[href="https://www.example.com"]'));
  // The next question has no field.
  await page.click('#btnDel');
  assert.equal(await page.isVisible('#askInput'), false);
  await page.click('#askOk');
});

test('several signatures: picked with Ctrl / Shift + click, deleted together (asked first), Undo brings them back', { skip }, async () => {
  await page.goto(url);
  await page.waitForSelector('html[data-ready="1"]');
  for (let i = 0; i < 3; i++) { await page.click('#btnNew'); await page.waitForSelector('#viewSig:not([hidden])'); await page.fill('#sigName', 'Multi ' + i); await page.dispatchEvent('#sigName', 'input'); }
  const names = () => page.$$eval('#sigList li a', (as) => as.map((a) => a.firstChild.textContent));
  const namesAre = (want) => page.waitForFunction((w) => JSON.stringify([...document.querySelectorAll('#sigList li a')].map((a) => a.firstChild.textContent)) === JSON.stringify(w), want);
  await page.waitForFunction(() => /Multi 2/.test(document.querySelector('#sigList').textContent));
  const before = await names();
  await page.click('#sigList li a:has-text("Multi 0")');
  await page.click('#sigList li a:has-text("Multi 2")', { modifiers: ['Shift'] });
  assert.equal(await page.textContent('#sigPickedN'), '3 selected');
  await page.click('#sigList li a:has-text("Multi 1")', { modifiers: ['Control'] });
  assert.equal(await page.textContent('#sigPickedN'), '2 selected');
  await page.click('#sigPickedDel');
  assert.match(await page.textContent('#askText'), /Delete 2 signatures: «Multi 0», «Multi 2»/);
  assert.equal(await page.textContent('#askOk'), 'Delete 2');
  await page.click('#askOk');
  await namesAre(before.filter((n) => n !== 'Multi 0' && n !== 'Multi 2'));
  assert.equal(await page.isVisible('#sigPicked'), false);
  await page.click('#toast button');
  await namesAre(before);
  // All three away again.
  await page.click('#sigList li a:has-text("Multi 0")');
  await page.click('#sigList li a:has-text("Multi 2")', { modifiers: ['Shift'] });
  await page.click('#sigPickedDel');
  await page.click('#askOk');
  await namesAre(before.filter((n) => !/^Multi /.test(n)));
});
