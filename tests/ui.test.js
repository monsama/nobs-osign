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
  browser = await chromium.launch();
  context = await signedIn();
  page = await context.newPage();
  page.on('pageerror', (e) => errors.push(e.message));
  page.on('console', (m) => { if (m.type() === 'error' && !/Failed to load resource/.test(m.text())) errors.push(m.text()); });
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
  await page.click('#sigList a');
  await page.click('.tab[data-mode=html]');
  assert.match(await page.inputValue('#source'), /Abteilung @@abteilung/);
  assert.match(await page.textContent('#status'), /Unpublished|Saving/);
});

test('import: UTF-16 templates with their images, text-block files left out, groups set', { skip }, async () => {
  const utf16 = (s) => Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(s, 'utf16le')]);
  const tpl = '<!DOCTYPE html><html><head><title></title></head><body><table><tbody><tr>' +
    '<td><img src="images/logo.png" width="90" /><br /><img src="images/@@sAMAccountName.jpg" width="100" /></td>' +
    '<td><strong>@@givenName @@sn</strong><br />@@title<br />Mobile @@mobile</td></tr></tbody></table>@@info{##notice}{##legal}</body></html>';
  const blocks = '<html><body>{##notice}{dats 01.01.2020}{date 31.01.2020}<p>Closed in January</p>{end ##notice}{##legal}<p>Contoso AG</p>{end ##legal}</body></html>';
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
  assert.match(items[0].meta, /photo per person images\/@@sAMAccountName\.jpg → photos\\@@sAMAccountName\.jpg/);
  assert.match(items[1].meta, /text blocks/);
  assert.match(items[0].meta, /text blocks notice, legal \(with days: a campaign each\)/);
  await page.click('#impGo');
  await page.waitForSelector('#dlgImport:not([open])', { state: 'attached' });
  assert.equal(await page.inputValue('#sigName'), 'Vertrieb');
  // The text block without days in the template; the one with days a campaign of its own.
  let imported = '';
  for (let i = 0; i < 50 && !/Vertrieb – notice/.test(imported); i++) { await new Promise((r) => setTimeout(r, 100)); imported = (await api('draft/get')).configJson || ''; }
  const sigs = JSON.parse(imported).signatures;
  const vt = sigs.find((x) => x.name === 'Vertrieb'), camp = sigs.find((x) => x.name === 'Vertrieb – notice');
  assert.match(vt.html, /Contoso AG/);
  assert.doesNotMatch(vt.html, /Closed in January|\{##/);
  assert.match(camp.html, /Closed in January.*Contoso AG|Contoso AG.*Closed in January/s);
  assert.deepEqual([camp.validFrom, camp.validUntil], ['2020-01-01', '2020-01-31']);
  await page.click('#sigList a:has-text("Vertrieb – notice")');
  await page.click('#btnDel');
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
  assert.match(checks, /GroupMember\.Read\.All/);
  assert.match(checks, /Photos per person \(Vertrieb\) come from photos\\<username>\.jpg/);
  await page.click('#dlgPublish button[value=cancel]');
  assert.deepEqual(errors, []);
  // Only the editor's own design view tried the photo with its placeholder (it shows as missing there).
  assert.deepEqual([...new Set(notFound)], ['/photos/@@sAMAccountName.jpg']);
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
  assert.match(await de.textContent('.side'), /Signaturen.*Einrichtung.*Felder/s);
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
  assert.deepEqual(rows[0], ['new', 'anna@example.com', 'Dipl. Ing.', '+41 41 000 00 01', '', '']);
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
  await api('settings/save', { settings: { setupDone: '', clientId: '' } });
  await page.goto(url); await page.reload();
  await page.waitForSelector('#dlgWizard[open]');
  assert.equal(await page.textContent('#wizTitle'), 'Welcome to NOBS OSign');
  await page.click('#wizLater');
  assert.equal(await page.isVisible('#dlgWizard'), false);

  // From the top bar, step by step; each step keeps what it asks for.
  await page.click('#btnWizard');
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
  await page.click('#dlgSettings button.help[data-help=hSpSite]');
  await page.waitForSelector('#helpPop:popover-open');
  assert.match(await page.textContent('#helpPop'), /sharing link/);
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
  await page.click('#setAudit');
  await page.waitForFunction(() => /signin/.test(document.querySelector('#auditRows').textContent));
  await page.selectOption('#auditWhat', 'publish');
  await page.waitForFunction(() => /First signatures/.test(document.querySelector('#auditRows').textContent));
  await page.keyboard.press('Escape');
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
  await page.click('#setConn');
  await page.waitForSelector('#dlgConn[open]');
  await page.click('#dlgConn [data-ctab=graph]');
  assert.equal(await page.isVisible('#connScriptGet'), true);
  assert.equal(await page.isVisible('#cn_graphTenantId'), false, 'by hand only when opened');
  const lead = await page.textContent('#connGraph');
  assert.match(lead, /setup script below/);
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
  // Without the folder and the add-in's address the script cannot be made: the assistant asks.
  await api('settings/save', { settings: { addinUrl: '' } });
  await page.click('#connScriptGet');
  await page.waitForSelector('#connScriptWizard:not([hidden])');
  assert.match(await page.textContent('#connScriptNote'), /first/);
  await page.keyboard.press('Escape'); await page.keyboard.press('Escape');
  assert.deepEqual(errors, []);
});
