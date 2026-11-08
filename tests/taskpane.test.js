// node --test tests/taskpane.test.js   (needs Playwright's Chromium)
//
// The add-in's pane in a browser: office.js replaced by a stand-in Outlook, Graph and SharePoint
// by tests/support/fake-graph.js (behind the browser's own fetch), the sign-in by the stub.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { JSDOM } = require('jsdom');
const core = require('../shared/signature-core.js');
// A signature for the entire organization (a new one is not assigned to anybody).
const sigAll = (id, name) => { const s = core.newSignature(id, name); s.audience.everyone = true; return s; };
const buildTestAddin = require('./support/build-test-addin.js');
const fakeGraph = require('./support/fake-graph.js');

let chromium = null;
try { ({ chromium } = require('playwright')); } catch (e) { /* skipped below */ }
const skip = !chromium && 'playwright is not installed';
const HOST = 'https://addin.example/';
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');

const doc = new JSDOM('').window.document;
const cfg = core.defaultConfig();
cfg.updated = '2026-10-30T13:02:00Z';
cfg.signatures = [
  Object.assign(sigAll('std', 'Standard'), { defaultNew: true, defaultReply: true, html: '<p><b>@@givenName @@sn</b></p><p>Mobile @@mobile</p><p><img src="assets/logo.png"></p>' }),
  Object.assign(sigAll('short', 'Short'), { html: '<p>@@givenName</p>' }),
  Object.assign(sigAll('it', 'IT only'), { audience: { departments: ['IT'], domains: [], groups: [] }, html: '<p>IT</p>' })
];
cfg.signatures.forEach((s) => { s.compiled = core.compile(s.html, doc); });

// office.js, as far as the pane uses it. What it is asked to do lands in window.__calls.
const OFFICE_STUB = `
window.__calls = []; window.__roaming = {};
const ok = (v) => (cb) => cb({ status: 'succeeded', value: v });
window.Office = {
  AsyncResultStatus: { Succeeded: 'succeeded' }, CoercionType: { Html: 'html', Text: 'text' },
  MailboxEnums: { ItemType: { Message: 'message' } },
  onReady: (fn) => setTimeout(fn, 0),
  context: {
    requirements: { isSetSupported: () => true },
    roamingSettings: { get: (k) => window.__roaming[k], set: (k, v) => { window.__roaming[k] = v; }, remove: (k) => { delete window.__roaming[k]; },
      // window.__refuse(settings), when a test sets it: true has the save refused (as a mailbox that is full).
      saveAsync: (cb) => cb && cb({ status: window.__refuse && window.__refuse(window.__roaming) ? 'failed' : 'succeeded' }) },
    mailbox: {
      userProfile: { displayName: 'Anna Muster', emailAddress: 'anna@example.com' },
      item: {
        itemType: 'message',
        getComposeTypeAsync: (cb) => ok({ composeType: 'newMail', coercionType: 'html' })(cb),
        from: { getAsync: (cb) => ok({ emailAddress: 'anna@example.com' })(cb) },
        disableClientSignatureAsync: (cb) => { window.__calls.push(['disable']); ok()(cb); },
        getAttachmentsAsync: (cb) => ok([])(cb),
        addFileAttachmentFromBase64Async: (b, n, o, cb) => { window.__calls.push(['attach', n]); ok('a1')(cb); },
        body: { setSignatureAsync: (c, o, cb) => { window.__calls.push(['set', c, o.coercionType]); ok()(cb); } }
      }
    }
  }
};`;

async function openPane(browser, dir, graph, displayLanguage, platform, init) {
  const page = await browser.newPage();
  if (init) await page.addInitScript(init);
  await page.route('https://appsforoffice.microsoft.com/**', (r) => r.fulfill({ contentType: 'text/javascript', body: OFFICE_STUB.replace("context: {", "context: { displayLanguage: '" + displayLanguage + "', diagnostics: { platform: '" + (platform || 'OfficeOnline') + "' },") }));
  const viaFake = async (r) => {
    const res = await graph.fetch(r.request().url(), { headers: { Authorization: r.request().headers().authorization } });
    await r.fulfill({ status: res.status, body: Buffer.from(await res.arrayBuffer()), headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Authorization', 'Content-Type': 'application/json' } });
  };
  await page.route('https://graph.microsoft.com/**', viaFake);
  await page.route('https://download.example/**', viaFake);
  await page.route(HOST + '**', (r) => {
    const file = path.join(dir, new URL(r.request().url()).pathname.slice(1));
    return fs.existsSync(file) ? r.fulfill({ path: file }) : r.fulfill({ status: 404, body: '' });
  });
  await page.goto(HOST + 'taskpane.html');
  await page.waitForSelector('.card');
  return page;
}

test('the pane speaks German in a German Outlook', { skip }, async () => {
  const dir = await buildTestAddin();
  const graph = fakeGraph({ files: { 'signatures.json': JSON.stringify(cfg), 'assets/logo.png': PNG }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    const page = await openPane(browser, dir, graph, 'de-CH');
    assert.equal(await page.textContent('h1'), 'OSign Signaturen');
    assert.match(await page.textContent('#choice'), /Für neue Mails.*Für Antworten und Weiterleitungen/s);
    assert.deepEqual(await page.$$eval('.card button', (b) => b.map((x) => x.textContent)), ['Einfügen', 'Einfügen']);
    assert.equal((await page.$$eval('#prefNew option', (o) => o.map((x) => x.textContent)))[0], 'Firmenstandard (Standard)');
    assert.equal(await page.textContent('#refresh'), 'Meine Angaben neu lesen');
    // Which add-in, and since when the signatures are as they are.
    assert.match(await page.textContent('#ver'), /^OSign 9\.9\.9 · Signaturen veröffentlicht 30\.10\.(20)?26/);
  } finally { await browser.close(); }
});

test('the pane lists what is open to the user, with images, inserts, and saves the choice', { skip }, async () => {
  const dir = await buildTestAddin();
  const graph = fakeGraph({
    files: { 'signatures.json': JSON.stringify(cfg), 'assets/logo.png': PNG },
    me: { givenName: 'Anna', surname: 'Muster', mobilePhone: '', department: 'Sales', mail: 'anna@example.com' }
  });
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('https://appsforoffice.microsoft.com/**', (r) => r.fulfill({ contentType: 'text/javascript', body: OFFICE_STUB }));
    // Graph and SharePoint downloads go to the stand-in, with the headers the page sent.
    const viaFake = async (r) => {
      const res = await graph.fetch(r.request().url(), { headers: { Authorization: r.request().headers().authorization } });
      const body = Buffer.from(await res.arrayBuffer());
      await r.fulfill({ status: res.status, body, headers: { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'Authorization', 'Content-Type': 'application/json' } });
    };
    await page.route('https://graph.microsoft.com/**', viaFake);
    await page.route('https://download.example/**', viaFake);
    await page.route(HOST + '**', (r) => {
      const file = path.join(dir, new URL(r.request().url()).pathname.slice(1));
      return fs.existsSync(file) ? r.fulfill({ path: file }) : r.fulfill({ status: 404, body: '' });
    });
    await page.goto(HOST + 'taskpane.html');
    await page.waitForSelector('.card', { timeout: 8000 }).catch(async (e) => { throw new Error('No cards. Message: ' + await page.textContent('#msg') + ' | ' + errors.join(' / ')); });
    assert.deepEqual(await page.$$eval('.card b', (b) => b.map((x) => x.textContent)), ['Standard', 'Short']);
    assert.match(await page.textContent('#who'), /Anna Muster · anna@example\.com/);
    assert.deepEqual(await page.$$eval('#prefNew option', (o) => o.map((x) => x.textContent)), ['Company default (Standard)', 'No signature', 'Standard', 'Short']);

    // The preview shows the logo, embedded.
    const frame = page.frameLocator('.card:nth-child(1) iframe');
    await frame.locator('img').waitFor();
    assert.match(await frame.locator('img').getAttribute('src'), /^data:image\/png;base64,/);

    await page.click('.card:nth-child(2) button');
    await page.waitForFunction(() => window.__calls.some((c) => c[0] === 'set'));
    assert.deepEqual((await page.evaluate(() => window.__calls)).find((c) => c[0] === 'set'), ['set', '<p>Anna</p>', 'html']);

    await page.click('.card:nth-child(1) button');
    await page.waitForFunction(() => window.__calls.filter((c) => c[0] === 'set').length === 2);
    const calls = await page.evaluate(() => window.__calls);
    assert.deepEqual(calls.filter((c) => c[0] === 'attach'), [['attach', 'osign-assets-logo.png']]);
    assert.match(calls.filter((c) => c[0] === 'set')[1][1], /<img src="cid:osign-assets-logo\.png">/);

    await page.selectOption('#prefReply', 'short');
    await page.waitForFunction(() => /Saved/.test(document.querySelector('#msg').textContent));
    assert.deepEqual(JSON.parse(await page.evaluate(() => window.__roaming['osign.prefs'])), { new: '', reply: 'short' });
    assert.deepEqual(errors, []);
  } finally { await browser.close(); }
});

test('own variants: made from a signature with editable fields, chosen, inserted, deleted', { skip }, async () => {
  const c = JSON.parse(JSON.stringify(cfg));
  c.fields.push({ key: 'greeting', label: 'Greeting', ad: '', graph: '', sample: '', userEditable: true });
  c.signatures[1].html = '<p>@@greeting</p><p>@@givenName</p><p><a href="mailto:@@greeting">@@greeting</a></p>';
  c.signatures[1].compiled = core.compile(c.signatures[1].html, doc);
  const dir = await buildTestAddin();
  const graph = fakeGraph({ files: { 'signatures.json': JSON.stringify(c), 'assets/logo.png': PNG }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    const page = await openPane(browser, dir, graph, 'en-US');
    const errors = []; page.on('pageerror', (e) => errors.push(e.message));
    // Only "Short" uses an editable field.
    assert.deepEqual(await page.$$eval('.card .actions', (a) => a.map((x) => x.textContent)), ['Insert', 'Own variantInsert']);
    // Which fields are the user's to change: on the card, in the form, marked in its preview.
    assert.match(await page.textContent('.card:nth-child(2) .card-head'), /Yours to change: Greeting/);
    await page.click('.card:nth-child(2) button.secondary');
    assert.match(await page.textContent('.variant small'), /Yours to change: Greeting: in the boxes below, or right in the signature/);
    // Written right in the signature: the box above follows.
    await page.waitForSelector('.own-edit.live .fld[data-k=greeting]');
    assert.equal(await page.$$eval('.own-edit.live .fld[data-k=greeting]', (x) => x.length), 2);
    assert.equal(await page.$$eval('.own-edit.live a[href]', (x) => x.length), 0);
    await page.click('.own-edit.live .fld[data-k=greeting] >> nth=0');
    await page.keyboard.type('Hi');
    assert.equal(await page.inputValue('.variant label:nth-of-type(2) input'), 'Hi');
    assert.equal(await page.textContent('.own-edit.live .fld[data-k=greeting] >> nth=1'), 'Hi');
    await page.fill('.variant label:nth-of-type(1) input', 'Warm');
    await page.fill('.variant label:nth-of-type(2) input', 'Kind regards');
    await page.click('.variant button:not(.secondary)');
    await page.waitForFunction(() => document.querySelectorAll('.card').length === 3);
    const saved = JSON.parse(await page.evaluate(() => window.__roaming['osign.variants']));
    assert.equal(saved.length, 1);
    assert.deepEqual([saved[0].name, saved[0].base, saved[0].values], ['Warm', 'short', { greeting: 'Kind regards' }]);
    assert.ok((await page.$$eval('#prefNew option', (o) => o.map((x) => x.textContent))).includes('Warm'));
    // Inserted with the user's text.
    await page.click('.card:nth-child(3) button:not(.secondary)');
    await page.waitForFunction(() => window.__calls.some((x) => x[0] === 'set'));
    assert.equal((await page.evaluate(() => window.__calls)).find((x) => x[0] === 'set')[1], '<p>Kind regards</p><p>Anna</p><p><a href="mailto:Kind regards">Kind regards</a></p>');
    // Picked for new mail, then deleted (asked twice): the choice goes back to the company's.
    await page.selectOption('#prefNew', saved[0].id);
    await page.waitForFunction(() => /Saved/.test(document.querySelector('#msg').textContent));
    const del = page.locator('.card:nth-child(3) button.secondary', { hasText: 'Delete' });
    await del.click();
    await page.locator('.card:nth-child(3) button', { hasText: 'Really delete?' }).click();
    await page.waitForFunction(() => document.querySelectorAll('.card').length === 2);
    assert.deepEqual(JSON.parse(await page.evaluate(() => window.__roaming['osign.variants'])), []);
    assert.equal(JSON.parse(await page.evaluate(() => window.__roaming['osign.prefs'])).new, '');
    assert.deepEqual(errors, []);
  } finally { await browser.close(); }
});

test('classic Outlook: the pane keeps what goes in by itself in the mailbox, its images as small as shown', { skip }, async () => {
  const dir = await buildTestAddin();
  const graph = fakeGraph({ files: { 'signatures.json': JSON.stringify(cfg), 'assets/logo.png': PNG }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    const page = await openPane(browser, dir, graph, 'en-US', 'PC');
    await page.waitForFunction(() => !!window.__roaming['osign.classic'], null, { timeout: 8000 });
    const kept = JSON.parse(await page.evaluate(() => window.__roaming['osign.classic']));
    assert.deepEqual(kept.cfg.signatures.map((x) => x.id), ['std']);
    assert.ok(kept.images['assets/logo.png'] && kept.images['assets/logo.png'].b64.length > 0);
    // The fields, as far as reading the details needs them: without, they would be read empty.
    assert.deepEqual(kept.cfg.fields.map((f) => f.key), cfg.fields.map((f) => f.key));
    assert.equal(kept.cfg.fields.some((f) => 'label' in f || 'sample' in f), false);
    // All the add-in keeps in the mailbox stays below Outlook's 32 KB.
    const all = await page.evaluate(() => Object.entries(window.__roaming).reduce((n, [k, v]) => n + k.length + String(v).length, 0));
    assert.ok(all < 32 * 1024, all);
    // A campaign running now: kept with what goes in once it ends.
    const withCampaign = JSON.parse(JSON.stringify(cfg));
    withCampaign.signatures.push(Object.assign(sigAll('camp', 'Campaign'), { defaultNew: true, validFrom: '2020-01-01', validUntil: '2099-12-31', html: '<p>Fair</p>', compiled: core.compile('<p>Fair</p>', doc) }));
    const graph2 = fakeGraph({ files: { 'signatures.json': JSON.stringify(withCampaign), 'assets/logo.png': PNG }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
    const page3 = await openPane(browser, dir, graph2, 'en-US', 'PC');
    await page3.waitForFunction(() => !!window.__roaming['osign.classic'], null, { timeout: 8000 });
    assert.deepEqual(JSON.parse(await page3.evaluate(() => window.__roaming['osign.classic'])).cfg.signatures.map((x) => x.id), ['camp', 'std']);
    // Not in new Outlook or on the web: nothing kept there.
    const page2 = await openPane(browser, dir, graph, 'en-US');
    await page2.waitForTimeout(500);
    assert.equal(await page2.evaluate(() => window.__roaming['osign.classic']), undefined);
  } finally { await browser.close(); }
});

test('a signature of one\'s own (when the editor allows it): all of it changed in the pane, kept, inserted with the company\'s image', { skip }, async () => {
  const c = JSON.parse(JSON.stringify(cfg));
  c.options.allowOwnSignatures = true;
  const dir = await buildTestAddin();
  const graph = fakeGraph({ files: { 'signatures.json': JSON.stringify(c), 'assets/logo.png': PNG }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    const page = await openPane(browser, dir, graph, 'en-US');
    const errors = []; page.on('pageerror', (e) => errors.push(e.message));
    // Every company signature offers it (no editable fields needed).
    await page.click('.card:nth-child(1) button.secondary:has-text("Own variant")');
    await page.waitForFunction(() => /Anna Muster/.test(document.querySelector('.own-edit').textContent));
    // Filled in with her details, the logo shown.
    assert.equal(await page.$$eval('.own-edit img[data-osign-src="assets/logo.png"]', (x) => x.length), 1);
    await page.click('.own-edit b');
    await page.keyboard.press('End');
    await page.keyboard.type(', Team Lead');
    await page.click('.variant button:not(.secondary):has-text("Save")');
    await page.waitForFunction(() => /Your own signature/.test(document.body.textContent));
    const saved = JSON.parse(await page.evaluate(() => window.__roaming['osign.variants']));
    assert.equal(saved.length, 1);
    assert.match(saved[0].own, /Anna Muster, Team Lead/);
    assert.match(saved[0].own, /<img src="assets\/logo\.png"/);
    assert.doesNotMatch(saved[0].own, /data:|data-osign-src|@@/);
    // Inserted: as she wrote it, the logo embedded.
    const card = page.locator('.card', { hasText: 'Your own signature' });
    await card.locator('button:not(.secondary)', { hasText: 'Insert' }).click();
    await page.waitForFunction(() => window.__calls.some((x) => x[0] === 'set'));
    const set = (await page.evaluate(() => window.__calls)).find((x) => x[0] === 'set')[1];
    assert.match(set, /Anna Muster, Team Lead/);
    assert.match(set, /cid:osign-assets-logo\.png/);
    // Edited again: the same box, with what she wrote.
    await card.locator('button.secondary', { hasText: 'Edit' }).click();
    await page.waitForFunction(() => /Team Lead/.test(document.querySelector('.own-edit').textContent));
    assert.deepEqual(errors, []);
  } finally { await browser.close(); }
});

test('classic Outlook: an image too large for the mailbox is made smaller step by step for that copy, so it still goes in', { skip }, async () => {
  const dir = await buildTestAddin();
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    // A banner hard to make small (noise), 900 pixels wide, shown 300 wide.
    const maker = await browser.newPage();
    const jpg = Buffer.from(await maker.evaluate(() => { const c = document.createElement('canvas'); c.width = 900; c.height = 500; const g = c.getContext('2d'); const d = g.createImageData(900, 500); for (let i = 0; i < d.data.length; i += 4) { d.data[i] = Math.random() * 255; d.data[i + 1] = Math.random() * 255; d.data[i + 2] = Math.random() * 255; d.data[i + 3] = 255; } g.putImageData(d, 0, 0); return c.toDataURL('image/jpeg', 0.9).split(',')[1]; }), 'base64');
    await maker.close();
    const c = JSON.parse(JSON.stringify(cfg));
    c.signatures[0].html = '<p><b>@@givenName @@sn</b></p><p><img src="assets/logo.png"></p><p><img src="assets/banner.jpg" width="300"></p>';
    c.signatures[0].compiled = core.compile(c.signatures[0].html, doc);
    const graph = fakeGraph({ files: { 'signatures.json': JSON.stringify(c), 'assets/logo.png': PNG, 'assets/banner.jpg': jpg }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
    const page = await openPane(browser, dir, graph, 'en-US', 'PC');
    await page.waitForFunction(() => !!window.__roaming['osign.classic'], null, { timeout: 15000 });
    const kept = JSON.parse(await page.evaluate(() => window.__roaming['osign.classic']));
    assert.ok(jpg.length * 4 / 3 > 30 * 1024, 'the banner alone is more than the room');
    assert.ok(kept.images['assets/banner.jpg'], 'the banner is in the copy, made smaller');
    assert.ok(kept.images['assets/banner.jpg'].b64.length < jpg.length * 4 / 3 / 3);
    assert.ok(kept.images['assets/logo.png'], 'the logo too');
    const all = await page.evaluate(() => Object.entries(window.__roaming).reduce((n, [k, v]) => n + k.length + String(v).length, 0));
    assert.ok(all < 32 * 1024, all);
  } finally { await browser.close(); }
});

test('the steps an image is made smaller by for classic Outlook are the same in the add-in and in the editor\'s estimate', () => {
  const steps = (file) => (/CLASSIC_SQUEEZE = (\[[^\]]+\])/.exec(fs.readFileSync(path.join(__dirname, '..', file), 'utf8')) || [])[1];
  assert.ok(steps('addin/src/common.js'));
  assert.equal(steps('editor/ui/app/10-publish.js'), steps('addin/src/common.js'));
});

test('a choice the mailbox does not keep: said so, not "Saved", and the list shows what is in force', { skip }, async () => {
  const dir = await buildTestAddin();
  const graph = fakeGraph({ files: { 'signatures.json': JSON.stringify(cfg), 'assets/logo.png': PNG }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    const page = await openPane(browser, dir, graph, 'de-CH');
    await page.evaluate(() => { window.__refuse = () => true; });
    await page.selectOption('#prefReply', 'short');
    await page.waitForFunction(() => /nicht gespeichert/.test(document.querySelector('#msg').textContent));
    assert.match(await page.getAttribute('#msg', 'class'), /error/);
    assert.equal(await page.evaluate(() => window.__roaming['osign.prefs']), undefined);
    assert.equal(await page.inputValue('#prefReply'), '');
  } finally { await browser.close(); }
});

test('classic Outlook: a new choice is kept for its event handler too', { skip }, async () => {
  const dir = await buildTestAddin();
  const graph = fakeGraph({ files: { 'signatures.json': JSON.stringify(cfg), 'assets/logo.png': PNG }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    const page = await openPane(browser, dir, graph, 'en-US', 'PC');
    await page.waitForFunction(() => !!window.__roaming['osign.classic'], null, { timeout: 8000 });
    assert.deepEqual(JSON.parse(await page.evaluate(() => window.__roaming['osign.classic'])).cfg.signatures.map((x) => x.id), ['std']);
    await page.selectOption('#prefReply', 'short');
    await page.waitForFunction(() => /"short"/.test(window.__roaming['osign.classic'] || ''), null, { timeout: 8000 });
    assert.deepEqual(JSON.parse(await page.evaluate(() => window.__roaming['osign.classic'])).cfg.signatures.map((x) => x.id), ['std', 'short']);
  } finally { await browser.close(); }
});

test('classic Outlook: a copy the mailbox refuses is tried again without its largest image, and the refused one is not left behind', { skip }, async () => {
  const dir = await buildTestAddin();
  const graph = fakeGraph({ files: { 'signatures.json': JSON.stringify(cfg), 'assets/logo.png': PNG }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    // The mailbox refuses whatever carries the logo (as one that counts more than the add-in does).
    const page = await openPane(browser, dir, graph, 'en-US', 'PC', "window.__refuse = (r) => /assets\\/logo\\.png\":\\{\"b64/.test(r['osign.classic'] || '');");
    await page.waitForFunction(() => !!window.__roaming['osign.classic'], null, { timeout: 8000 });
    const kept = JSON.parse(await page.evaluate(() => window.__roaming['osign.classic']));
    assert.deepEqual(kept.cfg.signatures.map((x) => x.id), ['std']);
    assert.deepEqual(kept.images, {});
    await page.waitForFunction(() => /1 image\(s\) did not fit/.test(document.querySelector('#msg').textContent));
    // What is saved after it is not refused with it.
    await page.selectOption('#prefNew', 'short');
    await page.waitForFunction(() => !!window.__roaming['osign.prefs']);
    assert.deepEqual(JSON.parse(await page.evaluate(() => window.__roaming['osign.prefs'])), { new: 'short', reply: '' });
    assert.doesNotMatch(await page.textContent('#msg'), /did not save/);
  } finally { await browser.close(); }
});

test('a signature put into the pane to be changed: nothing in it runs, no address it loads from that is not one', { skip }, async () => {
  const c = JSON.parse(JSON.stringify(cfg));
  c.options.allowOwnSignatures = true;
  c.signatures[0].html = '<p><b>@@givenName @@sn</b></p><p><img src="' + HOST + 'missing.png" onerror="window.__xss=1"><a href="java&#9;script:window.__xss=2">site</a></p><p><img src="assets/logo.png"></p>';
  c.signatures[0].compiled = core.compile(c.signatures[0].html, doc);
  const dir = await buildTestAddin();
  const graph = fakeGraph({ files: { 'signatures.json': JSON.stringify(c), 'assets/logo.png': PNG }, me: { givenName: 'Anna', surname: 'Muster', department: 'Sales', mail: 'anna@example.com' } });
  const browser = await chromium.launch(process.env.OSIGN_TEST_CHROMIUM ? { executablePath: process.env.OSIGN_TEST_CHROMIUM } : {});
  try {
    const page = await openPane(browser, dir, graph, 'en-US');
    await page.click('.card:nth-child(1) button.secondary:has-text("Own variant")');
    await page.waitForFunction(() => /Anna Muster/.test(document.querySelector('.own-edit').textContent));
    await page.waitForTimeout(300);
    assert.equal(await page.evaluate(() => window.__xss), undefined);
    assert.equal(await page.$$eval('.own-edit [onerror]', (x) => x.length), 0);
    assert.deepEqual(await page.$$eval('.own-edit a', (x) => x.map((a) => a.getAttribute('href'))), [null]);
    // The company's image is still there, shown.
    assert.equal(await page.$$eval('.own-edit img[data-osign-src="assets/logo.png"]', (x) => x.length), 1);
  } finally { await browser.close(); }
});
