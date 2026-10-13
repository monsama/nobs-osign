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
const buildTestAddin = require('./support/build-test-addin.js');
const fakeGraph = require('./support/fake-graph.js');

let chromium = null;
try { ({ chromium } = require('playwright')); } catch (e) { /* skipped below */ }
const skip = !chromium && 'playwright is not installed';
const HOST = 'https://addin.example/';
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');

const doc = new JSDOM('').window.document;
const cfg = core.defaultConfig();
cfg.signatures = [
  Object.assign(core.newSignature('std', 'Standard'), { defaultNew: true, defaultReply: true, html: '<p><b>@@givenName @@sn</b></p><p>Mobile @@mobile</p><p><img src="assets/logo.png"></p>' }),
  Object.assign(core.newSignature('short', 'Short'), { html: '<p>@@givenName</p>' }),
  Object.assign(core.newSignature('it', 'IT only'), { audience: { departments: ['IT'], domains: [], groups: [] }, html: '<p>IT</p>' })
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
    roamingSettings: { get: (k) => window.__roaming[k], set: (k, v) => { window.__roaming[k] = v; }, saveAsync: (cb) => cb && cb({ status: 'succeeded' }) },
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

async function openPane(browser, dir, graph, displayLanguage) {
  const page = await browser.newPage();
  await page.route('https://appsforoffice.microsoft.com/**', (r) => r.fulfill({ contentType: 'text/javascript', body: OFFICE_STUB.replace("context: {", "context: { displayLanguage: '" + displayLanguage + "',") }));
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
  const browser = await chromium.launch();
  try {
    const page = await openPane(browser, dir, graph, 'de-CH');
    assert.equal(await page.textContent('h1'), 'Signaturen');
    assert.match(await page.textContent('#choice'), /Für neue Mails.*Für Antworten und Weiterleitungen/s);
    assert.deepEqual(await page.$$eval('.card button', (b) => b.map((x) => x.textContent)), ['Einfügen', 'Einfügen']);
    assert.equal((await page.$$eval('#prefNew option', (o) => o.map((x) => x.textContent)))[0], 'Firmenstandard (Standard)');
    assert.equal(await page.textContent('#refresh'), 'Meine Angaben neu lesen');
  } finally { await browser.close(); }
});

test('the pane lists what is open to the user, with images, inserts, and saves the choice', { skip }, async () => {
  const dir = await buildTestAddin();
  const graph = fakeGraph({
    files: { 'signatures.json': JSON.stringify(cfg), 'assets/logo.png': PNG },
    me: { givenName: 'Anna', surname: 'Muster', mobilePhone: '', department: 'Sales', mail: 'anna@example.com' }
  });
  const browser = await chromium.launch();
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
