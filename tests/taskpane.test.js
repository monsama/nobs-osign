// node --test tests/taskpane.test.js   (after: node addin/build.mjs; needs Playwright's Chromium)
//
// The built task pane in a browser, with office.js replaced by a stand-in Outlook and the
// published configuration served from memory.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { JSDOM } = require('jsdom');
const core = require('../shared/signature-core.js');

let chromium = null;
try { ({ chromium } = require('playwright')); } catch (e) { /* skipped below */ }
const dist = path.join(__dirname, '..', 'dist');
const skip = (!chromium && 'playwright is not installed') || (!fs.existsSync(path.join(dist, 'taskpane.js')) && 'run node addin/build.mjs first');
const BASE = 'https://monsama.github.io/osign/';

const doc = new JSDOM('').window.document;
const cfg = core.defaultConfig();
cfg.assetBaseUrl = BASE + 'config/';
cfg.signatures = [
  Object.assign(core.newSignature('std', 'Standard'), { defaultNew: true, defaultReply: true, html: '<p><b>@@givenName @@sn</b></p><p>Mobile @@mobile</p>' }),
  Object.assign(core.newSignature('short', 'Short'), { html: '<p>@@givenName</p>' }),
  Object.assign(core.newSignature('it', 'IT only'), { audience: { departments: ['IT'], domains: [] }, html: '<p>IT</p>' })
];
cfg.signatures.forEach((s) => { s.compiled = core.compile(s.html, doc); });

// office.js, as far as the pane uses it. What it is asked to do lands in window.__calls.
const OFFICE_STUB = `
window.__calls = []; window.__roaming = { 'osign.profile': JSON.stringify({ at: Date.now(), values: { givenName: 'Anna', sn: 'Muster', mobile: '', department: 'Sales', mail: 'anna@monsama.ch' } }) };
const ok = (v) => (cb) => cb({ status: 'succeeded', value: v });
window.Office = {
  AsyncResultStatus: { Succeeded: 'succeeded' }, CoercionType: { Html: 'html', Text: 'text' },
  MailboxEnums: { ItemType: { Message: 'message' } },
  onReady: (fn) => setTimeout(fn, 0),
  context: {
    requirements: { isSetSupported: () => false },
    roamingSettings: { get: (k) => window.__roaming[k], set: (k, v) => { window.__roaming[k] = v; }, saveAsync: (cb) => cb && cb({ status: 'succeeded' }) },
    mailbox: {
      userProfile: { displayName: 'Anna Muster', emailAddress: 'anna@monsama.ch' },
      item: {
        itemType: 'message',
        getComposeTypeAsync: (cb) => ok({ composeType: 'newMail', coercionType: 'html' })(cb),
        from: { getAsync: (cb) => ok({ emailAddress: 'anna@monsama.ch' })(cb) },
        disableClientSignatureAsync: (cb) => { window.__calls.push(['disable']); ok()(cb); },
        body: { setSignatureAsync: (c, o, cb) => { window.__calls.push(['set', c, o.coercionType]); ok()(cb); } }
      }
    }
  }
};`;

test('the pane lists what is open to the user, inserts, and saves the choice', { skip }, async () => {
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage();
    const errors = [];
    page.on('pageerror', (e) => errors.push(e.message));
    await page.route('https://appsforoffice.microsoft.com/**', (r) => r.fulfill({ contentType: 'text/javascript', body: OFFICE_STUB }));
    await page.route(BASE + '**', (r) => {
      const u = new URL(r.request().url());
      if (u.pathname === '/osign/config/signatures.json') return r.fulfill({ contentType: 'application/json', body: JSON.stringify(cfg) });
      const file = path.join(dist, u.pathname.replace(/^\/osign\//, ''));
      if (fs.existsSync(file)) return r.fulfill({ path: file });
      return r.fulfill({ status: 404, body: '' });
    });
    await page.goto(BASE + 'taskpane.html');
    await page.waitForSelector('.card');
    assert.deepEqual(await page.$$eval('.card b', (b) => b.map((x) => x.textContent)), ['Standard', 'Short']);
    assert.match(await page.textContent('#who'), /Anna Muster · anna@monsama\.ch/);
    assert.deepEqual(await page.$$eval('#prefNew option', (o) => o.map((x) => x.textContent)), ['Company default (Standard)', 'No signature', 'Standard', 'Short']);

    await page.click('.card:nth-child(2) button');
    await page.waitForFunction(() => window.__calls.some((c) => c[0] === 'set'));
    const calls = await page.evaluate(() => window.__calls);
    assert.deepEqual(calls.find((c) => c[0] === 'set'), ['set', '<p>Anna</p>', 'html']);

    await page.selectOption('#prefReply', 'short');
    await page.waitForFunction(() => /Saved/.test(document.querySelector('#msg').textContent));
    assert.deepEqual(JSON.parse(await page.evaluate(() => window.__roaming['osign.prefs'])), { new: '', reply: 'short' });
    assert.deepEqual(errors, []);
  } finally { await browser.close(); }
});
