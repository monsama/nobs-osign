// node --test tests/core.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const { JSDOM } = require('jsdom');
const core = require('../shared/signature-core.js');

const doc = new JSDOM('<!doctype html><body></body>').window.document;
const run = (html, values, opts) => core.render(core.compile(html, doc), values, opts);
// Rendered HTML read back as text, so the tests say what a reader of the mail sees.
const text = (html) => { const d = doc.createElement('div'); d.innerHTML = html; return d.textContent.replace(/\s+/g, ' ').trim(); };

test('placeholders are filled and escaped', () => {
  const out = run('<p>@@givenName @@sn</p>', { givenName: 'Anna', sn: '<b>&Co' });
  assert.equal(out, '<p>Anna &lt;b&gt;&amp;Co</p>');
});

test('a longer name is not read as a shorter one', () => {
  assert.equal(run('<p>@@mailNickname</p>', { mail: 'x' }), '<p>@@mailNickname</p>');
  assert.deepEqual(core.placeholdersIn('@@mail @@mailNickname @@mail'), ['mail', 'mailNickname']);
});

test('a paragraph whose fields are all empty is removed', () => {
  const out = run('<p>@@givenName</p><p>Mobile @@mobile</p><p>Fixed text</p>', { givenName: 'Anna', mobile: '' });
  assert.equal(out, '<p>Anna</p><p>Fixed text</p>');
});

test('a line stays when one of its fields has a value', () => {
  const out = run('<p>@@homePhone / @@mobile</p>', { homePhone: '', mobile: '079' });
  assert.equal(text(out), '/ 079');
});

test('lines split by <br> are judged each on its own', () => {
  const out = run('<p>@@givenName @@sn<br>Tel @@homePhone<br>Mobile @@mobile<br>Fixed</p>', { givenName: 'A', sn: 'B', homePhone: '', mobile: '079' });
  assert.equal(out, '<p>A B<br>Mobile 079<br>Fixed</p>');
});

test('a table row goes, never a single cell', () => {
  const html = '<table><tbody><tr><td>Mobile</td><td>@@mobile</td></tr><tr><td>Mail</td><td>@@mail</td></tr></tbody></table>';
  const out = run(html, { mobile: '', mail: 'a@b.ch' });
  assert.equal(text(out), 'Maila@b.ch');
  assert.ok(!/<td>Mobile/.test(out));
});

test('placeholders in attributes are filled and count for the line', () => {
  const out = run('<p><a href="mailto:@@mail">@@mail</a></p><p><a href="tel:@@mobile">call</a></p>', { mail: 'a"b@c.ch', mobile: '' });
  assert.equal(out, '<p><a href="mailto:a&quot;b@c.ch">a&quot;b@c.ch</a></p>');
});

test('nested lines: the outer line is judged by its own fields only', () => {
  const html = '<div>@@givenName<p>@@mobile</p></div>';
  assert.equal(run(html, { givenName: 'A', mobile: '' }), '<div>A</div>');
  assert.equal(run(html, { givenName: '', mobile: '079' }), '<div><p>079</p></div>');
});

test('removeEmptyLines false keeps every line', () => {
  assert.equal(run('<p>Mobile @@mobile</p>', { mobile: '' }, { removeEmptyLines: false }), '<p>Mobile </p>');
});

test('unknown placeholders stay as they are', () => {
  assert.equal(run('<p>@@nothing</p>', {}), '<p>@@nothing</p>');
});

test('compile is idempotent and strips scripts', () => {
  const once = core.compile('<p>@@mail</p><script>alert(1)</script>', doc);
  assert.equal(core.compile(once, doc), once);
  assert.ok(!/script/i.test(once));
});

test('images get the published address', () => {
  const out = core.render('<img src="assets/logo.png"><img src="https://x/y.png">', {}, { assetBase: 'https://monsama.github.io/osign' });
  assert.equal(out, '<img src="https://monsama.github.io/osign/assets/logo.png"><img src="https://x/y.png">');
});

test('plain text version', () => {
  const t = core.htmlToText(core.render(core.compile('<p>Anna &amp; Co<br>@@mobile</p><table><tr><td>Mail</td><td>@@mail</td></tr></table>', doc), { mobile: '', mail: 'a@b.ch' }));
  assert.equal(t, 'Anna & Co\nMail\ta@b.ch');
});

test('pick: user choice, then the default, filtered by audience', () => {
  const cfg = core.defaultConfig();
  cfg.signatures = [
    Object.assign(core.newSignature('std', 'Standard'), { defaultNew: true }),
    Object.assign(core.newSignature('short', 'Short'), { defaultReply: true }),
    Object.assign(core.newSignature('it', 'IT'), { audience: { departments: ['Informatik'], domains: [] } }),
    Object.assign(core.newSignature('other', 'Other firm'), { defaultNew: true, audience: { departments: [], domains: ['other.ch'] } })
  ];
  const me = { department: 'informatik ' };
  assert.equal(core.pick(cfg, me, 'a@monsama.ch', 'new', null).id, 'std');
  assert.equal(core.pick(cfg, me, 'a@monsama.ch', 'reply', null).id, 'short');
  assert.equal(core.pick(cfg, me, 'a@monsama.ch', 'new', { new: 'it' }).id, 'it');
  assert.equal(core.pick(cfg, { department: 'Sales' }, 'a@monsama.ch', 'new', { new: 'it' }).id, 'std');
  assert.equal(core.pick(cfg, me, 'a@monsama.ch', 'reply', { reply: 'none' }), null);
  assert.deepEqual(core.available(cfg, me, 'x@other.ch').map(s => s.id), ['std', 'short', 'it', 'other']);
  assert.equal(core.pick(cfg, me, 'x@other.ch', 'new', null).id, 'other');
  cfg.options.allowUserChoice = false;
  assert.equal(core.pick(cfg, me, 'a@monsama.ch', 'new', { new: 'it' }).id, 'std');
});

test('values from Graph, with paths', () => {
  const fields = [
    { key: 'title', graph: 'jobTitle' },
    { key: 'description', graph: 'onPremisesExtensionAttributes.extensionAttribute1' },
    { key: 'phone', graph: 'businessPhones[1]' },
    { key: 'first', graph: 'businessPhones' },
    { key: 'info', graph: '' }
  ];
  const user = { jobTitle: 'Boss', onPremisesExtensionAttributes: { extensionAttribute1: 'Dr.' }, businessPhones: ['1', '2'] };
  assert.deepEqual(core.valuesFromGraph(fields, user), { title: 'Boss', description: 'Dr.', phone: '2', first: '1', info: '' });
  assert.deepEqual(core.graphSelect(fields), ['jobTitle', 'onPremisesExtensionAttributes', 'businessPhones']);
});

test('validate', () => {
  const cfg = core.defaultConfig();
  assert.deepEqual(core.validate(cfg), []);
  cfg.signatures.push(Object.assign(core.newSignature('a b'), { name: ' ' }));
  cfg.fields.push({ key: 'mail' });
  const errs = core.validate(cfg);
  assert.equal(errs.length, 3);
});

test('groups: a signature for groups is open to their members only, and wins as default', () => {
  const cfg = core.defaultConfig();
  cfg.signatures = [
    Object.assign(core.newSignature('std', 'Standard'), { defaultNew: true }),
    Object.assign(core.newSignature('sales', 'Sales'), { defaultNew: true, audience: { departments: [], domains: [], groups: ['SG-Sales', 'SG-Field'] } })
  ];
  assert.deepEqual(core.groupsUsed(cfg), ['sg-sales', 'sg-field']);
  assert.equal(core.pick(cfg, {}, 'a@x.ch', 'new', null, ['sg-field']).id, 'sales');
  assert.equal(core.pick(cfg, {}, 'a@x.ch', 'new', null, ['Other']).id, 'std');
  // Groups not known (not readable): nothing limited to groups.
  assert.equal(core.pick(cfg, {}, 'a@x.ch', 'new', null, null).id, 'std');
  assert.deepEqual(core.available(cfg, {}, 'a@x.ch', ['SG-SALES']).map((s) => s.id), ['std', 'sales']);
});

test('images no mail program can load are left out', () => {
  const html = '<p><img src="images/logo.gif"><img src="images/@@sAMAccountName.jpg"><img src="assets/@@nope.png"><img src="assets/a.png"><img src="https://x/b.png"><img src="cid:c"></p>';
  assert.equal(core.render(html, { sAMAccountName: 'x' }, { assetBase: 'https://h/' }), '<p><img src="https://h/assets/a.png"><img src="https://x/b.png"><img src="cid:c"></p>');
  assert.equal(core.render('<img src="images/a.gif">', {}, { dropBrokenImages: false }), '<img src="images/a.gif">');
});

test('signature files are read in their own encoding', () => {
  const text = '<p>Grüsse @@givenName</p>';
  const le = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(text, 'utf16le')]);
  const be = Buffer.concat([Buffer.from([0xfe, 0xff]), Buffer.from(text, 'utf16le').swap16()]);
  const noBom = Buffer.from(text, 'utf16le');
  const cp1252 = Buffer.from('<meta charset="windows-1252"><p>Gr\xfcsse</p>', 'latin1');
  const bare1252 = Buffer.from('<p>Gr\xfcsse</p>', 'latin1');
  assert.equal(core.decodeHtmlFile(le), text);
  assert.equal(core.decodeHtmlFile(be), text);
  assert.equal(core.decodeHtmlFile(noBom), text);
  assert.equal(core.decodeHtmlFile(Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(text)])), text);
  assert.match(core.decodeHtmlFile(cp1252), /Grüsse/);
  assert.equal(core.decodeHtmlFile(bare1252), '<p>Grüsse</p>');
  assert.equal(core.assetName('C:\\x\\Logo Neu (2).JPG'), 'Logo-Neu-2.jpg');
});
