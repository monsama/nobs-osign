// node --test tests/core.test.js
const test = require('node:test');
const assert = require('node:assert/strict');
const { JSDOM } = require('jsdom');
const core = require('../shared/signature-core.js');
// A signature for the entire organization (a new one is not assigned to anybody).
const sigAll = (id, name) => { const s = core.newSignature(id, name); s.audience.everyone = true; return s; };

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

test('CI-Sign: an empty field and a banner in one row, lines inside a <font>: the field\'s line goes, the banner stays', () => {
  const html = '<table><tr><td>@@givenName</td></tr><tr><td><font size="2">@@info<br><br><img src="assets/banner.jpg" width="601"><br></font></td></tr></table>';
  const out = run(html, { givenName: 'Anna', info: '' });
  assert.match(out, /banner\.jpg/);
  assert.doesNotMatch(out, /@@info/);
  // With a note in the field, both.
  const out2 = run(html, { givenName: 'Anna', info: 'Away until Monday' });
  assert.match(out2, /Away until Monday[\s\S]*banner\.jpg/);
  // A row of single lines still goes as a whole.
  assert.equal(text(run('<table><tr><td>Mobile</td><td>@@mobile</td></tr></table>', { mobile: '' })), '');
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

test('a photo for someone with no login is the placeholder photo', () => {
  const out = core.render(core.compile('<p><img src="photos/@@sAMAccountName.jpg" width="100"></p>', doc), { sAMAccountName: '' });
  assert.match(out, /src="photos\/_default\.jpg"/);
  const ed = core.render(core.compile('<p><img src="photos/@@sAMAccountName.jpg"></p>', doc), { sAMAccountName: '' }, { assetBase: '/' });
  assert.match(ed, /src="\/photos\/_default\.jpg"/);
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
    Object.assign(sigAll('std', 'Standard'), { defaultNew: true }),
    Object.assign(sigAll('short', 'Short'), { defaultReply: true }),
    Object.assign(sigAll('it', 'IT'), { audience: { departments: ['Informatik'], domains: [] } }),
    Object.assign(sigAll('other', 'Other firm'), { defaultNew: true, audience: { departments: [], domains: ['other.ch'] } })
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
    Object.assign(sigAll('std', 'Standard'), { defaultNew: true }),
    Object.assign(sigAll('sales', 'Sales'), { defaultNew: true, audience: { departments: [], domains: [], groups: ['SG-Sales', 'SG-Field'] } })
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

test('value source: each field from Entra ID, from AD, or from nowhere', () => {
  const cfg = core.defaultConfig();
  const src = () => Object.fromEntries(cfg.fields.map((f) => [f.key, core.fieldSource(cfg, f)]));
  // New setups take both: Entra ID where a field has a Graph source, AD for the rest.
  assert.equal(core.valueSource(cfg), 'mixed');
  // A configuration that names none (made before) stays with Entra ID alone.
  assert.equal(core.valueSource({ options: {} }), 'graph');
  cfg.options.valueSource = 'graph';
  assert.equal(src().description, 'none');
  assert.equal(src().givenName, 'graph');
  assert.equal(core.usesAdExport(cfg), false);
  cfg.options.valueSource = 'mixed';
  assert.equal(src().description, 'ad');
  assert.equal(src().homePhone, 'ad');
  assert.equal(src().givenName, 'graph');
  assert.deepEqual(core.adFields(cfg).map((f) => f.key), ['description', 'homePhone', 'info']);
  cfg.options.valueSource = 'ad';
  assert.ok(Object.values(src()).every((v) => v === 'ad'));
  assert.equal(core.graphFields(cfg).length, 0);
  cfg.options.valueSource = 'mixed';
  assert.deepEqual(
    Object.entries(core.mergeValues(cfg, { givenName: 'Anna', description: 'from Graph?' }, { description: 'Dipl. Ing.', givenName: 'from AD?' })).filter(([, v]) => v),
    [['givenName', 'Anna'], ['description', 'Dipl. Ing.']]);
  assert.equal(core.userFileName(' Anna.Muster@Example.COM '), 'anna.muster@example.com.json');
  assert.equal(core.userFileName("o'neil+x@example.com"), 'o_neil_x@example.com.json');
  assert.equal(core.userFileName('not-a-mail'), '');
});

test('the images of a signature are listed, a stray % in a name as it is', () => {
  const list = core.localImages('<img src="assets/logo%20blue.png"><img src="photos/100%.jpg"><img src="https://x/y.png">');
  assert.deepEqual(list.map((i) => i.dir + '/' + i.name), ['assets/logo blue.png', 'photos/100%.jpg']);
});

test('campaigns: only on their days, and then they win as default; before and after, the usual one', () => {
  const cfg = { schemaVersion: 1, options: {}, signatures: [
    { id: 'std', name: 'Standard', defaultNew: true, audience: { groups: ['sg-sales'] } },
    { id: 'xmas', name: 'Weihnachten', defaultNew: true, validFrom: '2026-12-01', validUntil: '2026-12-31', audience: {} }
  ] };
  const on = (d) => core.pick(cfg, {}, 'a@x.ch', 'new', null, ['sg-sales'], new Date(d + 'T12:00:00')).id;
  assert.equal(on('2026-11-30'), 'std');
  assert.equal(on('2026-12-01'), 'xmas', 'the first day is included');
  assert.equal(on('2026-12-31'), 'xmas', 'the last day too');
  assert.equal(on('2027-01-01'), 'std');
  // Picked by the user in the pane, outside its days: the default instead.
  assert.equal(core.pick(cfg, {}, '', 'new', { new: 'xmas' }, ['sg-sales'], new Date('2027-02-01T12:00:00')).id, 'std');
  assert.equal(core.isCampaign(cfg.signatures[0]), false);
  assert.equal(core.inPeriod({ validFrom: '2026-12-01' }, new Date('2030-01-01T12:00:00')), true, 'open-ended');
  assert.deepEqual(core.validate({ schemaVersion: 1, signatures: [{ id: 'a', name: 'A', validFrom: '2026-12-31', validUntil: '2026-12-01' }] }), ['Signature 1: the campaign ends before it starts.']);
  assert.deepEqual(core.validate({ schemaVersion: 1, signatures: [{ id: 'a', name: 'A', validFrom: '1.12.2026' }] }), ['Signature 1: the campaign\'s days are dates like 2026-12-01.']);
});

test('variants: the user\'s own text only in editable fields, only of signatures still open', () => {
  const c = core.defaultConfig();
  c.fields.push({ key: 'greeting', label: 'Grussformel', ad: '', graph: '', sample: '', userEditable: true });
  c.signatures = [Object.assign(sigAll('std', 'Standard'), { defaultNew: true, html: '<p>@@greeting</p><p>@@givenName</p>' })];
  const variants = [
    { id: 'my-a', name: 'Herzlich', base: 'std', values: { greeting: 'Herzliche Grüsse', givenName: 'Hacker' } },
    { id: 'my-b', name: 'Gone', base: 'old', values: { greeting: 'x' } },
    { id: 'evil id', name: 'Bad', base: 'std', values: {} }
  ];
  const full = core.withVariants(c, variants);
  assert.deepEqual(full.signatures.map((s) => s.id), ['std', 'my-a']);
  const v = full.signatures[1];
  assert.equal(v.variantOf, 'std');
  assert.equal(v.defaultNew, false);
  // A field the editor did not open stays the user's own value.
  assert.deepEqual(core.valuesFor(v, { givenName: 'Anna', greeting: '' }), { givenName: 'Anna', greeting: 'Herzliche Grüsse' });
  assert.deepEqual(core.editableIn(c, c.signatures[0]), ['greeting']);
  // Picked by the user, it is used; not picked, the company default.
  assert.equal(core.pick(full, {}, 'a@b.ch', 'new', { new: 'my-a' }, []).id, 'my-a');
  assert.equal(core.pick(full, {}, 'a@b.ch', 'new', {}, []).id, 'std');
  // No editable field, or no choice for users: no variants.
  c.options.allowUserChoice = false;
  assert.equal(core.withVariants(c, variants).signatures.length, 1);
  c.options.allowUserChoice = true; c.fields.pop();
  assert.equal(core.withVariants(c, variants).signatures.length, 1);
});

test('a field for a placeholder from CI-Sign: as known (AD and Graph), whatever its case, a misspelling too', () => {
  assert.deepEqual(core.fieldFor('telephoneNumber'), { key: 'telephoneNumber', label: 'Telefon', ad: 'telephoneNumber', graph: 'businessPhones[0]', sample: '+41 44 123 45 00' });
  assert.equal(core.fieldFor('wWWHomepage').ad, 'wWWHomePage');
  assert.equal(core.fieldFor('wWWHomepage').key, 'wWWHomepage');
  assert.deepEqual([core.fieldFor('streetAdress').ad, core.fieldFor('streetAdress').graph], ['streetAddress', 'streetAddress']);
  assert.deepEqual([core.fieldFor('l').label, core.fieldFor('l').graph], ['Ort', 'city']);
  assert.deepEqual(core.fieldFor('extensionAttribute3'), { key: 'extensionAttribute3', label: 'extensionAttribute3', ad: 'extensionAttribute3', graph: '', sample: '' });
});

test('lines: a label row with several empty values goes whole; a label stays with its value; a field\'s line in a CI-Sign <font> goes, its banner stays', () => {
  const r = (h, v) => core.render(core.compile(h, doc), v);
  assert.equal(r('<table><tr><td>Phone</td><td>@@telephoneNumber<br>@@mobile</td></tr><tr><td>X</td></tr></table>', { telephoneNumber: '', mobile: '' }).replace(/<\/?tbody>/g, ''), '<table><tr><td>X</td></tr></table>');
  assert.equal(r('<p>@@givenName<br><b>Tel</b> <span>@@telephoneNumber<br>Mobile @@mobile</span></p>', { givenName: 'A', telephoneNumber: '', mobile: '079' }), '<p>A<br><b>Tel</b> <span><br>Mobile 079</span></p>');
  const banner = r('<table><tr><td><font>@@info<br><br><img src="assets/b.jpg"></font></td></tr></table>', { info: '' });
  assert.doesNotMatch(banner, /@@info/);
  assert.match(banner, /<td><font><br><img src="assets\/b\.jpg"><\/font><\/td>/);
  assert.equal(r('<p><a href="mailto:@@mail">Mail<br>me</a></p>', { mail: '' }), '');
});

test('a signature of one\'s own: only while the editor allows it, and not too long', () => {
  const c = core.defaultConfig();
  c.signatures = [Object.assign(sigAll('std', 'Standard'), { html: '<p>@@givenName</p>', compiled: '<p>@@givenName</p>' })];
  const v = [{ id: 'my-1', name: 'Mine', base: 'std', own: '<p>Anna, Team Lead</p>' }];
  assert.equal(core.withVariants(c, v).signatures.length, 1, 'off: not there');
  c.options.allowOwnSignatures = true;
  const s = core.withVariants(c, v).signatures[1];
  assert.deepEqual([s.id, s.own, s.compiled, s.variantOf], ['my-1', true, '<p>Anna, Team Lead</p>', 'std']);
  assert.equal(core.render(s.compiled, { givenName: 'X' }), '<p>Anna, Team Lead</p>');
  assert.equal(core.withVariants(c, [Object.assign({}, v[0], { own: 'x'.repeat(core.VARIANT_OWN_MAX + 1) })]).signatures.length, 1);
  // Users may not choose: none of their own either.
  c.options.allowUserChoice = false;
  assert.equal(core.ownAllowed(c), false);
});

test('who gets a signature: a new one is not assigned (nobody), the entire organization when ticked, one from before as it was (everyone)', () => {
  const s = core.newSignature('n', 'New');
  assert.equal(core.isAvailable(s, {}, 'a@b.ch', []), false);
  assert.equal(core.forNobody(s), true);
  s.audience.everyone = true;
  assert.equal(core.isAvailable(s, {}, 'a@b.ch', []), true);
  // From before the flag: everyone, as it was.
  assert.equal(core.isAvailable({ id: 'o', audience: { departments: [], domains: [], groups: [] } }, {}, 'a@b.ch', []), true);
  // A group given: the flag says nothing.
  const g = Object.assign(core.newSignature('g', 'G'), { audience: { departments: [], domains: [], groups: ['SG'], everyone: false } });
  assert.equal(core.isAvailable(g, {}, 'a@b.ch', ['sg']), true);
  assert.equal(core.forNobody(g), false);
});

test('building blocks: {{key}} put in on a line of its own or in the text, unknown ones left; not put in twice', () => {
  const blocks = [{ key: 'disclaimer', name: 'Disclaimer', html: '<p>Vertraulich @@mail</p>' }, { key: 'x', html: '{{disclaimer}}' }];
  assert.equal(core.expandBlocks('<p>Hi</p><p>{{Disclaimer}}</p>', blocks), '<p>Hi</p><p>Vertraulich @@mail</p>');
  assert.equal(core.expandBlocks('<div> {{disclaimer}} </div>', blocks), '<p>Vertraulich @@mail</p>');
  assert.equal(core.expandBlocks('a {{nope}} b', blocks), 'a {{nope}} b');
  assert.equal(core.expandBlocks('{{x}}', blocks), '{{disclaimer}}', 'a block in a block stays as typed (no loops)');
  assert.deepEqual(core.blocksIn('{{a}} {{b}} {{a}} @@c'), ['a', 'b']);
  assert.deepEqual(core.placeholdersIn(core.expandBlocks('{{disclaimer}}', blocks)), ['mail']);
  const c = core.defaultConfig(); c.blocks = [{ key: 'a' }, { key: 'A' }, { key: '1x' }];
  assert.equal(core.validate(c).filter((e) => /building block/i.test(e)).length, 2);
});

test('language variants: the default in the person\'s language wins; another language only when there is nothing else', () => {
  const cfg = { options: {}, signatures: [
    { id: 'de', name: 'Haupt', defaultNew: true, audience: { everyone: true } },
    { id: 'en', name: 'Main EN', language: 'en', defaultNew: true, audience: { everyone: true } },
    { id: 'fr', name: 'Réponse', language: 'fr', defaultReply: true, audience: { everyone: true } }] };
  assert.equal(core.pick(cfg, { preferredLanguage: 'en-US' }, 'a@x.ch', 'new').id, 'en');
  assert.equal(core.pick(cfg, { preferredLanguage: 'de-CH' }, 'a@x.ch', 'new').id, 'de');
  assert.equal(core.pick(cfg, {}, 'a@x.ch', 'new').id, 'de');
  assert.equal(core.pick(cfg, { preferredLanguage: 'de' }, 'a@x.ch', 'reply').id, 'fr', 'the only default for replies');
  assert.equal(core.langOf({ preferredLanguage: 'DE_ch' }), 'de');
  assert.equal(core.fieldFor('preferredLanguage').graph, 'preferredLanguage');
});

test('compiled says what html says: the same text, no address or tag of its own (the server refuses otherwise)', () => {
  const doc = new JSDOM('').window.document;
  const html = '<p>Hi <b>@@givenName</b></p><table><tr><td>Tel @@mobile</td></tr></table><p><a href="mailto:@@mail">@@mail</a><img src=\'assets/logo.png\' alt="x &amp; y"></p>';
  const compiled = core.compile(html, doc);
  assert.equal(core.compiledMatches(html, compiled), true);
  assert.equal(core.compiledMatches(html, ''), true, 'none: the add-in compiles nothing, it takes html');
  assert.equal(core.compiledMatches(html, compiled.replace('mailto:@@mail', 'https://evil.example')), false);
  assert.equal(core.compiledMatches(html, compiled + '<img src="https://t.example/p.gif">'), false);
  assert.equal(core.compiledMatches(html, compiled.replace('Tel', 'Call us at 0900')), false);
  assert.equal(core.compiledMatches(html, compiled.replace('<b>', '<script>x()</script><b>')), false);
});
