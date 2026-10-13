// node --test tests/addin.test.js
//
// Runs the add-in's event handler the way classic Outlook on Windows does: a bare JavaScript
// runtime with no window, no document, no localStorage and no btoa, and an Office object that
// records what the handler asks of it. Graph and SharePoint are tests/support/fake-graph.js; the
// sign-in is the stand-in in tests/support/auth-stub.js. Everything else is the real code.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { JSDOM } = require('jsdom');
const core = require('../shared/signature-core.js');
const buildTestAddin = require('./support/build-test-addin.js');
const fakeGraph = require('./support/fake-graph.js');

const doc = new JSDOM('').window.document;
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');
const JPG = Buffer.from('ffd8ffe000104a46494600010100000100010000ffd9', 'hex');
let bundle;
test.before(async () => { bundle = fs.readFileSync(path.join(await buildTestAddin(), 'launchevent.js'), 'utf8'); });

function sig(id, name, html, extra) {
  const s = Object.assign(core.newSignature(id, name), { html }, extra || {});
  s.compiled = core.compile(s.html, doc);
  return s;
}
function makeConfig(edit) {
  const c = core.defaultConfig();
  c.signatures = [
    sig('std', 'Standard', '<p><b>@@givenName @@sn</b></p><p>Mobile @@mobile</p><p><img src="assets/logo.png" width="90"><img src="photos/@@sAMAccountName.jpg" width="100"></p>', { defaultNew: true }),
    sig('short', 'Short', '<p>@@givenName</p>', { defaultReply: true }),
    sig('other', 'Other', '<p>Other @@mail</p>', { defaultNew: true, defaultReply: true, audience: { departments: [], domains: ['other.ch'], groups: [] } })
  ];
  if (edit) edit(c);
  return c;
}
const ME = { givenName: 'Anna', surname: 'Muster', mobilePhone: null, mail: 'anna@example.com', department: 'Sales', onPremisesSamAccountName: 'mustera' };

function runtime({ config, files, me = ME, groups = [], composeType = 'newMail', coercion = 'html', from = 'anna@example.com', roaming = {}, itemType = 'message', attachments = [], deny, signInFails } = {}) {
  const calls = [];
  const ok = (value) => (cb) => cb({ status: 'succeeded', value });
  const handlers = {};
  const atts = attachments.slice();
  let n = 0;
  const item = {
    itemType,
    getComposeTypeAsync: (cb) => ok({ composeType, coercionType: coercion })(cb),
    from: { getAsync: (cb) => ok({ emailAddress: from })(cb) },
    disableClientSignatureAsync: (cb) => { calls.push(['disable']); ok()(cb); },
    getAttachmentsAsync: (cb) => ok(atts.slice())(cb),
    removeAttachmentAsync: (id, cb) => { calls.push(['remove', id]); atts.splice(atts.findIndex((a) => a.id === id), 1); ok()(cb); },
    addFileAttachmentFromBase64Async: (b64, name, opts, cb) => { calls.push(['attach', name, b64, opts.isInline]); atts.push({ id: 'a' + (++n), name, isInline: opts.isInline }); ok('a' + n)(cb); },
    body: {
      getTypeAsync: (cb) => ok(coercion)(cb),
      setSignatureAsync: (content, opts, cb) => { calls.push(['set', content, opts.coercionType]); ok()(cb); }
    }
  };
  const settings = Object.assign({}, roaming);
  const Office = {
    AsyncResultStatus: { Succeeded: 'succeeded' },
    CoercionType: { Html: 'html', Text: 'text' },
    MailboxEnums: { ItemType: { Message: 'message', Appointment: 'appointment' } },
    actions: { associate: (name, fn) => { handlers[name] = fn; } },
    context: {
      requirements: { isSetSupported: () => true },
      roamingSettings: { get: (k) => settings[k], set: (k, v) => { settings[k] = v; }, saveAsync: (cb) => cb && cb({ status: 'succeeded' }) },
      mailbox: { item, userProfile: { displayName: 'Anna Muster', emailAddress: 'anna@example.com' } }
    }
  };
  const graph = fakeGraph({
    files: files || { 'signatures.json': JSON.stringify(config || makeConfig()), 'assets/logo.png': PNG, 'photos/mustera.jpg': JPG },
    me, groups, deny
  });
  const sandbox = { Office, console: { log: (m) => calls.push(['log', String(m)]) }, setTimeout, clearTimeout, fetch: graph.fetch, __osignSignInFails: !!signInFails };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(bundle, sandbox, { filename: 'launchevent.js' });
  const fire = (name) => new Promise((resolve) => handlers[name]({ completed: () => resolve() }));
  return { calls, graph, handlers, fire, settings, atts };
}
const setCall = (r) => r.calls.find((c) => c[0] === 'set');

test('registers its three handlers, in a runtime without window or DOM', () => {
  assert.deepEqual(Object.keys(runtime().handlers).sort(), ['onMessageFromChangedHandler', 'onNewAppointmentOrganizerHandler', 'onNewMessageComposeHandler']);
});

test('new mail: the signature from SharePoint, filled from Graph, its images embedded', async () => {
  const r = runtime();
  await r.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r)[1], '<p><b>Anna Muster</b></p><p><img src="cid:osign-assets-logo.png" width="90"><img src="cid:osign-photos-mustera.jpg" width="100"></p>');
  assert.equal(setCall(r)[2], 'html');
  const attached = r.calls.filter((c) => c[0] === 'attach');
  assert.deepEqual(attached.map((c) => [c[1], c[3]]), [['osign-assets-logo.png', true], ['osign-photos-mustera.jpg', true]]);
  assert.ok(r.calls.some((c) => c[0] === 'disable'));
  // Only Graph and SharePoint were asked: nothing anywhere else.
  assert.ok(r.graph.requests.every((u) => /^https:\/\/(graph\.microsoft\.com|download\.example)\//.test(u)), r.graph.requests.join('\n'));
  assert.ok(r.graph.requests.some((u) => u.endsWith('/sites/root/drive/root:/OSign/signatures.json')));
});

test('images are embedded byte for byte (base64 without btoa, every length)', async () => {
  const sizes = [1, 2, 3, 4, 5, 100, 1001];
  const files = { 'signatures.json': '' };
  let html = '<p>@@givenName';
  sizes.forEach((n) => {
    files['assets/i' + n + '.png'] = Buffer.from(Array.from({ length: n }, (_, i) => (i * 37 + n) & 255));
    html += '<img src="assets/i' + n + '.png">';
  });
  files['signatures.json'] = JSON.stringify(makeConfig((c) => { c.signatures = [sig('one', 'One', html + '</p>', { defaultNew: true })]; }));
  const r = runtime({ files });
  await r.fire('onNewMessageComposeHandler');
  const attached = Object.fromEntries(r.calls.filter((c) => c[0] === 'attach').map((c) => [c[1], c[2]]));
  sizes.forEach((n) => assert.equal(attached['osign-assets-i' + n + '.png'], files['assets/i' + n + '.png'].toString('base64'), 'length ' + n));
});

test('a user with no photo gets the signature without one', async () => {
  const r = runtime({ me: Object.assign({}, ME, { onPremisesSamAccountName: 'nobody' }) });
  await r.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r)[1], '<p><b>Anna Muster</b></p><p><img src="cid:osign-assets-logo.png" width="90"></p>');
});

test('the images of an earlier signature go when another one is put in; the user\'s attachments stay', async () => {
  const r = runtime({ attachments: [{ id: 'old1', name: 'osign-assets-logo.png', isInline: true }, { id: 'mine', name: 'offer.pdf', isInline: false }] });
  await r.fire('onMessageFromChangedHandler');
  assert.deepEqual(r.calls.filter((c) => c[0] === 'remove'), [['remove', 'old1']]);
  assert.deepEqual(r.atts.map((a) => a.name).sort(), ['offer.pdf', 'osign-assets-logo.png', 'osign-photos-mustera.jpg']);
});

test('a reply gets the reply signature; plain text gets text and no images', async () => {
  const r = runtime({ composeType: 'reply', coercion: 'text' });
  await r.fire('onNewMessageComposeHandler');
  assert.deepEqual(setCall(r).slice(1), ['Anna', 'text']);
  assert.ok(!r.calls.some((c) => c[0] === 'attach'));
});

test("the user's own choice wins", async () => {
  const roaming = { 'osign.prefs': JSON.stringify({ new: 'short', reply: 'none' }) };
  const r = runtime({ roaming });
  await r.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r)[1], '<p>Anna</p>');
  const r2 = runtime({ composeType: 'reply', roaming });
  await r2.fire('onNewMessageComposeHandler');
  assert.ok(!setCall(r2));
});

test('a changed sender swaps the signature', async () => {
  const r = runtime({ from: 'anna@other.ch' });
  await r.fire('onMessageFromChangedHandler');
  assert.equal(setCall(r)[1], '<p>Other anna@example.com</p>');
});

test('signatures for an AD group go to its members', async () => {
  const config = makeConfig((c) => c.signatures.push(sig('field', 'Field', '<p>Field @@givenName</p>', { defaultNew: true, audience: { departments: [], domains: [], groups: ['SG-Field'] } })));
  const member = runtime({ config, groups: [{ id: 'g1', displayName: 'Field staff', onPremisesSamAccountName: 'SG-Field' }] });
  await member.fire('onNewMessageComposeHandler');
  assert.equal(setCall(member)[1], '<p>Field Anna</p>');
  const other = runtime({ config, groups: [{ id: 'g2', displayName: 'Office', onPremisesSamAccountName: 'SG-Office' }] });
  await other.fire('onNewMessageComposeHandler');
  assert.match(setCall(other)[1], /Anna Muster/);
  // Groups not readable (no consent for GroupMember.Read.All): only the group signature is out of
  // reach; the user's details still are read, and the general signature goes in.
  const denied = runtime({ config, deny: ['transitiveMemberOf'] });
  await denied.fire('onNewMessageComposeHandler');
  assert.match(setCall(denied)[1], /Anna Muster/);
});

test('without the user details and without the option, nothing is inserted', async () => {
  const r = runtime({ deny: ['/me?'] });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(!setCall(r));
  assert.ok(r.calls.some((c) => c[0] === 'log' && /no user details/.test(c[1])));
});

test('meetings only when the option is on', async () => {
  const r = runtime({ itemType: 'appointment' });
  await r.fire('onNewAppointmentOrganizerHandler');
  assert.ok(!setCall(r));
  const r2 = runtime({ itemType: 'appointment', config: makeConfig((c) => { c.options.meetings = true; }) });
  await r2.fire('onNewAppointmentOrganizerHandler');
  assert.ok(setCall(r2));
});

test('nothing published yet, or no sign-in: no signature, and the event still completes', async () => {
  const r = runtime({ files: {} });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(!setCall(r));
  assert.ok(r.calls.some((c) => c[0] === 'log' && /No signatures are published yet/.test(c[1])));
  const r2 = runtime({ signInFails: true });
  await r2.fire('onNewMessageComposeHandler');
  assert.ok(!setCall(r2));
  assert.ok(r2.calls.some((c) => c[0] === 'log' && /sign-in failed/.test(c[1])));
});
