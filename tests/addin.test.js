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
// A signature for the entire organization (a new one is not assigned to anybody).
const sigAll = (id, name) => { const s = core.newSignature(id, name); s.audience.everyone = true; return s; };
const buildTestAddin = require('./support/build-test-addin.js');
const fakeGraph = require('./support/fake-graph.js');

const doc = new JSDOM('').window.document;
const PNG = Buffer.from('89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000d4944415478da63f8ffff3f0005fe02fea7d6a4d60000000049454e44ae426082', 'hex');
const JPG = Buffer.from('ffd8ffe000104a46494600010100000100010000ffd9', 'hex');
let bundle, siteBundle, idsBundle, upperSiteBundle, hostedBundle;
test.before(async () => {
  upperSiteBundle = fs.readFileSync(path.join(await buildTestAddin('contoso.sharepoint.com:/Sites/Signatures:'), 'launchevent.js'), 'utf8');
  bundle = fs.readFileSync(path.join(await buildTestAddin(), 'launchevent.js'), 'utf8');
  hostedBundle = fs.readFileSync(path.join(await buildTestAddin('', { base: 'https://osign.example/' }), 'launchevent.js'), 'utf8');
  siteBundle = fs.readFileSync(path.join(await buildTestAddin('contoso.sharepoint.com:/sites/Signatures:'), 'launchevent.js'), 'utf8');
  idsBundle = fs.readFileSync(path.join(await buildTestAddin('contoso.sharepoint.com:/sites/Signatures:', { drive: 'b!drive-1', item: 'OSIGN-ITEM' }), 'launchevent.js'), 'utf8');
});

function sig(id, name, html, extra) {
  const s = Object.assign(sigAll(id, name), { html }, extra || {});
  s.compiled = core.compile(s.html, doc);
  return s;
}
function makeConfig(edit) {
  const c = core.defaultConfig();
  // The values from Entra ID (Graph), as these tests' stand-in serves them.
  c.options.valueSource = 'graph';
  c.signatures = [
    sig('std', 'Standard', '<p><b>@@givenName @@sn</b></p><p>Mobile @@mobile</p><p><img src="assets/logo.png" width="90"><img src="photos/@@sAMAccountName.jpg" width="100"></p>', { defaultNew: true }),
    sig('short', 'Short', '<p>@@givenName</p>', { defaultReply: true }),
    sig('other', 'Other', '<p>Other @@mail</p>', { defaultNew: true, defaultReply: true, audience: { departments: [], domains: ['other.ch'], groups: [] } })
  ];
  if (edit) edit(c);
  return c;
}
const ME = { givenName: 'Anna', surname: 'Muster', mobilePhone: null, mail: 'anna@example.com', department: 'Sales', onPremisesSamAccountName: 'mustera' };

function runtime({ config, files, me = ME, groups = [], composeType = 'newMail', coercion = 'html', from = 'anna@example.com', roaming = {}, itemType = 'message', attachments = [], session, notices, deny, signInFails, signInHangs, site, byIds, lang, runtimeStorage, platform, failing, saveFails, upperSite, hosted } = {}) {
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
    // The notices above the message, by key, when the test asks for them.
    notificationMessages: notices && {
      replaceAsync: (k, m, cb) => { notices[k] = m; calls.push(['notice', k, m.message]); ok()(cb); },
      removeAsync: (k, cb) => { delete notices[k]; calls.push(['unnotice', k]); ok()(cb); }
    },
    // sessionData (Mailbox 1.11), when the test gives one: a missing key answers an error, as in Outlook.
    sessionData: session && {
      getAsync: (k, cb) => cb(k in session ? { status: 'succeeded', value: session[k] } : { status: 'failed', error: new Error('not found') }),
      setAsync: (k, v, cb) => { session[k] = v; ok()(cb); }
    },
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
      diagnostics: { platform: platform || 'OfficeOnline' },
      displayLanguage: lang || 'en-US',
      roamingSettings: { get: (k) => settings[k], set: (k, v) => { settings[k] = v; }, remove: (k) => { delete settings[k]; }, saveAsync: (cb) => cb && cb({ status: saveFails ? 'failed' : 'succeeded' }) },
      mailbox: { item, userProfile: { displayName: 'Anna Muster', emailAddress: 'anna@example.com' } }
    }
  };
  const graph = fakeGraph({
    files: files || { 'signatures.json': JSON.stringify(config || makeConfig()), 'assets/logo.png': PNG, 'photos/mustera.jpg': JPG },
    me, groups, deny, site, byIds
  });
  // The sign-in's time limit (the only one this long) runs out at once in the tests.
  const fastTimeout = (fn, ms) => setTimeout(fn, ms >= 15000 ? 5 : ms);
  // failing(url): 'network' (fetch fails, as with no network), a status to answer, or nothing.
  const fetch = async (url, init) => {
    const f = failing && failing(url);
    if (f === 'network') { graph.requests.push(url); throw new TypeError('Failed to fetch'); }
    if (f) { graph.requests.push(url); return { ok: false, status: f, json: async () => ({}), text: async () => '', arrayBuffer: async () => new ArrayBuffer(0) }; }
    // The add-in's own address (hosted: its files there), as Cloudflare Pages serves them.
    if (hosted && String(url).startsWith('https://osign.example/')) {
      graph.requests.push(url);
      const body = hosted[String(url).slice('https://osign.example/'.length)];
      if (body === undefined) return { ok: false, status: 404, json: async () => ({}), arrayBuffer: async () => new ArrayBuffer(0) };
      const buf = Buffer.isBuffer(body) ? body : Buffer.from(typeof body === 'string' ? body : JSON.stringify(body));
      return { ok: true, status: 200, json: async () => JSON.parse(buf.toString()), arrayBuffer: async () => buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.length) };
    }
    return graph.fetch(url, init);
  };
  const sandbox = { Office, console: { log: (m) => calls.push(['log', String(m)]) }, setTimeout: fastTimeout, clearTimeout, fetch, __osignSignInHangs: !!signInHangs };
  // How often the sign-in was asked (the stand-in reads this once each time).
  const signIns = { n: 0 };
  Object.defineProperty(sandbox, '__osignSignInFails', { get: () => { signIns.n++; return !!signInFails; } });
  // OfficeRuntime.storage, when the test gives one (classic Outlook's event runtime has it).
  if (runtimeStorage) sandbox.OfficeRuntime = { storage: { getItem: async (k) => (k in runtimeStorage ? runtimeStorage[k] : null), setItem: async (k, v) => { runtimeStorage[k] = v; } } };
  sandbox.globalThis = sandbox;
  vm.createContext(sandbox);
  vm.runInContext(hosted ? hostedBundle : upperSite ? upperSiteBundle : byIds ? idsBundle : site ? siteBundle : bundle, sandbox, { filename: 'launchevent.js' });
  const fire = (name) => new Promise((resolve) => handlers[name]({ completed: () => resolve() }));
  return { calls, graph, handlers, fire, settings, atts, signIns };
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

test('with a placeholder photo (photos/_default.jpg), a user with none, or with no login, gets that one', async () => {
  const files = { 'signatures.json': JSON.stringify(makeConfig()), 'assets/logo.png': PNG, 'photos/mustera.jpg': JPG, 'photos/_default.jpg': Buffer.concat([JPG, Buffer.from([1, 2, 3])]) };
  for (const login of ['nobody', null]) {
    const r = runtime({ files, me: Object.assign({}, ME, { onPremisesSamAccountName: login }) });
    await r.fire('onNewMessageComposeHandler');
    assert.match(setCall(r)[1], /<img src="cid:osign-photos-[^"]+\.jpg" width="100">/, String(login));
    const att = r.calls.filter((c) => c[0] === 'attach' && /photos-/.test(c[1]));
    assert.equal(Buffer.from(att[0][2], 'base64').length, JPG.length + 3);
  }
});

test('the images of an earlier signature go when another one is put in; the user\'s attachments stay', async () => {
  const r = runtime({ attachments: [{ id: 'old1', name: 'osign-assets-logo.png', isInline: true }, { id: 'mine', name: 'offer.pdf', isInline: false }] });
  await r.fire('onMessageFromChangedHandler');
  assert.deepEqual(r.calls.filter((c) => c[0] === 'remove'), [['remove', 'old1']]);
  assert.deepEqual(r.atts.map((a) => a.name).sort(), ['offer.pdf', 'osign-assets-logo.png', 'osign-photos-mustera.jpg']);
});

test('a reply keeps the images of the OSign signature it quotes; its own carry a tag of the message', async () => {
  const session = {};
  const quoted = [{ id: 'q1', name: 'osign-assets-logo.png', isInline: true }, { id: 'q2', name: 'osign-1a2b3c4d-photos-muellerb.jpg', isInline: true }];
  const config = makeConfig((c) => { c.signatures[0].defaultReply = true; c.signatures[1].defaultReply = false; });
  const r = runtime({ config, composeType: 'reply', attachments: quoted, session });
  await r.fire('onNewMessageComposeHandler');
  assert.match(session['osign.tag'], /^[0-9a-f]{8}$/);
  const p = 'osign-' + session['osign.tag'] + '-';
  assert.ok(!r.calls.some((c) => c[0] === 'remove'));
  assert.equal(setCall(r)[1], '<p><b>Anna Muster</b></p><p><img src="cid:' + p + 'assets-logo.png" width="90"><img src="cid:' + p + 'photos-mustera.jpg" width="100"></p>');
  // Put in again (a changed sender): its own images are replaced, under the same tag; the quoted ones stay.
  await r.fire('onMessageFromChangedHandler');
  assert.deepEqual(r.calls.filter((c) => c[0] === 'remove').map((c) => c[1]).sort(), ['a1', 'a2']);
  assert.deepEqual(r.atts.map((a) => a.name).sort(), ['osign-1a2b3c4d-photos-muellerb.jpg', 'osign-assets-logo.png', p + 'assets-logo.png', p + 'photos-mustera.jpg'].sort());
});

test('when no signature could be put in, a notice on the message says why; one that works takes it away', async () => {
  const notices = {};
  const r = runtime({ notices, signInFails: true });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(!setCall(r));
  assert.match(notices.osign.message, /^OSign inserted no signature; see Signatures\. .*sign-in failed/);
  assert.equal(notices.osign.type, 'informationalMessage');
  assert.ok(notices.osign.message.length <= 150);
  // No signature meant for this message is no failure: no notice.
  const quiet = {};
  const q = runtime({ notices: quiet, config: makeConfig((c) => { c.signatures.forEach((s) => { s.defaultNew = false; }); }) });
  await q.fire('onNewMessageComposeHandler');
  assert.deepEqual(quiet, {});
  // Signed: a notice an earlier attempt left goes.
  const left = { osign: { message: 'old' } };
  const w = runtime({ notices: left });
  await w.fire('onNewMessageComposeHandler');
  assert.ok(setCall(w));
  assert.deepEqual(left, {});
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

test('sign-in refused (classic Outlook\'s event runtime): the signature and images kept before are used', async () => {
  const runtimeStorage = {}, roaming = {};
  const first = runtime({ runtimeStorage, roaming });
  await first.fire('onNewMessageComposeHandler');
  assert.ok(setCall(first));
  // Its mailbox settings (the details) and storage (signatures, images) carry over.
  const notices = {};
  const r = runtime({ runtimeStorage, roaming: first.settings, signInFails: true, notices });
  await r.fire('onNewMessageComposeHandler');
  const set = setCall(r);
  assert.ok(set, r.calls.filter((c) => c[0] === 'log').map((c) => c[1]).join('\n'));
  assert.match(set[1], /Anna Muster/);
  assert.ok(r.calls.some((c) => c[0] === 'attach' && /logo/.test(c[1])), 'the logo from the copy');
  assert.ok(r.calls.some((c) => c[0] === 'log' && /from the copy kept before/.test(c[1])));
  assert.deepEqual(notices, {});
});

test('classic Outlook: refused the sign-in, with nothing kept, it says what to do instead of the code', async () => {
  const notices = {};
  const r = runtime({ signInFails: true, notices, platform: 'PC', lang: 'de-CH' });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(!setCall(r));
  assert.match(notices.osign.message, /Im klassischen Outlook einmal OSign → Signaturen öffnen/);
});

test('classic Outlook: the copy the pane kept in the mailbox goes in, with the images that fit', async () => {
  // What the pane keeps (keepForClassic), here written by hand: one signature, its logo.
  const config = makeConfig();
  const std = config.signatures[0];
  const roaming = {
    'osign.profile': JSON.stringify({ at: Date.now() - 60 * 60 * 1000, values: { givenName: 'Anna', sn: 'Muster', mobile: '', sAMAccountName: 'mustera', mail: 'anna@example.com' }, groups: [], groupsFor: 'graph\n' }),
    'osign.classic': JSON.stringify({ at: config.updated, cfg: { schemaVersion: config.schemaVersion, options: config.options, fields: [], signatures: [{ id: std.id, name: std.name, compiled: std.compiled, defaultNew: true, audience: std.audience }] },
      images: { 'assets/logo.png': { b64: PNG.toString('base64'), type: 'image/png', file: 'assets-logo.png' } } })
  };
  const notices = {};
  const r = runtime({ signInFails: true, notices, roaming, platform: 'PC' });
  await r.fire('onNewMessageComposeHandler');
  const set = setCall(r);
  assert.ok(set, r.calls.filter((c) => c[0] === 'log').map((c) => c[1]).join('\n'));
  assert.match(set[1], /Anna Muster/);
  assert.ok(r.calls.some((c) => c[0] === 'attach' && /logo/.test(c[1])), 'the logo from the mailbox');
  assert.ok(r.calls.some((c) => c[0] === 'log' && /copy in the mailbox/.test(c[1])));
  assert.deepEqual(notices, {});
});

test('classic Outlook: the images as they are from next to the add-in\'s files (img/), not the smaller copies in the mailbox', async () => {
  const config = makeConfig();
  const std = config.signatures[0];
  const small = Buffer.from('a smaller copy');
  const roaming = {
    'osign.profile': JSON.stringify({ at: Date.now() - 60 * 60 * 1000, values: { givenName: 'Anna', sn: 'Muster', mobile: '', sAMAccountName: 'mustera', mail: 'anna@example.com' }, groups: [], groupsFor: 'graph\n' }),
    'osign.classic': JSON.stringify({ at: config.updated, cfg: { schemaVersion: config.schemaVersion, options: config.options, fields: [], signatures: [{ id: std.id, name: std.name, compiled: std.compiled, defaultNew: true, audience: std.audience }] },
      images: { 'assets/logo.png': { b64: small.toString('base64'), type: 'image/png', file: 'assets-logo.png' } } })
  };
  const file = 'img/' + 'a1'.repeat(12) + '.png';
  const r = runtime({ signInFails: true, roaming, platform: 'PC', hosted: { 'img/index.json': { 'logo.png': file }, [file]: PNG } });
  await r.fire('onNewMessageComposeHandler');
  const logo = r.calls.find((c) => c[0] === 'attach' && /logo/.test(c[1]));
  assert.ok(logo, r.calls.filter((c) => c[0] === 'log').map((c) => c[1]).join('\n'));
  assert.equal(logo[2], PNG.toString('base64'), 'the image as published, not the copy');
  assert.equal(logo[3], true, 'embedded');
  assert.ok(r.graph.requests.includes('https://osign.example/img/index.json'));
  // Kept in this Outlook for the next mail, only the images published now.
  const storage = { 'osign.hosted': JSON.stringify({ ['img/' + 'b2'.repeat(12) + '.png']: 'b2xk' }) };
  const again = runtime({ signInFails: true, roaming, platform: 'PC', runtimeStorage: storage, hosted: { 'img/index.json': { 'logo.png': file }, [file]: PNG } });
  await again.fire('onNewMessageComposeHandler');
  assert.deepEqual(Object.keys(JSON.parse(storage['osign.hosted'])), [file], 'the image no longer published is gone');
  // Not to be had there (no network): the copy kept in the mailbox.
  const off = runtime({ signInFails: true, roaming, platform: 'PC', hosted: {}, failing: (u) => (/osign\.example/.test(u) ? 'network' : '') });
  await off.fire('onNewMessageComposeHandler');
  assert.equal(off.calls.find((c) => c[0] === 'attach' && /logo/.test(c[1]))[2], small.toString('base64'));
  // An index naming something that is not an image file of its own: not followed.
  const odd = runtime({ signInFails: true, roaming, platform: 'PC', hosted: { 'img/index.json': { 'logo.png': '../manifest.xml' } } });
  await odd.fire('onNewMessageComposeHandler');
  assert.equal(odd.calls.find((c) => c[0] === 'attach' && /logo/.test(c[1]))[2], small.toString('base64'));
});

test('a sign-in that never answers: after a while a notice, and the event completes', async () => {
  const notices = {};
  const r = runtime({ signInHangs: true, notices });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(!setCall(r));
  assert.ok(r.calls.some((c) => c[0] === 'log' && /Signing in took too long/.test(c[1])), r.calls.filter((c) => c[0] === 'log').map((c) => c[1]).join('\n'));
  assert.match(notices.osign.message, /Signing in took too long/);
});

test('each step is logged, so the console shows where it stopped', async () => {
  const r = runtime();
  await r.fire('onNewMessageComposeHandler');
  const logs = r.calls.filter((c) => c[0] === 'log').map((c) => c[1]);
  for (const rx of [/^OSign [\d.]+: event handler ready/, /new started \(sign-in through Outlook: yes\)/, /signatures loaded/, /signature "Standard" put in/]) {
    assert.ok(logs.some((l) => rx.test(l)), rx + '\n' + logs.join('\n'));
  }
});

test('a variant the user picked: their text in the editable field, the rest the company\'s', async () => {
  const config = makeConfig((c) => {
    c.fields.push({ key: 'greeting', label: 'Greeting', ad: '', graph: '', sample: '', userEditable: true });
    c.signatures[0] = sig('std', 'Standard', '<p>@@greeting</p><p><b>@@givenName @@sn</b></p>', { defaultNew: true });
  });
  const roaming = {
    'osign.variants': JSON.stringify([{ id: 'my-x', name: 'Warm', base: 'std', values: { greeting: 'Herzliche <Grüsse>', givenName: 'Not me' } }]),
    'osign.prefs': JSON.stringify({ new: 'my-x', reply: '' })
  };
  const r = runtime({ config, roaming });
  await r.fire('onNewMessageComposeHandler');
  const html = setCall(r)[1];
  assert.match(html, /Herzliche &lt;Grüsse&gt;/);
  assert.match(html, /<b>Anna Muster<\/b>/);
  assert.ok(r.calls.some((c) => c[0] === 'log' && /signature "Warm" put in/.test(c[1])));
});

test('values from AD: mixed takes the fields without a Graph source, and the groups, from the user file', async () => {
  const config = makeConfig((c) => {
    c.options.valueSource = 'mixed';
    c.signatures = [
      sig('std', 'Standard', '<p>@@givenName @@sn</p><p>@@description</p><p>Direkt @@homePhone</p>', { defaultNew: true }),
      sig('field', 'Field', '<p>Field @@givenName, @@homePhone</p>', { defaultNew: true, audience: { departments: [], domains: [], groups: ['SG-Field'] } })
    ];
  });
  const user = (groups) => JSON.stringify({ values: { description: 'Dipl. Ing.', homePhone: '+41 41 000 00 01', info: '' }, groups });
  const base = { 'signatures.json': JSON.stringify(config) };
  const r = runtime({ files: Object.assign({}, base, { 'users/anna@example.com.json': user([]) }) });
  await r.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r)[1], '<p>Anna Muster</p><p>Dipl. Ing.</p><p>Direkt +41 41 000 00 01</p>');
  // Groups from the AD file: no GroupMember.Read.All, no Graph group call.
  assert.ok(!r.graph.requests.some((u) => /memberOf/i.test(u)));
  const r2 = runtime({ files: Object.assign({}, base, { 'users/anna@example.com.json': user(['sg-field']) }) });
  await r2.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r2)[1], '<p>Field Anna, +41 41 000 00 01</p>');
  // Not in the export: Graph's fields still come, the AD ones stay empty (and their lines go).
  const r3 = runtime({ files: base });
  await r3.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r3)[1], '<p>Anna Muster</p>');
});

test('values from AD only: no /me, and no signature for someone not in the export', async () => {
  const config = makeConfig((c) => {
    c.options.valueSource = 'ad';
    c.fields = c.fields.map((f) => Object.assign({}, f));
    c.signatures = [sig('std', 'Standard', '<p>@@givenName @@sn, @@title</p>', { defaultNew: true })];
  });
  const file = JSON.stringify({ values: { givenName: 'Anna', sn: 'Muster', title: 'Leiterin' }, groups: [] });
  const r = runtime({ files: { 'signatures.json': JSON.stringify(config), 'users/anna@example.com.json': file } });
  await r.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r)[1], '<p>Anna Muster, Leiterin</p>');
  assert.ok(!r.graph.requests.some((u) => /\/me(\?|$)/.test(u)), r.graph.requests.join('\n'));
  const r2 = runtime({ files: { 'signatures.json': JSON.stringify(config) } });
  await r2.fire('onNewMessageComposeHandler');
  assert.ok(!setCall(r2));
  assert.ok(r2.calls.some((c) => c[0] === 'log' && /not in the AD export/.test(c[1])));
});

test('a site of its own: found once by its id, then the folder in its library (not the site\'s path and a path after it)', async () => {
  const r = runtime({ site: 'contoso.sharepoint.com:/sites/Signatures:' });
  await r.fire('onNewMessageComposeHandler');
  assert.match(setCall(r)[1], /Anna Muster/);
  const found = r.graph.requests.filter((u) => u.includes('/sites/contoso.sharepoint.com:/sites/Signatures:'));
  assert.deepEqual(found, ['https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Signatures:?$select=id']);
  assert.ok(r.graph.requests.some((u) => u.endsWith('/sites/' + encodeURIComponent('contoso.sharepoint.com,1111,2222') + '/drive/root:/OSign/signatures.json')));
});

test('the folder by its ids (as the server made the files): read without asking for the site, which a user who may only read the folder cannot', async () => {
  const r = runtime({ byIds: true, deny: ['/sites/'] });
  await r.fire('onNewMessageComposeHandler');
  assert.match(setCall(r)[1], /Anna Muster/);
  assert.ok(!r.graph.requests.some((u) => u.includes('/sites/')), r.graph.requests.join('\n'));
  assert.ok(r.graph.requests.some((u) => u.endsWith('/drives/b!drive-1/items/OSIGN-ITEM:/signatures.json')));
});

test('the folder not readable for the user (the consent is there): said as that, with what to share', async () => {
  const r = runtime({ byIds: true, deny: ['/drives/'] });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(!setCall(r));
  assert.ok(r.calls.some((c) => c[0] === 'log' && /does not let you read the folder OSign.*Everyone except external users/.test(c[1])), r.calls.filter((c) => c[0] === 'log').map((c) => c[1]).join('\n'));
});

test('in a German Outlook, what went wrong is said in German', async () => {
  const r = runtime({ byIds: true, deny: ['/drives/'], notices: {}, lang: 'de-CH' });
  await r.fire('onNewMessageComposeHandler');
  const n = r.calls.find((c) => c[0] === 'notice');
  assert.ok(n, 'a notice');
  assert.match(n[2], /^OSign hat keine Signatur eingefügt; mehr unter Signaturen\. Signaturen suchen: SharePoint lässt Sie den Ordner OSign nicht lesen/);
});

test('a banner campaign running today goes in where the signature has its place, with its image; none running: the place goes; a reply without it when it is for new mail only', async () => {
  const today = core.today(new Date());
  const banner = (c) => {
    c.signatures[1] = sig('short', 'Short', '<p>@@givenName</p><div>{{campaign}}</div>', { defaultReply: true });
    c.campaigns = [{ id: 'xmas', name: 'Xmas', validFrom: today, validUntil: today, html: '<a href="https://ex.ch/"><img src="assets/logo.png" width="300"></a>', compiled: '<a href="https://ex.ch/"><img src="assets/logo.png" width="300"></a>', signatures: [], newOnly: false }];
  };
  // The reply signature, with its place: the banner there, its image embedded.
  let r = runtime({ config: makeConfig(banner), composeType: 'reply' });
  await r.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r)[1], '<p>Anna</p><div><a href="https://ex.ch/"><img src="cid:osign-assets-logo.png" width="300"></a></div>');
  // New mail, the standard signature without a place: below it.
  r = runtime({ config: makeConfig(banner) });
  await r.fire('onNewMessageComposeHandler');
  assert.match(setCall(r)[1], /<\/p><div><a href="https:\/\/ex\.ch\/"><img src="cid:osign-assets-logo\.png" width="300"><\/a><\/div>$/);
  // New mail only: a reply without it, its place gone.
  r = runtime({ config: makeConfig((c) => { banner(c); c.campaigns[0].newOnly = true; }), composeType: 'reply' });
  await r.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r)[1], '<p>Anna</p>');
  // Over: the place goes.
  r = runtime({ config: makeConfig((c) => { banner(c); c.campaigns[0].validFrom = c.campaigns[0].validUntil = '2020-01-01'; }), composeType: 'reply' });
  await r.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r)[1], '<p>Anna</p>');
});

test('classic Outlook: a banner campaign kept with the copy in the mailbox goes in too', async () => {
  const config = makeConfig();
  const std = config.signatures[0];
  const roaming = {
    'osign.profile': JSON.stringify({ at: Date.now() - 60 * 60 * 1000, values: { givenName: 'Anna', sn: 'Muster', mobile: '', sAMAccountName: 'mustera', mail: 'anna@example.com' }, groups: [], groupsFor: 'graph\n' }),
    'osign.classic': JSON.stringify({ at: config.updated, cfg: { schemaVersion: config.schemaVersion, options: config.options, fields: [], signatures: [{ id: std.id, name: std.name, compiled: '<p>@@givenName</p>{{campaign}}', defaultNew: true, audience: std.audience }],
      campaigns: [{ id: 'c', compiled: '<b>BANNER</b>', signatures: [] }] }, images: {} })
  };
  const r = runtime({ signInFails: true, roaming, platform: 'PC' });
  await r.fire('onNewMessageComposeHandler');
  assert.equal(setCall(r)[1], '<p>Anna</p><div><b>BANNER</b></div>');
});

test('signed in, but SharePoint cannot be read now (offline, busy): the images kept before still go in', async () => {
  const runtimeStorage = {};
  const first = runtime({ runtimeStorage });
  await first.fire('onNewMessageComposeHandler');
  assert.match(setCall(first)[1], /cid:osign-assets-logo\.png/);
  // The listing answers 503, the photo cannot be found, every download fails as with no network.
  const failing = (u) => (/assets:\/children/.test(u) ? 503 : /photos\//.test(u) ? 429 : /download\.example/.test(u) ? 'network' : null);
  const r = runtime({ runtimeStorage, roaming: first.settings, failing });
  await r.fire('onNewMessageComposeHandler');
  const set = setCall(r);
  assert.ok(set, r.calls.filter((c) => c[0] === 'log').map((c) => c[1]).join('\n'));
  assert.equal(set[1], '<p><b>Anna Muster</b></p><p><img src="cid:osign-assets-logo.png" width="90"><img src="cid:osign-photos-mustera.jpg" width="100"></p>');
  const attached = Object.fromEntries(r.calls.filter((c) => c[0] === 'attach').map((c) => [c[1], c[2]]));
  assert.equal(attached['osign-assets-logo.png'], PNG.toString('base64'));
  assert.equal(attached['osign-photos-mustera.jpg'], JPG.toString('base64'));
  // Kept nowhere: the image is left out, the signature still goes in.
  const r2 = runtime({ runtimeStorage: {}, failing: (u) => (/assets:\/children/.test(u) ? 503 : null) });
  await r2.fire('onNewMessageComposeHandler');
  assert.match(setCall(r2)[1], /^<p><b>Anna Muster<\/b><\/p><p><img src="cid:osign-photos-mustera\.jpg" width="100"><\/p>$/);
});

test('no network: said as that, in Outlook\'s language', async () => {
  for (const [lang, rx] of [['en-US', /Finding the signatures: cannot be reached\./], ['de-CH', /Signaturen suchen: nicht erreichbar\./], ['fr-CH', /Recherche des signatures : impossible à joindre\./], ['it-CH', /Ricerca delle firme: non raggiungibile\./]]) {
    const notices = {};
    const r = runtime({ notices, lang, failing: () => 'network' });
    await r.fire('onNewMessageComposeHandler');
    assert.ok(!setCall(r));
    assert.match(notices.osign.message, rx, lang);
  }
});

test('the text the handler shows when the details could not be read, with no error to say, is translated', async () => {
  // The text the handler falls back on is in the translations.
  const { LANG } = await import(path.join(__dirname, '..', 'addin', 'src', 'lang.js'));
  const common = fs.readFileSync(path.join(__dirname, '..', 'addin', 'src', 'common.js'), 'utf8');
  assert.match(common, /\[\/\^your details could not be read\\\.\$\/, 'Ihre Angaben konnten nicht gelesen werden\.'\]/);
  assert.ok(LANG.fr.msg.includes('Vos données n\'ont pas pu être lues.'));
  assert.ok(LANG.it.msg.includes('Non è stato possibile leggere i suoi dati.'));
});

test('a sign-in without a window that failed is not tried again for each thing one message reads', async () => {
  const runtimeStorage = {};
  const first = runtime({ runtimeStorage });
  await first.fire('onNewMessageComposeHandler');
  const r = runtime({ runtimeStorage, roaming: first.settings, signInFails: true });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(setCall(r), 'from the copies kept before');
  assert.equal(r.signIns.n, 1);
  // The next message, a moment later: not tried again either.
  await r.fire('onNewMessageComposeHandler');
  assert.equal(r.signIns.n, 1);
});

test('a setting the mailbox refuses is not left in memory, so later saves are not refused with it', async () => {
  const r = runtime({ saveFails: true, roaming: { 'osign.prefs': JSON.stringify({ new: 'short', reply: '' }) } });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(setCall(r));
  assert.equal(r.settings['osign.profile'], undefined);
  assert.equal(r.settings['osign.prefs'], JSON.stringify({ new: 'short', reply: '' }));
});

test('the site as written in any case (Sites, ROOT): found as fill.mjs accepts it', async () => {
  const r = runtime({ upperSite: true, site: 'contoso.sharepoint.com:/sites/Signatures:' });
  await r.fire('onNewMessageComposeHandler');
  assert.match(setCall(r)[1], /Anna Muster/, r.calls.filter((c) => c[0] === 'log').map((c) => c[1]).join('\n'));
  assert.ok(r.graph.requests.includes('https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/Signatures:?$select=id'), r.graph.requests.join('\n'));
});

test('the manifest names the add-in in German, French and Italian, for each of their locales', () => {
  const xml = fs.readFileSync(path.join(__dirname, '..', 'addin', 'manifest.template.xml'), 'utf8');
  const d = new JSDOM(xml, { contentType: 'text/xml' }).window.document;
  assert.equal(d.getElementsByTagName('parsererror').length, 0);
  const locales = ['de-DE', 'de-CH', 'fr-FR', 'fr-CH', 'it-IT', 'it-CH'];
  const over = (el) => Object.fromEntries([...el.children].filter((x) => x.localName === 'Override').map((x) => [x.getAttribute('Locale'), x.getAttribute('Value')]));
  const desc = over(d.getElementsByTagName('Description')[0]);
  const strings = Object.fromEntries([...d.getElementsByTagName('bt:String')].map((x) => [x.getAttribute('id'), over(x)]));
  for (const o of [desc, strings['Pane.Label'], strings['Pane.Tooltip']]) {
    assert.deepEqual(Object.keys(o).sort(), locales.slice().sort());
    assert.equal(o['de-CH'], o['de-DE']); assert.equal(o['fr-CH'], o['fr-FR']); assert.equal(o['it-CH'], o['it-IT']);
  }
  assert.equal(strings['Pane.Label']['fr-FR'], 'Signatures');
  assert.equal(strings['Pane.Label']['it-IT'], 'Firme');
});

test('signature HTML put into the pane: nothing in it runs or loads from a strange address', async () => {
  const { safeNodes, safeUrl } = await import(path.join(__dirname, '..', 'addin', 'src', 'safehtml.js'));
  const d = new JSDOM('<!doctype html><body></body>').window.document;
  const clean = (h) => { const box = d.createElement('div'); box.replaceChildren(...safeNodes(h, d)); return box.innerHTML; };
  assert.equal(clean('<p style="color:red" onclick="x()">Hi <b onmouseover="y()">you</b></p><style>p{margin:0}</style>'), '<p style="color:red">Hi <b>you</b></p><style>p{margin:0}</style>');
  assert.equal(clean('<script>x()</script><iframe src="https://x"></iframe><object data="x"></object><embed src="x"><p>ok</p>'), '<p>ok</p>');
  assert.equal(clean('<a href="java\tscript:alert(1)">a</a><a href=" &#106;avascript:alert(1)">b</a><a href="data:text/html,<b>x</b>">c</a>'), '<a>a</a><a>b</a><a>c</a>');
  assert.equal(clean('<a href="https://ex.ch/">a</a><a href="mailto:a@b.ch">b</a><a href="tel:+41">c</a><a href="kontakt.html">d</a><img src="data:image/png;base64,AA" data-osign-src="assets/logo.png"><img src="cid:x">'),
    '<a href="https://ex.ch/">a</a><a href="mailto:a@b.ch">b</a><a href="tel:+41">c</a><a href="kontakt.html">d</a><img src="data:image/png;base64,AA" data-osign-src="assets/logo.png"><img src="cid:x">');
  // Read differently a second time (noscript): it is never written as text again.
  assert.equal(clean('<noscript><p title="</noscript><img src=x onerror=alert(1)>"></p></noscript><i>x</i>'), '<i>x</i>');
  assert.equal(clean('<p style="width:expression(alert(1))">x</p><svg><a xlink:href="javascript:alert(1)"><text>t</text></a></svg>'), '<p>x</p><svg><a><text>t</text></a></svg>');
  for (const bad of ['javascript:x', 'JaVaScRiPt:x', 'vbscript:x', '\u0001javascript:x', 'java\nscript:x', 'data:text/html,x', 'data:application/x', 'file:///c:/x']) assert.equal(safeUrl(bad), false, JSON.stringify(bad));
  for (const ok of ['https://x', 'http://x', 'mailto:a@b', 'tel:+41', 'cid:a', 'data:image/png;base64,AA', 'assets/logo.png', '#top', '']) assert.equal(safeUrl(ok), true, ok);
});
