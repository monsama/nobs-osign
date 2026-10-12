// node --test tests/addin.test.js   (after: node addin/build.mjs)
//
// Runs the built dist/launchevent.js the way classic Outlook on Windows does: a bare JavaScript
// runtime with no window, no document and no localStorage, and an Office object that records what
// the handler asks of it.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { JSDOM } = require('jsdom');
const core = require('../shared/signature-core.js');

const dist = path.join(__dirname, '..', 'dist');
const doc = new JSDOM('').window.document;

function makeConfig(over) {
  const c = core.defaultConfig();
  c.assetBaseUrl = 'https://monsama.github.io/osign/config/';
  c.signatures = [
    Object.assign(core.newSignature('std', 'Standard'), { defaultNew: true, html: '<p><b>@@givenName @@sn</b></p><p>Mobile @@mobile</p><p><img src="assets/logo.png"></p>' }),
    Object.assign(core.newSignature('short', 'Short'), { defaultReply: true, html: '<p>@@givenName</p>' }),
    Object.assign(core.newSignature('other', 'Other'), { defaultNew: true, defaultReply: true, audience: { departments: [], domains: ['other.ch'] }, html: '<p>Other @@mail</p>' })
  ];
  c.signatures.forEach((s) => { s.compiled = core.compile(s.html, doc); });
  return Object.assign(c, over || {});
}

function runtime({ config, composeType = 'newMail', coercion = 'html', from = 'anna@monsama.ch', roaming = {}, itemType = 'message', fetchFails = false }) {
  const calls = [];
  const ok = (value) => (cb) => cb({ status: 'succeeded', value });
  const handlers = {};
  const item = {
    itemType,
    getComposeTypeAsync: (cb) => ok({ composeType, coercionType: coercion })(cb),
    from: { getAsync: (cb) => ok({ emailAddress: from })(cb) },
    disableClientSignatureAsync: (cb) => { calls.push(['disable']); ok()(cb); },
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
      requirements: { isSetSupported: () => false },
      roamingSettings: { get: (k) => settings[k], set: (k, v) => { settings[k] = v; }, saveAsync: (cb) => cb && cb({ status: 'succeeded' }) },
      mailbox: { item, userProfile: { displayName: 'Anna Muster', emailAddress: 'anna@monsama.ch' } }
    }
  };
  const fetched = [];
  const sandbox = {
    Office, console: { log: (m) => calls.push(['log', String(m)]) }, setTimeout, clearTimeout,
    fetch: async (url) => {
      fetched.push(url);
      if (fetchFails) throw new Error('offline');
      return { ok: true, status: 200, text: async () => JSON.stringify(config), json: async () => config };
    }
  };
  vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(dist, 'launchevent.js'), 'utf8'), sandbox, { filename: 'launchevent.js' });
  const fire = (name) => new Promise((resolve) => handlers[name]({ completed: () => resolve() }));
  return { calls, fetched, handlers, fire, settings };
}

const cachedProfile = (values) => ({ 'osign.profile': JSON.stringify({ at: Date.now(), values }) });
const anna = { givenName: 'Anna', sn: 'Muster', mobile: '', mail: 'anna@monsama.ch', department: 'Informatik' };

test('the build exists', () => {
  assert.ok(fs.existsSync(path.join(dist, 'launchevent.js')), 'run node addin/build.mjs first');
  const m = fs.readFileSync(path.join(dist, 'manifest.xml'), 'utf8');
  assert.match(m, /<Id>[0-9a-f-]{36}<\/Id>/);
  assert.match(m, /DefaultValue="https:\/\/[^"]+\/launchevent\.js"/);
  assert.doesNotMatch(m, /\{\{/);
});

test('registers its three handlers, in a runtime without window or DOM', () => {
  const r = runtime({ config: makeConfig() });
  assert.deepEqual(Object.keys(r.handlers).sort(), ['onMessageFromChangedHandler', 'onNewAppointmentOrganizerHandler', 'onNewMessageComposeHandler']);
});

test('new mail gets the default signature, filled, with empty lines left out', async () => {
  const r = runtime({ config: makeConfig(), roaming: cachedProfile(anna) });
  await r.fire('onNewMessageComposeHandler');
  const set = r.calls.find((c) => c[0] === 'set');
  assert.ok(set, JSON.stringify(r.calls));
  assert.equal(set[1], '<p><b>Anna Muster</b></p><p><img src="https://monsama.github.io/osign/config/assets/logo.png"></p>');
  assert.equal(set[2], 'html');
  assert.ok(r.calls.some((c) => c[0] === 'disable'));
  assert.match(r.fetched[0], /^https:\/\/monsama\.github\.io\/osign\/config\/signatures\.json\?m=\d+$/);
});

test('a reply gets the reply signature; plain text gets text', async () => {
  const r = runtime({ config: makeConfig(), composeType: 'reply', coercion: 'text', roaming: cachedProfile(anna) });
  await r.fire('onNewMessageComposeHandler');
  assert.deepEqual(r.calls.find((c) => c[0] === 'set').slice(1), ['Anna', 'text']);
});

test("the user's own choice wins", async () => {
  const roaming = Object.assign(cachedProfile(anna), { 'osign.prefs': JSON.stringify({ new: 'short', reply: 'none' }) });
  const r = runtime({ config: makeConfig(), roaming });
  await r.fire('onNewMessageComposeHandler');
  assert.equal(r.calls.find((c) => c[0] === 'set')[1], '<p>Anna</p>');
  const r2 = runtime({ config: makeConfig(), composeType: 'reply', roaming });
  await r2.fire('onNewMessageComposeHandler');
  assert.ok(!r2.calls.some((c) => c[0] === 'set'));
});

test('a changed sender swaps the signature', async () => {
  const r = runtime({ config: makeConfig(), from: 'anna@other.ch', roaming: cachedProfile(anna) });
  await r.fire('onMessageFromChangedHandler');
  assert.equal(r.calls.find((c) => c[0] === 'set')[1], '<p>Other anna@monsama.ch</p>');
});

test('without the user details and without the option, nothing is inserted', async () => {
  const r = runtime({ config: makeConfig() });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(!r.calls.some((c) => c[0] === 'set'));
  assert.ok(r.calls.some((c) => c[0] === 'log' && /no user details/.test(c[1])));
});

test('with insertWithoutProfile, the signature goes in with what Outlook knows', async () => {
  const cfg = makeConfig(); cfg.options.insertWithoutProfile = true;
  const r = runtime({ config: cfg });
  await r.fire('onNewMessageComposeHandler');
  // The name line had only empty fields, so it is left out.
  assert.equal(r.calls.find((c) => c[0] === 'set')[1], '<p><img src="https://monsama.github.io/osign/config/assets/logo.png"></p>');
});

test('meetings only when the option is on', async () => {
  const r = runtime({ config: makeConfig(), itemType: 'appointment', roaming: cachedProfile(anna) });
  await r.fire('onNewAppointmentOrganizerHandler');
  assert.ok(!r.calls.some((c) => c[0] === 'set'));
  const cfg = makeConfig(); cfg.options.meetings = true;
  const r2 = runtime({ config: cfg, itemType: 'appointment', roaming: cachedProfile(anna) });
  await r2.fire('onNewAppointmentOrganizerHandler');
  assert.ok(r2.calls.some((c) => c[0] === 'set'));
});

test('the event always completes, even when everything fails', async () => {
  const r = runtime({ config: makeConfig(), fetchFails: true });
  await r.fire('onNewMessageComposeHandler');
  assert.ok(r.calls.some((c) => c[0] === 'log' && /offline/.test(c[1])));
});
