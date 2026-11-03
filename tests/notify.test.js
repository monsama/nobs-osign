'use strict';
// Notifications to a webhook (server/notify.js): only https, each problem once a day, the address sealed.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createNotify } = require('../server/notify');

test('notifications: https only, each problem once a day while it lasts, the address kept sealed', async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-notify-'));
  const sent = [];
  let problems = [{ key: 'backup:1', text: 'The automatic backup failed.' }];
  const n = createNotify({ dataDir, seal: (k, v) => 'sealed:' + Buffer.from(v).toString('base64'), open: (k, v) => Buffer.from(v.slice(7), 'base64').toString(),
    problems: () => problems, log: { warn() {} }, fetchImpl: async (url, o) => { sent.push({ url, body: JSON.parse(o.body) }); return { ok: true, status: 200 }; } });
  assert.match(n.save({ url: 'http://hooks.example/x' }).error, /https/);
  assert.match(n.save({ url: 'not a url' }).error, /web address/);
  assert.equal((await n.test()).error, 'Set the webhook first.');
  assert.equal(n.save({ url: 'https://hooks.example/x' }).on, true);
  assert.ok(!fs.readFileSync(path.join(dataDir, 'notify.json'), 'utf8').includes('hooks.example'), 'sealed');
  const t0 = Date.parse('2026-11-03T10:00:00Z');
  assert.equal((await n.run(t0)).sent, 1);
  assert.match(sent[0].body.text, /automatic backup failed/);
  assert.equal(sent[0].body.attachments[0].contentType, 'application/vnd.microsoft.card.adaptive');
  assert.equal((await n.run(t0 + 3600000)).sent, 0, 'not again within a day');
  assert.equal((await n.run(t0 + 86400000)).sent, 1, 'again the next day');
  problems = [];
  assert.equal((await n.run(t0 + 2 * 86400000)).sent, 0);
  problems = [{ key: 'backup:1', text: 'The automatic backup failed.' }];
  assert.equal((await n.run(t0 + 2 * 86400000 + 1000)).sent, 1, 'gone and back: told again');
  assert.equal((await n.test()).ok, true);
  assert.equal(n.save({ url: '' }).on, false);
  n.stop();
});
