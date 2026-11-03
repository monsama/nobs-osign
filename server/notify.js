/*
 * Telling someone when something needs doing, without them opening the editor: a message to a
 * webhook (a Teams channel's Workflows "When a Teams webhook request is received", Power Automate,
 * Slack or anything that takes JSON). Only outgoing, over HTTPS; set by the local account admin.
 *
 * What is told (problems(), from server/app.js): an automatic backup that failed, the server's
 * certificate or secret running out, the daily AD export that failed, a publish set for a time that
 * did not go out. Each one once a day at most while it lasts.
 *
 * notify.json: { url: <sealed>, sent: { <problem key>: <ISO time> } }
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { writeAtomic } = require('./store');

const DAY = 86400000;
const EVERY = 60 * 60 * 1000;

// Notifications to a webhook: its address (sealed), the problems told once a day, the timer.
function createNotify({ dataDir, seal, open, problems, log, fetchImpl, title = 'NOBS OSign' }) {
  const file = path.join(dataDir, 'notify.json');
  let timer = null;
  const read = () => { try { const s = JSON.parse(fs.readFileSync(file, 'utf8')); return s && typeof s === 'object' ? s : {}; } catch (e) { return {}; } };
  const write = (s) => writeAtomic(file, JSON.stringify(s, null, 2));
  const urlOf = (s) => { try { return s.url ? open('notifyWebhook', s.url) : ''; } catch (e) { return ''; } };
  const doFetch = fetchImpl || ((...a) => fetch(...a));

  // A Teams adaptive card, with the text also as "text" (Slack, Power Automate, most others).
  function body(lines) {
    const text = title + ': ' + lines.join(' · ');
    return {
      text,
      type: 'message',
      attachments: [{ contentType: 'application/vnd.microsoft.card.adaptive', content: {
        $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', type: 'AdaptiveCard', version: '1.4',
        body: [{ type: 'TextBlock', text: title, weight: 'Bolder', size: 'Medium' }].concat(lines.map((l) => ({ type: 'TextBlock', text: l, wrap: true })))
      } }]
    };
  }
  // One message to the webhook, at most 10 seconds, no redirects: '' or why it did not go.
  async function send(url, lines) {
    const ctl = new AbortController(), t = setTimeout(() => ctl.abort(), 10000);
    try {
      const r = await doFetch(url, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body(lines)), signal: ctl.signal, redirect: 'error' });
      if (!r.ok) return 'The webhook answered ' + r.status + '.';
      return '';
    } catch (e) { return 'The webhook could not be reached: ' + (e.name === 'AbortError' ? 'no answer within 10 seconds' : e.message); }
    finally { clearTimeout(t); }
  }
  // The settings for the editor: on or off, the last message, never the address.
  function answer() { const s = read(); return { on: !!urlOf(s), lost: !!s.url && !urlOf(s), last: s.last || null }; }
  // What is wrong with a webhook address, or ''.
  function check(u) {
    const x = String(u || '').trim();
    if (!x) return '';
    let p; try { p = new URL(x); } catch (e) { return 'The webhook is a web address (https://…).'; }
    if (p.protocol !== 'https:') return 'The webhook is an https:// address.';
    if (x.length > 2000) return 'That address is too long.';
    return '';
  }
  // { url: '' (off) | 'https://…' | undefined (kept) }
  function save(d) {
    const s = read();
    if (typeof d.url === 'string') {
      const bad = check(d.url); if (bad) return { error: bad };
      if (d.url.trim()) s.url = seal('notifyWebhook', d.url.trim()); else { delete s.url; s.sent = {}; }
    }
    write(s); schedule();
    return answer();
  }
  async function test() {
    const url = urlOf(read());
    if (!url) return { error: 'Set the webhook first.' };
    const err = await send(url, ['A test message: problems will be told here.']);
    return err ? { error: err } : { ok: true };
  }
  // The problems now, each told once a day at most; the ones gone are forgotten.
  async function run(now = Date.now()) {
    const s = read(), url = urlOf(s);
    if (!url) return { sent: 0 };
    let list = [];
    try { list = problems() || []; } catch (e) { log.warn('notify: ' + e.message); }
    const sent = s.sent || {}, keep = {}, due = [];
    for (const p of list) {
      const at = sent[p.key] ? Date.parse(sent[p.key]) : 0;
      if (!at || now - at >= DAY) due.push(p); else keep[p.key] = sent[p.key];
    }
    if (due.length) {
      const err = await send(url, due.map((p) => p.text));
      s.last = { at: new Date(now).toISOString(), ok: !err, error: err || '', count: due.length };
      if (!err) due.forEach((p) => { keep[p.key] = new Date(now).toISOString(); });
      else { log.warn('notify: ' + err); due.forEach((p) => { if (sent[p.key]) keep[p.key] = sent[p.key]; }); }
    }
    s.sent = keep; write(s);
    return { sent: due.length };
  }
  // The hourly check armed (and one a minute after the start), none while it is off.
  function schedule() {
    if (timer) { clearInterval(timer); timer = null; }
    if (!read().url) return;
    timer = setInterval(() => { run().catch((e) => log.warn('notify: ' + e.message)); }, EVERY);
    if (timer.unref) timer.unref();
    const first = setTimeout(() => { run().catch(() => {}); }, 60000);
    if (first.unref) first.unref();
  }
  function stop() { if (timer) clearInterval(timer); timer = null; }
  return { answer, save, test, run, schedule, stop };
}

module.exports = { createNotify };
