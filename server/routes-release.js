/*
 * Publishing at a time set beforehand. Routes for server/app.js, given what they need from it.
 *
 *   schedule/get    what is set to go out later: { schedule: { at, by, message, error } | null }
 *   schedule/set    { configJson, at (ISO), message }: published then, by the server, with the
 *                   draft's images as they were when it was set
 *   schedule/cancel
 *
 * A publisher's: not with four eyes on (the approver publishes then). What goes out then leaves the
 * draft alone and goes with its own images, never with what the draft has by then.
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const core = require('../shared/signature-core.js');
const { writeAtomic } = require('./store');

const MAX_AHEAD_MS = 90 * 86400000;
const IMG = /^[A-Za-z0-9_-]{1,60}(\.[a-z0-9]{1,5})?$/;

function createReleaseRoutes({ dataDir, store, publishNow, publisher, needsApproval, compiledProblem, stillPublisher, audit, log }) {
  const scheduleFile = path.join(dataDir, 'schedule.json');
  const scheduleImages = path.join(dataDir, 'schedule-assets');
  let timer = null;
  const by = (s) => s.user.name || s.user.login;
  const parse = (text) => { try { const c = JSON.parse(text); return c && typeof c === 'object' ? c : null; } catch (e) { return null; } };
  function may(s) {
    if (!publisher(s)) return 'Only the publishing group releases a version.';
    if (needsApproval()) return 'With approval on, a version goes out when it is approved: a time set beforehand is not for that.';
    return '';
  }
  // The images a configuration uses, out of the ones given.
  function imagesOf(config, from) {
    const text = JSON.stringify(config), out = new Map();
    for (const m of text.matchAll(/assets\/([^"'\s)?#\\]+)/g)) {
      let n = m[1]; try { n = decodeURIComponent(n); } catch (e) { /* as it is */ }
      if (from.has(n)) out.set(n, from.get(n));
    }
    return out;
  }

  const readSchedule = () => { try { return JSON.parse(fs.readFileSync(scheduleFile, 'utf8')); } catch (e) { return null; } };
  const scheduleAnswer = () => { const x = readSchedule(); return { schedule: x ? { at: x.at, by: x.by, message: x.message, error: x.error || '' } : null }; };
  function keptImages() {
    const m = new Map();
    try { for (const n of fs.readdirSync(scheduleImages)) if (IMG.test(n)) m.set(n, fs.readFileSync(path.join(scheduleImages, n))); } catch (e) { /* none */ }
    return m;
  }
  function dropSchedule() { fs.rmSync(scheduleFile, { force: true }); fs.rmSync(scheduleImages, { recursive: true, force: true }); }
  async function runSchedule() {
    const x = readSchedule();
    if (!x) return;
    const fail = (error) => {
      writeAtomic(scheduleFile, JSON.stringify(Object.assign({}, x, { error, failedAt: new Date().toISOString() })));
      log.warn('scheduled publish of ' + x.at + ' not done: ' + error);
      audit.add({ who: x.login || x.by, ip: 'server', what: 'schedule/run', ok: false, detail: error });
    };
    // As it is now: four eyes turned on since, or its setter no longer publishing, it does not go.
    if (needsApproval()) return fail('Four eyes is on now: a version goes out when it is approved.');
    if (x.login && (await stillPublisher(x.login)) === false) return fail(x.by + ' is no longer in the publishing group.');
    const r = await publishNow(x.configJson, x.by + ' (set for ' + x.at + ')', x.message, { images: imagesOf(parse(x.configJson) || {}, keptImages()), draft: false });
    if (r.error) return fail(r.error);
    dropSchedule();
    log.info('scheduled publish of ' + x.at + ' done');
    audit.add({ who: x.login || x.by, ip: 'server', what: 'schedule/run', ok: true, detail: 'published as set for ' + x.at });
  }
  function arm() {
    if (timer) { clearTimeout(timer); timer = null; }
    const x = readSchedule();
    if (!x || x.error) return;
    // setTimeout holds at most about 24 days: waited for in steps.
    const wait = Math.max(0, Date.parse(x.at) - Date.now());
    timer = setTimeout(() => {
      if (Date.parse(x.at) > Date.now() + 1000) { arm(); return; }
      runSchedule().catch((e) => log.error('scheduled publish: ' + (e && e.stack || e))).finally(arm);
    }, Math.min(wait, 86400000));
    if (timer.unref) timer.unref();
  }
  arm();

  const routes = {
    'schedule/get': () => scheduleAnswer(),
    'schedule/set': (d, s) => {
      const no = may(s); if (no) return { error: no };
      const at = Date.parse(String(d.at || ''));
      if (!at || at < Date.now() + 60000) return { error: 'The time is in the future, a minute from now at least.' };
      if (at > Date.now() + MAX_AHEAD_MS) return { error: 'At most 90 days ahead.' };
      const text = typeof d.configJson === 'string' ? d.configJson : '';
      const config = parse(text);
      if (!config) return { error: 'The signatures are not readable.' };
      const errs = core.validate(config); if (errs.length) return { error: errs[0] };
      const bad = compiledProblem(text); if (bad) return { error: bad };
      // Its images as they are now: what goes out then, whatever the draft has by that time.
      fs.rmSync(scheduleImages, { recursive: true, force: true });
      fs.mkdirSync(scheduleImages, { recursive: true });
      for (const [n, b] of imagesOf(config, store.draftImages())) if (IMG.test(n)) writeAtomic(path.join(scheduleImages, n), b);
      writeAtomic(scheduleFile, JSON.stringify({ at: new Date(at).toISOString(), by: by(s), login: s.user.local ? '' : s.user.login, message: String(d.message || '').slice(0, 200), configJson: text }));
      arm();
      return scheduleAnswer();
    },
    'schedule/cancel': (d, s) => {
      if (!publisher(s)) return { error: 'Only the publishing group releases a version.' };
      dropSchedule(); arm();
      return scheduleAnswer();
    }
  };
  const auditText = {
    'schedule/set': (d) => 'set to publish at ' + String(d.at || '').slice(0, 40) + (d.message ? ': ' + d.message : ''),
    'schedule/cancel': () => 'scheduled publish cancelled'
  };
  return { routes, audit: auditText, stop: () => { if (timer) clearTimeout(timer); }, runSchedule };
}

module.exports = { createReleaseRoutes };
