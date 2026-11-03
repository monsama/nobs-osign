/*
 * Releasing a version in steps: first to a pilot group, then to everyone (or back); and at a time
 * set beforehand. Routes for server/app.js, given what they need from it.
 *
 *   pilot/get       what runs as a pilot: { pilot: { group, at, by } | null }
 *   pilot/start     { configJson, group, message }: the new version for the group, the published
 *                   one stays for everyone else (both in signatures.json, see core.forUser)
 *   pilot/promote   the pilot's version for everyone
 *   pilot/end       the pilot stopped: everyone back on the version for all, as it was
 *   schedule/get    what is set to go out later: { schedule: { at, by, message, pilot, error } | null }
 *   schedule/set    { configJson, at (ISO), message, group? }: published then, by the server, with
 *                   the draft's images as they were when it was set
 *   schedule/cancel
 *
 * Publishing in steps is a publisher's: not with four eyes on (the approver publishes then). A
 * version that is not the draft (given to all, stopped, set for later) leaves the draft alone and
 * goes out with its own images, never with what the draft has by then.
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const core = require('../shared/signature-core.js');
const { writeAtomic } = require('./store');

const MAX_AHEAD_MS = 90 * 86400000;
const IMG = /^[A-Za-z0-9_-]{1,60}(\.[a-z0-9]{1,5})?$/;

function createReleaseRoutes({ dataDir, history, store, publishNow, publisher, needsApproval, compiledProblem, stillPublisher, audit, log }) {
  const scheduleFile = path.join(dataDir, 'schedule.json');
  const scheduleImages = path.join(dataDir, 'schedule-assets');
  let timer = null;
  const by = (s) => s.user.name || s.user.login;
  const parse = (text) => { try { const c = JSON.parse(text); return c && typeof c === 'object' ? c : null; } catch (e) { return null; } };
  const published = () => parse(history.latest() || '');
  function may(s) {
    if (!publisher(s)) return 'Only the publishing group releases a version.';
    if (needsApproval()) return 'With approval on, a version goes out when it is approved: a pilot or a time is not for that.';
    return '';
  }
  const groupOk = (g) => /^[^\u0000-\u001f]{1,256}$/.test(String(g || '').trim());
  // The images a configuration uses (its signatures, and a pilot's), out of the ones given.
  function imagesOf(config, from) {
    const text = JSON.stringify(config), out = new Map();
    for (const m of text.matchAll(/assets\/([^"'\s)?#\\]+)/g)) {
      let n = m[1]; try { n = decodeURIComponent(n); } catch (e) { /* as it is */ }
      if (from.has(n)) out.set(n, from.get(n));
    }
    return out;
  }

  // The version for the group inside the published one; both read the same fields, and the ones
  // for all are kept apart to come back when the pilot stops.
  function withPilot(forAll, config, group, who, message) {
    const cfg = Object.assign({}, config); delete cfg.pilot;
    const out = Object.assign({}, forAll, { pilot: { group: String(group).trim(), at: new Date().toISOString(), by: who, message: String(message || '').slice(0, 200), baseFields: forAll.fields || [], config: cfg } });
    return core.withPilotFields(out);
  }
  // The version for all, as it was before any pilot.
  function forAllOf(base) {
    const all = Object.assign({}, base);
    if (all.pilot && Array.isArray(all.pilot.baseFields)) all.fields = all.pilot.baseFields;
    delete all.pilot;
    return all;
  }
  async function startPilot(configJson, group, who, message, draftImages, opts) {
    const base = published();
    if (!base) return { error: 'Publish once for everyone first: a pilot needs a version for all to keep.' };
    const config = parse(configJson);
    if (!config) return { error: 'The signatures are not readable.' };
    const errs = core.validate(config);
    if (errs.length) return { error: errs[0] };
    const all = forAllOf(base);
    const text = JSON.stringify(withPilot(all, config, group, who, message));
    // The images of both: the version for all keeps its own, the pilot brings its own.
    const images = new Map([...imagesOf(all, history.latestImageBytes()), ...imagesOf(config, draftImages)]);
    const r = await publishNow(text, who + ' (pilot: ' + String(group).trim() + ')', message || 'pilot for ' + group, Object.assign({ images }, opts || {}));
    return r.error ? r : Object.assign(r, { pilot: { group: String(group).trim() } });
  }

  // ---- the time set beforehand
  const readSchedule = () => { try { return JSON.parse(fs.readFileSync(scheduleFile, 'utf8')); } catch (e) { return null; } };
  const scheduleAnswer = () => { const x = readSchedule(); return { schedule: x ? { at: x.at, by: x.by, message: x.message, pilot: x.group || '', error: x.error || '' } : null }; };
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
    const images = keptImages();
    const r = x.group
      ? await startPilot(x.configJson, x.group, x.by, x.message, images, { draft: false })
      : await publishNow(x.configJson, x.by + ' (set for ' + x.at + ')', x.message, { images: imagesOf(parse(x.configJson) || {}, images), draft: false });
    if (r.error) return fail(r.error);
    dropSchedule();
    log.info('scheduled publish of ' + x.at + ' done' + (x.group ? ' (pilot ' + x.group + ')' : ''));
    audit.add({ who: x.login || x.by, ip: 'server', what: 'schedule/run', ok: true, detail: 'published as set for ' + x.at + (x.group ? ' (pilot ' + x.group + ')' : '') });
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
    'pilot/get': () => { const p = core.pilotOf(published() || {}); return { pilot: p ? { group: p.group, at: p.at, by: p.by, message: p.message || '' } : null }; },
    'pilot/start': async (d, s) => {
      const no = may(s); if (no) return { error: no };
      if (!groupOk(d.group)) return { error: 'Name the pilot group (as in AD), like OSign-Pilot.' };
      return startPilot(typeof d.configJson === 'string' ? d.configJson : '', d.group, by(s), d.message, store.draftImages());
    },
    'pilot/promote': async (d, s) => {
      const no = may(s); if (no) return { error: no };
      const base = published(), p = core.pilotOf(base || {});
      if (!p) return { error: 'No pilot is running.' };
      const all = Object.assign({}, p.config); delete all.pilot;
      return publishNow(JSON.stringify(all), by(s) + ' (pilot ' + p.group + ' for everyone)', d.message || 'pilot given to everyone', { images: imagesOf(all, history.latestImageBytes()), draft: false });
    },
    'pilot/end': async (d, s) => {
      const no = may(s); if (no) return { error: no };
      const base = published();
      if (!core.pilotOf(base || {})) return { error: 'No pilot is running.' };
      const all = forAllOf(base);
      return publishNow(JSON.stringify(all), by(s) + ' (pilot stopped)', d.message || 'pilot stopped, everyone on the version before', { images: imagesOf(all, history.latestImageBytes()), draft: false });
    },
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
      if (d.group && !groupOk(d.group)) return { error: 'Name the pilot group (as in AD), like OSign-Pilot.' };
      // Its images as they are now: what goes out then, whatever the draft has by that time.
      fs.rmSync(scheduleImages, { recursive: true, force: true });
      fs.mkdirSync(scheduleImages, { recursive: true });
      for (const [n, b] of imagesOf(config, store.draftImages())) if (IMG.test(n)) writeAtomic(path.join(scheduleImages, n), b);
      writeAtomic(scheduleFile, JSON.stringify({ at: new Date(at).toISOString(), by: by(s), login: s.user.local ? '' : s.user.login, message: String(d.message || '').slice(0, 200), group: d.group ? String(d.group).trim() : '', configJson: text }));
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
    'pilot/start': (d) => 'pilot for ' + String(d.group || '').slice(0, 80) + (d.message ? ': ' + d.message : ''),
    'pilot/promote': () => 'pilot given to everyone',
    'pilot/end': () => 'pilot stopped',
    'schedule/set': (d) => 'set to publish at ' + String(d.at || '').slice(0, 40) + (d.group ? ' (pilot ' + String(d.group).slice(0, 80) + ')' : '') + (d.message ? ': ' + d.message : ''),
    'schedule/cancel': () => 'scheduled publish cancelled'
  };
  return { routes, audit: auditText, stop: () => { if (timer) clearTimeout(timer); }, runSchedule };
}

module.exports = { createReleaseRoutes };
