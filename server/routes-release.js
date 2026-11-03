/*
 * Releasing a version in steps: first to a pilot group, then to everyone (or back); and at a time
 * set beforehand. Routes for server/app.js, given what they need from it.
 *
 *   pilot/get       what runs as a pilot: { pilot: { group, at, by } | null }
 *   pilot/start     { configJson, group, message }: the new version for the group, the published
 *                   one stays for everyone else (both in signatures.json, see core.forUser)
 *   pilot/promote   the pilot's version for everyone
 *   pilot/end       the pilot stopped: everyone back on the version for all
 *   schedule/get    what is set to go out later: { schedule: { at, by, message, pilot } | null }
 *   schedule/set    { configJson, at (ISO), message, group? }: published then, by the server
 *   schedule/cancel
 *
 * Publishing in steps is a publisher's: not with four eyes on (the approver publishes then).
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const core = require('../shared/signature-core.js');
const { writeAtomic } = require('./store');

const MAX_AHEAD_MS = 90 * 86400000;

function createReleaseRoutes({ dataDir, history, publishNow, publisher, needsApproval, log }) {
  const scheduleFile = path.join(dataDir, 'schedule.json');
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

  // The version for the group inside the published one: both read the same fields.
  function withPilot(base, config, group, who, message) {
    const out = Object.assign({}, base, { pilot: { group: String(group).trim(), at: new Date().toISOString(), by: who, message: String(message || '').slice(0, 200), config } });
    delete out.pilot.config.pilot;
    return core.withPilotFields(out);
  }
  async function startPilot(configJson, group, who, message) {
    const base = published();
    if (!base) return { error: 'Publish once for everyone first: a pilot needs a version for all to keep.' };
    const config = parse(configJson);
    if (!config) return { error: 'The signatures are not readable.' };
    const errs = core.validate(config);
    if (errs.length) return { error: errs[0] };
    // The version for all is the one before any pilot.
    const forAll = Object.assign({}, base); delete forAll.pilot;
    const r = await publishNow(JSON.stringify(withPilot(forAll, config, group, who, message)), who + ' (pilot: ' + String(group).trim() + ')', message || 'pilot for ' + group);
    return r.error ? r : Object.assign(r, { pilot: { group: String(group).trim() } });
  }

  // ---- the time set beforehand
  const readSchedule = () => { try { return JSON.parse(fs.readFileSync(scheduleFile, 'utf8')); } catch (e) { return null; } };
  const scheduleAnswer = () => { const x = readSchedule(); return { schedule: x ? { at: x.at, by: x.by, message: x.message, pilot: x.group || '' } : null }; };
  async function runSchedule() {
    const x = readSchedule();
    if (!x) return;
    fs.rmSync(scheduleFile, { force: true });
    const r = x.group ? await startPilot(x.configJson, x.group, x.by, x.message) : await publishNow(x.configJson, x.by + ' (set for ' + x.at + ')', x.message);
    if (r.error) log.warn('scheduled publish of ' + x.at + ' failed: ' + r.error); else log.info('scheduled publish of ' + x.at + ' done' + (x.group ? ' (pilot ' + x.group + ')' : ''));
  }
  function arm() {
    if (timer) { clearTimeout(timer); timer = null; }
    const x = readSchedule();
    if (!x) return;
    // setTimeout holds at most about 24 days: waited for in steps.
    const wait = Math.max(0, Date.parse(x.at) - Date.now());
    timer = setTimeout(() => { if (Date.parse(x.at) <= Date.now() + 1000) runSchedule().finally(arm); else arm(); }, Math.min(wait, 86400000));
    if (timer.unref) timer.unref();
  }
  arm();

  const routes = {
    'pilot/get': () => { const p = core.pilotOf(published() || {}); return { pilot: p ? { group: p.group, at: p.at, by: p.by, message: p.message || '' } : null }; },
    'pilot/start': async (d, s) => {
      const no = may(s); if (no) return { error: no };
      if (!groupOk(d.group)) return { error: 'Name the pilot group (as in AD), like OSign-Pilot.' };
      return startPilot(typeof d.configJson === 'string' ? d.configJson : '', d.group, by(s), d.message);
    },
    'pilot/promote': async (d, s) => {
      const no = may(s); if (no) return { error: no };
      const p = core.pilotOf(published() || {});
      if (!p) return { error: 'No pilot is running.' };
      const all = Object.assign({}, p.config); delete all.pilot;
      return publishNow(JSON.stringify(all), by(s) + ' (pilot ' + p.group + ' for everyone)', d.message || 'pilot given to everyone');
    },
    'pilot/end': async (d, s) => {
      const no = may(s); if (no) return { error: no };
      const base = published();
      if (!core.pilotOf(base || {})) return { error: 'No pilot is running.' };
      const all = Object.assign({}, base); delete all.pilot;
      return publishNow(JSON.stringify(all), by(s) + ' (pilot stopped)', d.message || 'pilot stopped, everyone on the version before');
    },
    'schedule/get': () => scheduleAnswer(),
    'schedule/set': (d, s) => {
      const no = may(s); if (no) return { error: no };
      const at = Date.parse(String(d.at || ''));
      if (!at || at < Date.now() + 60000) return { error: 'The time is in the future, a minute from now at least.' };
      if (at > Date.now() + MAX_AHEAD_MS) return { error: 'At most 90 days ahead.' };
      const config = parse(typeof d.configJson === 'string' ? d.configJson : '');
      if (!config) return { error: 'The signatures are not readable.' };
      const errs = core.validate(config); if (errs.length) return { error: errs[0] };
      if (d.group && !groupOk(d.group)) return { error: 'Name the pilot group (as in AD), like OSign-Pilot.' };
      writeAtomic(scheduleFile, JSON.stringify({ at: new Date(at).toISOString(), by: by(s), login: s.user.login, message: String(d.message || '').slice(0, 200), group: d.group ? String(d.group).trim() : '', configJson: d.configJson }));
      arm();
      return scheduleAnswer();
    },
    'schedule/cancel': (d, s) => {
      if (!publisher(s)) return { error: 'Only the publishing group releases a version.' };
      fs.rmSync(scheduleFile, { force: true }); arm();
      return scheduleAnswer();
    }
  };
  const audit = {
    'pilot/start': (d) => 'pilot for ' + String(d.group || '').slice(0, 80) + (d.message ? ': ' + d.message : ''),
    'pilot/promote': () => 'pilot given to everyone',
    'pilot/end': () => 'pilot stopped',
    'schedule/set': (d) => 'set to publish at ' + String(d.at || '').slice(0, 40) + (d.group ? ' (pilot ' + String(d.group).slice(0, 80) + ')' : '') + (d.message ? ': ' + d.message : ''),
    'schedule/cancel': () => 'scheduled publish cancelled'
  };
  return { routes, audit, stop: () => { if (timer) clearTimeout(timer); }, runSchedule };
}

module.exports = { createReleaseRoutes };
