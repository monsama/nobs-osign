/*
 * The users' values from Active Directory, for signatures whose values come from AD
 * (options.valueSource 'ad' or 'mixed'): one file per user, users/<mail>.json in the publish
 * target, with the fields' values and the groups (of those the signatures name) the user is in.
 * The same rules as editor/OSign.ps1 (New-UserSnapshot, Write-UserSnapshot): read first and keep
 * what was read (a snapshot) for checking, then write exactly that, once.
 *
 * What the target holds is mirrored in the data volume (users-mirror.json: each file's content and
 * tag), so a read compares against it and asks SharePoint only for the files that changed since.
 */
'use strict';
const path = require('node:path');
const core = require('../shared/signature-core.js');
const { writeAtomic, readText } = require('./store');
const { userRecords, attrOk } = require('./directory');

const byKey = (a, b) => a.localeCompare(b, 'en', { sensitivity: 'base' });

// What a read is for: the fields (key and AD attribute) and the groups. Another set makes it
// another export: the signature (sig) tells them apart. (Read-ExportRequest)
function exportRequest(data) {
  const fields = [];
  for (const f of (data && data.fields) || []) {
    if (f && String(f.key || '') && attrOk(String(f.ad || ''))) fields.push({ key: String(f.key), ad: String(f.ad) });
  }
  const groups = [...new Set(((data && data.groups) || []).map((g) => String(g).trim().toLowerCase()).filter(Boolean))].sort(byKey);
  const sig = fields.slice().sort((a, b) => byKey(a.key, b.key)).map((f) => f.key + '=' + f.ad).join(';') + '|' + groups.join(';');
  return { fields, groups, sig };
}
// The same for what the published configuration asks for (the scheduled export).
function requestFor(config) {
  return exportRequest({ fields: core.adFields(config).map((f) => ({ key: f.key, ad: f.ad })), groups: core.groupsUsed(config) });
}

function createUsers({ dataDir, directory, log }) {
  const snapFile = path.join(dataDir, 'users-snapshot.json');
  const mirrorFile = path.join(dataDir, 'users-mirror.json');
  const load = (file) => { try { return JSON.parse(readText(file) || 'null'); } catch (e) { return null; } };

  // The files in users/ now, each with its content: from the mirror where the tag is the same.
  async function current(target) {
    const mirror = load(mirrorFile) || {};
    const now = {};
    for (const f of await target.listInfo('users')) {
      if (!/\.json$/i.test(f.name)) continue;
      const m = mirror[f.name];
      if (m && m.tag && m.tag === f.tag) { now[f.name] = m; continue; }
      const bytes = await target.read('users/' + f.name);
      now[f.name] = { tag: f.tag, json: bytes ? bytes.toString('utf8') : '' };
    }
    writeAtomic(mirrorFile, JSON.stringify(now));
    return now;
  }

  // What an export would do, without doing it: for every user the file it would write, and
  // whether that is new, changed or the same as there now; and the files that would go.
  async function read(target, req) {
    const records = await userRecords(directory, [...new Set(req.fields.map((f) => f.ad))], req.groups);
    const there = await current(target);
    const seen = new Set(), entries = [];
    const fields = req.fields.slice().sort((a, b) => byKey(a.key, b.key));
    for (const r of records) {
      const name = core.userFileName(r.mail);
      if (!name || seen.has(name)) continue;
      seen.add(name);
      const values = {};
      for (const f of fields) values[f.key] = String(r.attrs[f.ad.toLowerCase()] || '');
      const groups = r.groups.slice().sort(byKey);
      const json = JSON.stringify({ values, groups });
      const old = there[name] ? there[name].json : null;
      const status = old === null ? 'new' : old === json ? 'unchanged' : 'changed';
      let oldValues = null;
      if (status === 'changed') { try { oldValues = JSON.parse(old); } catch (e) { /* not JSON */ } }
      entries.push({ file: name, mail: r.mail, status, values, groups, old: oldValues, json });
    }
    const removed = Object.keys(there).filter((n) => !seen.has(n)).sort();
    const snap = { entries, removed, readAt: Date.now(), sig: req.sig, written: false };
    writeAtomic(snapFile, JSON.stringify(snap));
    return snap;
  }

  // Writes what a snapshot holds: the new and changed files, and removes the ones that went.
  async function write(target, snap) {
    // A read that found nobody (a directory that answered nothing) would remove every user's file.
    if (!snap.entries.length && snap.removed.length) throw new Error('Active Directory returned no users: the files in users/ are left as they are.');
    const mirror = load(mirrorFile) || {};
    let written = 0, same = 0, removed = 0;
    for (const e of snap.entries) {
      if (mirror[e.file] && mirror[e.file].json === e.json) { same++; continue; }
      const tag = await target.write('users/' + e.file, Buffer.from(e.json, 'utf8'), 'application/json');
      mirror[e.file] = { tag, json: e.json };
      written++;
      // Kept as it goes, so a write that stops halfway still knows what is there.
      if (written % 200 === 0) writeAtomic(mirrorFile, JSON.stringify(mirror));
    }
    for (const n of snap.removed) { await target.remove('users/' + n); delete mirror[n]; removed++; }
    writeAtomic(mirrorFile, JSON.stringify(mirror));
    return { users: snap.entries.length, written, unchanged: same, removed };
  }

  // For the UI: the users as read, without the file contents. (Get-SnapshotSummary)
  function summary(snap) {
    const count = (st) => snap.entries.filter((e) => e.status === st).length;
    return {
      readAt: snap.readAt, sig: snap.sig, written: !!snap.written,
      users: snap.entries.map((e) => ({ mail: e.mail, status: e.status, values: e.values, groups: e.groups, old: e.old ? e.old.values : null })),
      removed: snap.removed.map((n) => n.replace(/\.json$/, '')),
      counts: { new: count('new'), changed: count('changed'), unchanged: count('unchanged'), removed: snap.removed.length }
    };
  }

  return {
    // ad/read: read and keep, nothing written.
    async apiRead(target, data) { return summary(await read(target, exportRequest(data))); },
    // ad/snapshot: the last read, and whether it was read for these fields and groups.
    apiSnapshot(data) {
      const snap = load(snapFile);
      if (!snap || !Array.isArray(snap.entries)) return { snapshot: null, matches: false };
      return { snapshot: summary(snap), matches: snap.sig === exportRequest(data).sig };
    },
    // ad/export: exactly the last read, once; refused when it was read for other fields or
    // groups, or written already (an old read must not go over what a later export wrote).
    async apiExport(target, data) {
      const snap = load(snapFile);
      if (!snap || !Array.isArray(snap.entries)) return { error: 'Read the users from AD and check them first (Users).' };
      if (snap.sig !== exportRequest(data).sig) return { error: 'The fields or groups changed since the users were read: read them again (Users).' };
      if (snap.written) return { error: 'These users were written already: read them again to write newer values (Users).' };
      const r = await write(target, snap);
      snap.written = true;
      writeAtomic(snapFile, JSON.stringify(snap));
      return r;
    },
    // The scheduled export: read and write in one go, as the published configuration asks.
    async exportPublished(target) {
      const text = await target.read('signatures.json');
      if (!text) return { skipped: 'nothing is published yet' };
      const config = JSON.parse(text.toString('utf8'));
      if (!core.usesAdExport(config)) return { skipped: 'the signatures take their values from Entra ID' };
      const r = await write(target, await read(target, requestFor(config)));
      // The read the editor holds is older than what was just written: it must not go out.
      const snap = load(snapFile);
      if (snap) { snap.written = true; writeAtomic(snapFile, JSON.stringify(snap)); }
      return r;
    },
    log
  };
}

// The next time of day hh:mm (in the server's time zone, TZ) after now.
function nextAt(hhmm, now) {
  const m = /^([01]?\d|2[0-3]):([0-5]\d)$/.exec(String(hhmm || '').trim());
  if (!m) return null;
  const t = new Date(now);
  t.setHours(+m[1], +m[2], 0, 0);
  if (t.getTime() <= now) t.setDate(t.getDate() + 1);
  return t.getTime();
}

module.exports = { createUsers, exportRequest, requestFor, nextAt };
