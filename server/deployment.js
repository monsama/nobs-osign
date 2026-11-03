/*
 * Where the Outlook add-in is deployed in Microsoft 365 (Integrated apps): a script only for
 * reading, for the local account admin, that asks Microsoft's add-in deployment with the admin's
 * own sign-in (Get-OrganizationAddIn) and sends what it found back here, with a one-time code
 * valid for an hour (deploy/finish, no session: it is not the browser). Microsoft offers no way
 * for the server to ask this itself.
 *
 * Kept in the data volume: addin-deployment.json, the last answer and when it came.
 */
'use strict';
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { writeAtomic, readText } = require('./store');

const TEMPLATE = path.join(__dirname, 'deployment-template.ps1');
const VALID_MS = 60 * 60 * 1000;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const psq = (s) => "'" + String(s).replace(/['‘’‚‛]/g, '$&$&') + "'";
const same = (a, b) => crypto.timingSafeEqual(crypto.createHash('sha256').update(String(a)).digest(), crypto.createHash('sha256').update(String(b)).digest());
// A property name as PowerShell gives them, and its value as text (control characters out).
const NAME = /^[A-Za-z][A-Za-z0-9_]{0,60}$/;
const clean = (v) => String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]/g, ' ').slice(0, 500);

// What the answer says, in a few words: found, on, version, who.
function summary(r) {
  if (!r) return null;
  if (!r.found) return { found: false };
  const p = r.props || {};
  const pick = (...names) => { for (const n of names) { const k = Object.keys(p).find((x) => x.toLowerCase() === n.toLowerCase()); if (k && p[k] !== '') return p[k]; } return ''; };
  const enabled = pick('Enabled');
  return {
    found: true,
    name: pick('DisplayName', 'Name'),
    version: pick('Version'),
    enabled: enabled === '' ? null : /^true$/i.test(enabled),
    assigned: pick('AssignedUsers', 'Members', 'AssignedTo'),
    everyone: /^true$/i.test(pick('AssignedToEveryone', 'Everyone')) || /^everyone$/i.test(pick('AssignedUsers'))
  };
}

// Where the add-in is deployed in Microsoft 365: the read-only script made, and its answer taken.
function createDeployment({ dataDir, now = Date.now }) {
  const file = path.join(dataDir, 'addin-deployment.json');
  let pending = null;

  function start({ origin, addinId }) {
    if (!/^https:\/\/[A-Za-z0-9.-]+(:\d+)?$|^http:\/\/(127\.0\.0\.1|localhost)(:\d+)?$/.test(String(origin || ''))) return { error: 'The editor\'s address is not one the script can send to (https://…).' };
    if (!GUID.test(String(addinId || ''))) return { error: 'The add-in has no ID.' };
    pending = { id: crypto.randomBytes(12).toString('base64url'), token: crypto.randomBytes(32).toString('base64url'), until: now() + VALID_MS, origin, addinId };
    return { id: pending.id, until: new Date(pending.until).toISOString() };
  }

  // The PowerShell script that reads the add-in's deployment, with the one-time token to send its answer.
  function script(id) {
    if (!pending || pending.id !== id || now() > pending.until) return null;
    const values = { __OSIGN_SERVER__: psq(pending.origin), __OSIGN_TOKEN__: psq(pending.token), __OSIGN_ADDIN_ID__: psq(pending.addinId) };
    let text = fs.readFileSync(TEMPLATE, 'utf8');
    for (const [k, v] of Object.entries(values)) text = text.split(k).join(v);
    return '﻿' + text;
  }

  // The script's answer: kept (the code is used up), or what went wrong.
  function finish(body) {
    if (!pending || now() > pending.until || !body || !same(String(body.token || ''), pending.token)) return { error: 'This script is not valid (any more): download it again under Settings → Server.', status: 403 };
    const at = new Date(now()).toISOString();
    let r;
    if (typeof body.error === 'string' && body.error) r = { at, error: clean(body.error) };
    else {
      const props = {};
      if (body.props && typeof body.props === 'object' && !Array.isArray(body.props)) {
        for (const [k, v] of Object.entries(body.props).slice(0, 40)) if (NAME.test(k)) props[k] = clean(v);
      }
      r = { at, found: body.found === true, props };
    }
    writeAtomic(file, JSON.stringify(r));
    pending = null;
    return { result: r };
  }

  // The last answer, with its summary; null before the first.
  function last() {
    let r = null;
    try { r = JSON.parse(readText(file) || 'null'); } catch (e) { r = null; }
    if (!r) return null;
    return Object.assign({}, r, { summary: r.error ? null : summary(r), waiting: !!(pending && now() <= pending.until) });
  }
  const waiting = () => !!(pending && now() <= pending.until);

  return { start, script, finish, last, waiting };
}

module.exports = { createDeployment, summary };
