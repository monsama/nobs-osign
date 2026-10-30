/*
 * The setup script (the setup assistant's step "Microsoft 365", for the local account admin): a
 * PowerShell script for the admin's Windows PC that signs in with Microsoft once, makes both app
 * registrations (the add-in's, the server's) and gives the server's the SharePoint site, then
 * sends what the server needs back here and takes back the rights it used.
 *
 * The script is made for one setup: it carries this server's address, the site, the add-in's
 * address and a one-time code. With that code, and only for an hour, it may send its result
 * (setup/finish, no session: it is not the browser). The result goes where the admin would type
 * it: the add-in's client ID into the settings, the server's client ID into the server connections.
 * The server signs in with its own certificate (server/cert.js): the script gets only its public
 * half and puts it on the server's app. Where that is refused, it makes a client secret instead,
 * sent here and sealed (server/connections.js).
 */
'use strict';
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const TEMPLATE = path.join(__dirname, 'setup-template.ps1');
const VALID_MS = 60 * 60 * 1000;
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
// A site's name: the letters a SharePoint address has, nothing else (it goes into a script).
const SITE = /^root$|^[A-Za-z0-9.-]+\.sharepoint\.com:\/(sites|teams|personal)\/[A-Za-z0-9._-]+:$/;
// A PowerShell string in single quotes: nothing in it is run. PowerShell ends such a string at
// the typographic single quotes too (‘ ’ ‚ ‛), so each of them is doubled as well.
const psq = (s) => "'" + String(s).replace(/['\u2018\u2019\u201A\u201B]/g, '$&$&') + "'";
const same = (a, b) => crypto.timingSafeEqual(crypto.createHash('sha256').update(String(a)).digest(), crypto.createHash('sha256').update(String(b)).digest());
const line = (v, max) => typeof v === 'string' && v.trim() && v.length <= max && !/[\u0000-\u001f]/.test(v.trim());

function createSetup({ now = Date.now } = {}) {
  // One setup at a time: the last one made.
  let pending = null, last = null;

  // spec: { origin, site, addinUrl, groups, knownAddin, knownServer } → { id } or { error }. The
  // known apps (client IDs from an earlier run) are the only ones the script uses again.
  function start(spec) {
    for (const k of ['knownAddin', 'knownServer']) if (spec[k] && !GUID.test(spec[k])) spec[k] = '';
    if (!/^https:\/\/[A-Za-z0-9.-]+(:\d+)?$|^http:\/\/(127\.0\.0\.1|localhost)(:\d+)?$/.test(String(spec.origin || ''))) return { error: 'The editor\'s address is not one the script can send to (https://…).' };
    if (!SITE.test(String(spec.site || ''))) return { error: 'Choose the SharePoint folder first (the step "Where the signatures are kept").' };
    if (!/^https:\/\/[^\s"<>'&]+$/.test(String(spec.addinUrl || ''))) return { error: 'Set up Cloudflare Pages first (its tab, before SharePoint), or enter the add-in\'s address in the setup assistant (the step "Where the add-in is served").' };
    pending = { id: crypto.randomBytes(12).toString('base64url'), token: crypto.randomBytes(32).toString('base64url'), until: now() + VALID_MS, spec: Object.assign({}, spec), error: '' };
    return { id: pending.id, until: new Date(pending.until).toISOString() };
  }

  // The script for that setup, or null.
  function script(id) {
    if (!pending || pending.id !== id || now() > pending.until) return null;
    const s = pending.spec;
    const values = {
      __OSIGN_SERVER__: psq(s.origin), __OSIGN_TOKEN__: psq(pending.token), __OSIGN_SITE__: psq(s.site),
      __OSIGN_ADDIN_HOST__: psq(new URL(s.addinUrl).host), __OSIGN_GROUPS__: s.groups ? '$true' : '$false', __OSIGN_USERS__: s.users ? '$true' : '$false',
      __OSIGN_KNOWN_ADDIN__: psq(s.knownAddin || ''), __OSIGN_KNOWN_SERVER__: psq(s.knownServer || ''),
      __OSIGN_CERT__: psq(s.cert || ''), __OSIGN_CERT_THUMBPRINT__: psq(s.certThumbprint || '')
    };
    let text = fs.readFileSync(TEMPLATE, 'utf8');
    for (const [k, v] of Object.entries(values)) text = text.split(k).join(v);
    // With a BOM, so Windows PowerShell 5.1 reads it as UTF-8.
    return '﻿' + text;
  }

  // What the script sends: its result, or what went wrong. → { data } to keep, or { error }.
  function finish(body) {
    if (!pending || now() > pending.until || !body || !same(String(body.token || ''), pending.token)) return { error: 'This setup script is not valid (any more): download it again in the setup assistant.', status: 403 };
    if (typeof body.error === 'string' && body.error) { pending.error = body.error.slice(0, 500); return { failed: true }; }
    const s = body.server || {}, a = body.addin || {};
    // The server signs in with its certificate (its thumbprint, as put on the app), or with a secret.
    const cert = /^[0-9A-F]{40}$/i.test(String(s.certThumbprint || '')) ? String(s.certThumbprint).toUpperCase() : '';
    if (!GUID.test(String(body.tenantId || '')) || !GUID.test(String(a.clientId || '')) || !GUID.test(String(s.clientId || '')) || (!cert && !line(s.secret, 500))) return { error: 'The script sent something that is not a setup result.', status: 400 };
    const data = { tenantId: body.tenantId, addinClientId: a.clientId, serverClientId: s.clientId, certThumbprint: cert, secret: cert ? '' : s.secret.trim(), secretExpires: String(s.secretExpires || '').slice(0, 40), siteUrl: String(body.siteUrl || '').slice(0, 500) };
    return { data };
  }
  // The result taken: once only, the code is used up. (Until then, a result the server could not
  // take, such as another certificate than this setup's, leaves the code as it was.)
  function done(data) {
    last = { at: new Date(now()).toISOString(), tenantId: data.tenantId, addinClientId: data.addinClientId, serverClientId: data.serverClientId, credential: data.certThumbprint ? 'certificate' : 'secret', secretExpires: data.secretExpires, siteUrl: data.siteUrl };
    pending = null;
  }

  function status() {
    if (pending && now() <= pending.until) return { state: pending.error ? 'failed' : 'waiting', problem: pending.error, until: new Date(pending.until).toISOString() };
    return last ? { state: 'done', result: last } : { state: 'none' };
  }

  return { start, script, finish, done, status };
}

module.exports = { createSetup };
