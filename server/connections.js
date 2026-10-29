/*
 * The server's connections, set in the editor by the local account admin (Settings → Server
 * connections): the sign-in with Active Directory, the service account, Microsoft Graph and
 * Cloudflare Pages. The same values as the OSIGN_* settings in .env; what is saved here wins, and
 * a value left empty here comes from .env. So .env needs no more than OSIGN_HOST.
 *
 * The values are kept in connections.json in the data volume. The passwords and the token are
 * sealed there with the server's key (OSIGN_APP_KEY_FILE, secrets/app_key, made by the deploy
 * script), with AES-256-GCM; they are never sent back to the page, only whether one is set. A CA
 * certificate for the domain controller is kept beside it (ldap-ca.pem).
 *
 * The service account's password from secrets/ldap_service_password goes only to the domain
 * controller .env names (OSIGN_LDAP_URL): an address changed here does not get it, so it has to
 * be entered here again. Only the local account admin may change any of this: whoever sets the
 * directory's address gets the passwords people sign in with.
 */
'use strict';
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { writeAtomic, readText } = require('./store');

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DN = /^[^\u0000-\u001f]{1,500}$/;
// The ports a directory address set here may use: LDAP's own (389, 636; 3268, 3269 for the global
// catalog) and those OSIGN_LDAP_PORTS allows. So nobody signed in can make the server knock on
// other services in the network.
const LDAP_PORTS = [389, 636, 3268, 3269];
const hostOf = (u) => { try { return new URL(u).hostname.toLowerCase(); } catch (e) { return ''; } };

// What can be set, the setting in .env it stands for, and what a value has to look like.
const FIELDS = {
  ldapUrl: { env: 'OSIGN_LDAP_URL', check: (v, env) => {
    if (!/^ldaps?:\/\/[^\s/]+\/?$/i.test(v)) return 'The address of the domain controller is like ldaps://dc01.example.ch.';
    let port;
    try { port = Number(new URL(v).port) || (/^ldaps:/i.test(v) ? 636 : 389); } catch (e) { return 'The address of the domain controller is like ldaps://dc01.example.ch.'; }
    const extra = String(env.OSIGN_LDAP_PORTS || '').split(',').map((x) => Number(x.trim())).filter(Boolean);
    if (![...LDAP_PORTS, ...extra].includes(port)) return 'Port ' + port + ' is not a directory port. Use 389 or 636 (3268 or 3269 for the global catalog), or allow it on the server with OSIGN_LDAP_PORTS=' + port + '.';
    return '';
  } },
  // ldap://: StartTLS or not; ldaps:// is encrypted anyway.
  ldapStartTls: { env: 'OSIGN_LDAP_STARTTLS', check: (v) => (['on', 'off'].includes(v) ? '' : 'StartTLS is on or off.'), toEnv: (v) => (v === 'on' ? '1' : '0') },
  ldapBind: { env: 'OSIGN_LDAP_BIND', check: (v) => (v.includes('{login}') && DN.test(v) ? '' : 'The sign-in name names where the login goes, like {login}@example.ch.') },
  ldapBase: { env: 'OSIGN_LDAP_BASE', check: (v) => (DN.test(v) ? '' : 'Where the users are is a DN, like OU=Users,DC=example,DC=ch.') },
  ldapGroup: { env: 'OSIGN_LDAP_GROUP', check: (v) => (DN.test(v) ? '' : 'The sign-in group is a DN, like CN=NOBS-OSign-Admins,OU=Groups,DC=example,DC=ch.') },
  ldapLoginAttr: { env: 'OSIGN_LDAP_LOGIN_ATTR', check: (v) => (/^[A-Za-z][A-Za-z0-9-]{0,63}$/.test(v) ? '' : 'The login attribute is the name of one attribute, such as sAMAccountName.') },
  ldapServiceDn: { env: 'OSIGN_LDAP_SERVICE_DN', check: (v) => (DN.test(v) ? '' : 'The service account is a DN, like CN=svc-osign,OU=Service,DC=example,DC=ch.') },
  ldapDirectoryBase: { env: 'OSIGN_LDAP_DIRECTORY_BASE', check: (v) => (DN.test(v) ? '' : 'Where users and groups are read is a DN, like DC=example,DC=ch.') },
  adExportAt: { env: 'OSIGN_AD_EXPORT_AT', check: (v) => (/^([01]?\d|2[0-3]):[0-5]\d$/.test(v) ? '' : 'The daily export is a time of day, like 06:00.') },
  graphTenantId: { env: 'OSIGN_GRAPH_TENANT_ID', check: (v) => (GUID.test(v) ? '' : 'The tenant ID is the Directory (tenant) ID, a GUID.') },
  graphClientId: { env: 'OSIGN_GRAPH_CLIENT_ID', check: (v) => (GUID.test(v) ? '' : "The client ID is the server's app registration's Application (client) ID, a GUID.") },
  cfAccountId: { env: 'OSIGN_CF_ACCOUNT_ID', check: (v) => (/^[0-9a-f]{32}$/i.test(v) ? '' : "The Cloudflare account ID is 32 letters and digits, on the account's overview page.") },
  cfProject: { env: 'OSIGN_CF_PROJECT', check: (v) => (/^[a-z0-9][a-z0-9-]{0,57}$/.test(v) ? '' : 'The Pages project is a name of lower-case letters, digits and -.') }
};
// The secrets, and their file in ./secrets.
const SECRETS = {
  ldapServicePassword: { env: 'OSIGN_LDAP_SERVICE_PASSWORD', file: 'secrets/ldap_service_password' },
  graphClientSecret: { env: 'OSIGN_GRAPH_CLIENT_SECRET', file: 'secrets/graph_client_secret' },
  cfApiToken: { env: 'OSIGN_CF_API_TOKEN', file: 'secrets/cf_api_token' }
};
const MAX_SECRET = 2000;

function readFileValue(file) { try { return fs.readFileSync(file, 'utf8').trim(); } catch (e) { return ''; } }
// A secret as .env gives it: its file (a Docker secret) or, without one, the value itself.
const envSecret = (env, name) => (env[name + '_FILE'] ? readFileValue(env[name + '_FILE']) : String(env[name] || '').trim());

// The CA certificates in a file as Windows exports it (.cer, DER or Base-64) or a PEM with the
// chain. Only CA certificates: the domain controller's own cannot vouch for itself once renewed.
function parseCa(base64) {
  const buf = Buffer.from(String(base64 || ''), 'base64');
  if (!buf.length || buf.length > 200000) return { error: 'Choose the CA certificate file (.cer, .crt or .pem).' };
  const text = buf.toString('latin1');
  const blocks = /-----BEGIN CERTIFICATE-----/.test(text) ? text.match(/-----BEGIN CERTIFICATE-----[\s\S]+?-----END CERTIFICATE-----/g) || [] : [buf];
  const certs = [];
  for (const b of blocks.slice(0, 10)) {
    try { certs.push(new crypto.X509Certificate(b)); } catch (e) { return { error: 'This file is not a certificate. Export the CA certificate as .cer (DER or Base-64).' }; }
  }
  if (!certs.length) return { error: 'This file has no certificate in it.' };
  const cn = (n) => (String(n || '').match(/CN=([^\n,]+)/) || [])[1] || String(n || '').split('\n')[0];
  const leaf = certs.find((c) => !c.ca);
  if (leaf) return { error: cn(leaf.subject) + " is a server's own certificate, not a CA's. Export the certificate of the CA that issued it (" + cn(leaf.issuer) + ').' };
  return {
    pem: certs.map((c) => c.toString().trim()).join('\n') + '\n',
    certs: certs.map((c) => ({ name: cn(c.subject), expires: new Date(c.validTo).toISOString() }))
  };
}

function createConnections({ dataDir, env }) {
  const file = path.join(dataDir, 'connections.json');
  const caFile = path.join(dataDir, 'ldap-ca.pem');
  // The key: whatever secrets/app_key holds, made into 32 bytes.
  function key() {
    const raw = env.OSIGN_APP_KEY_FILE ? readFileValue(env.OSIGN_APP_KEY_FILE) : String(env.OSIGN_APP_KEY || '').trim();
    return raw.length >= 16 ? crypto.createHash('sha256').update(raw).digest() : null;
  }
  function seal(name, value) {
    const iv = crypto.randomBytes(12);
    const c = crypto.createCipheriv('aes-256-gcm', key(), iv);
    c.setAAD(Buffer.from(name));
    const data = Buffer.concat([c.update(value, 'utf8'), c.final()]);
    return { iv: iv.toString('base64'), tag: c.getAuthTag().toString('base64'), data: data.toString('base64'), at: new Date().toISOString() };
  }
  // The value, or '' when it cannot be opened (no key, or another one).
  function open(name, s) {
    const k = key();
    if (!k || !s) return '';
    try {
      const d = crypto.createDecipheriv('aes-256-gcm', k, Buffer.from(s.iv, 'base64'));
      d.setAAD(Buffer.from(name)); d.setAuthTag(Buffer.from(s.tag, 'base64'));
      return Buffer.concat([d.update(Buffer.from(s.data, 'base64')), d.final()]).toString('utf8');
    } catch (e) { return ''; }
  }
  function load() {
    let saved = {};
    try { saved = JSON.parse(readText(file) || '{}') || {}; } catch (e) { saved = {}; }
    const values = {};
    for (const k of Object.keys(FIELDS)) if (typeof (saved.values || {})[k] === 'string' && saved.values[k]) values[k] = saved.values[k];
    const secrets = {};
    for (const k of Object.keys(SECRETS)) if ((saved.secrets || {})[k] && typeof saved.secrets[k].data === 'string') secrets[k] = saved.secrets[k];
    return { values, secrets, ca: saved.ca && fs.existsSync(caFile) ? saved.ca : null };
  }

  // The settings as the rest of the server reads them: .env, with what is saved here on top.
  function effectiveEnv() {
    const saved = load();
    const e = Object.assign({}, env);
    for (const [k, f] of Object.entries(FIELDS)) if (saved.values[k]) e[f.env] = f.toEnv ? f.toEnv(saved.values[k]) : saved.values[k];
    const here = {};
    for (const [k, s] of Object.entries(SECRETS)) {
      here[k] = saved.secrets[k] ? open(k, saved.secrets[k]) : '';
      if (here[k]) { e[s.env] = here[k]; delete e[s.env + '_FILE']; }
    }
    if (saved.ca) e.OSIGN_LDAP_CA = caFile;
    // The file's service password, only for the domain controller .env names.
    const envHost = hostOf(env.OSIGN_LDAP_URL || '');
    if (!here.ldapServicePassword && envHost && hostOf(e.OSIGN_LDAP_URL) !== envHost && envSecret(env, SECRETS.ldapServicePassword.env)) {
      e.OSIGN_LDAP_SERVICE_PASSWORD = ''; delete e.OSIGN_LDAP_SERVICE_PASSWORD_FILE;
      e.OSIGN_PASSWORD_PINNED = envHost;
    }
    return e;
  }

  // For the page: what is saved, what .env says, and for each secret only where it comes from.
  function answer() {
    const saved = load();
    const fields = {};
    for (const [k, f] of Object.entries(FIELDS)) {
      let fromEnv = String(env[f.env] || '').trim();
      if (k === 'ldapStartTls') fromEnv = fromEnv === '1' ? 'on' : '';
      fields[k] = { value: saved.values[k] || '', env: fromEnv };
    }
    const secrets = {};
    for (const [k, s] of Object.entries(SECRETS)) {
      const here = saved.secrets[k];
      secrets[k] = { saved: here ? here.at : '', readable: here ? !!open(k, here) : false, file: !!envSecret(env, s.env), fileName: s.file };
    }
    const e = effectiveEnv();
    return {
      keyOk: !!key(), fields, secrets,
      ca: { saved: saved.ca, env: env.OSIGN_LDAP_CA && readFileValue(env.OSIGN_LDAP_CA) ? env.OSIGN_LDAP_CA : '' },
      passwordPinned: e.OSIGN_PASSWORD_PINNED || ''
    };
  }

  // values: field → text ('' goes back to .env); secrets: name → text to set, or null to remove.
  function save(sent) {
    const saved = load();
    const values = Object.assign({}, saved.values);
    for (const [k, f] of Object.entries(FIELDS)) {
      if (!sent.values || typeof sent.values[k] !== 'string') continue;
      const v = sent.values[k].trim();
      if (!v) { delete values[k]; continue; }
      const err = f.check(v, env);
      if (err) return { error: err };
      values[k] = v;
    }
    const secrets = Object.assign({}, saved.secrets);
    for (const k of Object.keys(SECRETS)) {
      const v = sent.secrets ? sent.secrets[k] : undefined;
      if (v === null) { delete secrets[k]; continue; }
      if (typeof v !== 'string' || !v.trim()) continue;
      if (v.length > MAX_SECRET || /[\u0000-\u001f]/.test(v.trim())) return { error: 'A password or token is one line of text.' };
      if (!key()) return { error: 'The server has no key to keep passwords with (secrets/app_key). Run sudo ./deploy-nobs-osign.sh again, or put them in ./secrets.' };
      secrets[k] = seal(k, v.trim());
    }
    let ca = saved.ca;
    if (sent.ca === null) { ca = null; try { fs.unlinkSync(caFile); } catch (e) { /* none */ } }
    else if (typeof sent.ca === 'string' && sent.ca) {
      const r = parseCa(sent.ca);
      if (r.error) return { error: r.error };
      writeAtomic(caFile, r.pem);
      ca = { certs: r.certs, at: new Date().toISOString() };
    }
    writeAtomic(file, JSON.stringify({ values, secrets, ca }));
    return { saved: true };
  }

  return { effectiveEnv, answer, save };
}

module.exports = { createConnections, parseCa, FIELDS, SECRETS };
