/*
 * Active Directory for the editor, read with a service account over LDAP (the server version's
 * take on what editor/OSign.ps1 does with ADSI on a domain PC): a user's values and groups for the
 * preview, group names to pick from, and every user's values for signatures whose values come
 * from AD (users/<mail>.json next to signatures.json, see docs/API.md, ad/read and ad/export).
 *
 *   OSIGN_LDAP_SERVICE_DN             a read-only account: its login (svc-osign, signed in as
 *                                     svc-osign@<domain>), or its DN
 *   OSIGN_LDAP_SERVICE_PASSWORD_FILE  its password, as a Docker secret
 *   OSIGN_LDAP_DIRECTORY_BASE         where users and groups are searched; empty: the whole
 *                                     domain the sign-in base (OSIGN_LDAP_BASE) is in
 *
 * The connection settings (address, TLS, CA) are the sign-in's (server/ldap.js).
 */
'use strict';
const fs = require('node:fs');
const crypto = require('node:crypto');
const { connect, close, bind, search, escapeFilter, chain, ldapProblem } = require('./ldap');

// An attribute a field may name. The LAPS password and the BitLocker recovery data never: the
// values go into the preview and, exported, into SharePoint for every employee to read.
const SECRET_ATTR = /ms-?mcs-?admpwd|mslaps-|msfve-|unicodepwd|userpassword|supplementalcredentials|password|pwd|msds-managed|mstpm-|mspki|dpapi|credential|secret|bitlocker|recovery/i;
const attrOk = (a) => /^[A-Za-z][A-Za-z0-9-]{0,63}$/.test(a) && !SECRET_ATTR.test(a);
// Users with a mail address whose account is not disabled (userAccountControl flag 2).
const ACTIVE_USERS = '(&(objectCategory=person)(objectClass=user)(mail=*)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))';

function directoryConfig(env, ldapCfg) {
  let pw = '';
  if (env.OSIGN_LDAP_SERVICE_PASSWORD_FILE) { try { pw = fs.readFileSync(env.OSIGN_LDAP_SERVICE_PASSWORD_FILE, 'utf8').trim(); } catch (e) { pw = ''; } }
  else pw = String(env.OSIGN_LDAP_SERVICE_PASSWORD || '').trim();
  // The domain the sign-in base is in: its DC= parts.
  const dcs = String(ldapCfg.searchBase || '').split(',').map((x) => x.trim()).filter((x) => /^DC=/i.test(x)).join(',');
  // A login without a domain (svc-osign): as svc-osign@<domain>, as people sign in.
  let serviceDn = String(env.OSIGN_LDAP_SERVICE_DN || '').trim();
  if (serviceDn && !/[=@\\]/.test(serviceDn) && ldapCfg.domain) serviceDn += '@' + ldapCfg.domain;
  return Object.assign({}, ldapCfg, { serviceDn, servicePassword: pw, base: String(env.OSIGN_LDAP_DIRECTORY_BASE || '').trim() || dcs || ldapCfg.searchBase,
    // The password in ./secrets is for this domain controller only (server/connections.js).
    pinnedTo: String(env.OSIGN_PASSWORD_PINNED || '') });
}
const directoryEnabled = (c) => !!(c.url && c.serviceDn);
function directoryProblem(c) {
  if (!c.url) return 'Active Directory is not set up on the server (Settings → Server connections).';
  if (!c.serviceDn) return 'Enter the service account and its password (Settings → Server connections).';
  // An empty password is an anonymous sign-in to AD, which reads nothing or the wrong things.
  if (!c.servicePassword && c.pinnedTo) return "The service account's password in secrets/ldap_service_password is for " + c.pinnedTo + ', not for this domain controller. Enter it in Server connections.';
  if (!c.servicePassword) return "The service account's password is missing (secrets/ldap_service_password, or Server connections).";
  if (!c.base) return 'Where users and groups are read is not set (like DC=example,DC=ch).';
  return ldapProblem(c);
}

async function withService(c, fn) {
  const problem = directoryProblem(c);
  if (problem) throw new Error(problem);
  const client = await connect(c);
  try {
    try { await bind(client, c.serviceDn, c.servicePassword); }
    catch (e) { throw new Error('The service account could not sign in to Active Directory: ' + (e && e.message)); }
    return await fn(client);
  } finally { close(client); }
}
const first = (attrs, name) => { const v = attrs[String(name).toLowerCase()]; return v && v.length ? String(v[0]) : ''; };

// One user by sAMAccountName, mail or UPN: { values: { <key>: value }, displayName, groups }, the
// groups by name and sAMAccountName, nested ones included.
async function lookup(c, query, fields) {
  const q = String(query || '').trim();
  if (!q) throw new Error('Enter a user name or mail address.');
  const attrs = [...new Set((fields || []).map((f) => String(f.ad || '')).filter(attrOk).map((a) => a.toLowerCase()))];
  return withService(c, async (client) => {
    const v = escapeFilter(q);
    const found = await search(client, c.base, '(&(objectCategory=person)(objectClass=user)(|(sAMAccountName=' + v + ')(mail=' + v + ')(userPrincipalName=' + v + ')))', ['displayname', 'samaccountname', ...attrs], 2);
    if (!found.length) throw new Error('No user ' + q + ' in Active Directory.');
    if (found.length > 1) throw new Error('More than one user matches ' + q + '.');
    const u = found[0];
    const values = {};
    for (const f of fields || []) { if (f && f.key) values[f.key] = attrOk(String(f.ad || '')) ? first(u.attrs, f.ad) : ''; }
    const groups = new Set();
    // The rule IN_CHAIN follows group in group, as the add-in reads groups from Entra ID.
    for (const g of await search(client, c.base, '(&(objectClass=group)(member:1.2.840.113556.1.4.1941:=' + escapeFilter(u.dn) + '))', ['name', 'samaccountname'])) {
      for (const a of ['samaccountname', 'name']) { const n = first(g.attrs, a); if (n) groups.add(n); }
    }
    return { values, displayName: first(u.attrs, 'displayname') || q, groups: [...groups].sort((a, b) => a.localeCompare(b)) };
  });
}

// Up to 30 group names starting with what was typed.
async function groupsStartingWith(c, query) {
  const q = String(query || '').trim();
  if (q.length < 2) return [];
  return withService(c, async (client) => {
    const v = escapeFilter(q);
    const list = await search(client, c.base, '(&(objectClass=group)(|(sAMAccountName=' + v + '*)(name=' + v + '*)))', ['samaccountname'], 30);
    return [...new Set(list.map((g) => first(g.attrs, 'samaccountname')).filter(Boolean))].sort((a, b) => a.localeCompare(b));
  });
}

// Every active user with a mail address: { mail, attrs: { <attribute, lower case>: value },
// groups: [<the named groups they are in, nested ones included, lower case>] }. One search per
// group for its members, not one per user.
async function userRecords(c, attributes, groupNames) {
  const attrs = [...new Set((attributes || []).filter(attrOk).map((a) => a.toLowerCase()))];
  return withService(c, async (client) => {
    const memberOf = new Map();
    for (const name of groupNames || []) {
      const v = escapeFilter(name);
      for (const g of await search(client, c.base, '(&(objectClass=group)(|(sAMAccountName=' + v + ')(name=' + v + ')))', ['distinguishedname'])) {
        for (const u of await search(client, c.base, '(&(objectCategory=person)(objectClass=user)' + chain(g.dn) + ')', ['distinguishedname'])) {
          const key = u.dn.toLowerCase();
          if (!memberOf.has(key)) memberOf.set(key, new Set());
          memberOf.get(key).add(name);
        }
      }
    }
    return (await search(client, c.base, ACTIVE_USERS, ['mail', ...attrs])).map((u) => {
      const a = {};
      for (const x of attrs) a[x] = first(u.attrs, x);
      return { mail: first(u.attrs, 'mail'), attrs: a, groups: [...(memberOf.get(u.dn.toLowerCase()) || [])] };
    });
  });
}

const sha256 = (s) => crypto.createHash('sha256').update(s).digest('hex');

// Whether the service account signs in and sees users ("Test" in Server connections): how many,
// up to 5 (the first few are enough to know).
async function testService(c) {
  return withService(c, async (client) => ({ users: (await search(client, c.base, ACTIVE_USERS, ['samaccountname'], 5)).length }));
}

module.exports = { testService, attrOk, directoryConfig, directoryEnabled, directoryProblem, lookup, groupsStartingWith, userRecords, sha256 };
