/*
 * Sign-in against Active Directory: the user's own password is checked by binding with it, and
 * never stored. Only members of the sign-in group may use the editor. The same rules as
 * nobs-inventory's server/ldap.js, so both tools are set up alike.
 *
 *   OSIGN_LDAP_URL         ldaps://dc01.example.ch (636), or ldap://… with OSIGN_LDAP_STARTTLS=1
 *   OSIGN_LDAP_BIND        how a login becomes a bind name: {login}@example.ch
 *   OSIGN_LDAP_BASE        where users are searched: OU=Users,DC=example,DC=ch
 *   OSIGN_LDAP_GROUP       the sign-in group. As a DN (CN=OSign-Admins,OU=…), nested groups count;
 *                          as a plain name, only direct membership.
 *   OSIGN_LDAP_LOGIN_ATTR  the attribute a login is (sAMAccountName)
 *   OSIGN_LDAP_CA          a file with the CA certificate(s) the domain controller's comes from
 *   OSIGN_LDAP_ALLOW_PLAIN=1  allows ldap:// without StartTLS, which sends passwords readable
 */
'use strict';
const fs = require('node:fs');
const tls = require('node:tls');
const net = require('node:net');
const ldap = require('ldapjs');

const LOGIN_OK = /^[\w.@-]{1,64}$/;
const escapeFilter = (v) => String(v).replace(/[\\*()\0]/g, (c) => '\\' + c.charCodeAt(0).toString(16).padStart(2, '0'));
// The name in a DN's first part: CN=Doe\, John,OU=Staff → Doe, John.
function cnOf(dn) {
  const first = (String(dn).match(/^(?:[^,\\]|\\.)*/) || [''])[0];
  return first.replace(/^CN=/i, '').replace(/\\([0-9a-f]{2})/gi, (_, h) => String.fromCharCode(parseInt(h, 16))).replace(/\\(.)/g, '$1');
}
const isDn = (g) => /=/.test(String(g || ''));
// Active Directory's "in chain" rule: in the group directly or through groups nested in it.
const chain = (dn) => '(memberOf:1.2.840.113556.1.4.1941:=' + escapeFilter(dn) + ')';

function ldapConfig(env) {
  return {
    url: env.OSIGN_LDAP_URL || '',
    bindTemplate: env.OSIGN_LDAP_BIND || '',
    searchBase: env.OSIGN_LDAP_BASE || '',
    group: env.OSIGN_LDAP_GROUP || '',
    loginAttr: env.OSIGN_LDAP_LOGIN_ATTR || 'sAMAccountName',
    caFile: env.OSIGN_LDAP_CA || '',
    startTls: env.OSIGN_LDAP_STARTTLS === '1',
    allowPlain: env.OSIGN_LDAP_ALLOW_PLAIN === '1'
  };
}
// What is missing or wrong, in a sentence, or '' when sign-in with AD can work.
function ldapProblem(c) {
  if (!c.url) return '';
  if (!/^ldaps?:\/\/[^\s/]+\/?$/i.test(c.url)) return 'OSIGN_LDAP_URL is an address like ldaps://dc01.example.ch.';
  if (!c.bindTemplate.includes('{login}')) return 'OSIGN_LDAP_BIND names where the login goes, like {login}@example.ch.';
  if (!c.searchBase) return 'OSIGN_LDAP_BASE is not set (where the users are, like OU=Users,DC=example,DC=ch).';
  if (!c.group) return 'OSIGN_LDAP_GROUP is not set: without it, every account in the directory could sign in.';
  if (!/^[A-Za-z][A-Za-z0-9-]{0,63}$/.test(c.loginAttr)) return 'OSIGN_LDAP_LOGIN_ATTR is the name of one attribute, such as sAMAccountName.';
  if (/^ldap:\/\//i.test(c.url) && !c.startTls && !c.allowPlain) return 'OSIGN_LDAP_URL starts with ldap://, which sends passwords unencrypted. Use ldaps://, or set OSIGN_LDAP_STARTTLS=1.';
  return '';
}

async function connect(c) {
  // The CA named, on top of the public ones Node trusts; an empty file (no CA given) adds none.
  const extra = c.caFile ? fs.readFileSync(c.caFile, 'utf8').trim() : '';
  const ca = extra ? [...tls.rootCertificates, extra] : null;
  const opts = { url: c.url, timeout: 8000, connectTimeout: 8000 };
  const plain = /^ldap:\/\//i.test(c.url);
  if (!plain && ca) opts.tlsOptions = { ca };
  const client = (c.createClient || ldap.createClient)(opts);
  // Connection errors also reach the pending call; without a listener they would stop the server.
  client.on('error', () => {});
  if (plain && c.startTls) {
    const host = new URL(c.url).hostname.replace(/^\[|\]$/g, '');
    try {
      await new Promise((resolve, reject) => client.starttls(Object.assign(ca ? { ca } : {}, { host }, net.isIP(host) ? {} : { servername: host }), [], (e) => (e ? reject(e) : resolve())));
    } catch (e) { close(client); throw e; }
  }
  return client;
}
function close(client) { try { client.unbind(() => {}); } catch (e) { /* closed */ } }
const bind = (client, dn, password) => new Promise((resolve, reject) => client.bind(dn, password, (e) => (e ? reject(e) : resolve())));
// sizeLimit: at most that many (a list to pick from); the directory then says it has more, which
// is no error here.
function search(client, base, filter, attributes, sizeLimit) {
  return new Promise((resolve, reject) => {
    const out = [];
    const opts = sizeLimit ? { scope: 'sub', filter, attributes, sizeLimit } : { scope: 'sub', filter, attributes, paged: { pageSize: 500 } };
    client.search(base, opts, (err, res) => {
      if (err) return reject(err);
      res.on('searchEntry', (e) => {
        const attrs = {};
        for (const a of (e.pojo && e.pojo.attributes) || []) attrs[a.type.toLowerCase()] = [].concat(a.values);
        out.push({ dn: (e.pojo && e.pojo.objectName) || '', attrs });
      });
      res.on('pageError', () => {});
      res.on('error', (e) => (sizeLimit && e instanceof ldap.SizeLimitExceededError ? resolve(out) : reject(e)));
      res.on('end', () => resolve(out));
    });
  });
}

// { login, name, allowed } for a right password, null for a wrong one. Throws when the directory
// cannot be asked, so the caller can tell the two apart.
async function ldapLogin(c, login, password) {
  if (!password || !LOGIN_OK.test(String(login || ''))) return null;
  const client = await connect(c);
  try {
    await bind(client, c.bindTemplate.replace('{login}', login), password);
    const attr = c.loginAttr;
    // Asked for in lower case: AD does not mind, and a directory stand-in that compares names as
    // written still answers.
    const [entry] = await search(client, c.searchBase, '(' + attr + '=' + escapeFilter(login) + ')', ['displayname', 'memberof', attr.toLowerCase()]);
    if (!entry) return null;
    let allowed;
    if (isDn(c.group)) allowed = (await search(client, c.searchBase, '(&(' + attr + '=' + escapeFilter(login) + ')' + chain(c.group) + ')', ['cn'])).length > 0;
    else allowed = (entry.attrs.memberof || []).some((g) => cnOf(g).toLowerCase() === c.group.toLowerCase());
    return { login, name: (entry.attrs.displayname || [])[0] || login, allowed };
  } catch (e) {
    if (e instanceof ldap.InvalidCredentialsError) return null;
    throw e;
  } finally { close(client); }
}

module.exports = { ldapConfig, ldapProblem, ldapLogin, escapeFilter, cnOf, connect, close, bind, search, chain };
