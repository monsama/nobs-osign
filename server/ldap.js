/*
 * Sign-in against Active Directory: the user's own password is checked by binding with it, and
 * never stored. Only members of the sign-in group may use the editor. The same rules as
 * nobs-inventory's server/ldap.js, so both tools are set up alike.
 *
 *   OSIGN_LDAP_URL         ldaps://dc01.example.ch (636), or ldap://… with OSIGN_LDAP_STARTTLS=1
 *   OSIGN_AD_DOMAIN        the domain's DNS name, example.ch (read from the domain controller by
 *                          "Detect"). The rest follows from it unless set:
 *   OSIGN_LDAP_BIND        how a login becomes a bind name: {login}@example.ch
 *   OSIGN_LDAP_BASE        where users are searched: DC=example,DC=ch (or an OU in it)
 *   OSIGN_LDAP_GROUP       the sign-in group, by name (NOBS-OSign-Admins) or as a DN; nested
 *                          groups count either way.
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

// example.ch → DC=example,DC=ch
const dnOfDomain = (d) => (String(d || '').trim() ? String(d).trim().split('.').map((p) => 'DC=' + p).join(',') : '');
// DC=example,DC=ch (the DC= parts of any DN) → example.ch
const domainOfDn = (dn) => String(dn || '').split(',').map((x) => x.trim()).filter((x) => /^DC=/i.test(x)).map((x) => x.slice(3)).join('.');

function ldapConfig(env) {
  const domain = String(env.OSIGN_AD_DOMAIN || '').trim().toLowerCase();
  const searchBase = env.OSIGN_LDAP_BASE || dnOfDomain(domain);
  return {
    url: env.OSIGN_LDAP_URL || '',
    domain,
    // Without a format of its own: login@domain, which Active Directory takes for any account of
    // that domain, whatever UPN suffix the account has.
    bindTemplate: env.OSIGN_LDAP_BIND || (domain ? '{login}@' + domain : ''),
    searchBase,
    // The whole domain, where groups are found.
    domainBase: dnOfDomain(domainOfDn(searchBase)) || dnOfDomain(domain),
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
  if (!/^ldaps?:\/\/[^\s/]+\/?$/i.test(c.url)) return 'The address of the domain controller is like ldaps://dc01.example.ch.';
  if (!c.bindTemplate) return 'Enter the domain (like example.ch), or use Detect.';
  if (!c.bindTemplate.includes('{login}')) return 'The sign-in name format names where the login goes, like {login}@example.ch (Advanced). Leave it empty to use the domain.';
  if (!c.searchBase) return 'Enter the domain (like example.ch), or use Detect.';
  if (!c.group) return 'The sign-in group is not set: without it, every account in the directory could sign in.';
  if (!/^[A-Za-z][A-Za-z0-9-]{0,63}$/.test(c.loginAttr)) return 'The login attribute is the name of one attribute, such as sAMAccountName.';
  if (/^ldap:\/\//i.test(c.url) && !c.startTls && !c.allowPlain) return "The domain controller's address starts with ldap://, which sends passwords unencrypted. Use ldaps://, or turn on StartTLS.";
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

// The sign-in group's DN: as set, or found by its name (sAMAccountName or CN) in the domain. null
// when there is no such group.
async function groupDn(client, c) {
  if (isDn(c.group)) return c.group;
  const g = escapeFilter(c.group);
  const [hit] = await search(client, c.domainBase || c.searchBase, '(&(objectClass=group)(|(sAMAccountName=' + g + ')(cn=' + g + ')))', ['cn'], 2);
  return hit ? hit.dn : null;
}

// What the domain controller says of itself, without signing in (its RootDSE, which AD lets anyone
// read): the domain's DNS name and DN. For "Detect" in Server connections.
async function detect(c) {
  if (!/^ldaps?:\/\/[^\s/]+\/?$/i.test(c.url || '')) throw new Error('The address of the domain controller is like ldaps://dc01.example.ch.');
  const client = await connect(c);
  try {
    const [root] = await new Promise((resolve, reject) => client.search('', { scope: 'base', filter: '(objectClass=*)', attributes: ['defaultNamingContext', 'dnsHostName'] }, (err, res) => {
      if (err) return reject(err);
      const out = [];
      res.on('searchEntry', (e) => { const a = {}; for (const x of (e.pojo && e.pojo.attributes) || []) a[x.type.toLowerCase()] = [].concat(x.values); out.push(a); });
      res.on('error', reject);
      res.on('end', () => resolve(out));
    }));
    const dn = root && (root.defaultnamingcontext || [])[0];
    if (!dn) throw new Error('The domain controller answers, but does not say its domain. Enter it by hand.');
    return { domain: domainOfDn(dn).toLowerCase(), base: dn, host: (root.dnshostname || [])[0] || '' };
  } finally { close(client); }
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
    // In the group, directly or through nested groups (found by name, when a name is set).
    const group = await groupDn(client, c);
    let allowed = false;
    if (group) allowed = (await search(client, c.searchBase, '(&(' + attr + '=' + escapeFilter(login) + ')' + chain(group) + ')', ['cn'])).length > 0;
    else allowed = (entry.attrs.memberof || []).some((g) => cnOf(g).toLowerCase() === c.group.toLowerCase());
    return { login, name: (entry.attrs.displayname || [])[0] || login, allowed };
  } catch (e) {
    if (e instanceof ldap.InvalidCredentialsError) return null;
    throw e;
  } finally { close(client); }
}

// Whether the domain controller answers, with TLS as set ("Test" in Server connections); and,
// with an account to look with, whether the sign-in group is there.
async function testSignIn(c, lookWith) {
  const problem = ldapProblem(c); if (problem) throw new Error(problem);
  if (!c.url) throw new Error('Signing in with Active Directory is not set up.');
  const client = await connect(c);
  try {
    if (!lookWith) return { reached: true };
    await bind(client, lookWith.dn, lookWith.password);
    if (!isDn(c.group)) return { reached: true, group: !!(await groupDn(client, c)) };
    const found = await new Promise((resolve, reject) => client.search(c.group, { scope: 'base', filter: '(objectClass=*)', attributes: ['cn'] }, (err, res) => {
      if (err) return reject(err);
      let n = 0;
      res.on('searchEntry', () => { n++; });
      res.on('error', (e) => (e instanceof ldap.NoSuchObjectError ? resolve(0) : reject(e)));
      res.on('end', () => resolve(n));
    }));
    return { reached: true, group: found > 0 };
  } finally { close(client); }
}

module.exports = { testSignIn, detect, dnOfDomain, domainOfDn, ldapConfig, ldapProblem, ldapLogin, escapeFilter, cnOf, connect, close, bind, search, chain };
