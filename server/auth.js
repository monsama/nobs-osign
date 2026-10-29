/*
 * Who may use the editor: members of the AD sign-in group (server/ldap.js), and the local account
 * "admin" for the first setup and for when the directory cannot be reached. Its password is
 * OSIGN_ADMIN_PASSWORD_FILE (a Docker secret, made by the deploy script); without one there is no
 * local account.
 *
 * A session is a random id in an HttpOnly, SameSite=Strict cookie, kept in memory: a restart signs
 * everyone out. It ends 2 hours after the last request, and after 12 hours at the latest.
 */
'use strict';
const crypto = require('node:crypto');
const fs = require('node:fs');
const { ldapLogin, ldapProblem } = require('./ldap');

const COOKIE = 'osign_session';
const IDLE_MS = 2 * 60 * 60 * 1000;
const MAX_MS = 12 * 60 * 60 * 1000;
// Wrong passwords for one login: after this many within the window, that login waits.
const MAX_FAILS = 5;
const FAIL_WINDOW_MS = 15 * 60 * 1000;
const LOCAL_LOGIN = 'admin';

function readSecret(file, value) {
  if (file) { try { return fs.readFileSync(file, 'utf8').trim(); } catch (e) { return ''; } }
  return String(value || '').trim();
}
// Compared in constant time, through hashes so the lengths may differ.
function same(a, b) {
  const x = crypto.createHash('sha256').update(String(a)).digest(), y = crypto.createHash('sha256').update(String(b)).digest();
  return crypto.timingSafeEqual(x, y);
}
function cookieOf(req) {
  const m = new RegExp('(?:^|;\\s*)' + COOKIE + '=([A-Za-z0-9_-]{43})(?:;|$)').exec(String(req.headers.cookie || ''));
  return m ? m[1] : null;
}

function createAuth({ env, ldap, now = Date.now, log = console }) {
  const sessions = new Map();
  const fails = new Map();
  const adminPassword = () => readSecret(env.OSIGN_ADMIN_PASSWORD_FILE, env.OSIGN_ADMIN_PASSWORD);
  const secure = env.OSIGN_SECURE_COOKIE !== '0';

  function session(req) {
    const id = cookieOf(req);
    const s = id && sessions.get(id);
    if (!s) return null;
    const t = now();
    if (t - s.last > IDLE_MS || t - s.started > MAX_MS) { sessions.delete(id); return null; }
    s.last = t;
    return s;
  }
  function cookie(id, maxAgeSeconds) {
    return COOKIE + '=' + id + '; Path=/; HttpOnly; SameSite=Strict' + (secure ? '; Secure' : '') + '; Max-Age=' + maxAgeSeconds;
  }

  function throttled(key) {
    const f = fails.get(key);
    if (!f) return false;
    if (now() - f.first > FAIL_WINDOW_MS) { fails.delete(key); return false; }
    return f.count >= MAX_FAILS;
  }
  function failed(key) {
    const f = fails.get(key);
    if (!f || now() - f.first > FAIL_WINDOW_MS) fails.set(key, { first: now(), count: 1 });
    else f.count++;
  }

  // Ended sessions and old failures go, so nobody can fill the memory with made-up names.
  function sweep() {
    const t = now();
    for (const [id, s] of sessions) if (t - s.last > IDLE_MS || t - s.started > MAX_MS) sessions.delete(id);
    for (const [k, f] of fails) if (t - f.first > FAIL_WINDOW_MS) fails.delete(k);
  }

  // { cookie, user } when signed in; { status, error } when not.
  // Wrong passwords count per name and address: someone elsewhere cannot lock out a name (admin,
  // above all, the way in when AD is down) by typing wrong passwords for it.
  async function login(name, password, from) {
    sweep();
    const loginName = String(name || '').trim().slice(0, 256);
    const key = loginName.toLowerCase() + '|' + String(from || '');
    if (!loginName || !password) return { status: 400, error: 'Enter your sign-in name and password.' };
    if (throttled(key)) return { status: 429, error: 'Too many wrong passwords for this name. Wait a quarter of an hour, then try again.' };
    let user = null;
    if (loginName.toLowerCase() === LOCAL_LOGIN) {
      const pw = adminPassword();
      if (pw && same(pw, String(password))) user = { login: LOCAL_LOGIN, name: 'admin', local: true };
    } else if (ldap.url) {
      const problem = ldapProblem(ldap);
      if (problem) return { status: 503, error: 'Signing in with Active Directory is not set up right: ' + problem };
      let r;
      try { r = await ldapLogin(ldap, loginName, String(password)); } catch (e) {
        log.warn('sign-in: the directory could not be asked: ' + (e && e.message));
        return { status: 503, error: 'Active Directory cannot be reached. The local account admin still works.' };
      }
      if (r && !r.allowed) return { status: 403, error: 'You are not in the group that may use NOBS OSign. Ask IT to add you.' };
      if (r) user = { login: r.login, name: r.name, local: false };
    }
    if (!user) { failed(key); return { status: 401, error: 'The sign-in name or the password is wrong.' }; }
    fails.delete(key);
    const id = crypto.randomBytes(32).toString('base64url');
    sessions.set(id, { id, user, started: now(), last: now() });
    log.info('sign-in: ' + user.login);
    return { cookie: cookie(id, MAX_MS / 1000), user };
  }
  function logout(req) {
    const id = cookieOf(req);
    if (id) sessions.delete(id);
    return cookie('', 0);
  }
  // What the sign-in page offers.
  const methods = () => ({ local: !!adminPassword(), ad: !!ldap.url });
  return { session, login, logout, methods };
}

module.exports = { createAuth, COOKIE };
