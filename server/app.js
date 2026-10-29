/*
 * The editor's backend in the container: the same HTTP API as editor/OSign.ps1 (docs/API.md), for
 * the same web UI (editor/ui), behind a sign-in instead of a start token.
 *
 * It listens on plain HTTP inside the container; HTTPS comes from the shared proxy in front
 * (nobs-proxy, as for nobs-inventory). Nothing here is meant to be reachable from the internet.
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { createStore, ASSET_TYPES } = require('./store');
const { createSettings } = require('./settings');
const { createAuth } = require('./auth');
const { ldapConfig } = require('./ldap');

const VERSION = require('../package.json').version;
const ROOT = path.resolve(__dirname, '..');
const UI_DIR = path.join(ROOT, 'editor', 'ui');
const CORE_FILE = path.join(ROOT, 'shared', 'signature-core.js');
const STATIC_TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon' };
// A request body larger than this is refused: a configuration is at most 1.5 MB, an image 1 MB
// (as base64, 1.4 MB).
const MAX_BODY = 4 * 1024 * 1024;

// Never inside another site's frame, never read as another type; the page loads nothing but its
// own files, plus images from https (an image a signature links to, shown in the preview).
const SECURITY = {
  'X-Frame-Options': 'DENY',
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'no-referrer',
  'Content-Security-Policy': "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob: https:; connect-src 'self'; frame-src 'self'; object-src 'none'; base-uri 'self'; form-action 'none'; frame-ancestors 'self'"
};

function send(res, status, type, body, extra) {
  const buf = Buffer.isBuffer(body) ? body : Buffer.from(String(body || ''), 'utf8');
  res.writeHead(status, Object.assign({ 'Content-Type': type, 'Content-Length': buf.length, 'Cache-Control': 'no-store' }, SECURITY, extra || {}));
  res.end(buf);
}
const json = (res, status, obj, extra) => send(res, status, 'application/json; charset=utf-8', JSON.stringify(obj), extra);
const ok = (res, obj, extra) => json(res, 200, Object.assign({}, obj, { ok: true }), extra);
const fail = (res, error, status) => json(res, status || 200, { ok: false, error });

// A file under root, by a path from the request; 404 for anything outside it or of another type.
function sendFile(res, root, rel, types) {
  const base = path.resolve(root) + path.sep;
  const full = path.resolve(root, '.' + path.sep + String(rel || ''));
  const type = types[path.extname(full).toLowerCase()];
  let stat = null;
  try { stat = fs.statSync(full); } catch (e) { /* none */ }
  if (!full.startsWith(base) || !type || !stat || !stat.isFile()) return send(res, 404, 'text/plain; charset=utf-8', '404 Not Found');
  send(res, 200, type, fs.readFileSync(full));
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    const parts = []; let size = 0;
    req.on('data', (d) => { size += d.length; if (size > MAX_BODY) { reject(Object.assign(new Error('too large'), { status: 413 })); req.destroy(); } else parts.push(d); });
    req.on('end', () => resolve(Buffer.concat(parts).toString('utf8')));
    req.on('error', reject);
  });
}

// A request from this site: the browser names its origin on every POST (Origin), and the fetch
// metadata (Sec-Fetch-Site) when it has it. The cookie is SameSite=Strict as well; a JSON body
// can only come from another site after a preflight, which this server never answers.
function sameSite(req) {
  const site = req.headers['sec-fetch-site'];
  if (site && site !== 'same-origin' && site !== 'none') return false;
  const origin = req.headers.origin;
  if (!origin) return true;
  try { return new URL(origin).host === String(req.headers.host || ''); } catch (e) { return false; }
}

function createApp(opts) {
  const env = (opts && opts.env) || process.env;
  const log = (opts && opts.log) || console;
  const dataDir = env.OSIGN_DATA || '/data';
  const publishDir = env.OSIGN_PUBLISH_DIR || path.join(dataDir, 'publish');
  fs.mkdirSync(dataDir, { recursive: true });
  fs.mkdirSync(publishDir, { recursive: true });
  const store = createStore(dataDir);
  const settings = createSettings(dataDir, publishDir);
  const auth = createAuth({ env, ldap: Object.assign(ldapConfig(env), (opts && opts.ldap) || {}), log });

  const routes = {
    ping: () => ({}),
    info: (data, s) => ({
      backend: 'node', version: VERSION, dataDir,
      // Not here yet: AD lookup and export (LDAP), and the add-in's files (Cloudflare Pages).
      features: { adLookup: false, addinFiles: false, pickFolder: false, signOut: true },
      user: { login: s.user.login, name: s.user.name }
    }),
    'draft/get': () => store.getDraft(),
    'draft/save': (d) => store.saveDraft(typeof d.configJson === 'string' ? d.configJson : ''),
    'assets/upload': (d) => store.uploadAsset(d.name, d.data),
    'assets/delete': (d) => store.deleteAsset(d.name),
    'settings/get': () => settings.answer(),
    'settings/save': (d) => settings.save(d.settings || {}),
    publish: (d) => {
      const r = store.publish(publishDir, typeof d.configJson === 'string' ? d.configJson : '');
      return r.error ? r : { changed: r.changed, folder: publishDir };
    },
    'load-published': () => store.loadPublished(publishDir)
  };
  // What only Windows, OneDrive or (for now) the PowerShell editor can do: said plainly.
  const NOT_HERE = {
    'pick-folder': 'Choosing a folder needs Windows.',
    'sp/resolve': 'The server does not use a synced folder.',
    'sp/local': 'The server does not use a synced folder.',
    'ad/lookup': 'Looking up users in Active Directory is not in the server version yet.',
    'ad/groups': 'Looking up groups in Active Directory is not in the server version yet.',
    'ad/read': 'Reading users from Active Directory is not in the server version yet.',
    'ad/snapshot': null,
    'ad/export': 'Writing users from Active Directory is not in the server version yet.',
    'addin/build': "Making the add-in's files is not in the server version yet.",
    'addin/check': "Checking the add-in's files is not in the server version yet.",
    quit: 'The server keeps running; sign out instead.'
  };

  async function api(req, res, name) {
    if (req.method !== 'POST' || !/^application\/json\b/i.test(String(req.headers['content-type'] || '')) || !sameSite(req)) return send(res, 403, 'text/plain; charset=utf-8', '403 Forbidden');
    let data = {};
    try { const body = await readBody(req); data = body ? JSON.parse(body) : {}; } catch (e) {
      if (e.status === 413) return fail(res, 'The request is too large.', 413);
      return fail(res, 'The request is not valid JSON.');
    }
    if (!data || typeof data !== 'object' || Array.isArray(data)) data = {};

    if (name === 'login') {
      const r = await auth.login(data.login, data.password);
      if (r.error) return fail(res, r.error, r.status);
      return ok(res, { user: { login: r.user.login, name: r.user.name } }, { 'Set-Cookie': r.cookie });
    }
    if (name === 'login-methods') return ok(res, auth.methods());
    if (name === 'logout') return ok(res, {}, { 'Set-Cookie': auth.logout(req) });

    const s = auth.session(req);
    if (!s) return fail(res, 'Not signed in.', 401);
    if (name === 'ad/snapshot') return ok(res, { snapshot: null, matches: false });
    if (Object.prototype.hasOwnProperty.call(NOT_HERE, name)) return fail(res, NOT_HERE[name]);
    const fn = Object.prototype.hasOwnProperty.call(routes, name) ? routes[name] : null;
    if (!fn) return fail(res, 'unknown', 404);
    let r;
    try { r = fn(data, s); } catch (e) {
      log.error('api ' + name + ': ' + (e && e.stack || e));
      return fail(res, 'The server could not do that: ' + (e && e.message), 500);
    }
    if (r && r.error) return fail(res, r.error);
    return ok(res, r || {});
  }

  return async function handle(req, res) {
    let pathname;
    try { pathname = decodeURIComponent(new URL(req.url, 'http://x').pathname); } catch (e) { return send(res, 400, 'text/plain; charset=utf-8', '400 Bad Request'); }
    try {
      if (pathname.startsWith('/api/')) return await api(req, res, pathname.slice(5));
      if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, 'text/plain; charset=utf-8', '405 Method Not Allowed');
      // For the proxy's and Docker's health check: answers without a sign-in, and says nothing.
      if (pathname === '/health') return send(res, 200, 'text/plain; charset=utf-8', 'ok');
      if (pathname === '/login') return sendFile(res, UI_DIR, 'login.html', STATIC_TYPES);
      if (pathname === '/ui/login.js' || pathname === '/ui/app.css' || pathname === '/ui/icon.svg' || pathname === '/ui/i18n.js') return sendFile(res, UI_DIR, pathname.slice(4), STATIC_TYPES);
      const s = auth.session(req);
      if (pathname === '/' || pathname === '/index.html') {
        if (!s) { res.writeHead(302, Object.assign({ Location: '/login', 'Cache-Control': 'no-store' }, SECURITY)); return res.end(); }
        return sendFile(res, UI_DIR, 'index.html', STATIC_TYPES);
      }
      // Everything else is the editor's, and only for who is signed in.
      if (!s) return send(res, 401, 'text/plain; charset=utf-8', '401 Not signed in');
      if (pathname === '/shared/signature-core.js') return sendFile(res, path.dirname(CORE_FILE), path.basename(CORE_FILE), STATIC_TYPES);
      if (pathname.startsWith('/ui/')) return sendFile(res, UI_DIR, pathname.slice(4), STATIC_TYPES);
      if (pathname.startsWith('/assets/')) return sendFile(res, store.assetDir, pathname.slice(8), ASSET_TYPES);
      if (pathname.startsWith('/photos/')) return sendFile(res, settings.photosDir(), pathname.slice(8), ASSET_TYPES);
      return send(res, 404, 'text/plain; charset=utf-8', '404 Not Found');
    } catch (e) {
      log.error('request ' + pathname + ': ' + (e && e.stack || e));
      if (!res.headersSent) send(res, 500, 'text/plain; charset=utf-8', '500 Internal Server Error');
    }
  };
}

module.exports = { createApp };
