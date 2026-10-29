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
const { createSettings, check: checkSettings } = require('./settings');
const { graphConfig, graphEnabled, graphProblem, createGraph, folderTarget } = require('./graph');
const { createAuth } = require('./auth');
const { ldapConfig, ldapProblem } = require('./ldap');
const { directoryConfig, directoryEnabled, directoryProblem, lookup, groupsStartingWith } = require('./directory');
const { createUsers, nextAt } = require('./users');
const { createConnections } = require('./connections');
const { createSetup } = require('./setup');
const addin = require('./addin');

const VERSION = require('../package.json').version;
const ROOT = path.resolve(__dirname, '..');
const UI_DIR = path.join(ROOT, 'editor', 'ui');
const CORE_FILE = path.join(ROOT, 'shared', 'signature-core.js');
const STATIC_TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon' };
// A person's photo by name (photos/<username>.jpg): one file name, an image.
const PHOTO_NAME = /^[^/\\:*?"<>|\u0000-\u001f]{1,120}\.(png|jpe?g|gif)$/i;
const PHOTO_CACHE_MS = 5 * 60 * 1000;
// A request body larger than this is refused: a configuration is at most 1.5 MB, an image 1 MB
// (as base64, 1.4 MB).
const MAX_BODY = 4 * 1024 * 1024;

// Never inside another site's frame, never read as another type; the page loads nothing but its
// own files, plus images from https (an image a signature links to, shown in the preview). It
// talks to its server only, and to Microsoft's sign-in and Graph for the one step the admin does
// with their own account (giving the server access to the site, in the setup assistant).
const SECURITY = {
  'X-Frame-Options': 'DENY',
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'no-referrer',
  'Content-Security-Policy': "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob: https:; connect-src 'self' https://login.microsoftonline.com https://graph.microsoft.com; frame-src 'self'; object-src 'none'; base-uri 'self'; form-action 'none'; frame-ancestors 'self'"
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

// Who is asking: the shared proxy (Caddy) puts the client's address in X-Forwarded-For, replacing
// what the client sent; without a proxy, the connection's own.
function clientAddress(req) {
  const fwd = String(req.headers['x-forwarded-for'] || '').split(',').map((x) => x.trim()).filter(Boolean);
  return fwd.length ? fwd[fwd.length - 1] : String(req.socket.remoteAddress || '');
}

function createApp(opts) {
  const env = (opts && opts.env) || process.env;
  const log = (opts && opts.log) || console;
  const dataDir = env.OSIGN_DATA || '/data';
  const publishDir = env.OSIGN_PUBLISH_DIR || path.join(dataDir, 'publish');
  fs.mkdirSync(dataDir, { recursive: true });
  fs.mkdirSync(publishDir, { recursive: true });
  const store = createStore(dataDir);
  // The connections (AD, Graph, Cloudflare): .env, with what the admin saved in the editor on top.
  // Each is read anew when that changes (configure, below), so no restart is needed.
  const connections = createConnections({ dataDir, env });
  // Filled in by configure(); the same objects throughout, so what holds them sees the change.
  const graphCfg = {}, ldapCfg = {}, dir = {}, pages = {};
  const refill = (obj, fresh) => { for (const k of Object.keys(obj)) delete obj[k]; Object.assign(obj, fresh); };
  // Where publishing goes: SharePoint through Graph when the server's app registration is set
  // up, its own folder otherwise (development, tests).
  let graph = null;
  const spFolder = (s) => String(s.spFolder || 'OSign').replace(/^\/+|\/+$/g, '');
  const local = folderTarget(publishDir);
  const targetFor = (s) => (graph ? graph.target(s.spSite || 'root', spFolder(s)) : local);
  const settings = createSettings(dataDir, targetFor);
  // Photos for the preview, read from the target: kept a few minutes, so a preview that is drawn
  // again does not ask SharePoint each time.
  const photos = new Map();
  async function photo(name) {
    const key = settings.target().describe() + '|' + name.toLowerCase();
    const hit = photos.get(key);
    if (hit && Date.now() - hit.at < PHOTO_CACHE_MS) return hit.bytes;
    const bytes = await settings.target().read('photos/' + name);
    if (photos.size > 300) photos.clear();
    photos.set(key, { at: Date.now(), bytes });
    return bytes;
  }
  const auth = createAuth({ env, ldap: ldapCfg, log });
  // Active Directory with the service account (dir): the preview's users and groups, and the
  // users' values for signatures that take them from AD.
  let adOn = false;
  const users = createUsers({ dataDir, directory: dir, log });
  // The add-in's files, as last made; and where they go (Cloudflare Pages).
  const addinDir = path.join(dataDir, 'addin');
  const fetchFn = (opts && opts.fetch) || fetch;
  // The scheduled export (OSIGN_AD_EXPORT_AT, hh:mm in TZ): every day, what the published
  // signatures ask for, so the values stay as current as AD without anyone publishing.
  let exportAt = '';
  let lastExport = null, exportTimer = null;
  async function runExport() {
    const at = Date.now();
    try {
      const r = await users.exportPublished(settings.target());
      lastExport = { at, ok: true, result: r };
      log.info('AD export: ' + (r.skipped ? 'skipped, ' + r.skipped : r.users + ' users, ' + r.written + ' written, ' + r.unchanged + ' unchanged, ' + r.removed + ' removed'));
    } catch (e) {
      lastExport = { at, ok: false, error: e.message };
      log.error('AD export failed: ' + e.message);
    }
    return lastExport;
  }
  function scheduleExport() {
    const next = nextAt(exportAt, Date.now());
    if (!next) return;
    exportTimer = setTimeout(async () => { await runExport(); scheduleExport(); }, next - Date.now());
    // The timer alone never keeps the process (or a test) running.
    if (exportTimer.unref) exportTimer.unref();
  }
  // Reads the connections (again): at the start, and when the admin saved them.
  function configure() {
    const e = connections.effectiveEnv();
    refill(graphCfg, graphConfig(e));
    graph = graphEnabled(graphCfg) && !graphProblem(graphCfg) ? createGraph(graphCfg, { fetch: opts && opts.fetch }) : null;
    if (graphProblem(graphCfg)) log.warn('Publishing to SharePoint is not set up right: ' + graphProblem(graphCfg));
    refill(ldapCfg, Object.assign(ldapConfig(e), (opts && opts.ldap) || {}));
    refill(dir, directoryConfig(e, ldapCfg));
    adOn = directoryEnabled(dir) && !directoryProblem(dir);
    if (directoryEnabled(dir) && directoryProblem(dir)) log.warn('Active Directory is not set up right: ' + directoryProblem(dir));
    refill(pages, addin.pagesConfig(e));
    exportAt = String(e.OSIGN_AD_EXPORT_AT || '').trim();
    if (exportTimer) { clearTimeout(exportTimer); exportTimer = null; }
    if (adOn && exportAt) {
      if (nextAt(exportAt, Date.now())) scheduleExport();
      else log.warn('OSIGN_AD_EXPORT_AT is a time of day like 06:00: ' + exportAt);
    }
    settings.forget(); photos.clear();
  }
  configure();
  // The setup script for Microsoft 365 (server/setup.js): one at a time, for an hour.
  const setup = createSetup();
  // Its result goes where the admin would type it, and takes effect at once.
  async function setupApply(r) {
    const c = connections.save({ values: { graphTenantId: r.tenantId, graphClientId: r.serverClientId }, secrets: { graphClientSecret: r.secret } });
    if (c.error) return c;
    const s = await settings.save({ clientId: r.addinClientId, tenantId: r.tenantId });
    if (s.error) return s;
    configure();
    log.info('setup: the setup script set up Microsoft 365 (add-in ' + r.addinClientId + ', server ' + r.serverClientId + ')');
    return {};
  }
  // The connections are the local account admin's alone: whoever sets the directory's address
  // gets the passwords people sign in with.
  const adminOnly = (s) => (s.user.local ? '' : 'Only the local account admin can change the server\'s connections.');
  const ADMIN_SETTINGS = ['addinUrl', 'clientId', 'tenantId', 'spSite', 'spFolder', 'addinHost'];
  function connectionsAnswer() {
    return Object.assign(connections.answer(), {
      problems: {
        signIn: ldapProblem(ldapCfg) || (ldapCfg.url ? '' : 'Signing in with Active Directory is not set up.'),
        directory: directoryProblem(dir), graph: graphProblem(graphCfg) || (graphEnabled(graphCfg) ? '' : 'Publishing to SharePoint is not set up.'),
        pages: addin.pagesProblem(pages)
      }
    });
  }

  const routes = {
    ping: () => ({}),
    info: (data, s) => ({
      backend: 'node', version: VERSION, dataDir,
      // Not here yet: AD lookup and export (LDAP), and the add-in's files (Cloudflare Pages).
      features: {
        adLookup: adOn, addinFiles: !!addin.templateDir(env), pickFolder: false, signOut: true,
        // The add-in's files go to Cloudflare Pages from here, or by hand as a ZIP.
        pagesUpload: !addin.pagesProblem(pages),
        pagesProblem: addin.pagesProblem(pages),
        pagesProject: pages.project,
        // 'sharepoint': publishing goes through Graph; 'folder': into the server's folder.
        publishTarget: graph ? 'sharepoint' : 'folder',
        graphClientId: graph ? graphCfg.clientId : '',
        graphTenantId: graph ? graphCfg.tenantId : '',
        graphProblem: graphProblem(graphCfg),
        // AD with the server's service account; the scheduled export, and how its last run went.
        adProblem: adOn ? '' : directoryProblem(dir),
        adExport: adOn && exportAt ? { at: exportAt, last: lastExport } : null
      },
      user: { login: s.user.login, name: s.user.name, local: !!s.user.local }
    }),
    'draft/get': () => store.getDraft(),
    'draft/save': (d) => store.saveDraft(typeof d.configJson === 'string' ? d.configJson : ''),
    'assets/upload': (d) => store.uploadAsset(d.name, d.data),
    'assets/delete': (d) => store.deleteAsset(d.name),
    'settings/get': () => settings.answer(),
    // Where the server publishes and what the add-in is made with: the local account admin's, like
    // the connections (the site goes into the setup script, the add-in's address is fetched).
    'settings/save': (d, s) => {
      const sent = d.settings || {}, now = settings.get();
      if (!s.user.local && ADMIN_SETTINGS.some((k) => sent[k] !== undefined && sent[k] !== null && String(sent[k]).trim().replace(/[\\/]+$/, '') !== now[k])) {
        return { error: 'Only the local account admin can change these settings.' };
      }
      return settings.save(sent);
    },
    // The setup script: made for this setup (the site and the add-in's address as saved), and
    // whether it has sent its result yet.
    'setup/start': (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      if (!connections.answer().keyOk) return { error: 'The server has no key to keep the client secret with (secrets/app_key): run sudo ./deploy-nobs-osign.sh again.' };
      const st = settings.get();
      return setup.start({ origin: String(d.origin || ''), site: st.spFolder ? st.spSite || 'root' : '', addinUrl: st.addinUrl, groups: !!d.groups,
        // The apps of an earlier run, the only ones the script may use again.
        knownAddin: st.clientId, knownServer: graphCfg.clientId || '' });
    },
    'setup/status': (d, s) => (adminOnly(s) ? { error: adminOnly(s) } : setup.status()),
    // The server's connections (Settings → Server connections), for the local account admin.
    'connections/get': (d, s) => (adminOnly(s) ? { error: adminOnly(s) } : connectionsAnswer()),
    'connections/save': (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      const r = connections.save({ values: d.values, secrets: d.secrets && typeof d.secrets === 'object' ? d.secrets : {}, ca: d.ca });
      if (r.error) return r;
      configure();
      log.info('connections: changed by ' + s.user.login);
      return connectionsAnswer();
    },
    publish: async (d) => {
      const t = settings.target();
      let r;
      try { r = await store.publish(t, typeof d.configJson === 'string' ? d.configJson : ''); }
      catch (e) { return { error: e.message }; }
      settings.forget(); photos.clear();
      return r.error ? r : { changed: r.changed, folder: t.describe() };
    },
    'load-published': async () => {
      try { return await store.loadPublished(settings.target()); } catch (e) { return { error: e.message }; }
    },
    // The add-in's files: made into the data volume, then uploaded or downloaded.
    'addin/build': async () => {
      const r = await addin.build(env, settings.get(), addinDir);
      return r.error ? r : Object.assign({ folder: '' }, r);
    },
    'addin/check': async () => {
      const s = settings.get();
      if (!s.addinUrl) return { error: 'Enter the address the add-in will be served from.' };
      if (!fs.existsSync(path.join(addinDir, 'manifest.xml'))) return { error: 'Create the add-in files first.' };
      return addin.check(s.addinUrl, addinDir, fetchFn);
    },
    'addin/upload': async () => {
      if (!fs.existsSync(path.join(addinDir, 'manifest.xml'))) return { error: 'Create the add-in files first.' };
      try { return await addin.pagesDeploy(pages, addinDir, fetchFn); } catch (e) { return { error: e.message }; }
    },
    // Active Directory, with the service account.
    'ad/lookup': async (d, s) => {
      if (!adOn) return { error: directoryProblem(dir) };
      // Nothing typed: the one signed in, as OSign.ps1 takes the Windows user.
      const q = String(d.query || '').trim() || (s.user.local ? '' : s.user.login);
      try { return await lookup(dir, q, Array.isArray(d.fields) ? d.fields : []); } catch (e) { return { error: e.message }; }
    },
    'ad/groups': async (d) => {
      if (!adOn) return { error: directoryProblem(dir) };
      try { return { groups: await groupsStartingWith(dir, d.query) }; } catch (e) { return { error: e.message }; }
    },
    'ad/read': async (d) => {
      if (!adOn) return { error: directoryProblem(dir) };
      try { return await users.apiRead(settings.target(), d); } catch (e) { return { error: 'Active Directory: ' + e.message }; }
    },
    'ad/snapshot': (d) => users.apiSnapshot(d),
    'ad/export': async (d) => {
      try { return await users.apiExport(settings.target(), d); } catch (e) { return { error: 'Writing the users: ' + e.message }; }
    },
    // A sharing link: the folder it leads to (its own address), for the setup assistant.
    'sp/share': async (d) => {
      const link = String(d.url || '').trim();
      if (!/^https:\/\/[a-z0-9-]+(-my)?\.sharepoint\.com\/:[a-z]:\//i.test(link)) return { error: 'That is not a sharing link of SharePoint or OneDrive.' };
      if (!graph) return { error: 'The server cannot open sharing links before publishing to SharePoint is set up.' };
      let r;
      try { r = await graph.share(link); }
      catch (e) {
        if (e.status === 401 || e.status === 403) return { error: 'The server may not open sharing links (its app registration may only reach the sites granted to it).' };
        if (e.status === 404) return { error: 'This sharing link leads nowhere: it was removed or has expired.' };
        return { error: e.message };
      }
      if (!r.folder) return { error: 'This link leads to a file (' + r.name + '), not to a folder.' };
      return { webUrl: r.webUrl, name: r.name };
    },
    // Whether the server reaches a SharePoint folder, before it is saved (the setup assistant).
    'sp/check': async (d) => {
      const s = Object.assign(settings.get(), { spSite: String(d.spSite || 'root').trim(), spFolder: String(d.spFolder || '').trim() });
      if (!s.spFolder) return { error: 'Name the folder in the site, like OSign.' };
      const err = checkSettings(s); if (err) return { error: err };
      if (!graph) return { error: graphProblem(graphCfg) || 'Publishing to SharePoint is not set up on the server (Settings → Server connections).' };
      const t = targetFor(s);
      try { return { found: await t.exists(''), photos: await t.exists('photos'), where: t.describe() }; }
      catch (e) {
        // The usual two, in words the assistant can show next to its instructions.
        if (e.status === 401 || e.status === 403) return { error: 'The server may not reach this site yet: its app registration needs write access to it (Sites.Selected).' };
        if (e.status === 404) return { error: 'There is no such SharePoint site. Check the address.' };
        return { error: e.message };
      }
    }
  };
  // What only Windows, OneDrive or (for now) the PowerShell editor can do: said plainly.
  const NOT_HERE = {
    'pick-folder': 'Choosing a folder needs Windows.',
    'sp/resolve': 'The server does not use a synced folder.',
    'sp/local': 'The server does not use a synced folder.',
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
      const r = await auth.login(data.login, data.password, clientAddress(req));
      if (r.error) return fail(res, r.error, r.status);
      return ok(res, { user: { login: r.user.login, name: r.user.name } }, { 'Set-Cookie': r.cookie });
    }
    if (name === 'login-methods') return ok(res, auth.methods());
    // The setup script sends its result with its one-time code, not with a session.
    if (name === 'setup/finish') {
      const r = setup.finish(data);
      if (r.error) return fail(res, r.error, r.status);
      if (r.data) { const a = await setupApply(r.data); if (a.error) return fail(res, a.error); }
      return ok(res, {});
    }
    if (name === 'logout') return ok(res, {}, { 'Set-Cookie': auth.logout(req) });

    const s = auth.session(req);
    if (!s) return fail(res, 'Not signed in.', 401);
    if (Object.prototype.hasOwnProperty.call(NOT_HERE, name)) return fail(res, NOT_HERE[name]);
    const fn = Object.prototype.hasOwnProperty.call(routes, name) ? routes[name] : null;
    if (!fn) return fail(res, 'unknown', 404);
    let r;
    try { r = await fn(data, s); } catch (e) {
      log.error('api ' + name + ': ' + (e && e.stack || e));
      return fail(res, 'The server could not do that: ' + (e && e.message), 500);
    }
    if (r && r.error) return fail(res, r.error);
    return ok(res, r || {});
  }

  async function handle(req, res) {
    let pathname;
    try { pathname = decodeURIComponent(new URL(req.url, 'http://x').pathname); } catch (e) { return send(res, 400, 'text/plain; charset=utf-8', '400 Bad Request'); }
    try {
      if (pathname.startsWith('/api/')) return await api(req, res, pathname.slice(5));
      if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, 'text/plain; charset=utf-8', '405 Method Not Allowed');
      // For the proxy's and Docker's health check: answers without a sign-in, and says nothing.
      if (pathname === '/health') return send(res, 200, 'text/plain; charset=utf-8', 'ok');
      if (pathname === '/login') return sendFile(res, UI_DIR, 'login.html', STATIC_TYPES);
      if (pathname === '/ui/login.js' || pathname === '/ui/app.css' || pathname === '/ui/icon.svg' || pathname === '/ui/i18n.js') return sendFile(res, UI_DIR, pathname.slice(4), STATIC_TYPES);
      // Where Microsoft's sign-in comes back to: a cross-site navigation, so without the session
      // cookie (SameSite=Strict). It holds nothing; it hands the answer to the editor that opened it.
      if (pathname === '/ui/signin.html' || pathname === '/ui/signin.js') return sendFile(res, UI_DIR, pathname.slice(4), STATIC_TYPES);
      const s = auth.session(req);
      if (pathname === '/' || pathname === '/index.html') {
        if (!s) { res.writeHead(302, Object.assign({ Location: '/login', 'Cache-Control': 'no-store' }, SECURITY)); return res.end(); }
        return sendFile(res, UI_DIR, 'index.html', STATIC_TYPES);
      }
      // Everything else is the editor's, and only for who is signed in.
      if (!s) return send(res, 401, 'text/plain; charset=utf-8', '401 Not signed in');
      if (pathname === '/shared/signature-core.js') return sendFile(res, path.dirname(CORE_FILE), path.basename(CORE_FILE), STATIC_TYPES);
      if (pathname.startsWith('/ui/')) return sendFile(res, UI_DIR, pathname.slice(4), STATIC_TYPES);
      // The setup script, for the local account admin (made by setup/start).
      if (pathname === '/setup/nobs-osign-setup.ps1') {
        const text = s.user.local ? setup.script(new URL(req.url, 'http://x').searchParams.get('id') || '') : null;
        if (!text) return send(res, 404, 'text/plain; charset=utf-8', 'Download the setup script again in the setup assistant.');
        return send(res, 200, 'text/plain; charset=utf-8', text, { 'Content-Disposition': 'attachment; filename="nobs-osign-setup.ps1"' });
      }
      // The add-in's files as last made, to upload by hand.
      if (pathname === '/addin.zip') {
        if (!fs.existsSync(path.join(addinDir, 'manifest.xml'))) return send(res, 404, 'text/plain; charset=utf-8', 'Create the add-in files first.');
        return send(res, 200, 'application/zip', addin.zip(addinDir), { 'Content-Disposition': 'attachment; filename="osign-addin.zip"' });
      }
      if (pathname.startsWith('/assets/')) return sendFile(res, store.assetDir, pathname.slice(8), ASSET_TYPES);
      if (pathname.startsWith('/photos/')) {
        const name = pathname.slice(8);
        if (!PHOTO_NAME.test(name) || name.startsWith('.')) return send(res, 404, 'text/plain; charset=utf-8', '404 Not Found');
        let bytes = null;
        try { bytes = await photo(name); } catch (e) { log.warn('photo ' + name + ': ' + e.message); }
        if (!bytes) return send(res, 404, 'text/plain; charset=utf-8', '404 Not Found');
        return send(res, 200, ASSET_TYPES[path.extname(name).toLowerCase()], bytes, { 'Cache-Control': 'private, max-age=300' });
      }
      return send(res, 404, 'text/plain; charset=utf-8', '404 Not Found');
    } catch (e) {
      log.error('request ' + pathname + ': ' + (e && e.stack || e));
      if (!res.headersSent) send(res, 500, 'text/plain; charset=utf-8', '500 Internal Server Error');
    }
  }
  // For the tests and a manual run: the scheduled export, now.
  handle.runExport = runExport;
  return handle;
}

module.exports = { createApp };
