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
const { graphConfig, graphEnabled, graphProblem, createGraph, folderTarget } = require('./graph');
const { createAuth } = require('./auth');
const { ldapConfig, ldapProblem } = require('./ldap');
const { directoryConfig, directoryEnabled, directoryProblem, allowedAttrs, lookup } = require('./directory');
const { createUsers, nextAt } = require('./users');
const { createConnections } = require('./connections');
const { createSetup } = require('./setup');
const { createDeployment } = require('./deployment');
const { createHistory } = require('./history');
const { createCerts } = require('./cert');
const { createAudit } = require('./audit');
const { createReleaseRoutes } = require('./routes-release');
const { createNotify } = require('./notify');
const { createBackupRoutes } = require('./routes-backup');
const { createHistoryRoutes } = require('./routes-history');
const { createAddinRoutes } = require('./routes-addin');
const { createConnectionRoutes } = require('./routes-connections');
const { createPublishRoutes } = require('./routes-publish');
const { createPhotoRoutes } = require('./routes-photos');
const { createDirectoryRoutes } = require('./routes-directory');
const { createUpdates } = require('./updates');
const tlsInfo = require('./tls');
const { writeAtomic, readText } = require('./store');
const core = require('../shared/signature-core.js');
const addin = require('./addin');

const VERSION = require('../package.json').version;
const ROOT = path.resolve(__dirname, '..');
const UI_DIR = path.join(ROOT, 'editor', 'ui');
const CORE_FILE = path.join(ROOT, 'shared', 'signature-core.js');
const STATIC_TYPES = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon' };
// A person's photo by name (photos/<username>.jpg): one file name, an image.
const PHOTO_NAME = /^[^/\\:*?"<>|\u0000-\u001f]{1,120}\.(png|jpe?g|gif)$/i;
const PHOTO_CACHE_MS = 5 * 60 * 1000;
const PHOTO_PLACEHOLDER = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 80" preserveAspectRatio="xMidYMid slice"><rect width="100" height="80" fill="#e9ecef"/><circle cx="50" cy="31" r="14" fill="#b8bec6"/><ellipse cx="50" cy="82" rx="30" ry="26" fill="#b8bec6"/></svg>';
// A request body larger than this is refused: a configuration is at most 1.5 MB, an image 1 MB
// (as base64, 1.4 MB).
const MAX_BODY = 4 * 1024 * 1024;
// A backup brought back may be larger: the history with its images.
const MAX_BACKUP_BODY = 96 * 1024 * 1024;

// Never inside another site's frame, never read as another type; the page loads nothing but its
// own files, plus images from https (an image a signature links to, shown in the preview). It
// talks to its server only, and to Microsoft's sign-in and Graph for the one step the admin does
// with their own account (giving the server access to the site, in the setup assistant).
const SECURITY = {
  'X-Frame-Options': 'DENY',
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'no-referrer',
  // Always over HTTPS once seen so (the proxy serves it; plain http, as in tests, ignores it), no
  // camera, microphone or location, and a window of its own (popups allowed: the Microsoft sign-in
  // of the setup assistant is one).
  'Strict-Transport-Security': 'max-age=31536000',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=(), payment=(), usb=()',
  'Cross-Origin-Opener-Policy': 'same-origin-allow-popups',
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

// The editor's script: one file for the browser, kept as parts by topic (editor/ui/app/01-base.js
// … 11-start.js, in that order, one function scope across them); put together once, and again
// when a part changed (while developing).
const APP_PARTS = path.join(UI_DIR, 'app');
let appJs = null;
function editorScript() {
  const files = fs.readdirSync(APP_PARTS).filter((f) => /^\d\d-[a-z-]+\.js$/.test(f)).sort();
  const stamp = files.map((f) => f + ':' + fs.statSync(path.join(APP_PARTS, f)).mtimeMs).join('|');
  if (!appJs || appJs.stamp !== stamp) appJs = { stamp, body: Buffer.from(files.map((f) => fs.readFileSync(path.join(APP_PARTS, f), 'utf8')).join(''), 'utf8') };
  return appJs.body;
}

// The body of a request, at most max bytes: a larger one is refused before it is all read.
function readBody(req, max = MAX_BODY) {
  return new Promise((resolve, reject) => {
    const parts = []; let size = 0;
    req.on('data', (d) => { size += d.length; if (size > max) { reject(Object.assign(new Error('too large'), { status: 413 })); req.destroy(); } else parts.push(d); });
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

// The server: its parts made from the environment, and the one handler for every request.
function createApp(opts) {
  const env = (opts && opts.env) || process.env;
  const log = (opts && opts.log) || console;
  const dataDir = env.OSIGN_DATA || '/data';
  const publishDir = env.OSIGN_PUBLISH_DIR || path.join(dataDir, 'publish');
  fs.mkdirSync(dataDir, { recursive: true });
  fs.mkdirSync(publishDir, { recursive: true });
  const store = createStore(dataDir);
  // Who did what (server/audit.js); what each request adds to it is in AUDIT, below.
  const audit = createAudit({ dataDir });
  // What was published, to bring back (server/history.js).
  const history = createHistory({ dataDir, assetDir: store.assetDir });
  // The connections (AD, Graph, Cloudflare): .env, with what the admin saved in the editor on top.
  // Each is read anew when that changes (configure, below), so no restart is needed.
  const connections = createConnections({ dataDir, env });
  // The server's certificate for Microsoft Entra ID (server/cert.js), sealed like the passwords.
  const certs = createCerts({ dataDir, seal: connections.seal, open: connections.open });
  // Filled in by configure(); the same objects throughout, so what holds them sees the change.
  const graphCfg = {}, ldapCfg = {}, dir = {}, pages = {};
  const refill = (obj, fresh) => { for (const k of Object.keys(obj)) delete obj[k]; Object.assign(obj, fresh); };
  // Where publishing goes: SharePoint through Graph when the server's app registration is set
  // up, its own folder otherwise (development, tests).
  let graph = null;
  // When the client secret (or the certificate) runs out, and how many days are left (null: not known).
  let secretExpires = '';
  const secretDaysLeft = () => (secretExpires && !isNaN(Date.parse(secretExpires)) ? Math.floor((Date.parse(secretExpires) - Date.now()) / 86400000) : null);
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
  // The sessions, kept sealed with the server's key (secrets/app_key): an update signs nobody out.
  const sessionsFile = path.join(dataDir, 'sessions.json');
  const auth = createAuth({ env, ldap: ldapCfg, log, store: {
    load: () => { if (!connections.keyOk()) return []; const s = JSON.parse(readText(sessionsFile) || 'null'); const v = s ? connections.open('sessions', s) : ''; return v ? JSON.parse(v) : []; },
    save: (list) => { if (connections.keyOk()) writeAtomic(sessionsFile, JSON.stringify(connections.seal('sessions', JSON.stringify(list)))); }
  } });
  // Active Directory with the service account (dir): the preview's users and groups, and the
  // users' values for signatures that take them from AD.
  let adOn = false;
  const users = createUsers({ dataDir, directory: dir, log });
  // The add-in's files, as last made; and where they go (Cloudflare Pages).
  const addinDir = path.join(dataDir, 'addin');
  // The add-in's files as last made (their version), and as last uploaded to Cloudflare Pages.
  const addinUploadFile = path.join(dataDir, 'addin-upload.json');
  const readJson = (f) => { try { return JSON.parse(fs.readFileSync(f, 'utf8')); } catch (e) { return null; } };
  const addinMade = () => (fs.existsSync(path.join(addinDir, 'manifest.xml')) && readJson(path.join(addinDir, 'osign-addin.json'))) || { version: '' };
  const fetchFn = (opts && opts.fetch) || fetch;
  // The scheduled export (OSIGN_AD_EXPORT_AT, hh:mm in TZ): every day, what the published
  // signatures ask for, so the values stay as current as AD without anyone publishing.
  let exportAt = '';
  let lastExport = null, exportTimer = null;
  // The daily export of the users from AD: read and written in one go, the outcome kept for the status.
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
  // The next daily export armed at its time of day (OSIGN_AD_EXPORT_AT); none without one.
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
    const cert = certs.active();
    refill(graphCfg, graphConfig(e, cert));
    secretExpires = cert ? cert.expires.slice(0, 10) : String(e.OSIGN_GRAPH_CLIENT_SECRET_EXPIRES || '').trim();
    const left = secretDaysLeft();
    const what = cert ? 'certificate' : 'client secret';
    if (left !== null && left <= 30) log.warn(left < 0 ? 'The server\'s ' + what + ' ran out on ' + secretExpires + ': publishing to SharePoint fails until a new one is set (run the setup script again).' : 'The server\'s ' + what + ' runs out on ' + secretExpires + ' (in ' + left + ' days): run the setup script again.');
    graph = graphEnabled(graphCfg) && !graphProblem(graphCfg) ? createGraph(graphCfg, { fetch: opts && opts.fetch, nextCert: () => certs.waiting(), onNextCert: (thumb) => {
      // The setup script's answer did not arrive, but its certificate is on the app: in use now.
      if (certs.activate(thumb)) { log.info('graph: the certificate of the last setup script (' + thumb + ') signs in; it is used from now on'); setImmediate(configure); }
    } }) : null;
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
  // Where the add-in is deployed in Microsoft 365: asked by a script of its own, only reading.
  const deployment = createDeployment({ dataDir });
  const addinId = () => { try { return JSON.parse(fs.readFileSync(path.join(addin.templateDir(env), 'osign-addin.json'), 'utf8')).id || ''; } catch (e) { return ''; } };
  // Its result goes where the admin would type it, and takes effect at once.
  async function setupApply(r) {
    let c;
    if (r.certThumbprint) {
      // The certificate the script put on the app: the server signs in with it, no secret kept.
      if (!certs.activate(r.certThumbprint)) return { error: 'The script put another certificate on the app than this setup made: download the setup script again.' };
      c = connections.save({ values: { graphTenantId: r.tenantId, graphClientId: r.serverClientId, graphSecretExpires: '' }, secrets: { graphClientSecret: null } });
    } else {
      const when = Date.parse(r.secretExpires);
      const expires = isNaN(when) ? '' : new Date(when).toISOString().slice(0, 10);
      certs.remove();
      c = connections.save({ values: { graphTenantId: r.tenantId, graphClientId: r.serverClientId, graphSecretExpires: expires }, secrets: { graphClientSecret: r.secret } });
    }
    if (c.error) return c;
    const s = await settings.save({ clientId: r.addinClientId, tenantId: r.tenantId });
    if (s.error) return s;
    configure();
    log.info('setup: the setup script set up Microsoft 365 (add-in ' + r.addinClientId + ', server ' + r.serverClientId + ', ' + (r.certThumbprint ? 'certificate ' + r.certThumbprint : 'client secret') + ')');
    return {};
  }
  // With Cloudflare Pages set up and no add-in address yet: the project's own (<name>.pages.dev),
  // so the setup script does not wait for it. An address set by hand is never replaced.
  async function addinUrlFromPages() {
    const st = settings.get();
    if (st.addinUrl || !/^[a-z0-9][a-z0-9-]{0,57}$/.test(pages.project || '')) return;
    await settings.save({ addinUrl: 'https://' + pages.project + '.pages.dev/' });
    log.info('settings: the add-in\'s address set to https://' + pages.project + '.pages.dev/ (Cloudflare Pages)');
  }
  // The connections are the local account admin's alone: whoever sets the directory's address
  // gets the passwords people sign in with.
  const adminOnly = (s) => (s.user.local ? '' : 'Only the local account admin can change the server\'s connections.');
  // Who saved the draft last (since the server started), for the answer to a save that came too late.
  let draftBy = null;
  const ADMIN_SETTINGS = ['addinUrl', 'clientId', 'tenantId', 'spSite', 'spFolder', 'addinHost', 'requireApproval', 'addinMobile'];
  // Whether a newer version is out (server/updates.js; OSIGN_UPDATE_CHECK=1).
  const updates = createUpdates({ env, version: VERSION, fetchFn, log });

  // Four eyes: with requireApproval on, publishing is asked for by one person and done by another.
  // The request is for the draft as it was then (its hash over the configuration and the images):
  // a change to the draft after that withdraws it.
  const approvalFile = path.join(dataDir, 'approval.json');
  const approval = () => { try { return JSON.parse(readText(approvalFile) || 'null'); } catch (e) { return null; } };
  const needsApproval = () => settings.get().requireApproval === 'on';
  // Who publishes and approves: with a publishing group set, only its members (and the local admin).
  const publisher = (s) => !!(s && s.user && (s.user.local || s.user.publisher !== false));
  // Approval needed for this person: four eyes on, or not one who publishes.
  const needsApprovalFor = (s) => needsApproval() || !publisher(s);
  // Taken out of the publishing group (or the sign-in group), someone loses publishing at once, not
  // with their next sign-in: asked again of AD (with the service account) before what publishes,
  // at most once a minute per session. AD not reachable: as it was at sign-in.
  const RECHECK = new Set(['info', 'approval/get', 'publish', 'approval/request', 'approval/approve', 'ad/export', 'photos/upload', 'photos/delete', 'photos/list',
    'schedule/set', 'schedule/cancel']);
  const cnOf = (g) => { const m = /^CN=((?:\\,|[^,])+)/i.exec(String(g || '')); return (m ? m[1].replace(/\\,/g, ',') : String(g || '')).trim().toLowerCase(); };
  // What AD says of a login now: { signIn, publisher } (in the sign-in group, in the publishing
  // group); signIn false for an account that is gone; null when AD cannot be asked.
  async function adRoles(login) {
    try {
      const r = await lookup(dir, login, []);
      const mine = r.groups.map((g) => g.toLowerCase());
      const signIn = mine.includes(cnOf(ldapCfg.group));
      return { signIn, publisher: signIn && (!ldapCfg.publishGroup || mine.includes(cnOf(ldapCfg.publishGroup))) };
    } catch (e) {
      if (/^No user /.test(e.message)) return { signIn: false, publisher: false };
      log.warn('roles: could not ask AD about ' + login + ': ' + e.message);
      return null;
    }
  }
  // For a publish set for later: whether its setter still publishes (null: not known).
  async function stillPublisher(login) {
    if (!login || !ldapCfg.publishGroup || !adOn) return null;
    const r = await adRoles(login);
    return r ? r.publisher : null;
  }
  // Whether someone signed in with AD still publishes, and may still sign in: asked of AD at most once a minute.
  async function recheckPublisher(s) {
    if (!s || !s.user || s.user.local || !ldapCfg.publishGroup || !adOn) return;
    if (s.user.publisherAt && Date.now() - s.user.publisherAt < 60000) return;
    const r = await adRoles(s.user.login);
    if (!r) return;
    const was = s.user.publisher;
    s.user.publisher = r.publisher; s.user.publisherAt = Date.now();
    // Gone from AD, or out of the sign-in group: signed out.
    if (!r.signIn) { s.user.gone = true; log.info('roles: ' + s.user.login + ' is no longer allowed to sign in: signed out'); }
    else if (was !== s.user.publisher) log.info('roles: ' + s.user.login + (s.user.publisher ? ' publishes again' : ' no longer publishes'));
  }
  // The same person, however they typed their login (AD does not mind the case).
  // A signature whose compiled HTML (what the add-in renders) says other than its html is refused:
  // what was looked at (and approved) is what goes out.
  function compiledProblem(configJson) {
    let c; try { c = JSON.parse(configJson); } catch (e) { return ''; }
    for (const s of (c && c.signatures) || []) {
      if (s && !core.compiledMatches(s.html, s.compiled)) return 'The signature "' + (s.name || s.id) + '" is not as its HTML says (compiled). Open it in the editor and publish again.';
    }
    // A campaign's banner too, in each of its languages.
    for (const x of (c && c.campaigns) || []) {
      const parts = [x].concat(Object.values((x && x.langs) || {}));
      if (parts.some((p) => p && !core.compiledMatches(p.html, p.compiled))) return 'The campaign "' + (x.name || x.id) + '" is not as its HTML says (compiled). Open it in the editor and publish again.';
    }
    return '';
  }
  // While publishing, the draft's images and the draft stay as they are (what was checked goes out).
  let publishing = false;
  // opts: { images, draft } as store.publish takes them (a version that is not the draft).
  async function publishNow(configJson, by, message, opts) {
    const t = settings.target();
    const bad = compiledProblem(typeof configJson === 'string' ? configJson : '');
    if (bad) return { error: bad };
    if (publishing) return { error: 'Another publish is running: try again in a moment.' };
    let r;
    publishing = true;
    try { r = await store.publish(t, typeof configJson === 'string' ? configJson : '', opts); }
    catch (e) { return { error: e.message }; }
    finally { publishing = false; }
    settings.forget(); photos.clear();
    // A version that differs from what was there: kept, with who and the note.
    if (!r.error && r.changed) history.add({ by, message, configJson, images: r.images });
    // The draft's version after it (it was written as published), so the editor goes on from it.
    return r.error ? r : { changed: r.changed, folder: t.describe(), rev: store.getDraft().rev };
  }
  // The server's connections for the editor: what is set, where from (GUI or .env), never a secret.
  function connectionsAnswer() {
    const cert = certs.active();
    return Object.assign(connections.answer(), {
      // What follows from the domain, for the fields left empty (shown as their placeholders).
      derived: { ldapBind: ldapCfg.bindTemplate, ldapBase: ldapCfg.searchBase, ldapDirectoryBase: dir.base, ldapServiceDn: dir.serviceDn },
      // The certificate in use, if any: it wins over a client secret.
      cert: cert ? { thumbprint: cert.thumbprint, expires: cert.expires } : null,
      problems: {
        signIn: ldapProblem(ldapCfg) || (ldapCfg.url ? '' : 'Not set up yet: enter the domain controller, then Detect.'),
        directory: directoryProblem(dir), graph: graphProblem(graphCfg) || (graphEnabled(graphCfg) ? '' : 'Not set up yet: the setup script does it (Server connections → SharePoint).'),
        pages: addin.pagesEnabled(pages) ? addin.pagesProblem(pages) : 'Not set up yet: three steps (Server connections → Cloudflare Pages).'
      }
    });
  }

  // Publishing at a time set beforehand (server/routes-release.js).
  const release = createReleaseRoutes({ dataDir, store, publishNow, publisher, needsApproval, compiledProblem, stillPublisher, audit, log });
  const routes = {
    ping: () => ({}),
    info: (data, s) => ({
      // The data volume's path only for the local admin (no one else needs where the server keeps it).
      backend: 'node', version: VERSION, dataDir: s.user.local ? dataDir : '',
      // Not here yet: AD lookup and export (LDAP), and the add-in's files (Cloudflare Pages).
      features: {
        adLookup: adOn, addinFiles: !!addin.templateDir(env), pickFolder: false, signOut: true,
        // Made already: its manifest.xml to download (Settings → Outlook add-in).
        addinBuilt: fs.existsSync(path.join(addinDir, 'manifest.xml')),
        // Their version, and the one on Cloudflare Pages: older than the server's, they are made again.
        addinVersion: addinMade().version || '',
        addinUploaded: readJson(addinUploadFile),
        // The add-in's files go to Cloudflare Pages from here, or by hand as a ZIP.
        pagesUpload: !addin.pagesProblem(pages),
        pagesProblem: addin.pagesProblem(pages),
        pagesProject: pages.project,
        // 'sharepoint': publishing goes through Graph; 'folder': into the server's folder.
        publishTarget: graph ? 'sharepoint' : 'folder',
        graphClientId: graph ? graphCfg.clientId : '',
        // The client secret: until when, and the days left (warned of from 30 on).
        secretExpires: graph ? secretExpires : '', secretDaysLeft: graph ? secretDaysLeft() : null,
        // How the server signs in to Microsoft Entra ID: 'certificate' or 'secret'.
        graphCredential: graph ? (graphCfg.cert ? 'certificate' : 'secret') : '',
        graphTenantId: graph ? graphCfg.tenantId : '',
        graphProblem: graphProblem(graphCfg),
        // AD with the server's service account; the scheduled export, and how its last run went.
        adProblem: adOn ? '' : directoryProblem(dir),
        // The AD attributes fields may read.
        adAttributes: allowedAttrs(dir),
        adExport: adOn && exportAt ? { at: exportAt, last: lastExport } : null,
        // A newer release, when the server is allowed to look (OSIGN_UPDATE_CHECK).
        update: updates.status(),
        // Four eyes: publishing waits for a second person.
        requireApproval: needsApprovalFor(s),
        // Others than the local admin can sign in (Active Directory).
        adSignIn: !!ldapCfg.url && !ldapProblem(ldapCfg),
        // The key the passwords and the certificate's key are kept with (secrets/app_key).
        keyOk: connections.keyOk()
      },
      user: { login: s.user.login, name: s.user.name, local: !!s.user.local, publisher: publisher(s) }
    }),
    'draft/get': () => store.getDraft(),
    // Saved from the version it was made from (rev): a newer one there is not written over unless
    // force; the answer says who saved it last.
    'draft/save': (d, s) => {
      if (publishing) return { error: 'Publishing right now: saved in a moment.' };
      const r = store.saveDraft(typeof d.configJson === 'string' ? d.configJson : '', { rev: typeof d.rev === 'string' ? d.rev : '', force: d.force === true });
      if (r.conflict) return Object.assign(r, { by: draftBy });
      if (!r.error) draftBy = { name: s.user.name || s.user.login, at: new Date().toISOString() };
      return r;
    },
    'assets/upload': (d) => (publishing ? { error: 'Publishing right now: try again in a moment.' } : store.uploadAsset(d.name, d.data)),
    'assets/delete': (d) => (publishing ? { error: 'Publishing right now: try again in a moment.' } : store.deleteAsset(d.name)),
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
    // The web certificate the proxy serves for the editor's address (as the browser asked for it).
    'server/tls': (d) => tlsInfo.inspect({ addr: env.OSIGN_PROXY_ADDR || 'nobs-proxy:443', host: String(d.host || ''), caFile: ldapCfg.caFile }),
    // The audit log, for the local account admin.
    // Who is signed in now, and a session ended (that person signs in again). The local admin's.
    'sessions/list': (d, s) => (adminOnly(s) ? { error: 'Only the local account admin sees who is signed in.' } : { sessions: auth.list().map((x) => Object.assign(x, { mine: x.ref === auth.ref(s.id) })) }),
    'sessions/end': (d, s) => {
      if (adminOnly(s)) return { error: 'Only the local account admin sees who is signed in.' };
      if (String(d.ref || '') === auth.ref(s.id)) return { error: 'That is your own session: sign out instead.' };
      const who = auth.end(String(d.ref || ''));
      return who ? { sessions: auth.list().map((x) => Object.assign(x, { mine: x.ref === auth.ref(s.id) })), ended: who } : { error: 'That session has ended already.' };
    },
    'audit/list': (d, s) => (adminOnly(s) ? { error: 'Only the local account admin can read the audit log.' } : { entries: audit.list({ limit: d.limit, what: String(d.what || ''), failed: !!d.failed }) }),

  };
  Object.assign(routes, release.routes);
  // The backups (server/routes-backup.js).
  const backups = createBackupRoutes({ dataDir, env, connections, store, configure, adminOnly, version: VERSION, log, busy: () => publishing });
  Object.assign(routes, backups.routes);
  // Telling a webhook (Teams, Power Automate) when something needs doing (server/notify.js).
  const problemsNow = () => {
    const out = [];
    const ab = backups.autoBackup.answer();
    if (ab.on && ab.last && ab.last.ok === false) out.push({ key: 'backup:' + ab.last.at, text: 'The automatic backup of ' + ab.last.at.slice(0, 16).replace('T', ' ') + ' failed: ' + ab.last.error });
    if (ab.passLost) out.push({ key: 'backup-pass', text: 'The automatic backups are not made: their passphrase cannot be opened. Set it again (Settings → Server → Backup).' });
    const left = graph ? secretDaysLeft() : null;
    if (left !== null && left <= 14) out.push({ key: 'secret:' + secretExpires + (left < 0 ? ':gone' : ':soon'), text: left < 0 ? 'The server\'s sign-in to Microsoft 365 ran out on ' + secretExpires + ': publishing fails. Run the setup script again.' : 'The server\'s sign-in to Microsoft 365 runs out on ' + secretExpires + ' (in ' + left + ' days): run the setup script again.' });
    if (adOn && exportAt && lastExport && lastExport.ok === false) out.push({ key: 'adexport:' + lastExport.at, text: 'The daily export of the users from AD failed: ' + lastExport.error });
    const sch = release.routes['schedule/get']().schedule;
    if (sch && sch.error) out.push({ key: 'schedule:' + sch.at, text: 'The publish set for ' + sch.at.slice(0, 16).replace('T', ' ') + ' did not go out: ' + sch.error });
    // A campaign of what is published ending in 3 days, with none after it for the same signatures.
    try {
      const pub = history.latest(), cfg = pub ? JSON.parse(pub) : null;
      for (const e of core.campaignsEnding(cfg, 3)) out.push({ key: 'campaign:' + e.c.id + ':' + e.c.validUntil, text: 'The campaign "' + (e.c.name || e.c.id) + '" ends ' + (e.left ? 'in ' + e.left + ' day(s), on ' + e.c.validUntil : 'today') + ', and none follows it: plan the next one in the editor (Campaigns), or let the signatures go without a banner.' });
    } catch (e) { /* no history yet */ }
    return out;
  };
  const notify = createNotify({ dataDir, seal: connections.seal, open: connections.open, problems: problemsNow, log, fetchImpl: fetchFn });
  notify.schedule();
  Object.assign(routes, {
    'notify/get': (d, s) => (adminOnly(s) ? { error: 'Only the local account admin sets the notifications.' } : Object.assign(notify.answer(), { now: problemsNow().map((p) => p.text) })),
    'notify/save': (d, s) => (adminOnly(s) ? { error: 'Only the local account admin sets the notifications.' } : notify.save(d)),
    'notify/test': (d, s) => (adminOnly(s) ? { error: 'Only the local account admin sets the notifications.' } : notify.test())
  });
  // Publishing and four eyes (server/routes-publish.js).
  Object.assign(routes, createPublishRoutes({ store, history, settings, users, approvalFile, approval, needsApproval, needsApprovalFor, publisher, publishNow, compiledProblem, live: () => ({ adOn }) }).routes);
  // The server's connections, the setup script, SharePoint for the assistant (server/routes-connections.js).
  Object.assign(routes, createConnectionRoutes({ connections, connectionsAnswer, settings, certs, setup, deployment, addinId, addinUrlFromPages, configure, targetFor,
    graphCfg, ldapCfg, dir, pages, fetchFn, adminOnly, live: () => ({ graph }), log }).routes);
  // The add-in's files (server/routes-addin.js).
  Object.assign(routes, createAddinRoutes({ env, settings, pages, addinDir, addinUploadFile, addinMade, fetchFn, adminOnly, live: () => ({ graph }), log }).routes);
  // What was published (server/routes-history.js).
  Object.assign(routes, createHistoryRoutes({ history, store, settings, busy: () => publishing }).routes);
  // Photos (server/routes-photos.js) and the directory, AD and Entra ID (server/routes-directory.js).
  Object.assign(routes, createPhotoRoutes({ settings, photos, publisher, needsApproval, opts }).routes);
  Object.assign(routes, createDirectoryRoutes({ live: () => ({ adOn, graph }), dir, graphCfg, users, settings, store, history, publisher, needsApproval }).routes);
  // What goes into the audit log, by request: the words for it, from what was sent and answered.
  const AUDIT = {
    publish: (d, r) => (r.changed ? 'changed' : 'nothing changed') + (d.message ? ': ' + d.message : ''),
    'sessions/end': (d, r) => 'session of ' + String(r.ended || '?') + ' ended',
    'notify/save': (d) => 'notifications ' + (typeof d.url === 'string' ? (d.url.trim() ? 'to a webhook set' : 'off') : 'kept'),
    'notify/test': () => 'test notification sent',
    'history/restore': (d) => 'version of ' + new Date(Number(d.id) || 0).toISOString() + ' into the draft',
    'connections/save': (d) => 'saved' + Object.keys(d.secrets || {}).map((k) => (d.secrets[k] === null ? ', ' + k + ' removed' : ', ' + k + ' set')).join('') + (d.removeCert ? ', certificate removed' : '') + (d.ca !== undefined ? ', CA certificate ' + (d.ca === null ? 'removed' : 'set') : ''),
    'settings/save': (d) => Object.keys(d.settings || {}).join(', '),
    'setup/start': () => 'setup script made',
    'connections/cert': (d, r) => 'certificate downloaded (' + (r.thumbprint || '') + ')',
    'addin/build': () => 'add-in files made',
    'deploy/start': () => 'deployment script made',
    'photos/upload': (d) => 'photo ' + String(d.name || '').slice(0, 120),
    'photos/delete': (d) => 'photo removed ' + String(d.name || '').slice(0, 120),
    'photos/fetch': (d) => 'photo fetched from ' + String(d.url || '').slice(0, 200),
    'addin/upload': (d, r) => 'uploaded to Cloudflare Pages' + (r.url ? ' ' + r.url : ''),
    'ad/export': () => 'users exported',
    'approval/request': (d) => 'asked for approval' + (d.message ? ': ' + d.message : ''),
    'approval/approve': () => 'approved and published',
    'approval/cancel': () => 'request withdrawn'
  };
  Object.assign(AUDIT, release.audit, backups.audit);

  // What only Windows, OneDrive or (for now) the PowerShell editor can do: said plainly.
  const NOT_HERE = {
    'pick-folder': 'Choosing a folder needs Windows.',
    'sp/resolve': 'The server does not use a synced folder.',
    'sp/local': 'The server does not use a synced folder.',
    quit: 'The server keeps running; sign out instead.'
  };

  // One API request: signed in, from this site, allowed for this account; answered as JSON, written to the audit log.
  async function api(req, res, name) {
    if (req.method !== 'POST' || !/^application\/json\b/i.test(String(req.headers['content-type'] || '')) || !sameSite(req)) return send(res, 403, 'text/plain; charset=utf-8', '403 Forbidden');
    // Signed in first, then the body is read: nobody who is not signed in makes the server hold a
    // large one, and only the local account admin may send a backup.
    const open = ['login', 'login-methods', 'setup/finish', 'deploy/finish', 'logout'].includes(name);
    const s = open ? null : auth.session(req);
    if (!open && !s) return fail(res, 'Not signed in.', 401);
    let data = {};
    try { const body = await readBody(req, (name === 'backup/restore' || name === 'backup/check') && s && s.user.local ? MAX_BACKUP_BODY : MAX_BODY); data = body ? JSON.parse(body) : {}; } catch (e) {
      if (e.status === 413) return fail(res, 'The request is too large.', 413);
      return fail(res, 'The request is not valid JSON.');
    }
    if (!data || typeof data !== 'object' || Array.isArray(data)) data = {};

    const ip = clientAddress(req);
    if (name === 'login') {
      const r = await auth.login(data.login, data.password, ip);
      audit.add({ who: String(data.login || '').slice(0, 120), ip, what: 'signin', ok: !r.error, detail: r.error || '' });
      if (r.error) return fail(res, r.error, r.status);
      return ok(res, { user: { login: r.user.login, name: r.user.name } }, { 'Set-Cookie': r.cookie });
    }
    if (name === 'login-methods') return ok(res, auth.methods());
    // The setup script sends its result with its one-time code, not with a session.
    if (name === 'setup/finish') {
      const r = setup.finish(data);
      const note = (ok, detail) => audit.add({ who: 'setup script', ip, what: 'setup/finish', ok, detail });
      if (r.error) { note(false, r.error); return fail(res, r.error, r.status); }
      if (r.failed) note(false, 'the script stopped: ' + String(data.error || ''));
      if (r.data) {
        const a = await setupApply(r.data);
        if (a.error) { note(false, a.error); return fail(res, a.error); }
        setup.done(r.data);
        note(true, 'Microsoft 365 set up (server ' + r.data.serverClientId + ', ' + (r.data.certThumbprint ? 'certificate' : 'client secret') + ')');
      }
      return ok(res, {});
    }
    // The deployment script's answer, with its one-time code.
    if (name === 'deploy/finish') {
      const r = deployment.finish(data);
      audit.add({ who: 'deployment script', ip, what: 'deploy/finish', ok: !r.error && !(r.result && r.result.error), detail: r.error || (r.result.error ? 'the script stopped: ' + r.result.error : r.result.found ? 'the add-in is deployed' : 'the add-in is not deployed') });
      if (r.error) return fail(res, r.error, r.status);
      return ok(res, {});
    }
    if (name === 'logout') {
      const was = auth.session(req);
      if (was) audit.add({ who: was.user.login, ip, what: 'signout' });
      return ok(res, {}, { 'Set-Cookie': auth.logout(req) });
    }

    if (Object.prototype.hasOwnProperty.call(NOT_HERE, name)) return fail(res, NOT_HERE[name]);
    const fn = Object.prototype.hasOwnProperty.call(routes, name) ? routes[name] : null;
    if (!fn) return fail(res, 'unknown', 404);
    let r;
    // What publishes or goes live at once: whether they are still in the publishing group, asked now.
    if (RECHECK.has(name)) await recheckPublisher(s);
    if (s.user.gone) { auth.logout(req); return fail(res, 'You are not in the group that may use NOBS OSign any more.', 401); }
    const logged = Object.prototype.hasOwnProperty.call(AUDIT, name);
    try { r = await fn(data, s, ip); } catch (e) {
      log.error('api ' + name + ': ' + (e && e.stack || e));
      if (logged) audit.add({ who: s.user.login, ip, what: name, ok: false, detail: String(e && e.message) });
      return fail(res, 'The server could not do that: ' + (e && e.message), 500);
    }
    if (logged) audit.add({ who: s.user.login, ip, what: name, ok: !(r && r.error), detail: r && r.error ? r.error : AUDIT[name](data, r || {}) });
    // A refusal with more to say (a draft changed in the meantime: its version, who saved it) keeps it.
    if (r && r.error) return r.conflict ? json(res, 200, { ok: false, error: r.error, conflict: true, rev: r.rev, by: r.by || null }) : fail(res, r.error);
    return ok(res, r || {});
  }

  // Every request: the API, the editor's files, the published images and photos, the add-in's files.
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
      if (pathname === '/ui/app.js') return send(res, 200, STATIC_TYPES['.js'], editorScript());
      if (pathname.startsWith('/ui/')) return sendFile(res, UI_DIR, pathname.slice(4), STATIC_TYPES);
      // The setup script, for the local account admin (made by setup/start).
      if (pathname === '/setup/nobs-osign-setup.ps1') {
        const text = s.user.local ? setup.script(new URL(req.url, 'http://x').searchParams.get('id') || '') : null;
        if (!text) return send(res, 404, 'text/plain; charset=utf-8', 'Download the setup script again in the setup assistant.');
        return send(res, 200, 'text/plain; charset=utf-8', text, { 'Content-Disposition': 'attachment; filename="nobs-osign-setup.ps1"' });
      }
      // The script that reads where the add-in is deployed (made by deploy/start).
      if (pathname === '/setup/nobs-osign-deployment.ps1') {
        const text = s.user.local ? deployment.script(new URL(req.url, 'http://x').searchParams.get('id') || '') : null;
        if (!text) return send(res, 404, 'text/plain; charset=utf-8', 'Download the script again under Settings → Server.');
        return send(res, 200, 'text/plain; charset=utf-8', text, { 'Content-Disposition': 'attachment; filename="nobs-osign-deployment.ps1"' });
      }
      // The add-in's files as last made, to upload by hand.
      // Its manifest alone, for "Add from file" in Outlook or "Upload manifest file" in the admin center.
      if (pathname === '/addin/manifest.xml') {
        const m = path.join(addinDir, 'manifest.xml');
        if (!fs.existsSync(m)) return send(res, 404, 'text/plain; charset=utf-8', 'Create the add-in files first.');
        return send(res, 200, 'application/xml; charset=utf-8', fs.readFileSync(m), { 'Content-Disposition': 'attachment; filename="manifest.xml"' });
      }
      if (pathname === '/addin.zip') {
        if (!fs.existsSync(path.join(addinDir, 'manifest.xml'))) return send(res, 404, 'text/plain; charset=utf-8', 'Create the add-in files first.');
        return send(res, 200, 'application/zip', addin.zip(addinDir), { 'Content-Disposition': 'attachment; filename="osign-addin.zip"' });
      }
      if (pathname.startsWith('/assets/')) return sendFile(res, store.assetDir, pathname.slice(8), ASSET_TYPES);
      if (pathname.startsWith('/photos/')) {
        const name = pathname.slice(8);
        // The design view asks for photos/@@sAMAccountName.jpg: a grey figure where a person's photo
        // goes (in every browser, not only where CSS can put a picture in place of an image).
        if (/@@|%40%40/i.test(name)) {
          // The placeholder photo that is set (the one the signatures take, by extension), else the grey figure.
          let ph = null;
          try { ph = await photo(core.PHOTO_DEFAULT + path.extname(name).toLowerCase()); } catch (e) { ph = null; }
          return ph ? send(res, 200, ASSET_TYPES[path.extname(name).toLowerCase()] || 'image/jpeg', ph) : send(res, 200, 'image/svg+xml', PHOTO_PLACEHOLDER);
        }
        if (!PHOTO_NAME.test(name) || name.startsWith('.')) return send(res, 404, 'text/plain; charset=utf-8', '404 Not Found');
        let bytes = null;
        try { bytes = await photo(name); } catch (e) { log.warn('photo ' + name + ': ' + e.message); }
        // None of their own: the placeholder photo, as the add-in takes it.
        if (!bytes && !name.startsWith(core.PHOTO_DEFAULT + '.')) {
          try { bytes = await photo(core.PHOTO_DEFAULT + path.extname(name).toLowerCase()); } catch (e) { bytes = null; }
        }
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
