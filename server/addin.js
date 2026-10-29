/*
 * The Outlook add-in's files in the server version: made from the template (dist/template, in the
 * image addin-template/) with addin/fill.mjs, the same code the build and the tests use; offered as
 * a ZIP; checked online; and uploaded to Cloudflare Pages (Direct Upload), so nothing on this
 * server has to be reachable from outside.
 *
 *   OSIGN_CF_ACCOUNT_ID        the Cloudflare account
 *   OSIGN_CF_PROJECT           the Pages project (made on the first upload when it is not there)
 *   OSIGN_CF_API_TOKEN_FILE    an API token with the permission Cloudflare Pages: Edit
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const zlib = require('node:zlib');
const { pathToFileURL } = require('node:url');
const { blake3 } = require('@noble/hashes/blake3');
const { bytesToHex } = require('@noble/hashes/utils');

const ROOT = path.resolve(__dirname, '..');

// Where the template is: next to the server in the image, dist/template in the repository.
function templateDir(env) {
  for (const d of [env.OSIGN_ADDIN_TEMPLATE, path.join(ROOT, 'addin-template'), path.join(ROOT, 'dist', 'template')]) {
    if (d && fs.existsSync(path.join(d, 'osign-addin.json'))) return d;
  }
  return '';
}

// Every file under dir, as paths with / ("assets/icon-16.png").
function walk(dir, rel = '') {
  const out = [];
  for (const e of fs.readdirSync(path.join(dir, rel), { withFileTypes: true }).sort((a, b) => a.name.localeCompare(b.name))) {
    const r = rel ? rel + '/' + e.name : e.name;
    if (e.isDirectory()) out.push(...walk(dir, r)); else if (e.isFile()) out.push(r);
  }
  return out;
}

// Makes the files for the saved settings into outDir (emptied first): what addin/build answers.
async function build(env, settings, outDir) {
  const from = templateDir(env);
  if (!from) return { error: "The add-in's template is not on the server (addin-template)." };
  if (!settings.addinUrl) return { error: 'Enter the address the add-in will be served from.' };
  if (!settings.clientId) return { error: "Enter the app registration's client ID: without it the add-in cannot sign in." };
  const fill = await import(pathToFileURL(path.join(ROOT, 'addin', 'fill.mjs')).href);
  fs.rmSync(outDir, { recursive: true, force: true });
  let s;
  try {
    s = fill.fillTemplate(from, outDir, { base: settings.addinUrl, clientId: settings.clientId, tenantId: settings.tenantId, site: settings.spSite, folder: settings.spFolder, host: settings.addinHost });
  } catch (e) { return { error: e.message }; }
  const version = JSON.parse(fs.readFileSync(path.join(from, 'osign-addin.json'), 'utf8')).version;
  return { files: walk(outDir).length, version, manifestUrl: s.base + 'manifest.xml', redirectUri: 'brk-multihub://' + new URL(s.base).host };
}

// A ZIP of the folder, stored (the files are small and mostly compressed already). zlib.crc32
// is in Node since 22.2.
function zip(dir) {
  const parts = [], central = [];
  let offset = 0;
  const dosTime = (() => { const d = new Date(); return { time: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1), date: ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate() }; })();
  for (const rel of walk(dir)) {
    const data = fs.readFileSync(path.join(dir, rel)), name = Buffer.from(rel, 'utf8'), crc = zlib.crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x0800, 6); local.writeUInt16LE(0, 8);
    local.writeUInt16LE(dosTime.time, 10); local.writeUInt16LE(dosTime.date, 12); local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18); local.writeUInt32LE(data.length, 22); local.writeUInt16LE(name.length, 26); local.writeUInt16LE(0, 28);
    parts.push(local, name, data);
    const c = Buffer.alloc(46);
    c.writeUInt32LE(0x02014b50, 0); c.writeUInt16LE(20, 4); c.writeUInt16LE(20, 6); c.writeUInt16LE(0x0800, 8); c.writeUInt16LE(0, 10);
    c.writeUInt16LE(dosTime.time, 12); c.writeUInt16LE(dosTime.date, 14); c.writeUInt32LE(crc, 16); c.writeUInt32LE(data.length, 20);
    c.writeUInt32LE(data.length, 24); c.writeUInt16LE(name.length, 28); c.writeUInt32LE(offset, 42);
    central.push(c, name);
    offset += local.length + name.length + data.length;
  }
  const cd = Buffer.concat(central), end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0); end.writeUInt16LE(central.length / 2, 8); end.writeUInt16LE(central.length / 2, 10);
  end.writeUInt32LE(cd.length, 12); end.writeUInt32LE(offset, 16);
  return Buffer.concat([...parts, cd, end]);
}

// Whether the address serves the files made here, byte for byte, and the manifest with a caching
// header that lets updates through. (Test-AddinOnline in OSign.ps1)
async function check(base, dir, fetchFn) {
  base = String(base).replace(/\/+$/, '') + '/';
  const checks = [];
  for (const name of ['manifest.xml', 'launchevent.js', 'taskpane.js', 'taskpane.html']) {
    try {
      const r = await fetchFn(base + name, { headers: { 'Cache-Control': 'no-cache' }, redirect: 'follow', signal: AbortSignal.timeout(20000) });
      if (!r.ok) throw new Error('the server answered ' + r.status);
      const body = Buffer.from(await r.arrayBuffer());
      let local = null; try { local = fs.readFileSync(path.join(dir, name)); } catch (e) { /* not made */ }
      const same = !!local && local.equals(body);
      checks.push({ what: name, ok: same, detail: same ? 'the same as made here' : 'there, but not the files made here: upload them again' });
      if (name === 'manifest.xml') {
        const cc = r.headers.get('cache-control') || '';
        checks.push({ what: 'Cache-Control', ok: /no-cache|max-age=0/.test(cc), detail: cc || 'not sent: updates can take a while to reach everyone' });
      }
    } catch (e) {
      checks.push({ what: name, ok: false, detail: 'cannot be loaded: ' + (e.cause && e.cause.message ? e.cause.message : e.message) });
    }
  }
  return { checks, allOk: !checks.some((c) => !c.ok && c.what !== 'Cache-Control') };
}

// ---------------------------------------------------------------------------------------------
// Cloudflare Pages, Direct Upload: the steps wrangler takes for `pages deploy`.
// ---------------------------------------------------------------------------------------------
function pagesConfig(env) {
  let token = '';
  if (env.OSIGN_CF_API_TOKEN_FILE) { try { token = fs.readFileSync(env.OSIGN_CF_API_TOKEN_FILE, 'utf8').trim(); } catch (e) { token = ''; } }
  else token = String(env.OSIGN_CF_API_TOKEN || '').trim();
  return { account: String(env.OSIGN_CF_ACCOUNT_ID || '').trim(), project: String(env.OSIGN_CF_PROJECT || '').trim(), token, api: env.OSIGN_CF_API_URL || 'https://api.cloudflare.com/client/v4' };
}
const pagesEnabled = (c) => !!(c.account || c.project);
function pagesProblem(c) {
  if (!pagesEnabled(c)) return 'Uploading to Cloudflare Pages is not set up on the server (Settings → Server connections).';
  if (!/^[0-9a-f]{32}$/i.test(c.account)) return "The Cloudflare account ID is 32 letters and digits, on the account's overview page.";
  if (!/^[a-z0-9][a-z0-9-]{0,57}$/.test(c.project)) return 'The Pages project is a name of lower-case letters, digits and -.';
  if (!c.token) return 'The Cloudflare API token is missing (secrets/cf_api_token, or Server connections).';
  return '';
}

const TYPES = { '.html': 'text/html', '.js': 'application/javascript', '.css': 'text/css', '.xml': 'application/xml', '.json': 'application/json', '.png': 'image/png', '.svg': 'image/svg+xml', '.txt': 'text/plain' };
// What Pages calls a file: the first 32 hex digits of BLAKE3 over its base64 and its extension.
function pagesHash(data, rel) { return bytesToHex(blake3(Buffer.from(data.toString('base64') + path.extname(rel).slice(1), 'utf8'))).slice(0, 32); }
// Files Pages reads as settings, not serves: they go with the deployment itself.
const SPECIAL = new Set(['_headers', '_redirects', '_routes.json', '_worker.js']);

async function pagesDeploy(c, dir, fetchFn) {
  const problem = pagesProblem(c); if (problem) throw new Error(problem);
  async function cf(method, url, { token, json, form } = {}) {
    const headers = { Authorization: 'Bearer ' + (token || c.token) };
    let body;
    if (json !== undefined) { headers['Content-Type'] = 'application/json'; body = JSON.stringify(json); }
    if (form) body = form;
    let r;
    try { r = await fetchFn(url, { method, headers, body, signal: AbortSignal.timeout(60000) }); }
    catch (e) { throw new Error('Cloudflare cannot be reached from the server: ' + e.message); }
    const j = await r.json().catch(() => ({}));
    if (!r.ok || j.success === false) {
      const why = ((j.errors || [])[0] || {}).message || r.status;
      const err = new Error(r.status === 401 || r.status === 403 ? 'Cloudflare refused the API token (' + why + '). It needs the permission Cloudflare Pages: Edit for this account.' : 'Cloudflare: ' + why);
      err.status = r.status; throw err;
    }
    return j.result;
  }
  const proj = c.api + '/accounts/' + c.account + '/pages/projects/' + c.project;
  // The project, made the first time.
  try { await cf('GET', proj); } catch (e) {
    if (e.status !== 404) throw e;
    await cf('POST', c.api + '/accounts/' + c.account + '/pages/projects', { json: { name: c.project, production_branch: 'main' } });
  }
  const files = walk(dir).filter((rel) => !SPECIAL.has(rel)).map((rel) => {
    const data = fs.readFileSync(path.join(dir, rel));
    return { rel, data, hash: pagesHash(data, rel), type: TYPES[path.extname(rel).toLowerCase()] || 'application/octet-stream' };
  });
  const jwt = (await cf('GET', proj + '/upload-token')).jwt;
  const missing = new Set(await cf('POST', c.api + '/pages/assets/check-missing', { token: jwt, json: { hashes: files.map((f) => f.hash) } }));
  const todo = files.filter((f) => missing.has(f.hash));
  // In parts, as wrangler does: a few files per request.
  for (let i = 0; i < todo.length; i += 20) {
    await cf('POST', c.api + '/pages/assets/upload', { token: jwt, json: todo.slice(i, i + 20).map((f) => ({ key: f.hash, value: f.data.toString('base64'), metadata: { contentType: f.type }, base64: true })) });
  }
  await cf('POST', c.api + '/pages/assets/upsert-hashes', { token: jwt, json: { hashes: files.map((f) => f.hash) } });
  const form = new FormData();
  form.append('manifest', JSON.stringify(Object.fromEntries(files.map((f) => ['/' + f.rel, f.hash]))));
  form.append('branch', 'main');
  for (const name of SPECIAL) {
    const p = path.join(dir, name);
    if (fs.existsSync(p)) form.append(name, new Blob([fs.readFileSync(p)]), name);
  }
  const d = await cf('POST', proj + '/deployments', { form });
  return { id: d.id, url: d.url, uploaded: todo.length, files: files.length };
}

module.exports = { templateDir, build, zip, check, pagesConfig, pagesEnabled, pagesProblem, pagesDeploy, pagesHash, walk };
