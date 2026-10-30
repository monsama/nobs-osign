/*
 * Publishing into SharePoint through Microsoft Graph, with the server's own app registration
 * (client credentials): the same site and folder the add-in reads (settings spSite, spFolder).
 *
 *   OSIGN_GRAPH_TENANT_ID           the Directory (tenant) ID
 *   OSIGN_GRAPH_CLIENT_ID           the server's app registration (not the add-in's)
 *   OSIGN_GRAPH_CLIENT_SECRET_FILE  its client secret, as a Docker secret
 *
 * Or, instead of the secret, the server's own certificate (server/cert.js), which the setup
 * script puts on the app: then the server signs in with a signed client assertion.
 *
 * The app needs the application permission Sites.Selected with write access granted on that one
 * site (docs/SETUP.md), or Sites.ReadWrite.All, which reaches every site and is not recommended.
 *
 * A target is a folder with paths in it ("signatures.json", "assets/logo.png"): list, read, write,
 * remove, exists. server/store.js publishes into one, here or on the disk (folderTarget).
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const { writeAtomic } = require('./store');
const { clientAssertion } = require('./cert');

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function graphConfig(env, cert) {
  let secret = '';
  if (env.OSIGN_GRAPH_CLIENT_SECRET_FILE) { try { secret = fs.readFileSync(env.OSIGN_GRAPH_CLIENT_SECRET_FILE, 'utf8').trim(); } catch (e) { secret = ''; } }
  else secret = String(env.OSIGN_GRAPH_CLIENT_SECRET || '').trim();
  // A placeholder of all zeros (as in examples) counts as not set.
  const id = (v) => { const x = String(v || '').trim(); return /^0{8}-0{4}-0{4}-0{4}-0{12}$/.test(x) ? '' : x; };
  return {
    tenantId: id(env.OSIGN_GRAPH_TENANT_ID),
    clientId: id(env.OSIGN_GRAPH_CLIENT_ID),
    secret,
    // The certificate in use (server/cert.js): it wins over a secret.
    cert: cert || null,
    // Only the tests point these elsewhere (a stand-in Graph).
    graphUrl: env.OSIGN_GRAPH_URL || 'https://graph.microsoft.com/v1.0',
    loginUrl: env.OSIGN_LOGIN_URL || 'https://login.microsoftonline.com'
  };
}
const graphEnabled = (c) => !!(c.tenantId || c.clientId);
function graphProblem(c) {
  if (!graphEnabled(c)) return '';
  if (!GUID.test(c.tenantId)) return 'The tenant ID is the Directory (tenant) ID, a GUID.';
  if (!GUID.test(c.clientId)) return "The client ID is the server's app registration's Application (client) ID, a GUID.";
  if (!c.secret && !c.cert) return 'The server cannot sign in to Microsoft 365 yet: run the setup script, which gives it its own certificate.';
  return '';
}

// What Graph's answers mean, in the words of the setup.
function graphError(status, what, detail) {
  if (status === 401 || status === 403) return what + ': the server may not do this in SharePoint (' + status + "). Give its app registration write access to the site (Sites.Selected, see the setup assistant)." + (detail ? ' ' + detail : '');
  if (status === 404) return what + ': not found in SharePoint (404). Check the site and the folder.';
  if (status === 429 || status === 503) return what + ': SharePoint is busy (' + status + '). Try again in a minute.';
  return what + ': Microsoft Graph answered ' + status + '.' + (detail ? ' ' + detail : '');
}

function createGraph(c, opts) {
  const fetchFn = (opts && opts.fetch) || fetch;
  let tok = null;
  async function token() {
    if (tok && tok.until > Date.now()) return tok.value;
    const url = c.loginUrl + '/' + encodeURIComponent(c.tenantId) + '/oauth2/v2.0/token';
    const body = new URLSearchParams({ client_id: c.clientId, grant_type: 'client_credentials', scope: 'https://graph.microsoft.com/.default' });
    if (c.cert) {
      body.set('client_assertion_type', 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer');
      body.set('client_assertion', clientAssertion({ keyPem: c.cert.keyPem, x5t: c.cert.x5t, clientId: c.clientId, audience: url }));
    } else body.set('client_secret', c.secret);
    let r;
    try { r = await fetchFn(url, { method: 'POST', body, signal: AbortSignal.timeout(15000) }); }
    catch (e) { throw new Error('Microsoft Entra ID cannot be reached from the server: ' + e.message); }
    const j = await r.json().catch(() => ({}));
    if (!r.ok || !j.access_token) {
      const code = String(j.error_description || j.error || r.status).split(/\r?\n/)[0];
      throw new Error("The server's app registration could not sign in (" + code.slice(0, 200) + '). Check the tenant ID, the client ID and ' + (c.cert ? "that the server's certificate is on the app registration (run the setup script again)." : 'the client secret.'));
    }
    // Used until five minutes before it ends.
    tok = { value: j.access_token, until: Date.now() + Math.max(60, (Number(j.expires_in) || 3600) - 300) * 1000 };
    return tok.value;
  }
  // One request; a busy SharePoint (429, 503) is asked once more after the time it names.
  async function call(method, url, { body, type, what, okStatus = [], raw } = {}) {
    for (let attempt = 0; ; attempt++) {
      const headers = { Authorization: 'Bearer ' + await token() };
      if (type) headers['Content-Type'] = type;
      let r;
      try { r = await fetchFn(url, { method, headers, body, redirect: 'follow', signal: AbortSignal.timeout(30000) }); }
      catch (e) { throw new Error(what + ': Microsoft Graph cannot be reached from the server: ' + e.message); }
      if ((r.status === 429 || r.status === 503) && attempt === 0) {
        const wait = Math.min(30, Number(r.headers.get('retry-after')) || 2);
        await r.arrayBuffer().catch(() => {});
        await new Promise((res) => setTimeout(res, wait * 1000));
        continue;
      }
      if (r.ok || okStatus.includes(r.status)) return r;
      let detail = '';
      try { const j = await r.json(); detail = (j && j.error && j.error.message) || ''; } catch (e) { /* none */ }
      const err = new Error(graphError(r.status, what, String(detail).slice(0, 200)));
      err.status = r.status;
      throw err;
    }
  }

  // The site as Graph addresses it: root, or host:/sites/name: (each part encoded).
  function siteUrl(site) {
    if (!site || site === 'root') return c.graphUrl + '/sites/root';
    const m = /^([A-Za-z0-9.-]+):\/(sites|teams|personal)\/([^:/]+):$/.exec(site);
    if (!m) throw new Error('The SharePoint site is root, or <host>.sharepoint.com:/sites/<name>:');
    return c.graphUrl + '/sites/' + m[1] + ':/' + m[2] + '/' + encodeURIComponent(m[3]) + ':';
  }

  // A site by its id: Graph takes a site's path, or a path in its drive, but not both in one URL
  // ("Resource not found for the segment 'root:'"), so the path is asked for the id once.
  const siteIds = new Map();
  function siteBase(site) {
    if (!site || site === 'root') return Promise.resolve(c.graphUrl + '/sites/root');
    if (!siteIds.has(site)) {
      const p = call('GET', siteUrl(site) + '?$select=id', { what: 'Finding the site ' + site.replace(/:\//, '/').replace(/:$/, '') })
        .then((r) => r.json()).then((j) => { if (!j.id) throw new Error('Microsoft Graph did not name the site\'s id.'); return c.graphUrl + '/sites/' + j.id; });
      p.catch(() => siteIds.delete(site));
      siteIds.set(site, p);
    }
    return siteIds.get(site);
  }

  function target(site, folder) {
    const baseOf = async () => (await siteBase(site)) + '/drive/root';
    const parts = (rel) => (String(folder || '') + '/' + String(rel || '')).split('/').filter(Boolean);
    const itemUrl = async (rel) => (await baseOf()) + ':/' + parts(rel).map(encodeURIComponent).join('/') + ':';
    // As people read it: SharePoint contoso.sharepoint.com/sites/IT › OSign (root: the home site).
    const where = 'SharePoint ' + (!site || site === 'root' ? 'root' : site.replace(/:\//, '/').replace(/:$/, '')) + ' › ' + folder;

    async function meta(rel, what) {
      const r = await call('GET', (await itemUrl(rel)) + '?$select=id,name,folder,file,size', { what, okStatus: [404] });
      return r.status === 404 ? null : r.json();
    }
    // The folders on the way, made where they are missing (a new OSign folder, its assets/).
    async function ensureFolder(rel) {
      const p = parts(rel), base = await baseOf();
      for (let i = 1; i <= p.length; i++) {
        const sub = p.slice(0, i);
        const r = await call('GET', base + ':/' + sub.map(encodeURIComponent).join('/') + ':?$select=id,folder', { what: 'Finding the folder ' + sub.join('/'), okStatus: [404] });
        if (r.status !== 404) { await r.arrayBuffer().catch(() => {}); continue; }
        const parent = i === 1 ? base : base + ':/' + sub.slice(0, -1).map(encodeURIComponent).join('/') + ':';
        await call('POST', parent + '/children', {
          what: 'Making the folder ' + sub.join('/'), okStatus: [409], type: 'application/json',
          body: JSON.stringify({ name: sub[sub.length - 1], folder: {}, '@microsoft.graph.conflictBehavior': 'fail' })
        }).then((x) => x.arrayBuffer());
      }
    }
    // The folders known to be there, so writing many files asks for each folder once.
    const made = new Set();
    return {
      describe: () => where,
      // The folder by its ids (its library's, its own): what the add-in reads with, so a user who
      // may read the folder needs nothing of the site around it.
      async locate() {
        const r = await call('GET', (await itemUrl('')) + '?$select=id,parentReference', { what: 'Finding the folder ' + folder });
        const j = await r.json();
        return { driveId: String((j.parentReference || {}).driveId || ''), itemId: String(j.id || '') };
      },
      async exists(rel) { const m = await meta(rel || '', 'Finding ' + (rel || 'the folder')); return !!(m && m.folder); },
      // The names of the files in a folder ([] when it is not there).
      async list(rel) { return (await this.listInfo(rel)).map((f) => f.name); },
      // The same, each with a tag that changes when the file does (its eTag): what changed since
      // it was last read can be told without reading every file.
      async listInfo(rel) {
        const out = [];
        let url = (await itemUrl(rel)) + '/children?$select=name,file,eTag&$top=999';
        // A page holds up to 999: 200 pages are some 200 000 users.
        for (let page = 0; url && page < 200; page++) {
          const r = await call('GET', url, { what: 'Listing ' + rel, okStatus: [404] });
          if (r.status === 404) return [];
          const j = await r.json();
          for (const it of j.value || []) if (it.file) out.push({ name: it.name, tag: String(it.eTag || '') });
          url = j['@odata.nextLink'];
        }
        return out;
      },
      async read(rel) {
        const r = await call('GET', (await itemUrl(rel)) + '/content', { what: 'Reading ' + rel, okStatus: [404] });
        return r.status === 404 ? null : Buffer.from(await r.arrayBuffer());
      },
      // Simple upload: signatures.json and the images are far below its 250 MB.
      // Answers the file's new tag (as listInfo has it).
      async write(rel, bytes, type) {
        const dir = path.posix.dirname(rel) === '.' ? '' : path.posix.dirname(rel);
        if (!made.has(dir)) { await ensureFolder(dir); made.add(dir); }
        const r = await call('PUT', (await itemUrl(rel)) + '/content', { what: 'Writing ' + rel, body: bytes, type: type || 'application/octet-stream' });
        const j = await r.json().catch(() => ({}));
        return String(j.eTag || '');
      },
      async remove(rel) { await call('DELETE', await itemUrl(rel), { what: 'Removing ' + rel, okStatus: [404] }).then((x) => x.arrayBuffer().catch(() => {})); }
    };
  }
  // A sharing link (https://…/:f:/g/…): what it leads to, as Graph names it (/shares, with the
  // link encoded as u!<base64url>). Its webUrl is the item's own address, the folder in it.
  async function share(link) {
    const id = 'u!' + Buffer.from(String(link), 'utf8').toString('base64').replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');
    const r = await call('GET', c.graphUrl + '/shares/' + id + '/driveItem?$select=name,webUrl,folder,file', { what: 'Opening the sharing link' });
    const j = await r.json();
    return { name: String(j.name || ''), webUrl: String(j.webUrl || ''), folder: !!j.folder };
  }
  // A user from Entra ID, for the preview ("Person from Entra ID"): by UPN or mail, the properties
  // asked for, and the groups (nested ones too). Needs User.Read.All and GroupMember.Read.All
  // (application), which the setup script gives the server's app when asked to.
  // People to pick from, for the preview: name, mail or login starting with what is typed.
  async function users(query) {
    const q = String(query || '').trim();
    if (q.length < 2) return [];
    const v = q.replace(/'/g, "''");
    const f = ['displayName', 'givenName', 'surname', 'mail', 'userPrincipalName'].map((p) => "startswith(" + p + ",'" + v + "')").join(' or ');
    try {
      const r = await call('GET', c.graphUrl + '/users?$filter=' + encodeURIComponent(f) + '&$select=displayName,mail,userPrincipalName,department&$top=15', { what: 'Finding users' });
      return ((await r.json()).value || []).map((u) => ({ login: u.userPrincipalName || '', name: u.displayName || '', mail: u.mail || '', department: u.department || '' }))
        .sort((a, b) => (a.name || a.login).localeCompare(b.name || b.login));
    } catch (e) {
      if (e.status === 403 || e.status === 401) throw new Error("The server's app may not read users from Entra ID (User.Read.All, GroupMember.Read.All): run the setup script again with that box ticked.");
      throw e;
    }
  }
  async function user(query, select) {
    const q = String(query || '').trim();
    if (!q) throw new Error('Enter a user name or mail address.');
    const sel = '$select=' + encodeURIComponent(['id', 'displayName'].concat(select || []).join(','));
    let u = null;
    try {
      const r = await call('GET', c.graphUrl + '/users/' + encodeURIComponent(q) + '?' + sel, { what: 'Reading the user', okStatus: [404] });
      if (r.status !== 404) u = await r.json();
      if (!u) {
        const f = await call('GET', c.graphUrl + '/users?$filter=' + encodeURIComponent("mail eq '" + q.replace(/'/g, "''") + "'") + '&' + sel, { what: 'Reading the user' });
        const list = (await f.json()).value || [];
        if (list.length > 1) throw new Error('More than one user has the mail address ' + q + '.');
        u = list[0] || null;
      }
      if (!u) throw new Error('No user ' + q + ' in Entra ID.');
      const groups = new Set();
      let next = c.graphUrl + '/users/' + encodeURIComponent(u.id) + '/transitiveMemberOf/microsoft.graph.group?$select=displayName,onPremisesSamAccountName&$top=999';
      while (next) {
        const g = await (await call('GET', next, { what: 'Reading the user\'s groups' })).json();
        for (const x of g.value || []) for (const n of [x.onPremisesSamAccountName, x.displayName]) if (n) groups.add(n);
        next = g['@odata.nextLink'] || null;
      }
      return { user: u, groups: [...groups].sort((a, b) => a.localeCompare(b)) };
    } catch (e) {
      if (e.status === 403 || e.status === 401) throw new Error("The server's app may not read users from Entra ID (User.Read.All, GroupMember.Read.All): run the setup script again with that box ticked.");
      throw e;
    }
  }
  return { target, token, share, user, users };
}

// The same, in a folder on the disk: for a server without Graph (development, tests).
function folderTarget(root) {
  const full = (rel) => path.join(root, ...String(rel || '').split('/').filter(Boolean));
  return {
    describe: () => root,
    async exists(rel) { try { return fs.statSync(full(rel)).isDirectory(); } catch (e) { return false; } },
    async list(rel) { return (await this.listInfo(rel)).map((f) => f.name); },
    async listInfo(rel) {
      let names = [];
      try { names = fs.readdirSync(full(rel)); } catch (e) { return []; }
      const out = [];
      for (const n of names) { const st = fs.statSync(path.join(full(rel), n)); if (st.isFile()) out.push({ name: n, tag: st.size + ':' + st.mtimeMs }); }
      return out;
    },
    async read(rel) { try { return fs.readFileSync(full(rel)); } catch (e) { return null; } },
    async write(rel, bytes) { writeAtomic(full(rel), bytes); const st = fs.statSync(full(rel)); return st.size + ':' + st.mtimeMs; },
    async remove(rel) { fs.rmSync(full(rel), { force: true }); }
  };
}

module.exports = { graphConfig, graphEnabled, graphProblem, createGraph, folderTarget };
