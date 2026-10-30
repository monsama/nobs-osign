// Stands in for Microsoft Entra ID (the token endpoint) and Microsoft Graph's SharePoint drive API,
// for the server's publishing (server/graph.js). Like an app with Sites.Selected, it reaches only
// the sites it was granted; any other answers 403.
const http = require('node:http');
const crypto = require('node:crypto');
// A file's eTag changes when its content does.
const etag = (buf) => '"{' + crypto.createHash('sha1').update(buf).digest('hex').slice(0, 32) + '},1"';

// A site by its id (host,guid,guid, as Graph answers it for its path) or root, then its drive.
const SITE_RE = /^\/v1\.0\/sites\/(root|[^/:]+,[^/:]+,[^/:]+)\/drive\/root(?::\/([^:]*):)?(\/children|\/content)?$/;
// A site by its path: only on its own (Graph refuses a second path after it, as below).
const SITE_PATH_RE = /^\/v1\.0\/sites\/([^/]+:\/(?:sites|teams|personal)\/[^:/]+:)$/;
const siteId = (key) => key.split(':')[0] + ',' + crypto.createHash('md5').update(key).digest('hex').slice(0, 8) + ',1';

function fakeSharePoint({ tenantId, clientId, secret, granted }) {
  // Per site: files (path → Buffer) and folders (path), paths like "OSign/assets/logo.png".
  const sites = {};
  for (const s of granted) sites[s] = { files: new Map(), folders: new Set() };
  const requests = [];
  const shares = {};
  const users = {};
  let busy = 0, tokens = 0;
  // Certificates on the app (X509Certificate), for a client assertion.
  const certs = [];

  const server = http.createServer(async (req, res) => {
    const u = new URL(req.url, 'http://x');
    const body = await new Promise((resolve) => { const p = []; req.on('data', (d) => p.push(d)); req.on('end', () => resolve(Buffer.concat(p))); });
    requests.push(req.method + ' ' + decodeURIComponent(u.pathname));
    const json = (status, obj, headers) => { res.writeHead(status, Object.assign({ 'Content-Type': 'application/json' }, headers || {})); res.end(JSON.stringify(obj)); };

    if (req.method === 'POST' && u.pathname === '/' + tenantId + '/oauth2/v2.0/token') {
      const f = new URLSearchParams(body.toString());
      // A client assertion signed with a certificate on the app (certs), or the secret.
      const assertion = f.get('client_assertion');
      if (assertion) {
        const [h, b, sig] = assertion.split('.');
        const head = JSON.parse(Buffer.from(h, 'base64url')), claims = JSON.parse(Buffer.from(b, 'base64url'));
        const cert = certs.find((c) => crypto.createHash('sha1').update(c.raw).digest('base64url') === head.x5t);
        const good = cert && head.alg === 'RS256' && f.get('client_assertion_type') === 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer' &&
          crypto.verify('sha256', Buffer.from(h + '.' + b), cert.publicKey, Buffer.from(sig, 'base64url')) &&
          claims.iss === clientId && claims.sub === clientId && claims.aud === 'http://' + req.headers.host + u.pathname && claims.exp > Date.now() / 1000;
        if (!good || f.get('client_id') !== clientId) return json(401, { error: 'invalid_client', error_description: 'AADSTS700027: Client assertion failed signature validation.\r\nTrace ID: x' });
      } else if (f.get('client_id') !== clientId || f.get('client_secret') !== secret || f.get('grant_type') !== 'client_credentials') return json(401, { error: 'invalid_client', error_description: 'AADSTS7000215: Invalid client secret provided.\r\nTrace ID: x' });
      tokens++;
      return json(200, { access_token: 'tok-' + tokens, expires_in: 3600, token_type: 'Bearer' });
    }
    if (!/^Bearer tok-\d+$/.test(String(req.headers.authorization || ''))) return json(401, { error: { code: 'InvalidAuthenticationToken', message: 'Access token is empty.' } });
    if (busy > 0) { busy--; return json(429, { error: { code: 'TooManyRequests', message: 'slow down' } }, { 'Retry-After': '0' }); }

    // Users and their groups (the preview of a real person from Entra ID), as the tests set them up.
    let um;
    if (req.method === 'GET' && (um = /^\/v1\.0\/users\/([^/]+)$/.exec(u.pathname))) {
      if (users.denied) return json(403, { error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges' } });
      const x = users[decodeURIComponent(um[1]).toLowerCase()];
      return x ? json(200, x) : json(404, { error: { code: 'Request_ResourceNotFound', message: 'not found' } });
    }
    if (req.method === 'GET' && u.pathname === '/v1.0/users') {
      const mail = (/mail eq '([^']*)'/.exec(u.searchParams.get('$filter') || '') || [])[1] || '';
      return json(200, { value: Object.values(users).filter((x, i, all) => x.mail && x.mail.toLowerCase() === mail.toLowerCase() && all.indexOf(x) === i) });
    }
    if (req.method === 'GET' && (um = /^\/v1\.0\/users\/([^/]+)\/transitiveMemberOf\/microsoft\.graph\.group$/.exec(u.pathname))) {
      const x = Object.values(users).find((y) => y.id === decodeURIComponent(um[1]));
      return json(200, { value: (x && x.groups || []).map((g) => ({ displayName: g })) });
    }
    // A sharing link, as the tests set it up (shares: link → { webUrl, name, folder }).
    const sh = /^\/v1\.0\/shares\/u!([A-Za-z0-9_-]+)\/driveItem$/.exec(u.pathname);
    if (sh && req.method === 'GET') {
      const link = Buffer.from(sh[1].replace(/-/g, '+').replace(/_/g, '/'), 'base64').toString('utf8');
      const item = shares[link];
      if (shares.denied) return json(403, { error: { code: 'accessDenied', message: 'Access denied' } });
      if (!item) return json(404, { error: { code: 'itemNotFound', message: 'The sharing link no longer exists.' } });
      return json(200, Object.assign({ name: item.name, webUrl: item.webUrl }, item.folder ? { folder: { childCount: 0 } } : { file: {} }));
    }
    const path0 = decodeURIComponent(u.pathname);
    // As Graph: a site by its path, then a path in its drive, in one URL, is refused.
    if (/\/sites\/[^/]+:\/[^:]*:\/drive\/root:/.test(path0)) return json(400, { error: { code: 'BadRequest', message: "Resource not found for the segment 'root:'." } });
    const sp = SITE_PATH_RE.exec(path0);
    if (sp && req.method === 'GET') return sites[sp[1]] ? json(200, { id: siteId(sp[1]), webUrl: 'https://' + sp[1].replace(/:\//, '/').replace(/:$/, '') }) : json(403, { error: { code: 'accessDenied', message: 'Access denied' } });
    const m = SITE_RE.exec(path0);
    if (!m) return json(400, { error: { code: 'BadRequest', message: 'not understood: ' + u.pathname } });
    const site = m[1] === 'root' ? sites.root : sites[Object.keys(sites).find((k) => siteId(k) === m[1])];
    if (!site) return json(403, { error: { code: 'accessDenied', message: 'Access denied' } });
    const rel = (m[2] || '').split('/').filter(Boolean).join('/');
    const suffix = m[3] || '';
    const parentOf = (p) => p.split('/').slice(0, -1).join('/');
    const exists = (p) => !p || site.folders.has(p);

    if (req.method === 'GET' && !suffix) {
      if (site.files.has(rel)) return json(200, { id: rel, name: rel.split('/').pop(), file: {}, size: site.files.get(rel).length });
      if (rel && site.folders.has(rel)) return json(200, { id: rel, name: rel.split('/').pop(), folder: { childCount: 0 } });
      if (!rel) return json(200, { id: 'root', name: 'root', folder: {} });
      return json(404, { error: { code: 'itemNotFound', message: 'The resource could not be found.' } });
    }
    if (req.method === 'GET' && suffix === '/children') {
      if (!exists(rel)) return json(404, { error: { code: 'itemNotFound', message: 'not found' } });
      const value = [];
      for (const [p, buf] of site.files) if (parentOf(p) === rel) value.push({ name: p.split('/').pop(), file: {}, eTag: etag(buf) });
      for (const p of site.folders) if (parentOf(p) === rel) value.push({ name: p.split('/').pop(), folder: {} });
      return json(200, { value });
    }
    if (req.method === 'GET' && suffix === '/content') {
      if (!site.files.has(rel)) return json(404, { error: { code: 'itemNotFound', message: 'not found' } });
      res.writeHead(200, { 'Content-Type': 'application/octet-stream' }); return res.end(site.files.get(rel));
    }
    if (req.method === 'POST' && suffix === '/children') {
      if (!exists(rel)) return json(404, { error: { code: 'itemNotFound', message: 'parent not found' } });
      const j = JSON.parse(body.toString());
      const p = (rel ? rel + '/' : '') + j.name;
      if (site.folders.has(p) || site.files.has(p)) return json(409, { error: { code: 'nameAlreadyExists', message: 'exists' } });
      site.folders.add(p);
      return json(201, { id: p, name: j.name, folder: {} });
    }
    if (req.method === 'PUT' && suffix === '/content') {
      // As SharePoint: the folder the file goes in has to be there.
      if (!exists(parentOf(rel))) return json(404, { error: { code: 'itemNotFound', message: 'parent not found' } });
      site.files.set(rel, body);
      return json(201, { id: rel, name: rel.split('/').pop(), file: {}, eTag: etag(body) });
    }
    if (req.method === 'DELETE' && !suffix) {
      if (!site.files.delete(rel) && !site.folders.delete(rel)) return json(404, { error: { code: 'itemNotFound', message: 'not found' } });
      res.writeHead(204); return res.end();
    }
    return json(405, { error: { code: 'notAllowed', message: req.method + ' ' + suffix } });
  });
  return {
    sites, requests, shares, users, certs,
    busyOnce: () => { busy = 1; },
    listen: () => new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve('http://127.0.0.1:' + server.address().port))),
    close: () => server.close()
  };
}

module.exports = { fakeSharePoint };
