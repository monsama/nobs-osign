// Stands in for Microsoft Entra ID (the token endpoint) and Microsoft Graph's SharePoint drive API,
// for the server's publishing (server/graph.js). Like an app with Sites.Selected, it reaches only
// the sites it was granted; any other answers 403.
const http = require('node:http');

const SITE_RE = /^\/v1\.0\/sites\/(root|[^/]+:\/(?:sites|teams|personal)\/[^:/]+:)\/drive\/root(?::\/([^:]*):)?(\/children|\/content)?$/;

function fakeSharePoint({ tenantId, clientId, secret, granted }) {
  // Per site: files (path → Buffer) and folders (path), paths like "OSign/assets/logo.png".
  const sites = {};
  for (const s of granted) sites[s] = { files: new Map(), folders: new Set() };
  const requests = [];
  let busy = 0, tokens = 0;

  const server = http.createServer(async (req, res) => {
    const u = new URL(req.url, 'http://x');
    const body = await new Promise((resolve) => { const p = []; req.on('data', (d) => p.push(d)); req.on('end', () => resolve(Buffer.concat(p))); });
    requests.push(req.method + ' ' + decodeURIComponent(u.pathname));
    const json = (status, obj, headers) => { res.writeHead(status, Object.assign({ 'Content-Type': 'application/json' }, headers || {})); res.end(JSON.stringify(obj)); };

    if (req.method === 'POST' && u.pathname === '/' + tenantId + '/oauth2/v2.0/token') {
      const f = new URLSearchParams(body.toString());
      if (f.get('client_id') !== clientId || f.get('client_secret') !== secret || f.get('grant_type') !== 'client_credentials') return json(401, { error: 'invalid_client', error_description: 'AADSTS7000215: Invalid client secret provided.\r\nTrace ID: x' });
      tokens++;
      return json(200, { access_token: 'tok-' + tokens, expires_in: 3600, token_type: 'Bearer' });
    }
    if (!/^Bearer tok-\d+$/.test(String(req.headers.authorization || ''))) return json(401, { error: { code: 'InvalidAuthenticationToken', message: 'Access token is empty.' } });
    if (busy > 0) { busy--; return json(429, { error: { code: 'TooManyRequests', message: 'slow down' } }, { 'Retry-After': '0' }); }

    const m = SITE_RE.exec(decodeURIComponent(u.pathname));
    if (!m) return json(400, { error: { code: 'BadRequest', message: 'not understood: ' + u.pathname } });
    const site = sites[m[1]];
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
      for (const p of site.files.keys()) if (parentOf(p) === rel) value.push({ name: p.split('/').pop(), file: {} });
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
      return json(201, { id: rel, name: rel.split('/').pop(), file: {} });
    }
    if (req.method === 'DELETE' && !suffix) {
      if (!site.files.delete(rel) && !site.folders.delete(rel)) return json(404, { error: { code: 'itemNotFound', message: 'not found' } });
      res.writeHead(204); return res.end();
    }
    return json(405, { error: { code: 'notAllowed', message: req.method + ' ' + suffix } });
  });
  return {
    sites, requests,
    busyOnce: () => { busy = 1; },
    listen: () => new Promise((resolve) => server.listen(0, '127.0.0.1', () => resolve('http://127.0.0.1:' + server.address().port))),
    close: () => server.close()
  };
}

module.exports = { fakeSharePoint };
