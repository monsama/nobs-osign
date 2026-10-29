// Stands in for Microsoft's sign-in and Microsoft Graph, as far as the setup script uses them
// (server/setup-template.ps1): the sign-in with PKCE, app registrations, service principals,
// consents, app role assignments, client secrets and a site's permissions.
const http = require('node:http');
const crypto = require('node:crypto');

const TENANT = '11111111-2222-3333-4444-555555555555';
const ME = 'aaaaaaaa-0000-0000-0000-000000000001';
const CLI = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
const GRAPH_APP = '00000003-0000-0000-c000-000000000000';
const b64u = (x) => Buffer.from(typeof x === 'string' ? x : JSON.stringify(x)).toString('base64url');

function fakeMicrosoft() {
  const id = () => crypto.randomUUID();
  const state = {
    apps: [], sps: [{ id: 'sp-graph', appId: GRAPH_APP }, { id: 'sp-cli', appId: CLI }], grants: [], roleAssignments: [], sitePerms: [], passwords: [], codes: new Map(), log: []
  };
  const graphSp = {
    id: 'sp-graph', appRoles: [{ id: 'role-sites-selected', value: 'Sites.Selected' }, { id: 'role-other', value: 'Sites.Read.All' }],
    oauth2PermissionScopes: ['User.Read', 'Sites.Read.All', 'GroupMember.Read.All', 'Mail.Read'].map((v) => ({ id: 'scope-' + v, value: v }))
  };
  // The consent the sign-in gives, as Entra keeps it: for this account, with what it was asked for.
  state.grants.push({ id: 'grant-cli-me', clientId: 'sp-cli', consentType: 'Principal', principalId: ME, resourceId: 'sp-graph', scope: 'openid profile Application.ReadWrite.All AppRoleAssignment.ReadWrite.All DelegatedPermissionGrant.ReadWrite.All Sites.FullControl.All' });
  const server = http.createServer(async (req, res) => {
    const u = new URL(req.url, 'http://x');
    const body = await new Promise((r) => { const p = []; req.on('data', (d) => p.push(d)); req.on('end', () => r(Buffer.concat(p).toString())); });
    const json = (status, obj) => { res.writeHead(status, { 'Content-Type': 'application/json' }); res.end(obj === undefined ? '' : JSON.stringify(obj)); };
    const p = decodeURIComponent(u.pathname);
    state.log.push(req.method + ' ' + p);
    // Sign-in.
    if (p === '/organizations/oauth2/v2.0/authorize') {
      const q = u.searchParams;
      if (q.get('client_id') !== CLI || q.get('code_challenge_method') !== 'S256' || !/Sites\.FullControl\.All/.test(q.get('scope'))) return json(400, { error: 'bad authorize' });
      const code = id();
      state.codes.set(code, { challenge: q.get('code_challenge'), redirect: q.get('redirect_uri') });
      res.writeHead(302, { Location: q.get('redirect_uri') + '/?code=' + code + '&state=' + encodeURIComponent(q.get('state')) });
      return res.end();
    }
    if (p === '/organizations/oauth2/v2.0/token') {
      const f = new URLSearchParams(body);
      const c = state.codes.get(f.get('code'));
      const ok = c && f.get('client_id') === CLI && f.get('redirect_uri') === c.redirect && crypto.createHash('sha256').update(f.get('code_verifier') || '').digest('base64url') === c.challenge;
      if (!ok) return json(400, { error: 'invalid_grant', error_description: 'AADSTS70008: bad code' });
      state.codes.delete(f.get('code'));
      return json(200, { access_token: b64u({ alg: 'none' }) + '.' + b64u({ tid: TENANT, oid: ME }) + '.x', token_type: 'Bearer' });
    }
    if (!/^Bearer [\w-]+\.[\w-]+\.x$/.test(String(req.headers.authorization || ''))) return json(401, { error: { message: 'no token' } });
    const j = body ? JSON.parse(body) : null;
    const filter = u.searchParams.get('$filter') || '';
    const eq = (k) => (new RegExp(k + " eq '([^']*)'").exec(filter) || [])[1];
    let m;
    if (p === '/v1.0/me') return json(200, { id: ME, userPrincipalName: 'admin@contoso.com' });
    if ((m = /^\/v1\.0\/servicePrincipals\(appId='([^']+)'\)$/.exec(p))) {
      if (m[1] === GRAPH_APP) return json(200, graphSp);
      const sp = state.sps.find((s) => s.appId === m[1]);
      return sp ? json(200, sp) : json(404, { error: { message: 'Resource not found' } });
    }
    if (p === '/v1.0/servicePrincipals' && req.method === 'POST') { const sp = { id: 'sp-' + id(), appId: j.appId }; state.sps.push(sp); return json(201, sp); }
    if (p === '/v1.0/applications' && req.method === 'GET') return json(200, { value: state.apps.filter((a) => a.displayName === eq('displayName')) });
    if (p === '/v1.0/applications' && req.method === 'POST') { const a = Object.assign({ id: 'obj-' + id(), appId: id(), spa: { redirectUris: [] } }, j); state.apps.push(a); return json(201, a); }
    if ((m = /^\/v1\.0\/applications\/([^/]+)$/.exec(p)) && req.method === 'PATCH') { Object.assign(state.apps.find((a) => a.id === m[1]), j); return json(204); }
    if ((m = /^\/v1\.0\/applications\/([^/]+)\/addPassword$/.exec(p))) {
      const pw = { secretText: 'secret-' + id(), endDateTime: j.passwordCredential.endDateTime, app: m[1] };
      state.passwords.push(pw); return json(200, pw);
    }
    if (p === '/v1.0/oauth2PermissionGrants' && req.method === 'GET') {
      return json(200, { value: state.grants.filter((g) => g.clientId === eq('clientId') && (!eq('consentType') || g.consentType === eq('consentType')) && (!eq('resourceId') || g.resourceId === eq('resourceId'))) });
    }
    if (p === '/v1.0/oauth2PermissionGrants' && req.method === 'POST') { const g = Object.assign({ id: 'grant-' + id() }, j); state.grants.push(g); return json(201, g); }
    if ((m = /^\/v1\.0\/oauth2PermissionGrants\/([^/]+)$/.exec(p))) {
      const i = state.grants.findIndex((g) => g.id === m[1]);
      if (req.method === 'DELETE') state.grants.splice(i, 1); else Object.assign(state.grants[i], j);
      return json(204);
    }
    if ((m = /^\/v1\.0\/servicePrincipals\/([^/]+)\/appRoleAssignments$/.exec(p))) {
      if (req.method === 'GET') return json(200, { value: state.roleAssignments.filter((r) => r.principalId === m[1]) });
      state.roleAssignments.push(j); return json(201, j);
    }
    if (p === '/v1.0/sites/contoso.sharepoint.com:/sites/osign:') return json(200, { id: 'contoso.sharepoint.com,1,2', webUrl: 'https://contoso.sharepoint.com/sites/osign' });
    if (p === '/v1.0/sites/contoso.sharepoint.com,1,2/permissions') {
      if (req.method === 'GET') return json(200, { value: state.sitePerms });
      state.sitePerms.push({ id: id(), roles: j.roles, grantedToIdentitiesV2: j.grantedToIdentities }); return json(201, {});
    }
    return json(404, { error: { message: 'not here: ' + req.method + ' ' + p } });
  });
  return { state, listen: () => new Promise((r) => server.listen(0, '127.0.0.1', () => r('http://127.0.0.1:' + server.address().port))), close: () => server.close() };
}

module.exports = { fakeMicrosoft, TENANT, ME };
