// A stand-in for Microsoft Graph and SharePoint, as far as the add-in uses them: the OSign folder
// of the home site (signatures.json, assets/, photos/), /me and the user's groups.
//   files: { 'signatures.json': string, 'assets/logo.png': Buffer, 'photos/mustera.jpg': Buffer }
// Returns { fetch, requests } - fetch(url, init) answers like the real one.
const GRAPH = 'https://graph.microsoft.com/v1.0';
// The home site, or a site of its own (contoso.sharepoint.com:/sites/Signatures:), found by its id.
const SITE_PATH = GRAPH + '/sites/contoso.sharepoint.com:/sites/Signatures:';
const SITE_ID = 'contoso.sharepoint.com,1111,2222';

// Or the folder by its ids (drive, item), as the server puts them into the add-in's files.
const BY_IDS = GRAPH + '/drives/' + encodeURIComponent('b!drive-1') + '/items/OSIGN-ITEM:';

module.exports = function fakeGraph({ files = {}, me = {}, groups = [], deny = [], site = 'root', byIds = false } = {}) {
  const ROOT = byIds ? BY_IDS : GRAPH + '/sites/' + (site === 'root' ? 'root' : encodeURIComponent(SITE_ID)) + '/drive/root:/OSign';
  const requests = [];
  const etag = (k) => '"' + k + ':' + String(files[k] && files[k].length) + '"';
  const resp = (status, body, bytes) => {
    if (bytes === undefined && body !== undefined) bytes = JSON.stringify(body);
    return {
    ok: status >= 200 && status < 300, status,
    json: async () => body, text: async () => (typeof bytes === 'string' ? bytes : Buffer.from(bytes || '').toString('utf8')),
    arrayBuffer: async () => { const b = Buffer.isBuffer(bytes) ? bytes : Buffer.from(String(bytes || '')); return b.buffer.slice(b.byteOffset, b.byteOffset + b.length); }
    };
  };
  const meta = (k) => ({ name: k.split('/').pop(), eTag: etag(k), size: files[k].length, '@microsoft.graph.downloadUrl': 'https://download.example/' + encodeURIComponent(k) });
  async function fetch(url, init) {
    const auth = (init && init.headers && init.headers.Authorization) || '';
    requests.push(url);
    if (url.startsWith('https://download.example/')) {
      const k = decodeURIComponent(url.slice('https://download.example/'.length));
      return files[k] === undefined ? resp(404) : resp(200, null, files[k]);
    }
    if (!url.startsWith(GRAPH)) return resp(404);
    if (deny.some((d) => url.includes(d))) return resp(403, { error: { code: 'Authorization_RequestDenied' } });
    // As Graph does: a site's path, then a path into its library, is refused.
    if (url.startsWith(SITE_PATH + '/')) return resp(400, { error: { code: 'BadRequest', message: "Resource not found for the segment 'root:'." } });
    if (url.startsWith(SITE_PATH)) return site === 'root' ? resp(404, {}) : resp(200, { id: SITE_ID });
    if (url.startsWith(ROOT)) {
      if (!/Sites\.Read\.All/.test(auth)) return resp(403, {});
      const rest = decodeURIComponent(url.slice(ROOT.length + 1));
      if (rest.startsWith('assets:/children')) {
        return resp(200, { value: Object.keys(files).filter((k) => k.startsWith('assets/')).map(meta) });
      }
      return files[rest] === undefined ? resp(404, { error: { code: 'itemNotFound' } }) : resp(200, meta(rest));
    }
    if (url.startsWith(GRAPH + '/me/transitiveMemberOf')) {
      if (!/GroupMember\.Read\.All/.test(auth)) return resp(403, {});
      return resp(200, { value: groups });
    }
    if (url.startsWith(GRAPH + '/me')) return resp(200, me);
    return resp(404);
  }
  return { fetch, requests };
};
