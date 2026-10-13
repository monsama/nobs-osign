// Fills in the blanks of the add-in's template (dist/template, see addin/build.mjs) with a
// company's settings. editor/OSign.ps1 does the same (Test-AddinSettings, Build-AddinPackage); the
// tests check both come out byte for byte the same.
import fs from 'node:fs';
import path from 'node:path';

const xmlEscape = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// The add-in's settings, checked. The same rules as Test-AddinSettings in editor/OSign.ps1.
export function checkSetup(s) {
  let base = String(s.base || '').trim();
  if (!/^https:\/\/[^\s"<>'&]+$/.test(base) && !/^https?:\/\/localhost(:\d+)?\//.test(base)) throw new Error('The add-in address must be an https:// address: ' + base);
  if (!base.endsWith('/')) base += '/';
  const clientId = String(s.clientId || '').trim();
  const tenantId = String(s.tenantId || '').trim();
  if (clientId && !GUID.test(clientId)) throw new Error('The client ID is the app registration\'s Application (client) ID, a GUID.');
  if (tenantId && !GUID.test(tenantId)) throw new Error('The tenant ID is a Directory (tenant) ID, a GUID.');
  const site = String(s.site || 'root').trim();
  if (!/^root$|^[A-Za-z0-9.-]+\.sharepoint\.com:\/sites\/[^:/"\\\s]+:$/.test(site)) throw new Error('The SharePoint site is "root" or "<host>.sharepoint.com:/sites/<name>:".');
  const folder = String(s.folder || 'OSign').trim().replace(/^\/+|\/+$/g, '');
  if (!folder || /["\\\u0000-\u001f]/.test(folder) || folder.split('/').some((p) => !p || p === '.' || p === '..')) throw new Error('The SharePoint folder is a folder like "OSign".');
  return { base, clientId, tenantId, site, folder };
}

// Copies the template `from` into outDir, with the blanks filled in.
export function fillTemplate(from, outDir, setup) {
  const s = checkSetup(setup);
  const js = { '@@OSIGN_CLIENT_ID@@': s.clientId, '@@OSIGN_TENANT_ID@@': s.tenantId, '@@OSIGN_SP_SITE@@': s.site, '@@OSIGN_SP_FOLDER@@': s.folder };
  const xml = { '{{BASE}}': xmlEscape(s.base), '{{ORIGIN}}': xmlEscape(new URL(s.base).origin), '{{SUPPORT}}': xmlEscape(s.base) };
  fs.mkdirSync(outDir, { recursive: true });
  fs.cpSync(from, outDir, { recursive: true });
  const fill = (file, map) => {
    let text = fs.readFileSync(path.join(outDir, file), 'utf8');
    for (const [k, v] of Object.entries(map)) text = text.split(k).join(v);
    fs.writeFileSync(path.join(outDir, file), text);
  };
  fill('launchevent.js', js);
  fill('taskpane.js', js);
  fill('manifest.xml', xml);
  return s;
}

