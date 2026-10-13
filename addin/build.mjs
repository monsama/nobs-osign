// Builds the add-in into dist/: the two bundles, its pages, the icons and the manifest - the site
// GitHub Pages serves. No company data goes in: the signatures are read at run time from the
// company's own SharePoint.
//
//   BASE_URL=https://monsama.github.io/osign/ OSIGN_CLIENT_ID=<app id> node addin/build.mjs
//
// BASE_URL         where dist/ will be served (default: GitHub Pages of this repository)
// OSIGN_CLIENT_ID  the Entra ID app registration the add-in signs in with (not a secret)
// OSIGN_TENANT_ID  only for an app registration of a single organization; empty: the user's own
// OSIGN_SP_SITE    the SharePoint site whose document library holds the folder: "root" (default,
//                  the tenant's home site) or "<host>.sharepoint.com:/sites/<name>:"
// OSIGN_SP_FOLDER  the folder in that library (default: OSign)
import { build } from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const dist = path.join(root, 'dist');
const addin = JSON.parse(fs.readFileSync(path.join(root, 'addin', 'addin.json'), 'utf8'));
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));

let base = process.env.BASE_URL || addin.defaultBaseUrl;
if (!/^https:\/\/[^\s"<>]+$/.test(base) && !/^https?:\/\/localhost(:\d+)?\//.test(base)) throw new Error('BASE_URL must be an https:// address: ' + base);
if (!base.endsWith('/')) base += '/';
const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const clientId = (process.env.OSIGN_CLIENT_ID || '').trim();
const tenantId = (process.env.OSIGN_TENANT_ID || '').trim();
if (clientId && !GUID.test(clientId)) throw new Error('OSIGN_CLIENT_ID is the app registration\'s Application (client) ID, a GUID.');
if (tenantId && !GUID.test(tenantId)) throw new Error('OSIGN_TENANT_ID is a Directory (tenant) ID, a GUID.');
if (!clientId) console.warn('No OSIGN_CLIENT_ID: the add-in is built, but cannot sign in.');
const site = (process.env.OSIGN_SP_SITE || 'root').trim();
if (!/^root$|^[A-Za-z0-9.-]+\.sharepoint\.com:\/sites\/[^:/]+:$/.test(site)) throw new Error('OSIGN_SP_SITE is "root" or "<host>.sharepoint.com:/sites/<name>:".');
const folder = (process.env.OSIGN_SP_FOLDER || 'OSign').trim().replace(/^\/+|\/+$/g, '');
if (!folder || folder.split('/').some((p) => !p || p === '.' || p === '..')) throw new Error('OSIGN_SP_FOLDER is a folder like "OSign".');

fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(dist, 'assets'), { recursive: true });

const common = {
  bundle: true,
  minify: true,
  sourcemap: false,
  format: 'iife',
  // Classic Outlook's JavaScript-only runtime: syntax newer than this is turned into older syntax.
  target: 'es2016',
  define: { __CLIENT_ID__: JSON.stringify(clientId), __TENANT_ID__: JSON.stringify(tenantId), __SP_SITE__: JSON.stringify(site), __SP_FOLDER__: JSON.stringify(folder) },
  legalComments: 'none',
  logLevel: 'warning'
};
await build({ ...common, entryPoints: [path.join(root, 'addin/src/launchevent.js')], outfile: path.join(dist, 'launchevent.js') });
await build({ ...common, entryPoints: [path.join(root, 'addin/src/taskpane.js')], outfile: path.join(dist, 'taskpane.js') });

for (const f of ['taskpane.html', 'taskpane.css', 'commands.html']) fs.copyFileSync(path.join(root, 'addin', f), path.join(dist, f));
for (const f of fs.readdirSync(path.join(root, 'addin/assets'))) fs.copyFileSync(path.join(root, 'addin/assets', f), path.join(dist, 'assets', f));

const xmlEscape = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const version = (pkg.version.split('.').concat(['0', '0', '0', '0'])).slice(0, 4).map((n) => String(parseInt(n, 10) || 0)).join('.');
const manifest = fs.readFileSync(path.join(root, 'addin/manifest.template.xml'), 'utf8')
  .replace(/\{\{BASE\}\}/g, xmlEscape(base))
  .replace(/\{\{ORIGIN\}\}/g, xmlEscape(new URL(base).origin))
  .replace(/\{\{ID\}\}/g, addin.id)
  .replace(/\{\{VERSION\}\}/g, version)
  .replace(/\{\{PROVIDER\}\}/g, xmlEscape(addin.provider))
  .replace(/\{\{SUPPORT\}\}/g, xmlEscape(addin.supportUrl));
const left = /\{\{[A-Z]+\}\}/.exec(manifest);
if (left) throw new Error('manifest.template.xml: ' + left[0] + ' is not filled in.');
fs.writeFileSync(path.join(dist, 'manifest.xml'), manifest);

// A page for whoever opens the address in a browser.
fs.writeFileSync(path.join(dist, 'index.html'), `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>OSign</title>
<style>body{font:15px/1.5 "Segoe UI",system-ui,Arial,sans-serif;max-width:640px;margin:48px auto;padding:0 16px;color:#1d2330}code{background:#f2f4f7;padding:1px 5px;border-radius:4px}</style></head>
<body><h1><img src="assets/icon-64.png" alt="" width="32" height="32" style="vertical-align:-6px"> OSign</h1>
<p>The Outlook add-in that inserts the company signature. Version ${xmlEscape(pkg.version)}.</p>
<p>Admins: upload <a href="manifest.xml">manifest.xml</a> in the Microsoft 365 admin center, under Settings → Integrated apps → Upload custom apps.</p>
</body></html>`);

console.log('Built dist/ for ' + base + ' (signatures from SharePoint ' + site + ', folder ' + folder + (clientId ? '' : '; no app registration') + ').');
