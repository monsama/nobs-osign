// Builds the add-in. Two results, both without any company data (the signatures are read at run
// time from the company's own SharePoint):
//
//   dist/template/  always: the add-in with blanks for its settings. The editor fills them in
//                   (Settings → Outlook add-in) and writes the finished files to a folder, for any
//                   HTTPS web server. The release zip carries this as OSign\addin.
//   dist/site/      only with BASE_URL: the same, filled in here, ready to serve at BASE_URL.
//
//   BASE_URL=https://intranet.example.com/osign/ OSIGN_CLIENT_ID=<app id> node addin/build.mjs
//
// BASE_URL         where dist/site/ will be served (https://…/)
// OSIGN_CLIENT_ID  the Entra ID app registration the add-in signs in with (not a secret)
// OSIGN_TENANT_ID  only for an app registration of a single organization; empty: the user's own
// OSIGN_SP_SITE    the SharePoint site whose document library holds the folder: "root" (default,
//                  the tenant's home site) or "<host>.sharepoint.com:/sites/<name>:"
// OSIGN_SP_FOLDER  the folder in that library (default: OSign)
//
// The blanks: in launchevent.js and taskpane.js the strings @@OSIGN_CLIENT_ID@@,
// @@OSIGN_TENANT_ID@@, @@OSIGN_SP_SITE@@ and @@OSIGN_SP_FOLDER@@ (in the first line, var OSIGN_SETUP);
// in manifest.xml {{BASE}}, {{ORIGIN}} and {{SUPPORT}}. editor/OSign.ps1 (Build-AddinPackage)
// fills them the same way as fillTemplate in addin/fill.mjs.
import { build } from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { fillTemplate } from './fill.mjs';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const dist = path.join(root, 'dist');
const template = path.join(dist, 'template');
const addin = JSON.parse(fs.readFileSync(path.join(root, 'addin', 'addin.json'), 'utf8'));
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));

const xmlEscape = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(template, 'assets'), { recursive: true });

const common = {
  bundle: true,
  minify: true,
  sourcemap: false,
  format: 'iife',
  // Classic Outlook's JavaScript-only runtime: syntax newer than this is turned into older syntax.
  target: 'es2016',
  // The settings, outside the bundle so the minifier leaves them as they are; filled in later.
  banner: { js: 'var OSIGN_SETUP={clientId:"@@OSIGN_CLIENT_ID@@",tenantId:"@@OSIGN_TENANT_ID@@",site:"@@OSIGN_SP_SITE@@",folder:"@@OSIGN_SP_FOLDER@@"};' },
  legalComments: 'none',
  logLevel: 'warning'
};
await build({ ...common, entryPoints: [path.join(root, 'addin/src/launchevent.js')], outfile: path.join(template, 'launchevent.js') });
await build({ ...common, entryPoints: [path.join(root, 'addin/src/taskpane.js')], outfile: path.join(template, 'taskpane.js') });

for (const f of ['taskpane.html', 'taskpane.css', 'commands.html']) fs.copyFileSync(path.join(root, 'addin', f), path.join(template, f));
for (const f of fs.readdirSync(path.join(root, 'addin/assets'))) fs.copyFileSync(path.join(root, 'addin/assets', f), path.join(template, 'assets', f));

const version = (pkg.version.split('.').concat(['0', '0', '0', '0'])).slice(0, 4).map((n) => String(parseInt(n, 10) || 0)).join('.');
const manifest = fs.readFileSync(path.join(root, 'addin/manifest.template.xml'), 'utf8')
  .replace(/\{\{ID\}\}/g, addin.id)
  .replace(/\{\{VERSION\}\}/g, version)
  .replace(/\{\{PROVIDER\}\}/g, xmlEscape(addin.provider));
const left = /\{\{(?!BASE\}\}|ORIGIN\}\}|SUPPORT\}\})[A-Z]+\}\}/.exec(manifest);
if (left) throw new Error('manifest.template.xml: ' + left[0] + ' is not filled in.');
fs.writeFileSync(path.join(template, 'manifest.xml'), manifest);

// A page for whoever opens the address in a browser.
fs.writeFileSync(path.join(template, 'index.html'), `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>OSign</title>
<style>body{font:15px/1.5 "Segoe UI",system-ui,Arial,sans-serif;max-width:640px;margin:48px auto;padding:0 16px;color:#1d2330}code{background:#f2f4f7;padding:1px 5px;border-radius:4px}</style></head>
<body><h1><img src="assets/icon-64.png" alt="" width="32" height="32" style="vertical-align:-6px"> OSign</h1>
<p>The Outlook add-in that inserts the company signature. Version ${xmlEscape(pkg.version)}.</p>
<p>Admins: upload <a href="manifest.xml">manifest.xml</a> in the Microsoft 365 admin center, under Settings → Integrated apps → Upload custom apps.</p>
</body></html>`);
fs.writeFileSync(path.join(template, 'osign-addin.json'), JSON.stringify({ version: pkg.version, id: addin.id }) + '\n');
console.log('Built dist/template/ (version ' + pkg.version + ').');

if (process.env.BASE_URL) {
  const s = fillTemplate(template, path.join(dist, 'site'), {
    base: process.env.BASE_URL, clientId: process.env.OSIGN_CLIENT_ID, tenantId: process.env.OSIGN_TENANT_ID,
    site: process.env.OSIGN_SP_SITE, folder: process.env.OSIGN_SP_FOLDER
  });
  console.log('Built dist/site/ for ' + s.base + ' (signatures from SharePoint ' + s.site + ', folder ' + s.folder + (s.clientId ? '' : '; no app registration') + ').');
}
