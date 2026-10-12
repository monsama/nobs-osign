// Builds the add-in into dist/: the two bundles, its pages, the icons, the manifest, and the
// published configuration (config/) beside them - the whole site GitHub Pages serves.
//
//   BASE_URL=https://monsama.github.io/osign/ node addin/build.mjs
//
// BASE_URL is where dist/ will be served (default: GitHub Pages of this repository).
// CONFIG_PATH is the folder the editor publishes to (default: config), and where it is served.
// CONFIG_SRC is where that folder is on disk (default: <repository>/<CONFIG_PATH>); the Pages
// workflow checks the editor's branch out into data/ and points here.
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
const configPath = (process.env.CONFIG_PATH || 'config').replace(/^\/+|\/+$/g, '');
if (!/^[A-Za-z0-9._-]+(\/[A-Za-z0-9._-]+)*$/.test(configPath)) throw new Error('CONFIG_PATH is a folder like "config".');

fs.rmSync(dist, { recursive: true, force: true });
fs.mkdirSync(path.join(dist, 'assets'), { recursive: true });

const common = {
  bundle: true,
  minify: true,
  sourcemap: false,
  format: 'iife',
  // Classic Outlook's JavaScript-only runtime: syntax newer than this is turned into older syntax.
  target: 'es2016',
  define: { __BASE_URL__: JSON.stringify(base), __CONFIG_PATH__: JSON.stringify(configPath) },
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

// The published configuration, as the editor wrote it into the repository.
const cfgSrc = process.env.CONFIG_SRC ? path.resolve(process.env.CONFIG_SRC) : path.join(root, configPath);
if (fs.existsSync(cfgSrc)) fs.cpSync(cfgSrc, path.join(dist, configPath), { recursive: true });
else console.warn('No published signatures at ' + cfgSrc + ' yet: publish from the editor to create them.');

// A page for whoever opens the address in a browser.
fs.writeFileSync(path.join(dist, 'index.html'), `<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>OSign</title>
<style>body{font:15px/1.5 "Segoe UI",system-ui,Arial,sans-serif;max-width:640px;margin:48px auto;padding:0 16px;color:#1d2330}code{background:#f2f4f7;padding:1px 5px;border-radius:4px}</style></head>
<body><h1><img src="assets/icon-64.png" alt="" width="32" height="32" style="vertical-align:-6px"> OSign</h1>
<p>The Outlook add-in that inserts the company signature. Version ${xmlEscape(pkg.version)}.</p>
<p>Admins: upload <a href="manifest.xml">manifest.xml</a> in the Microsoft 365 admin center, under Settings → Integrated apps → Upload custom apps.</p>
</body></html>`);

console.log('Built dist/ for ' + base + ' (configuration in ' + configPath + '/).');
