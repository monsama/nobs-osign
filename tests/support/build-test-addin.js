// Builds the add-in for the tests into a temporary folder: the real code, with the sign-in
// replaced by tests/support/auth-stub.js. Returns the folder (launchevent.js, taskpane.js, pages).
const path = require('node:path');
const fs = require('node:fs');
const os = require('node:os');
const esbuild = require('esbuild');

const root = path.join(__dirname, '..', '..');
const stub = { name: 'auth-stub', setup(b) { b.onResolve({ filter: /[\\/]auth\.js$|^\.\/auth\.js$/ }, () => ({ path: path.join(__dirname, 'auth-stub.js') })); } };

module.exports = async function buildTestAddin(site) {
  const out = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-addin-'));
  const common = {
    bundle: true, format: 'iife', target: 'es2016', logLevel: 'error', plugins: [stub],
    banner: { js: 'var OSIGN_SETUP={clientId:"11111111-2222-3333-4444-555555555555",tenantId:"",site:' + JSON.stringify(site || 'root') + ',folder:"OSign"};' }
  };
  await esbuild.build({ ...common, entryPoints: [path.join(root, 'addin/src/launchevent.js')], outfile: path.join(out, 'launchevent.js') });
  await esbuild.build({ ...common, entryPoints: [path.join(root, 'addin/src/taskpane.js')], outfile: path.join(out, 'taskpane.js') });
  for (const f of ['taskpane.html', 'taskpane.css', 'commands.html']) fs.copyFileSync(path.join(root, 'addin', f), path.join(out, f));
  fs.cpSync(path.join(root, 'addin/assets'), path.join(out, 'assets'), { recursive: true });
  return out;
};
