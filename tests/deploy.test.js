// node --test tests/deploy.test.js
//
// The server's install (docker-compose.yml, Dockerfile, deploy-osign.sh, .env.example) against the
// server's code: every setting Compose hands over is one the server reads, every secret is made by
// the deploy script, and the image is the version of this release.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const ROOT = path.join(__dirname, '..');
const read = (f) => fs.readFileSync(path.join(ROOT, f), 'utf8');
const compose = read('docker-compose.yml');
const serverCode = fs.readdirSync(path.join(ROOT, 'server')).map((f) => read('server/' + f)).join('\n');
const envVars = [...new Set([...compose.matchAll(/^\s+(OSIGN_[A-Z_]+):/gm)].map((m) => m[1]))];

test('every setting Compose hands over is read by the server', () => {
  assert.ok(envVars.length > 15, envVars.join(', '));
  for (const v of envVars) assert.ok(serverCode.includes("'" + v + "'") || serverCode.includes('.' + v), v + ' is not read by server/');
});

test('every setting in .env.example reaches the container (or is the install\'s own)', () => {
  const own = ['OSIGN_HOST', 'OSIGN_IMAGE', 'TZ', 'BACKUP_PASSPHRASE_FILE'];
  const listed = [...read('.env.example').matchAll(/^#?\s*([A-Z_]+)=/gm)].map((m) => m[1]);
  for (const v of listed) assert.ok(own.includes(v) || compose.includes('${' + v + ':-'), v + ' is not handed to the container');
});

test('every secret is made by deploy-osign.sh, and read from its file', () => {
  const secrets = [...compose.matchAll(/^\s{2}([a-z_]+):\n\s+file: \.\/secrets\/\1$/gm)].map((m) => m[1]);
  assert.deepEqual(secrets.sort(), ['admin_password', 'cf_api_token', 'graph_client_secret', 'ldap_service_password']);
  const deploy = read('deploy-osign.sh');
  for (const s of secrets) {
    assert.ok(deploy.includes(s), s + ' is not made by deploy-osign.sh');
    assert.match(compose, new RegExp('_FILE: /run/secrets/' + s + '$', 'm'));
  }
});

test('the image is this release\'s, and the Dockerfile copies what the server needs', () => {
  const v = JSON.parse(read('package.json')).version;
  assert.match(compose, new RegExp('ghcr\\.io/monsama/osign:' + v.replace(/\./g, '\\.') + '\\}'));
  const docker = read('Dockerfile');
  for (const p of ['server', 'shared', 'editor/ui', 'addin/fill.mjs', 'dist/template']) assert.ok(docker.includes(p), 'Dockerfile: ' + p);
  // What the server requires at run time is in the image, and is a dependency (not a dev one).
  const deps = Object.keys(JSON.parse(read('package.json')).dependencies);
  const required = [...serverCode.matchAll(/require\('([^.'][^']*)'\)/g)].map((m) => m[1].split('/').slice(0, m[1].startsWith('@') ? 2 : 1).join('/')).filter((n) => !n.startsWith('node:'));
  for (const r of new Set(required)) assert.ok(deps.includes(r), r + ' is required by server/ but not a dependency');
});

test('the release archive names files that are there', () => {
  const wf = read('.github/workflows/release.yml');
  const list = /git archive[\s\S]*?HEAD \\\n([\s\S]*?)\n\s+\(cd out/.exec(wf)[1].replace(/\\\n/g, ' ').trim().split(/\s+/);
  for (const f of list) assert.ok(fs.existsSync(path.join(ROOT, f)), f);
});
