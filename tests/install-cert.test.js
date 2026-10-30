// node --test tests/install-cert.test.js   (needs bash and openssl; Linux and macOS)
//
// scripts/install-cert.sh: the web certificate for OSIGN_HOST, from a PFX or a PEM bundle, checked
// (its key, its names, its end) and put where the shared proxy takes it; --remove takes it back.
// Docker and root stand in: a fake docker and id on the PATH, the proxy's folder in a temp dir.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync, spawnSync } = require('node:child_process');

const ROOT = path.join(__dirname, '..');
let ok = process.platform !== 'win32';
try { execFileSync('openssl', ['version'], { stdio: 'ignore' }); } catch (e) { ok = false; }
const skip = !ok && 'needs bash and openssl';

let base, install, proxy, bin;
const sh = (cmd, cwd) => execFileSync('bash', ['-c', cmd], { cwd: cwd || base, stdio: 'pipe' }).toString();
function run(args, env) {
  return spawnSync('bash', [path.join(install, 'scripts', 'install-cert.sh'), ...args], {
    cwd: install, encoding: 'utf8', timeout: 60000,
    env: Object.assign({}, process.env, { PATH: bin + ':' + process.env.PATH, OSIGN_PROXY_DIR: proxy }, env || {})
  });
}

test.before(() => {
  if (skip) return;
  base = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-cert-'));
  install = path.join(base, 'install'); proxy = path.join(base, 'proxy'); bin = path.join(base, 'bin');
  for (const d of [path.join(install, 'scripts'), path.join(proxy, 'sites'), bin]) fs.mkdirSync(d, { recursive: true });
  for (const f of ['install-cert.sh', 'osign-lib.sh']) fs.copyFileSync(path.join(ROOT, 'scripts', f), path.join(install, 'scripts', f));
  fs.writeFileSync(path.join(install, '.env'), 'OSIGN_HOST=osign.argo.intern\n');
  fs.writeFileSync(path.join(bin, 'docker'), '#!/bin/sh\nexit 0\n', { mode: 0o755 });
  fs.writeFileSync(path.join(bin, 'id'), '#!/bin/sh\necho 0\n', { mode: 0o755 });
  // A CA, and a web server certificate from it for osign.argo.intern (and osign), as AD CS gives it.
  sh(`openssl req -x509 -newkey rsa:2048 -nodes -keyout ca.key -out ca.crt -days 30 -subj "/CN=argo-ARGODC02-CA-1" 2>/dev/null
openssl req -newkey rsa:2048 -nodes -keyout web.key -out web.csr -subj "/CN=osign.argo.intern" 2>/dev/null
printf 'subjectAltName=DNS:osign.argo.intern,DNS:osign\\n' > san.ext
openssl x509 -req -in web.csr -CA ca.crt -CAkey ca.key -CAcreateserial -out web.crt -days 30 -extfile san.ext 2>/dev/null
openssl pkcs12 -export -in web.crt -inkey web.key -certfile ca.crt -out web.pfx -passout pass:secret 2>/dev/null
cat ca.crt web.key web.crt > bundle.pem
openssl req -x509 -newkey rsa:2048 -nodes -keyout other.key -out other.crt -days 30 -subj "/CN=other.example.ch" -addext "subjectAltName=DNS:other.example.ch" 2>/dev/null
cat other.key other.crt > other.pem
cat other.key web.crt > mismatch.pem`);
});

test('a PFX from your CA: checked, the certificate first and the CA after it, the key private; the site file uses it', { skip }, () => {
  const r = run([path.join(base, 'web.pfx')], { OSIGN_PFX_PASSWORD: 'secret' });
  assert.equal(r.status, 0, r.stdout + r.stderr);
  assert.match(r.stdout, /Installed for osign\.argo\.intern: .*CN ?= ?osign\.argo\.intern, issued by .*argo-ARGODC02-CA-1/);
  const crt = fs.readFileSync(path.join(proxy, 'certificates', 'sites', 'nobs-osign.crt'), 'utf8');
  assert.equal((crt.match(/BEGIN CERTIFICATE/g) || []).length, 2, 'the certificate and its CA');
  assert.match(sh(`openssl x509 -noout -subject -in "${path.join(proxy, 'certificates', 'sites', 'nobs-osign.crt')}"`), /osign\.argo\.intern/);
  const key = path.join(proxy, 'certificates', 'sites', 'nobs-osign.key');
  assert.equal(fs.statSync(key).mode & 0o777, 0o600);
  assert.match(fs.readFileSync(path.join(proxy, 'sites', 'nobs-osign.caddy'), 'utf8'), /tls \/certs\/sites\/nobs-osign\.crt \/certs\/sites\/nobs-osign\.key/);
  assert.match(r.stdout, /Delete .* now/);
});

test('a PEM bundle works too; a wrong password, another name or a key of another certificate is refused', { skip }, () => {
  assert.equal(run([path.join(base, 'bundle.pem')]).status, 0);
  let r = run([path.join(base, 'web.pfx')], { OSIGN_PFX_PASSWORD: 'wrong' });
  assert.notEqual(r.status, 0); assert.match(r.stderr, /could not be opened/);
  r = run([path.join(base, 'other.pem')]);
  assert.notEqual(r.status, 0); assert.match(r.stderr, /not for osign\.argo\.intern/);
  r = run([path.join(base, 'mismatch.pem')]);
  assert.notEqual(r.status, 0); assert.match(r.stderr, /belongs to its key/);
  r = run([]);
  assert.notEqual(r.status, 0); assert.match(r.stderr, /Name the certificate file/);
});

test('--remove: back to the proxy\'s own', { skip }, () => {
  const r = run(['--remove']);
  assert.equal(r.status, 0, r.stderr);
  assert.equal(fs.existsSync(path.join(proxy, 'certificates', 'sites', 'nobs-osign.key')), false);
  assert.match(fs.readFileSync(path.join(proxy, 'sites', 'nobs-osign.caddy'), 'utf8'), /import tls_and_headers/);
});
