'use strict';
// Nothing outside the folders meant for it: ../, encoded slashes and dots, backslashes, NUL, double
// slashes, on every path the server serves files from.
const test = require('node:test');
const assert = require('node:assert/strict');
const http = require('node:http');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { createApp } = require('../server/app');

test('files: nothing outside the folders they come from, however the path is written', async () => {
  const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'osign-paths-'));
  fs.writeFileSync(path.join(dataDir, 'secret.txt'), 'TOPSECRET');
  const env = { OSIGN_DATA: dataDir, OSIGN_SECURE_COOKIE: '0', OSIGN_ADMIN_PASSWORD: 'admin-pw', OSIGN_APP_KEY: 'k', OSIGN_PUBLISH_DIR: fs.mkdtempSync(path.join(os.tmpdir(), 'osign-paths-pub-')) };
  const srv = http.createServer(createApp({ log: { info() {}, warn() {}, error() {} }, env }));
  await new Promise((r) => srv.listen(0, '127.0.0.1', r));
  const port = srv.address().port;
  try {
    const r0 = await fetch('http://127.0.0.1:' + port + '/api/login', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ login: 'admin', password: 'admin-pw' }) });
    const cookie = r0.headers.get('set-cookie').split(';')[0];
    const raw = (p) => new Promise((res) => { const req = http.request({ host: '127.0.0.1', port, path: p, headers: { Cookie: cookie } }, (r) => { let b = ''; r.on('data', (c) => (b += c)); r.on('end', () => res([r.statusCode, b])); }); req.on('error', (e) => res([0, e.message])); req.end(); });
    for (const p of ['/assets/../secret.txt', '/assets/..%2fsecret.txt', '/assets/%2e%2e/secret.txt', '/assets/..%5csecret.txt', '/photos/../secret.txt',
      '/photos/..%2f..%2fetc%2fpasswd', '/ui/../server/app.js', '/ui/..%2fserver%2fapp.js', '/ui/%2e%2e%2fpackage.json', '/addin/../secret.txt',
      '/addin/..%2fsecret.txt', '/shared/../server/app.js', '/assets/x%00.png', '/setup/nobs-osign-setup.ps1?id=../../secret.txt', '//etc/passwd', '/assets//etc/passwd']) {
      const [code, body] = await raw(p);
      assert.notEqual(code, 200, p);
      assert.doesNotMatch(body, /TOPSECRET|root:x:|require\(/, p);
    }
  } finally { srv.close(); }
});
