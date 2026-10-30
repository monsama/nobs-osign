/*
 * The web certificate people see: the one the shared proxy (nobs-proxy) serves for the editor's
 * address. The server asks the proxy for it itself, as a browser would (TLS with that name), and
 * says what it is: for which names, from whom, until when, and whether PCs trust it (the public
 * CAs, and your own CA as set for the domain controller).
 *
 *   OSIGN_PROXY_ADDR  where the proxy is reached from the container (nobs-proxy:443)
 */
'use strict';
const fs = require('node:fs');
const tls = require('node:tls');
const crypto = require('node:crypto');

const HOST = /^[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?)*$/;
const cn = (s) => (/CN=([^\n,]+)/.exec(String(s || '')) || [])[1] || String(s || '').split('\n')[0];

function inspect({ addr, host, caFile, timeout = 8000 }) {
  if (!HOST.test(String(host || ''))) return Promise.resolve({ error: 'That is not a host name.' });
  const [h, p] = String(addr || 'nobs-proxy:443').split(':');
  const extra = caFile ? (() => { try { return fs.readFileSync(caFile, 'utf8').trim(); } catch (e) { return ''; } })() : '';
  return new Promise((resolve) => {
    const socket = tls.connect({ host: h, port: Number(p) || 443, servername: host, rejectUnauthorized: false, ca: extra ? [...tls.rootCertificates, extra] : undefined, timeout }, () => {
      const c = socket.getPeerCertificate(true);
      const trusted = socket.authorized;
      const why = socket.authorizationError ? String(socket.authorizationError) : '';
      socket.end();
      if (!c || !c.raw) return resolve({ error: 'The proxy sent no certificate.' });
      const x = new crypto.X509Certificate(c.raw);
      const names = String(x.subjectAltName || '').split(',').map((n) => n.trim()).filter((n) => /^DNS:/.test(n)).map((n) => n.slice(4));
      const until = new Date(x.validTo);
      const nameOk = names.length ? names.some((n) => n.toLowerCase() === host.toLowerCase() || (n.startsWith('*.') && host.toLowerCase().endsWith(n.slice(1).toLowerCase()))) : cn(x.subject).toLowerCase() === host.toLowerCase();
      resolve({
        host, subject: cn(x.subject), issuer: cn(x.issuer), names, validFrom: new Date(x.validFrom).toISOString(), validTo: until.toISOString(),
        daysLeft: Math.floor((until.getTime() - Date.now()) / 86400000),
        selfSigned: x.issuer === x.subject, nameOk, trusted: trusted && nameOk, why,
        fingerprint: x.fingerprint256
      });
    });
    socket.on('timeout', () => { socket.destroy(); resolve({ error: 'The proxy did not answer (' + h + ':' + (p || 443) + ').' }); });
    socket.on('error', (e) => resolve({ error: 'The proxy could not be asked (' + h + ':' + (p || 443) + '): ' + e.message }));
  });
}

module.exports = { inspect };
