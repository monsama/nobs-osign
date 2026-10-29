/*
 * The server's certificate for signing in to Microsoft Entra ID, instead of a client secret: the
 * server makes the key itself and never lets it out; the setup script only puts the certificate
 * (the public half) on the server's app registration. With it, the server signs a short token
 * (a client assertion, RS256) each time it asks for an access token.
 *
 * In the data volume: graph-cert.json, the key sealed with the server's key (like the
 * connections' passwords). One is "pending" from the setup script's download until the script
 * says it put it on the app; then it is the one in use.
 */
'use strict';
const crypto = require('node:crypto');
const path = require('node:path');
const { writeAtomic, readText } = require('./store');

const YEARS = 2;

// DER, as much of it as a certificate needs.
function der(tag, body) {
  const n = body.length;
  const len = n < 0x80 ? Buffer.from([n]) : n < 0x100 ? Buffer.from([0x81, n]) : n < 0x10000 ? Buffer.from([0x82, n >> 8, n & 0xff]) : Buffer.from([0x83, n >> 16, (n >> 8) & 0xff, n & 0xff]);
  return Buffer.concat([Buffer.from([tag]), len, body]);
}
const seq = (...parts) => der(0x30, Buffer.concat(parts));
const set = (...parts) => der(0x31, Buffer.concat(parts));
const int = (buf) => der(0x02, buf[0] & 0x80 ? Buffer.concat([Buffer.from([0]), buf]) : buf);
const oid = (hex) => der(0x06, Buffer.from(hex, 'hex'));
const SHA256_RSA = seq(oid('2a864886f70d01010b'), der(0x05, Buffer.alloc(0)));
const CN = '550403';
function time(d) {
  // UTCTime up to 2049, GeneralizedTime after (RFC 5280).
  const iso = d.toISOString().replace(/[-:T]/g, '').slice(0, 14) + 'Z';
  return d.getUTCFullYear() < 2050 ? der(0x17, Buffer.from(iso.slice(2))) : der(0x18, Buffer.from(iso));
}

// A new key and a self-signed certificate for it, valid for YEARS. → { keyPem, certDer, ... }
function makeCert(name, now = new Date()) {
  const { privateKey, publicKey } = crypto.generateKeyPairSync('rsa', { modulusLength: 2048 });
  const until = new Date(now.getTime()); until.setUTCFullYear(until.getUTCFullYear() + YEARS);
  const from = new Date(now.getTime() - 5 * 60 * 1000);
  const dn = seq(set(seq(oid(CN), der(0x0c, Buffer.from(name, 'utf8')))));
  const serial = crypto.randomBytes(16); serial[0] &= 0x7f;
  const tbs = seq(der(0xa0, int(Buffer.from([2]))), int(serial), SHA256_RSA, dn, seq(time(from), time(until)), dn, publicKey.export({ type: 'spki', format: 'der' }));
  const sig = crypto.sign('sha256', tbs, privateKey);
  const certDer = seq(tbs, SHA256_RSA, der(0x03, Buffer.concat([Buffer.from([0]), sig])));
  return describe({ keyPem: privateKey.export({ type: 'pkcs8', format: 'pem' }), certDer: certDer.toString('base64') });
}
// What Entra and the page need to know of one: its thumbprints and how long it holds.
function describe(c) {
  const x = new crypto.X509Certificate(Buffer.from(c.certDer, 'base64'));
  const raw = Buffer.from(c.certDer, 'base64');
  return Object.assign({}, c, {
    x5t: crypto.createHash('sha1').update(raw).digest('base64url'),
    thumbprint: crypto.createHash('sha1').update(raw).digest('hex').toUpperCase(),
    expires: new Date(x.validTo).toISOString()
  });
}

// The client assertion for the token endpoint (RFC 7523), signed with the key.
function clientAssertion({ keyPem, x5t, clientId, audience, now = Date.now() }) {
  const b64 = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
  const t = Math.floor(now / 1000);
  const head = b64({ alg: 'RS256', typ: 'JWT', x5t });
  const body = b64({ aud: audience, iss: clientId, sub: clientId, jti: crypto.randomUUID(), nbf: t - 60, iat: t, exp: t + 600 });
  const sig = crypto.sign('sha256', Buffer.from(head + '.' + body), keyPem).toString('base64url');
  return head + '.' + body + '.' + sig;
}

// The certificates kept in the data volume; seal and open as the connections have them.
function createCerts({ dataDir, seal, open }) {
  const file = path.join(dataDir, 'graph-cert.json');
  const load = () => { try { return JSON.parse(readText(file) || '{}') || {}; } catch (e) { return {}; } };
  const unseal = (c) => {
    if (!c || !c.key) return null;
    const keyPem = open('graphCertKey', c.key);
    return keyPem ? describe({ keyPem, certDer: c.certDer }) : null;
  };
  const keep = (c) => ({ key: seal('graphCertKey', c.keyPem), certDer: c.certDer, at: new Date().toISOString() });
  return {
    // A new one for the setup script, kept until the script says it is on the app.
    pending() { const c = makeCert('NOBS OSign server'); const all = load(); all.pending = keep(c); writeAtomic(file, JSON.stringify(all)); return c; },
    // The script put it on the app (its thumbprint, to be sure it is this one): in use from now.
    activate(thumbprint) {
      const all = load(); const p = unseal(all.pending);
      if (!p || p.thumbprint !== String(thumbprint || '').toUpperCase()) return false;
      writeAtomic(file, JSON.stringify({ active: all.pending }));
      return true;
    },
    active() { return unseal(load().active); },
    remove() { writeAtomic(file, JSON.stringify({})); }
  };
}

module.exports = { makeCert, describe, clientAssertion, createCerts };
