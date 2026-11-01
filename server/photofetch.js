/*
 * A person's photo from an address on the web (e.g. the team page of the company's website), for
 * the editor to make small and keep in photos/. Only https, only addresses on the internet: the
 * name is looked up once, and an address in the company's own network (or this machine) is
 * refused, so nobody signed in to the editor can make the server fetch from inside. At most a
 * few redirects (each checked the same way), 10 seconds, 5 MB, and only a PNG, JPG, GIF or WebP.
 */
'use strict';
const https = require('node:https');
const dns = require('node:dns');
const net = require('node:net');

const MAX_BYTES = 5 * 1024 * 1024;
const TIMEOUT_MS = 10000;
const MAX_REDIRECTS = 3;

// Addresses that are not on the internet: private, loopback, link-local, shared, multicast, and
// the like (IPv4 and IPv6, IPv4 inside IPv6 too).
function isPrivate(ip) {
  if (net.isIPv6(ip)) {
    const v = ip.toLowerCase();
    const m = /^::ffff:(\d+\.\d+\.\d+\.\d+)$/.exec(v);
    if (m) return isPrivate(m[1]);
    return v === '::' || v === '::1' || /^f[cd]/.test(v) || /^fe[89ab]/.test(v) || /^ff/.test(v) || /^64:ff9b:/.test(v) || /^2001:db8:/.test(v) || /^100::/.test(v);
  }
  const p = ip.split('.').map(Number);
  if (p.length !== 4 || p.some((x) => !(x >= 0 && x <= 255))) return true;
  const [a, b] = p;
  return a === 0 || a === 10 || a === 127 || (a === 100 && b >= 64 && b <= 127) || (a === 169 && b === 254) ||
    (a === 172 && b >= 16 && b <= 31) || (a === 192 && b === 168) || (a === 192 && b === 0) || (a === 198 && (b === 18 || b === 19)) || a >= 224;
}

// dns.lookup as https wants it, refusing an address inside: checked at the moment of connecting,
// so a name that answers otherwise the second time gets nowhere either.
function publicLookup(host, opts, cb) {
  if (typeof opts === 'function') { cb = opts; opts = {}; }
  dns.lookup(host, Object.assign({}, opts, { all: true }), (err, list) => {
    if (err) return cb(err);
    const bad = list.find((x) => isPrivate(x.address));
    if (bad || !list.length) return cb(Object.assign(new Error('This address is in your own network, not on the internet.'), { code: 'OSIGN_PRIVATE' }));
    if (opts.all) return cb(null, list);
    return cb(null, list[0].address, list[0].family);
  });
}

const KINDS = [['89504e47', 'image/png'], ['ffd8ff', 'image/jpeg'], ['47494638', 'image/gif']];
function kindOf(bytes) {
  const hex = bytes.subarray(0, 12).toString('hex');
  for (const [m, t] of KINDS) if (hex.startsWith(m)) return t;
  if (hex.startsWith('52494646') && bytes.subarray(8, 12).toString('latin1') === 'WEBP') return 'image/webp';
  return '';
}

function get(url, left, opts) {
  return new Promise((resolve, reject) => {
    let u;
    try { u = new URL(url); } catch (e) { reject(new Error('That is not a web address.')); return; }
    if (u.protocol !== 'https:') { reject(new Error('Only https:// addresses.')); return; }
    if (u.username || u.password) { reject(new Error('An address with a password in it is not taken.')); return; }
    if (net.isIP(u.hostname.replace(/^\[|\]$/g, '')) && isPrivate(u.hostname.replace(/^\[|\]$/g, ''))) { reject(new Error('This address is in your own network, not on the internet.')); return; }
    const req = https.get(u, { lookup: (opts && opts.lookup) || publicLookup, timeout: TIMEOUT_MS, headers: { 'User-Agent': 'NOBS-OSign', Accept: 'image/*' }, ca: opts && opts.ca }, (res) => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        res.resume();
        if (!left) { reject(new Error('Too many redirects.')); return; }
        resolve(get(new URL(res.headers.location, u).href, left - 1, opts));
        return;
      }
      if (res.statusCode !== 200) { res.resume(); reject(new Error('The address answered ' + res.statusCode + '.')); return; }
      if (Number(res.headers['content-length'] || 0) > MAX_BYTES) { res.destroy(); reject(new Error('The image is larger than 5 MB.')); return; }
      const parts = []; let size = 0;
      res.on('data', (c) => { size += c.length; if (size > MAX_BYTES) { res.destroy(); reject(new Error('The image is larger than 5 MB.')); } else parts.push(c); });
      res.on('end', () => {
        const bytes = Buffer.concat(parts);
        const type = kindOf(bytes);
        if (!type) { reject(new Error('That address is not a PNG, JPG, GIF or WebP image.')); return; }
        resolve({ type, bytes });
      });
      res.on('error', reject);
    });
    req.on('timeout', () => req.destroy(new Error('The address took too long to answer.')));
    req.on('error', (e) => reject(e.code === 'OSIGN_PRIVATE' ? new Error(e.message) : e.code === 'ENOTFOUND' ? new Error('No such address: ' + u.hostname + '.') : e));
  });
}

// opts: only for the tests (a lookup and a CA of their own).
function fetchPhoto(url, opts) { return get(String(url || '').trim(), MAX_REDIRECTS, opts); }

module.exports = { fetchPhoto, isPrivate, kindOf, publicLookup };
