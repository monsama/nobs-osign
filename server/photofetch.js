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
// An IPv6 address as its eight 16-bit groups (an IPv4 at its end counted as two).
function groups6(ip) {
  let v = String(ip).toLowerCase().replace(/^\[|\]$/g, '').replace(/%.*$/, '');
  const v4 = /(\d+\.\d+\.\d+\.\d+)$/.exec(v);
  if (v4) {
    const p = v4[1].split('.').map(Number);
    v = v.slice(0, -v4[1].length) + ((p[0] << 8) | p[1]).toString(16) + ':' + ((p[2] << 8) | p[3]).toString(16);
  }
  const [head, tail] = v.split('::');
  const h = head ? head.split(':') : [], t = tail !== undefined && tail ? tail.split(':') : [];
  const fill = v.includes('::') ? 8 - h.length - t.length : 0;
  const all = [...h, ...Array(Math.max(0, fill)).fill('0'), ...t].map((x) => parseInt(x || '0', 16));
  return all.length === 8 && all.every((x) => x >= 0 && x <= 0xffff) ? all : null;
}
function isPrivate(ip) {
  if (net.isIPv6(ip)) {
    const g = groups6(ip);
    if (!g) return true;
    const v4 = (a, b) => [a >> 8, a & 255, b >> 8, b & 255].join('.');
    // An IPv4 inside: mapped (::ffff:a.b.c.d), compatible (::a.b.c.d), translated (::ffff:0:a.b.c.d,
    // 64:ff9b::a.b.c.d) or 6to4 (2002:aabb:ccdd::): judged as that IPv4.
    if (g.slice(0, 5).every((x) => x === 0) && (g[5] === 0xffff || g[5] === 0)) return g[5] === 0 && g[6] === 0 && g[7] <= 1 ? true : isPrivate(v4(g[6], g[7]));
    if (g.slice(0, 4).every((x) => x === 0) && g[4] === 0xffff && g[5] === 0) return isPrivate(v4(g[6], g[7]));
    if (g[0] === 0x64 && g[1] === 0xff9b) return isPrivate(v4(g[6], g[7]));
    if (g[0] === 0x2002) return isPrivate(v4(g[1], g[2]));
    const a = g[0];
    return (a & 0xfe00) === 0xfc00 || (a & 0xffc0) === 0xfe80 || (a & 0xffc0) === 0xfec0 || (a & 0xff00) === 0xff00 ||
      (a === 0x2001 && g[1] === 0x0db8) || (a === 0x100 && g[1] === 0 && g[2] === 0 && g[3] === 0) || (a === 0x2001 && g[1] < 0x200);
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
    const deadline = setTimeout(() => req.destroy(new Error('The address took too long to answer.')), TIMEOUT_MS);
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
    req.on('close', () => clearTimeout(deadline));
    // What went wrong in words that say nothing about the network behind (no port scan by errors).
    req.on('error', (e) => { clearTimeout(deadline); reject(e.code === 'OSIGN_PRIVATE' ? new Error(e.message) : e.code === 'ENOTFOUND' ? new Error('No such address: ' + u.hostname + '.') : /took too long|larger than/.test(e.message) ? e : new Error('The address could not be reached.')); });
  });
}

// opts: only for the tests (a lookup and a CA of their own). One fetch at a time for the whole
// server: an editor cannot keep many open.
let busy = false;
async function fetchPhoto(url, opts) {
  if (busy) throw new Error('Another photo is being fetched: try again in a moment.');
  busy = true;
  try { return await get(String(url || '').trim(), MAX_REDIRECTS, opts); } finally { busy = false; }
}

module.exports = { fetchPhoto, isPrivate, kindOf, publicLookup };
