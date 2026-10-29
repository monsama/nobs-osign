/*
 * A backup of the server's data, for the local account admin: one file, locked with a passphrase
 * (scrypt, AES-256-GCM), to keep somewhere else and to bring back on this server or a new one.
 *
 * It holds the data volume: the draft and its images, the settings, the server's connections,
 * the certificate, the history, the users as last read. Not the audit log (each server keeps its
 * own), nor what is made again anyway (the add-in's files, the local publish folder). The
 * passwords and the certificate's key are sealed with this server's key (secrets/app_key); that
 * key goes into the backup too, so a new server, with a key of its own, seals them again with it.
 */
'use strict';
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const zlib = require('node:zlib');
const { writeAtomic } = require('./store');

const FORMAT = 'nobs-osign-backup';
const SKIP = new Set(['audit.log', 'audit.1.log', 'addin', 'publish', 'lost+found']);
const MIN_PASS = 12;
// scrypt with 2^15 rounds: a second or so to try one passphrase, on purpose.
const KDF = { N: 32768, r: 8, p: 1, maxmem: 64 * 1024 * 1024 };
const key = (pass, salt) => crypto.scryptSync(String(pass), salt, 32, KDF);

function files(dir, rel = '') {
  const out = {};
  for (const e of fs.readdirSync(path.join(dir, rel), { withFileTypes: true })) {
    const r = rel ? rel + '/' + e.name : e.name;
    if (!rel && SKIP.has(e.name)) continue;
    if (e.name.endsWith('.tmp')) continue;
    if (e.isDirectory()) Object.assign(out, files(dir, r));
    else if (e.isFile()) out[r] = fs.readFileSync(path.join(dir, r)).toString('base64');
  }
  return out;
}

// → { name, data } (data: the file, base64) or { error }.
function exportBackup({ dataDir, passphrase, appKey, version, now = new Date() }) {
  if (String(passphrase || '').length < MIN_PASS) return { error: 'The passphrase has at least ' + MIN_PASS + ' characters.' };
  const plain = zlib.gzipSync(Buffer.from(JSON.stringify({ version, at: now.toISOString(), appKey: appKey || '', files: files(dataDir) })));
  const salt = crypto.randomBytes(16), iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv('aes-256-gcm', key(passphrase, salt), iv);
  c.setAAD(Buffer.from(FORMAT));
  const data = Buffer.concat([c.update(plain), c.final()]);
  const file = { format: FORMAT, v: 1, kdf: 'scrypt', N: KDF.N, r: KDF.r, p: KDF.p, salt: salt.toString('base64'), iv: iv.toString('base64'), tag: c.getAuthTag().toString('base64'), data: data.toString('base64') };
  return { name: 'nobs-osign-backup-' + now.toISOString().slice(0, 10) + '.osignbak', data: Buffer.from(JSON.stringify(file)).toString('base64') };
}

// The backup opened: { at, version, appKey, files } or { error }.
function openBackup(base64, passphrase) {
  let f;
  try { f = JSON.parse(Buffer.from(String(base64 || ''), 'base64').toString('utf8')); } catch (e) { f = null; }
  if (!f || f.format !== FORMAT || f.v !== 1 || f.kdf !== 'scrypt') return { error: 'This is not a backup of NOBS OSign.' };
  if (f.N !== KDF.N || f.r !== KDF.r || f.p !== KDF.p) return { error: 'This backup was made with other settings than this version knows.' };
  let plain;
  try {
    const d = crypto.createDecipheriv('aes-256-gcm', key(passphrase, Buffer.from(f.salt, 'base64')), Buffer.from(f.iv, 'base64'));
    d.setAAD(Buffer.from(FORMAT)); d.setAuthTag(Buffer.from(f.tag, 'base64'));
    plain = Buffer.concat([d.update(Buffer.from(f.data, 'base64')), d.final()]);
  } catch (e) { return { error: 'The passphrase is not the one of this backup (or the file is damaged).' }; }
  let b;
  try { b = JSON.parse(zlib.gunzipSync(plain).toString('utf8')); } catch (e) { return { error: 'This backup is damaged.' }; }
  if (!b || typeof b.files !== 'object') return { error: 'This backup is damaged.' };
  return b;
}

// What the backup holds, into the data volume, in place of what is there (the audit log stays).
function restoreFiles(dataDir, filesIn) {
  const base = path.resolve(dataDir) + path.sep;
  const paths = Object.keys(filesIn);
  for (const rel of paths) {
    const full = path.resolve(dataDir, rel);
    if (!full.startsWith(base) || rel.split('/').some((p) => !p || p === '.' || p === '..') || SKIP.has(rel.split('/')[0])) return { error: 'This backup names a file outside the data: ' + rel };
  }
  for (const e of fs.readdirSync(dataDir)) if (!SKIP.has(e)) fs.rmSync(path.join(dataDir, e), { recursive: true, force: true });
  for (const rel of paths) writeAtomic(path.resolve(dataDir, rel), Buffer.from(filesIn[rel], 'base64'));
  return { files: paths.length };
}

module.exports = { exportBackup, openBackup, restoreFiles, MIN_PASS };
