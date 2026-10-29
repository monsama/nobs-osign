/*
 * The draft, its images, the settings and what was last published: the same files and rules as
 * editor/OSign.ps1 keeps in %APPDATA%\OSign, here in the data volume (OSIGN_DATA, /data).
 *
 * Publishing goes into a folder (OSIGN_PUBLISH_DIR) for now, laid out as the add-in reads it in
 * SharePoint: signatures.json, assets/<image>, photos/<user>.jpg, users/<mail>.json.
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');

const MAX_ASSET_BYTES = 1024 * 1024;
// Windows PowerShell's limit decided this for OSign.ps1; the same here, so both answer alike.
const MAX_CONFIG_CHARS = 1500 * 1024;
const ASSET_TYPES = { '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.gif': 'image/gif' };
// What the first bytes of each kind of image are.
const MAGIC = { '.png': '89504e47', '.jpg': 'ffd8ff', '.jpeg': 'ffd8ff', '.gif': '47494638' };

const sha256 = (buf) => crypto.createHash('sha256').update(buf).digest('hex');

// An asset's file name: what the upload was called, cut down to what is safe in a path and a URL.
// null for anything that is not one of the image types Outlook shows. (Get-SafeAssetName)
function safeAssetName(name) {
  const base = String(name || '').split(/[\\/]/).pop();
  const dot = base.lastIndexOf('.');
  const ext = dot > 0 ? base.slice(dot).toLowerCase() : '';
  if (!ASSET_TYPES[ext]) return null;
  const stem = base.slice(0, dot).replace(/[^A-Za-z0-9_-]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'image';
  return stem + ext;
}

// Written to a temporary file and renamed over the old one, so a crash never leaves half a file.
function writeAtomic(file, data) {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  const tmp = file + '.tmp';
  fs.writeFileSync(tmp, data);
  fs.renameSync(tmp, file);
}
function readText(file) { try { return fs.readFileSync(file, 'utf8'); } catch (e) { return null; } }

function checkConfigJson(text) {
  if (!text) return 'No configuration was sent.';
  if (text.length > MAX_CONFIG_CHARS) return 'The configuration is larger than 1.5 MB: images belong in assets, not in the HTML.';
  let c;
  try { c = JSON.parse(text); } catch (e) { return 'The configuration is not valid JSON.'; }
  if (!c || typeof c !== 'object' || Array.isArray(c)) return 'The configuration is not a JSON object.';
  if (c.schemaVersion !== 1) return 'The configuration has an unknown schemaVersion.';
  return null;
}

function createStore(dataDir) {
  const draftFile = path.join(dataDir, 'draft.json');
  const assetDir = path.join(dataDir, 'assets');
  const stateFile = path.join(dataDir, 'state.json');
  fs.mkdirSync(assetDir, { recursive: true });

  function assets() {
    let names = [];
    try { names = fs.readdirSync(assetDir); } catch (e) { /* none */ }
    return names.filter((n) => safeAssetName(n) === n && fs.statSync(path.join(assetDir, n)).isFile()).sort()
      .map((n) => ({ name: n, size: fs.statSync(path.join(assetDir, n)).size }));
  }
  function state() {
    try { const s = JSON.parse(readText(stateFile)); if (s && typeof s === 'object') return s; } catch (e) { /* below */ }
    return { publishedHash: '', publishedAt: 0 };
  }
  function saveState(hash) { writeAtomic(stateFile, JSON.stringify({ publishedHash: hash, publishedAt: Date.now() })); }
  // One hash over the configuration and every image: does the draft differ from what was last
  // published or loaded? The same parts as Get-DraftHash.
  function draftHash() {
    const parts = [sha256(Buffer.from(readText(draftFile) || '', 'utf8'))];
    for (const a of assets()) parts.push(a.name + ':' + sha256(fs.readFileSync(path.join(assetDir, a.name))));
    return sha256(Buffer.from(parts.join('\n'), 'utf8'));
  }
  const dirty = () => draftHash() !== String(state().publishedHash || '');

  function getDraft() {
    const configJson = readText(draftFile);
    const st = state();
    return { configJson, assets: assets(), dirty: !!configJson && dirty(), publishedAt: typeof st.publishedAt === 'number' ? st.publishedAt : 0 };
  }
  function saveDraft(text) {
    const err = checkConfigJson(text); if (err) return { error: err };
    writeAtomic(draftFile, text);
    return { dirty: dirty() };
  }

  function uploadAsset(name, data) {
    let n = safeAssetName(name);
    if (!n) return { error: 'Only PNG, JPG and GIF images can be used: Outlook does not show SVG or WebP in a signature.' };
    const bytes = Buffer.from(String(data || ''), 'base64');
    // Buffer.from ignores what is not base64: the round trip tells whether it all was.
    if (bytes.toString('base64').replace(/=+$/, '') !== String(data || '').replace(/\s+/g, '').replace(/=+$/, '')) return { error: 'The image did not arrive whole.' };
    if (!bytes.length) return { error: 'The image is empty.' };
    if (bytes.length > MAX_ASSET_BYTES) return { error: 'The image is larger than 1 MB. A signature image is best well under 100 KB.' };
    const ext = path.extname(n);
    if (!bytes.subarray(0, 8).toString('hex').startsWith(MAGIC[ext])) return { error: 'The file is not the ' + ext.slice(1).toUpperCase() + ' image its name says.' };
    // Same name, other content: a new name, so a signature already using the old image keeps it.
    let file = path.join(assetDir, n);
    if (fs.existsSync(file) && sha256(fs.readFileSync(file)) !== sha256(bytes)) {
      const stem = n.slice(0, -ext.length);
      for (let i = 2; fs.existsSync(file); i++) { n = stem + '-' + i + ext; file = path.join(assetDir, n); }
    }
    fs.writeFileSync(file, bytes);
    return { path: 'assets/' + n, name: n, size: bytes.length };
  }
  function deleteAsset(name) {
    const n = safeAssetName(name);
    if (!n || n !== name) return { error: 'No such image.' };
    fs.rmSync(path.join(assetDir, n), { force: true });
    return {};
  }

  // Publishing: images first, the configuration last (the add-in never reads one that points at
  // images not there yet), and only what differs. photos/ and users/ are not touched.
  function publish(root, text) {
    const err = checkConfigJson(text); if (err) return { error: err };
    writeAtomic(draftFile, text);
    const dst = path.join(root, 'assets');
    fs.mkdirSync(dst, { recursive: true });
    let changed = false;
    const keep = new Set();
    for (const a of assets()) {
      keep.add(a.name);
      const src = fs.readFileSync(path.join(assetDir, a.name)), to = path.join(dst, a.name);
      if (!fs.existsSync(to) || sha256(fs.readFileSync(to)) !== sha256(src)) { writeAtomic(to, src); changed = true; }
    }
    for (const n of fs.readdirSync(dst)) {
      if (!keep.has(n) && safeAssetName(n) === n) { fs.rmSync(path.join(dst, n), { force: true }); changed = true; }
    }
    const cfgFile = path.join(root, 'signatures.json');
    if (readText(cfgFile) !== text) { writeAtomic(cfgFile, text); changed = true; }
    saveState(draftHash());
    return { changed };
  }

  function loadPublished(root) {
    const text = readText(path.join(root, 'signatures.json'));
    if (!text) return { error: 'Nothing has been published there yet.' };
    const err = checkConfigJson(text); if (err) return { error: 'The published configuration: ' + err };
    const published = new Map();
    let names = [];
    try { names = fs.readdirSync(path.join(root, 'assets')); } catch (e) { /* none */ }
    for (const n of names) if (safeAssetName(n) === n) published.set(n, fs.readFileSync(path.join(root, 'assets', n)));
    // The draft being replaced is kept, once, in case it held work that was never published.
    if (fs.existsSync(draftFile)) fs.copyFileSync(draftFile, draftFile + '.bak');
    writeAtomic(draftFile, text);
    for (const a of assets()) if (!published.has(a.name)) fs.rmSync(path.join(assetDir, a.name), { force: true });
    for (const [n, bytes] of published) fs.writeFileSync(path.join(assetDir, n), bytes);
    saveState(draftHash());
    return getDraft();
  }

  return { assetDir, getDraft, saveDraft, uploadAsset, deleteAsset, publish, loadPublished };
}

module.exports = { createStore, safeAssetName, checkConfigJson, ASSET_TYPES, writeAtomic, readText };
