/*
 * What was published, kept: each publish with when, who, what changed (the note given), the
 * configuration and its images, so an earlier version can come back into the draft.
 *
 * In the data volume: history/<time>.json per publish, and history/assets/<sha256><ext>, each image
 * once however many versions have it. The last KEEP versions are kept; images no version has go.
 */
'use strict';
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { writeAtomic, readText } = require('./store');

const KEEP = 50;
const sha256 = (b) => crypto.createHash('sha256').update(b).digest('hex');
const ID = /^\d{13}$/;

function createHistory({ dataDir, assetDir }) {
  const dir = path.join(dataDir, 'history');
  const blobs = path.join(dir, 'assets');
  fs.mkdirSync(blobs, { recursive: true });
  const ids = () => fs.readdirSync(dir).filter((f) => /^\d{13}\.json$/.test(f)).map((f) => f.slice(0, 13)).sort().reverse();
  const load = (id) => { try { return JSON.parse(readText(path.join(dir, id + '.json')) || 'null'); } catch (e) { return null; } };

  // After a publish: the version as it went out.
  function add({ by, message, configJson }) {
    const assets = [];
    for (const name of fs.readdirSync(assetDir).sort()) {
      const file = path.join(assetDir, name);
      if (!fs.statSync(file).isFile()) continue;
      const data = fs.readFileSync(file);
      const blob = sha256(data) + path.extname(name).toLowerCase();
      if (!fs.existsSync(path.join(blobs, blob))) writeAtomic(path.join(blobs, blob), data);
      assets.push({ name, blob });
    }
    let id = String(Date.now());
    while (fs.existsSync(path.join(dir, id + '.json'))) id = String(Number(id) + 1);
    writeAtomic(path.join(dir, id + '.json'), JSON.stringify({ at: Number(id), by: String(by || ''), message: String(message || '').slice(0, 200), configJson, assets }));
    prune();
    return id;
  }
  function prune() {
    const all = ids();
    for (const id of all.slice(KEEP)) { try { fs.unlinkSync(path.join(dir, id + '.json')); } catch (e) { /* gone */ } }
    const used = new Set();
    for (const id of all.slice(0, KEEP)) for (const a of (load(id) || {}).assets || []) used.add(a.blob);
    for (const b of fs.readdirSync(blobs)) if (!used.has(b)) { try { fs.unlinkSync(path.join(blobs, b)); } catch (e) { /* gone */ } }
  }

  // The versions, newest first: when, who, the note, how many signatures.
  function list() {
    return ids().map((id) => {
      const h = load(id); if (!h) return null;
      let signatures = 0; try { signatures = (JSON.parse(h.configJson).signatures || []).length; } catch (e) { /* none */ }
      return { id, at: h.at, by: h.by, message: h.message, signatures, images: (h.assets || []).length };
    }).filter(Boolean);
  }

  // That version into the draft: its configuration, and exactly its images.
  function restore(id) {
    if (!ID.test(String(id || ''))) return { error: 'There is no such version.' };
    const h = load(id);
    if (!h) return { error: 'There is no such version.' };
    const missing = (h.assets || []).filter((a) => !fs.existsSync(path.join(blobs, a.blob)));
    if (missing.length) return { error: 'Images of this version are gone: ' + missing.map((a) => a.name).join(', ') };
    for (const name of fs.readdirSync(assetDir)) { try { fs.unlinkSync(path.join(assetDir, name)); } catch (e) { /* gone */ } }
    for (const a of h.assets || []) if (path.basename(a.name) === a.name && !a.name.startsWith('.')) fs.copyFileSync(path.join(blobs, a.blob), path.join(assetDir, a.name));
    return { configJson: h.configJson, at: h.at };
  }

  return { add, list, restore };
}

module.exports = { createHistory };
