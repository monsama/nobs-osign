/*
 * What was published: the versions (server/history.js), what changed between them, one back into
 * the draft, the published one as it is, and one signature of it back. Routes for server/app.js.
 */
'use strict';
const path = require('node:path');
const { safeAssetName, writeAtomic } = require('./store');
const { diffConfigs } = require('./diff');

function createHistoryRoutes({ history, store, settings, busy }) {
  const routes = {
    // What changed: a version against the one before it, or against the draft.
    'history/diff': (d) => {
      const v = history.get(String(d.id || ''));
      if (!v) return { error: 'There is no such version.' };
      const draft = store.getDraft().configJson || '';
      return d.against === 'draft' ? { changes: diffConfigs(v.configJson, draft) } : { changes: diffConfigs(v.previous, v.configJson) };
    },
    // The versions published, and one of them back into the draft (with its images).
    'history/list': () => ({ versions: history.list() }),
    'history/restore': (d) => {
      if (busy()) return { error: 'Publishing right now: try again in a moment.' };
      const r = history.restore(String(d.id || ''));
      if (r.error) return r;
      const saved = store.saveDraft(r.configJson);
      if (saved && saved.error) return saved;
      return { at: r.at };
    },
    // What is published, to see what the draft changed: the configuration as it went out.
    'published/get': () => ({ configJson: history.latest() || '' }),
    // One signature back as it is published (its changes in the draft dropped), with the images it
    // needs that the draft no longer has.
    'published/revert': (d) => {
      if (busy()) return { error: 'Publishing right now: try again in a moment.' };
      let c = null; try { c = JSON.parse(history.latest() || 'null'); } catch (e) { c = null; }
      const sig = c && (c.signatures || []).find((x) => x && x.id === String(d.id || ''));
      if (!sig) return { error: 'This signature is not published yet: there is nothing to go back to.' };
      const have = new Set(store.draftImages().keys());
      for (const m of String((sig.source || '') + (sig.html || '')).matchAll(/assets\/([^"'\s)?#]+)/g)) {
        let n = m[1]; try { n = decodeURIComponent(n); } catch (e) { /* as it is */ }
        if (have.has(n) || safeAssetName(n) !== n) continue;
        const bytes = history.latestImage(n);
        if (bytes) { writeAtomic(path.join(store.assetDir, n), bytes); have.add(n); }
      }
      return { signature: sig, assets: store.getDraft().assets };
    },
    'load-published': async () => {
      if (busy()) return { error: 'Publishing right now: try again in a moment.' };
      try { return await store.loadPublished(settings.target()); } catch (e) { return { error: e.message }; }
    }
  };
  return { routes };
}

module.exports = { createHistoryRoutes };
