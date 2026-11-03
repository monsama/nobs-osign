/*
 * The add-in's files: made into the data volume (server/addin.js), checked where they are served,
 * and uploaded to Cloudflare Pages. Routes for server/app.js; live() gives what changes with the
 * connections (graph).
 */
'use strict';
const fs = require('node:fs');
const path = require('node:path');
const addin = require('./addin');
const { writeAtomic } = require('./store');

function createAddinRoutes({ env, settings, pages, addinDir, addinUploadFile, addinMade, fetchFn, adminOnly, live, log }) {
  const routes = {
    // The add-in's files: made into the data volume, then uploaded or downloaded.
    'addin/build': async (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      // With Cloudflare Pages and no web server chosen: Cloudflare's caching file goes with them.
      const st = settings.get();
      if (!st.addinHost && addin.pagesEnabled(pages)) st.addinHost = 'cloudflare';
      // The folder by its ids, when the server publishes to SharePoint: the add-in then reads it
      // without looking up the site (which a user who may only read the folder cannot).
      // Not reachable yet (before the site is granted, say): made without them; the add-in then
      // finds the folder through its site, and making the files again later adds them.
      let note = '';
      if (live().graph) {
        try { const at = await settings.target().locate(); st.spDrive = at.driveId; st.spItem = at.itemId; }
        catch (e) { note = 'The SharePoint folder could not be reached (' + e.message + '): make the files again once it can, so the add-in reads it directly.'; log.warn('addin/build: ' + note); }
      }
      const r = await addin.build(env, st, addinDir);
      if (!r.error && note) r.note = note;
      return r.error ? r : Object.assign({ folder: '' }, r);
    },
    'addin/check': async () => {
      const s = settings.get();
      if (!s.addinUrl) return { error: 'Enter the address the add-in will be served from.' };
      if (!fs.existsSync(path.join(addinDir, 'manifest.xml'))) return { error: 'Create the add-in files first.' };
      return addin.check(s.addinUrl, addinDir, fetchFn);
    },
    'addin/upload': async (d, s) => {
      if (adminOnly(s)) return { error: adminOnly(s) };
      if (!fs.existsSync(path.join(addinDir, 'manifest.xml'))) return { error: 'Create the add-in files first.' };
      try {
        const r = await addin.pagesDeploy(pages, addinDir, fetchFn);
        // Which version went up, and when: Status compares it with this server's.
        writeAtomic(addinUploadFile, JSON.stringify({ version: addinMade().version, at: Date.now(), url: r.url }));
        return r;
      } catch (e) { return { error: e.message }; }
    }
  };
  return { routes };
}

module.exports = { createAddinRoutes };
