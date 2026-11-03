/*
 * A person's photo (photos/<username>.jpg in the OSign folder): uploaded, listed, removed, or
 * fetched from an address on the web (server/photofetch.js). A photo counts at once, without
 * publishing: the publishing group's, and with four eyes on only the local admin's. Routes for
 * server/app.js.
 */
'use strict';
const path = require('node:path');
const { fetchPhoto } = require('./photofetch');

const PHOTO_NAME = /^[^/\\:*?"<>|\u0000-\u001f]{1,120}\.(png|jpe?g|gif)$/i;

function createPhotoRoutes({ settings, photos, publisher, needsApproval, opts }) {
  const routes = {
    // A person's photo (photos/<username>.jpg in the OSign folder), e.g. brought over from CI-Sign.
    'photos/upload': async (d, s) => {
      // A photo counts at once, without publishing: a publisher's, and with four eyes only the local admin's.
      if (!publisher(s)) return { error: 'A photo counts at once, without publishing: only the publishing group changes photos.' };
      if (needsApproval() && !s.user.local) return { error: 'With four eyes on, photos go into photos/ of the OSign folder by hand, or the local admin uploads them.' };
      const name = String(d.name || '');
      if (!PHOTO_NAME.test(name) || name !== path.basename(name)) return { error: 'A photo is a PNG, JPG or GIF named after the user, like mustera.jpg.' };
      const data = String(d.data || '');
      const bytes = Buffer.from(data, 'base64');
      if (!bytes.length || bytes.toString('base64').replace(/=+$/, '') !== data.replace(/\s+/g, '').replace(/=+$/, '')) return { error: 'The image did not arrive whole.' };
      if (bytes.length > 1024 * 1024) return { error: 'The image is larger than 1 MB. A signature image is best well under 100 KB.' };
      const magic = { png: '89504e47', jpg: 'ffd8ff', jpeg: 'ffd8ff', gif: '47494638' }[name.split('.').pop().toLowerCase()];
      if (!bytes.subarray(0, 8).toString('hex').startsWith(magic)) return { error: 'The file is not the image its name says.' };
      await settings.target().write('photos/' + name, bytes, 'image/' + (/png$/i.test(name) ? 'png' : /gif$/i.test(name) ? 'gif' : 'jpeg'));
      photos.clear();
      return { name, size: bytes.length };
    },
    // The photos there, for the Photos page: name and a tag that changes with the file.
    'photos/list': async (d, s) => {
      const t = settings.target();
      try {
        const list = (await t.listInfo('photos')).filter((f) => PHOTO_NAME.test(f.name) && !f.name.startsWith('.')).sort((a, b) => a.name.localeCompare(b.name, 'en', { sensitivity: 'base' }));
        return { photos: list.map((f) => ({ name: f.name, tag: f.tag })), where: t.describe() + '/photos', mayChange: publisher(s) && (!needsApproval() || !!s.user.local) };
      } catch (e) { return { error: e.message }; }
    },
    'photos/delete': async (d, s) => {
      if (!publisher(s)) return { error: 'A photo counts at once, without publishing: only the publishing group changes photos.' };
      if (needsApproval() && !s.user.local) return { error: 'With four eyes on, photos go into photos/ of the OSign folder by hand, or the local admin uploads them.' };
      const name = String(d.name || '');
      if (!PHOTO_NAME.test(name) || name !== path.basename(name) || name.startsWith('.')) return { error: 'There is no such photo.' };
      await settings.target().remove('photos/' + name);
      photos.clear();
      return { name };
    },
    // A photo from a web address (the team page, say): fetched here, as the browser may not fetch
    // it from another site; the editor makes it small and keeps it with photos/upload.
    'photos/fetch': async (d) => {
      try {
        const r = await ((opts && opts.photoFetch) || fetchPhoto)(d.url);
        return { type: r.type, data: r.bytes.toString('base64') };
      } catch (e) { return { error: e.message }; }
    }
  };
  return { routes };
}

module.exports = { createPhotoRoutes, PHOTO_NAME };
