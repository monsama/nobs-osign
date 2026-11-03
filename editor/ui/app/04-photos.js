  // ---------------------------------------------------------------------------------------------
  // Photos: one per person in photos/, named by a field's value (photos/@@sAMAccountName.jpg)
  // ---------------------------------------------------------------------------------------------
  // How the signatures name a photo: the field and the extension ('sAMAccountName', 'jpg').
  function photoRule() {
    let rule = null;
    for (const s of (S.config && S.config.signatures) || []) {
      const re = /<img[^>]+src="photos\/@@([A-Za-z][A-Za-z0-9_]*)\.(png|jpe?g|gif)"[^>]*>/gi;
      let m;
      while ((m = re.exec(sigFull(s)))) {
        // Its shape, from the size it is shown at: the first one shown at a width and a height
        // (one with a width only keeps the photo's own shape).
        const w = Number((/\swidth\s*=\s*["']?(\d+)/i.exec(m[0]) || [])[1]) || 0, h = Number((/\sheight\s*=\s*["']?(\d+)/i.exec(m[0]) || [])[1]) || 0;
        if (!rule) rule = { key: m[1], ext: m[2].toLowerCase(), sig: s.name, ratio: null, w: 0, h: 0 };
        if (!rule.ratio && w && h) Object.assign(rule, { ratio: h / w, w, h });
      }
    }
    return rule;
  }
  const isDefaultPhoto = (n) => n.toLowerCase().startsWith(core.PHOTO_DEFAULT + '.');
  const PHOTO_W = 400;
  let photoItems = [];
  async function renderPhotos() {
    const rule = photoRule();
    $('photosRule').textContent = rule ? t('The signature «{0}» takes photos/@@{1}.{2}: a photo is named after the person\'s {1}, e.g. mustera.{2}.', rule.sig, rule.key, rule.ext)
      : t('No signature shows a photo per person yet. Put an image photos/@@sAMAccountName.jpg into one (Insert image → Photo per person), then add the photos here.');
    $('photosRule').className = 'callout' + (rule ? '' : ' warn');
    photoNameShow();
    $('photoErr').textContent = '';
    $('photosInfo').textContent = t('Loading…');
    try {
      const r = await api('photos/list');
      photoItems = r.photos.filter((p) => !isDefaultPhoto(p.name));
      // The one the signatures take: _default with their extension (a _default of another kind is not used).
      photoDefaultShow(r.photos.find((p) => p.name.toLowerCase() === defaultName().toLowerCase()), r.mayChange);
      phRefresh();
      $('photoFile').disabled = $('photoUrlGo').disabled = !r.mayChange;
      if (!r.mayChange) $('photoErr').textContent = t('With four eyes on, photos go into photos/ of the OSign folder by hand, or the local admin uploads them.');
      photosShow(r.mayChange);
    } catch (e) { photoItems = []; $('photosInfo').textContent = ''; $('photoErr').textContent = t(e.message); $('photoList').replaceChildren(); }
  }
  function photosShow(mayChange) {
    const q = $('photosFilter').value.trim().toLowerCase();
    const list = photoItems.filter((p) => !q || p.name.toLowerCase().includes(q));
    $('photosInfo').textContent = t('{0} photo(s)', photoItems.length) + (q ? ' · ' + t('{0} shown', list.length) : '');
    const rule = photoRule();
    // A photo of another shape than the signature shows it in: Outlook stretches it. Said, and
    // made to fit with one click.
    const stretched = new Set();
    $('photosFitAll').hidden = true;
    const shapeCheck = (img, li, p) => {
      if (!rule || !rule.ratio || !img.naturalWidth) return;
      if (Math.abs(img.naturalHeight / img.naturalWidth - rule.ratio) < 0.02 || li.querySelector('.pwarn')) return;
      // Several of them: all made to fit with one click.
      stretched.add(p.name);
      if (mayChange !== false && stretched.size > 1) {
        $('photosFitAll').hidden = false;
        $('photosFitAll').textContent = t('Fit all {0}', stretched.size);
        $('photosFitAll').onclick = async () => {
          $('photosFitAll').disabled = true;
          try {
            for (const n of [...stretched]) {
              const res = await fetch('/photos/' + encodeURIComponent(n), { credentials: 'same-origin' });
              await photoSave(await res.blob(), n, { replace: true, saved: t('{0} photo(s) made to fit ({1} × {2}).', stretched.size, rule.w, rule.h) });
            }
          } finally { $('photosFitAll').disabled = false; }
        };
      }
      const fix = mayChange === false ? null : el('button', { type: 'button', class: 'small', onclick: async () => {
        fix.disabled = true;
        try {
          const res = await fetch('/photos/' + encodeURIComponent(p.name), { credentials: 'same-origin' });
          await photoSave(await res.blob(), p.name, { replace: true, saved: t('{0} made to fit ({1} × {2}).', p.name, rule.w, rule.h) });
        } finally { fix.disabled = false; }
      } }, t('Fit'));
      li.append(el('div', { class: 'pwarn', title: t('{0} × {1} pixels; the signature shows it at {2} × {3}: Outlook stretches it.', img.naturalWidth, img.naturalHeight, rule.w, rule.h) }, t('Gets stretched'), fix));
    };
    $('photoList').replaceChildren(...list.slice(0, 300).map((p) => el('li', null,
      el('img', { src: '/photos/' + encodeURIComponent(p.name) + '?v=' + encodeURIComponent(p.tag || ''), alt: '', loading: 'lazy', onload: (e) => shapeCheck(e.target, e.target.parentNode, p) }),
      el('div', { class: 'pf' }, el('span', { class: 'pn', title: p.name }, p.name),
      mayChange === false ? null : el('button', { type: 'button', class: 'ghost small pd-use', 'aria-label': t('Use {0} as the placeholder photo', p.name), title: t('As the placeholder photo, for whoever has none'), onclick: () => useAsDefault('/photos/' + encodeURIComponent(p.name), p.name) },
        el('span', { 'aria-hidden': 'true' }, '☆')),
      mayChange === false ? null : el('button', { type: 'button', class: 'ghost small danger', 'aria-label': t('Remove photo {0}', p.name), title: t('Remove'), onclick: async () => {
        if (!(await ask(t('Remove the photo {0}? The person\'s signature goes without one from now on.', p.name), t('Remove'), true))) return;
        try { await api('photos/delete', { name: p.name }); toast(t('Photo {0} removed.', p.name)); renderPhotos(); } catch (e) { $('photoErr').textContent = t(e.message); }
      } }, '✕')))));
    if (!photoItems.length) $('photoList').replaceChildren(el('li', { class: 'wide' }, emptyState('photo', t('No photos yet.'), t('Upload a person\'s photo above, or take it from an address on the web.'))));
  }
  $('photosFilter').addEventListener('input', () => photosShow(!$('photoFile').disabled));
  // The file name for the person typed: their value of the field, with the signatures' extension.
  function photoFileName() {
    const rule = photoRule() || { ext: 'jpg' };
    const who = $('photoWho').value.trim();
    return who && !/[\\/:*?"<>|]/.test(who) ? who + '.' + rule.ext : '';
  }
  function photoNameShow() { const n = photoFileName(); $('photoName').textContent = n ? '→ photos/' + n : ''; }
  $('photoWho').addEventListener('input', photoNameShow);
  suggest($('photoWho'), {
    fetch: async (q) => (adOn() ? (await api('ad/users', { query: q })).users : []).map((u) => {
      const rule = photoRule(), p = person(u);
      if (rule && /^mail$/i.test(rule.key)) p.value = u.mail || u.login;
      return p;
    }),
    pick: () => photoNameShow()
  });
  // The placeholder photo: for whoever has none of their own (photos/_default.jpg). None: their
  // signature goes without a photo.
  let photoDefault = null;
  function photoDefaultShow(p, mayChange) {
    photoDefault = p || null;
    $('photoDefaultImg').hidden = !p; $('photoDefaultNone').hidden = !!p;
    if (p) $('photoDefaultImg').src = '/photos/' + encodeURIComponent(p.name) + '?v=' + encodeURIComponent(p.tag || '');
    $('photoDefaultNote').textContent = p ? t('People without a photo get this one ({0}).', p.name) : t('People without a photo get the signature without one.');
    $('photoDefaultOff').hidden = !p;
    $('photoDefaultMake').disabled = $('photoDefaultFile').disabled = $('photoDefaultOff').disabled = $('photoDefaultAsset').disabled = mayChange === false;
    photoDefaultAssets();
    // A neutral one is drawn, and a drawing is no GIF: for signatures that take GIF photos, an image of your own.
    if (/\.gif$/i.test(defaultName())) { $('photoDefaultMake').disabled = true; $('photoDefaultMake').title = t('The signatures take GIF photos: choose a GIF of your own.'); }
    else $('photoDefaultMake').title = '';
  }
  // An image there is already as the placeholder: a person's photo, or one of the signatures'
  // images (the logo); made into the placeholder's format and size like any photo.
  // Where a problem is said: in the placeholder dialog when it is open, else on the Photos page.
  const errBox = () => ($('dlgPlaceholder').open ? $('phErr') : $('photoErr'));
  async function useAsDefault(src, label) {
    errBox().textContent = '';
    if (photoDefault && !(await ask(t('Use {0} as the placeholder photo instead of the one there is?', label), t('Use it')))) return;
    try {
      const res = await fetch(src, { credentials: 'same-origin' });
      if (!res.ok) throw new Error(t('{0} could not be read.', label));
      return await photoSave(await res.blob(), defaultName());
    } catch (e) { errBox().textContent = t(e.message); return false; }
  }
  // The placeholder as it is now (and the photos there), for the design view, the image bar and
  // the placeholder dialog.
  async function placeholderInfo() {
    const r = await api('photos/list');
    photoDefault = r.photos.find((p) => p.name.toLowerCase() === defaultName().toLowerCase()) || null;
    return { def: photoDefault, photos: r.photos.filter((p) => !isDefaultPhoto(p.name)), mayChange: r.mayChange };
  }
  const phUrl = (p) => '/photos/' + encodeURIComponent(p.name) + '?v=' + encodeURIComponent(p.tag || '');
  // The design view shows each person's photo as the placeholder that is set (the grey figure
  // when none is), and the image bar its small picture.
  async function phRefresh() {
    let def = null;
    try { def = (await placeholderInfo()).def; } catch (e) { return; }
    if (ddoc) {
      if (def) ddoc.documentElement.style.setProperty('--osign-ph', 'url("' + phUrl(def) + '")');
      else ddoc.documentElement.style.removeProperty('--osign-ph');
    }
    $('imgPhThumb').hidden = !def; $('imgPhNone').hidden = !!def;
    if (def) $('imgPhThumb').src = phUrl(def);
  }
  // The placeholder chosen from what is there: the grey figure, none, a person's photo, an image
  // of the signatures.
  async function openPlaceholderDialog() {
    $('phErr').textContent = '';
    $('phList').replaceChildren(el('p', { class: 'lead' }, t('Loading…')));
    $('dlgPlaceholder').showModal();
    let info;
    try { info = await placeholderInfo(); } catch (e) { $('phList').replaceChildren(); $('phErr').textContent = t(e.message); return; }
    const done = (ok) => { if (ok) $('dlgPlaceholder').close(); };
    const card = (img, name, run, cur) => el('button', { type: 'button', class: cur ? 'cur' : '', disabled: info.mayChange === false ? '' : null, onclick: async () => done(await run()) }, img, el('span', { class: 'nm', title: name }, name));
    const pic = (src) => el('img', { src, alt: '' });
    const neutral = el('img', { src: 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 80"><rect width="100" height="80" fill="#e9ecef"/><circle cx="50" cy="31" r="14" fill="#b8bec6"/><ellipse cx="50" cy="82" rx="30" ry="26" fill="#b8bec6"/></svg>'), alt: '' });
    $('phList').replaceChildren(
      card(neutral, t('Grey figure'), () => makeNeutral()),
      card(el('span', { class: 'ph-none' }, '–'), t('No placeholder'), () => (photoDefault ? removeDefault() : true), !info.def),
      ...(info.def ? [card(pic(phUrl(info.def)), t('As it is now'), async () => true, true)] : []),
      ...info.photos.map((p) => card(pic(phUrl(p)), p.name, () => useAsDefault('/photos/' + encodeURIComponent(p.name), p.name))),
      ...(S.assets || []).map((a) => card(pic('/assets/' + encodeURIComponent(a.name)), a.name, () => useAsDefault('/assets/' + encodeURIComponent(a.name), a.name))));
    if (info.mayChange === false) $('phErr').textContent = t('With four eyes on, photos go into photos/ of the OSign folder by hand, or the local admin uploads them.');
  }
  $('imgPhChange').addEventListener('click', () => openPlaceholderDialog());
  $('phManage').addEventListener('click', (e) => { e.preventDefault(); $('dlgPlaceholder').close(); hideImgBar(); goView('photos'); });
  // The signatures' images, offered as the placeholder.
  function photoDefaultAssets() {
    const sel = $('photoDefaultAsset');
    sel.replaceChildren(el('option', { value: '' }, t('An image of the signatures…')), ...(S.assets || []).map((a) => el('option', { value: a.name }, a.name)));
    sel.hidden = !(S.assets || []).length;
  }
  $('photoDefaultAsset').addEventListener('change', () => {
    const n = $('photoDefaultAsset').value; $('photoDefaultAsset').value = '';
    if (n) useAsDefault('/assets/' + encodeURIComponent(n), n);
  });
  const defaultName = () => core.PHOTO_DEFAULT + '.' + ((photoRule() || {}).ext || 'jpg');
  // A neutral one: a grey figure on a light ground, in the shape the signature shows photos.
  $('photoDefaultMake').addEventListener('click', () => makeNeutral());
  async function makeNeutral() {
    const ratio = (photoRule() || {}).ratio || 1, w = PHOTO_W, h = Math.round(PHOTO_W * ratio);
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const g = c.getContext('2d');
    g.fillStyle = '#e9ecef'; g.fillRect(0, 0, w, h);
    g.fillStyle = '#b8bec6';
    const r = Math.min(w, h) * 0.2, cx = w / 2, cy = h * 0.4;
    g.beginPath(); g.arc(cx, cy, r, 0, Math.PI * 2); g.fill();
    g.beginPath(); g.ellipse(cx, h + r * 0.35, r * 2.1, r * 1.9, 0, Math.PI, 0); g.fill();
    const blob = await new Promise((res) => c.toBlob(res, 'image/png'));
    return photoSave(blob, defaultName());
  }
  $('photoDefaultFile').addEventListener('change', async () => {
    const f = $('photoDefaultFile').files[0]; $('photoDefaultFile').value = '';
    if (f) await photoSave(f, defaultName());
  });
  $('photoDefaultOff').addEventListener('click', () => removeDefault());
  async function removeDefault() {
    if (!photoDefault || !(await ask(t('Remove the placeholder photo? People without a photo then get the signature without one.'), t('Remove'), true))) return false;
    try { await api('photos/delete', { name: photoDefault.name }); toast(t('Placeholder photo removed.')); photoDefault = null; renderPhotos(); schedulePreview(); phRefresh(); return true; } catch (e) { errBox().textContent = t(e.message); return false; }
  }
  // Any image, as the photo: at most 400 pixels wide, in the signatures' format (a JPEG for .jpg).
  // The part of an image of w × h with the shape (height / width) the signatures show photos in:
  // the middle across, a little above the middle down (where a face is).
  function cropTo(w, h, ratio) {
    if (!ratio || Math.abs(h / w - ratio) < 0.02) return { sx: 0, sy: 0, sw: w, sh: h };
    if (h / w > ratio) { const sh = Math.round(w * ratio); return { sx: 0, sy: Math.round((h - sh) * 0.3), sw: w, sh }; }
    const sw = Math.round(h / ratio); return { sx: Math.round((w - sw) / 2), sy: 0, sw, sh: h };
  }
  // opts.replace: no question when it is there already (made to fit again); opts.saved: the note.
  async function photoSave(blob, fixedName, opts) {
    opts = opts || {};
    const name = fixedName || photoFileName();
    if (!name) { errBox().textContent = t('Enter whose photo it is first.'); return; }
    errBox().textContent = '';
    let bmp;
    try { bmp = await createImageBitmap(blob); } catch (e) { errBox().textContent = t('This is not an image the browser can read.'); return; }
    // In the shape the signature shows it (its width and height): else Outlook stretches it.
    const rule = photoRule(), cut = cropTo(bmp.width, bmp.height, rule && rule.ratio);
    const scale = Math.min(1, PHOTO_W / cut.sw);
    const c = document.createElement('canvas'); c.width = Math.round(cut.sw * scale); c.height = Math.round(cut.sh * scale);
    const g = c.getContext('2d');
    if (/\.gif$/i.test(name) && blob.type !== 'image/gif') { errBox().textContent = t('The signatures take GIF photos: choose a GIF.'); return; }
    const type = /\.png$/i.test(name) ? 'image/png' : 'image/jpeg';
    // A JPEG has no transparency: white behind it, as in a mail.
    if (type === 'image/jpeg') { g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); }
    g.imageSmoothingQuality = 'high'; g.drawImage(bmp, cut.sx, cut.sy, cut.sw, cut.sh, 0, 0, c.width, c.height);
    const out = await new Promise((res) => c.toBlob(res, type, 0.86));
    // A GIF stays as it came (the canvas makes none).
    const data = /\.gif$/i.test(name) ? blob : out;
    const b64 = await new Promise((res) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(',')[1] || ''); fr.readAsDataURL(data); });
    if (!fixedName && isDefaultPhoto(name)) { errBox().textContent = t('That name is the placeholder photo\'s.'); return; }
    if (!opts.replace && photoItems.some((p) => p.name.toLowerCase() === name.toLowerCase()) && !(await ask(t('{0} is there already. Replace it?', name), t('Replace'), true))) return;
    try {
      await api('photos/upload', { name, data: b64 });
      toast(opts.saved || (fixedName ? t('Placeholder photo saved ({0}).', kb(data.size)) : t('Photo {0} saved ({1}).', name, kb(data.size))));
      if (!fixedName) { $('photoWho').value = ''; $('photoUrl').value = ''; photoNameShow(); }
      renderPhotos(); schedulePreview();
      if (fixedName) await phRefresh();
      return true;
    } catch (e) { errBox().textContent = t(e.message); return false; }
  }
  $('photoFile').addEventListener('change', async () => {
    const f = $('photoFile').files[0]; $('photoFile').value = '';
    if (f) await photoSave(f);
  });
  $('photoUrlGo').addEventListener('click', async () => {
    const u = $('photoUrl').value.trim();
    if (!u) return;
    if (!photoFileName()) { $('photoErr').textContent = t('Enter whose photo it is first.'); return; }
    $('photoUrlGo').disabled = true; $('photoErr').textContent = '';
    try {
      const r = await api('photos/fetch', { url: u });
      const bin = atob(r.data), bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      await photoSave(new Blob([bytes], { type: r.type }));
    } catch (e) { $('photoErr').textContent = t(e.message); }
    finally { $('photoUrlGo').disabled = false; }
  });
  $('photoUrl').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); $('photoUrlGo').click(); } });
  // From the preview's note on a missing photo: here, with the name filled in.
  function photoAddFor(name) {
    goView('photos');
    $('photoWho').value = name.replace(/\.[^.]+$/, ''); photoNameShow(); $('photoWho').focus();
  }

