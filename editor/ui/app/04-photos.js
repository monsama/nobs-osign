  // ---------------------------------------------------------------------------------------------
  // Photos: one per person in photos/, named by a field's value (photos/@@sAMAccountName.jpg)
  // ---------------------------------------------------------------------------------------------
  // How the signatures name a photo: the field and the extension ('sAMAccountName', 'jpg').
  function photoRule() {
    for (const s of (S.config && S.config.signatures) || []) {
      const m = /<img[^>]+src="photos\/@@([A-Za-z][A-Za-z0-9_]*)\.(png|jpe?g|gif)"[^>]*>/i.exec(s.html || '');
      if (m) {
        // Its shape, from the size it is shown at (for the neutral placeholder): 1:1 when not given.
        const w = Number((/\swidth\s*=\s*["']?(\d+)/i.exec(m[0]) || [])[1]) || 0, h = Number((/\sheight\s*=\s*["']?(\d+)/i.exec(m[0]) || [])[1]) || 0;
        return { key: m[1], ext: m[2].toLowerCase(), sig: s.name, ratio: w && h ? h / w : 1 };
      }
    }
    return null;
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
      photoDefaultShow(r.photos.find((p) => isDefaultPhoto(p.name)), r.mayChange);
      $('photoFile').disabled = $('photoUrlGo').disabled = !r.mayChange;
      if (!r.mayChange) $('photoErr').textContent = t('With four eyes on, photos go into photos/ of the OSign folder by hand, or the local admin uploads them.');
      photosShow(r.mayChange);
    } catch (e) { photoItems = []; $('photosInfo').textContent = ''; $('photoErr').textContent = t(e.message); $('photoList').replaceChildren(); }
  }
  function photosShow(mayChange) {
    const q = $('photosFilter').value.trim().toLowerCase();
    const list = photoItems.filter((p) => !q || p.name.toLowerCase().includes(q));
    $('photosInfo').textContent = t('{0} photo(s)', photoItems.length) + (q ? ' · ' + t('{0} shown', list.length) : '');
    $('photoList').replaceChildren(...list.slice(0, 300).map((p) => el('li', null,
      el('img', { src: '/photos/' + encodeURIComponent(p.name) + '?v=' + encodeURIComponent(p.tag || ''), alt: '', loading: 'lazy' }),
      el('div', { class: 'pf' }, el('span', { class: 'pn' }, p.name),
      mayChange === false ? null : el('button', { type: 'button', class: 'ghost small danger', 'aria-label': t('Remove photo {0}', p.name), title: t('Remove'), onclick: async () => {
        if (!(await ask(t('Remove the photo {0}? The person\'s signature goes without one from now on.', p.name), t('Remove'), true))) return;
        try { await api('photos/delete', { name: p.name }); toast(t('Photo {0} removed.', p.name)); renderPhotos(); } catch (e) { $('photoErr').textContent = t(e.message); }
      } }, '✕')))));
    if (!photoItems.length) $('photoList').replaceChildren(el('li', { class: 'empty' }, t('No photos yet.')));
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
    $('photoDefaultMake').disabled = $('photoDefaultFile').disabled = $('photoDefaultOff').disabled = mayChange === false;
  }
  const defaultName = () => core.PHOTO_DEFAULT + '.' + ((photoRule() || {}).ext || 'jpg');
  // A neutral one: a grey figure on a light ground, in the shape the signature shows photos.
  $('photoDefaultMake').addEventListener('click', async () => {
    const ratio = (photoRule() || {}).ratio || 1, w = PHOTO_W, h = Math.round(PHOTO_W * ratio);
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const g = c.getContext('2d');
    g.fillStyle = '#e9ecef'; g.fillRect(0, 0, w, h);
    g.fillStyle = '#b8bec6';
    const r = Math.min(w, h) * 0.2, cx = w / 2, cy = h * 0.4;
    g.beginPath(); g.arc(cx, cy, r, 0, Math.PI * 2); g.fill();
    g.beginPath(); g.ellipse(cx, h + r * 0.35, r * 2.1, r * 1.9, 0, Math.PI, 0); g.fill();
    const blob = await new Promise((res) => c.toBlob(res, 'image/png'));
    await photoSave(blob, defaultName());
  });
  $('photoDefaultFile').addEventListener('change', async () => {
    const f = $('photoDefaultFile').files[0]; $('photoDefaultFile').value = '';
    if (f) await photoSave(f, defaultName());
  });
  $('photoDefaultOff').addEventListener('click', async () => {
    if (!photoDefault || !(await ask(t('Remove the placeholder photo? People without a photo then get the signature without one.'), t('Remove'), true))) return;
    try { await api('photos/delete', { name: photoDefault.name }); toast(t('Placeholder photo removed.')); renderPhotos(); schedulePreview(); } catch (e) { $('photoErr').textContent = t(e.message); }
  });
  // Any image, as the photo: at most 400 pixels wide, in the signatures' format (a JPEG for .jpg).
  async function photoSave(blob, fixedName) {
    const name = fixedName || photoFileName();
    if (!name) { $('photoErr').textContent = t('Enter whose photo it is first.'); return; }
    $('photoErr').textContent = '';
    let bmp;
    try { bmp = await createImageBitmap(blob); } catch (e) { $('photoErr').textContent = t('This is not an image the browser can read.'); return; }
    const scale = Math.min(1, PHOTO_W / bmp.width);
    const c = document.createElement('canvas'); c.width = Math.round(bmp.width * scale); c.height = Math.round(bmp.height * scale);
    const g = c.getContext('2d');
    if (/\.gif$/i.test(name) && blob.type !== 'image/gif') { $('photoErr').textContent = t('The signatures take GIF photos: choose a GIF.'); return; }
    const type = /\.png$/i.test(name) ? 'image/png' : 'image/jpeg';
    // A JPEG has no transparency: white behind it, as in a mail.
    if (type === 'image/jpeg') { g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); }
    g.imageSmoothingQuality = 'high'; g.drawImage(bmp, 0, 0, c.width, c.height);
    const out = await new Promise((res) => c.toBlob(res, type, 0.86));
    // A GIF stays as it came (the canvas makes none).
    const data = /\.gif$/i.test(name) ? blob : out;
    const b64 = await new Promise((res) => { const fr = new FileReader(); fr.onload = () => res(String(fr.result).split(',')[1] || ''); fr.readAsDataURL(data); });
    if (!fixedName && isDefaultPhoto(name)) { $('photoErr').textContent = t('That name is the placeholder photo\'s.'); return; }
    if (photoItems.some((p) => p.name.toLowerCase() === name.toLowerCase()) && !(await ask(t('{0} is there already. Replace it?', name), t('Replace'), true))) return;
    try {
      await api('photos/upload', { name, data: b64 });
      toast(fixedName ? t('Placeholder photo saved ({0}).', kb(data.size)) : t('Photo {0} saved ({1}).', name, kb(data.size)));
      if (!fixedName) { $('photoWho').value = ''; $('photoUrl').value = ''; photoNameShow(); }
      renderPhotos(); schedulePreview();
    } catch (e) { $('photoErr').textContent = t(e.message); }
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

