  // ---------------------------------------------------------------------------------------------
  // The draft's images: which are used, made lighter (for the web and for classic Outlook), put in
  // another's place, looked at (a dark logo). The ways to classic Outlook: 10-classic.js.
  // ---------------------------------------------------------------------------------------------
  // Whether a PNG is a dark logo on nothing (darkCheck): by name, looked at once.
  const dark = {};
  const kb2 = (n) => (n >= 1024 * 1024 ? (n / 1024 / 1024).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB');
  const assetSize = (name) => ((S.assets || []).find((x) => x.name === name) || {}).size || 0;
  // An image's address, with its size: a file under a name used before is not taken from the cache.
  const assetUrl = (name) => '/assets/' + encodeURIComponent(name) + '?v=' + assetSize(name);
  const fetchAsset = async (name) => (await fetch(assetUrl(name), { credentials: 'same-origin' })).blob();
  // The HTML of what shows images: signatures, building blocks, campaign banners.
  const usedHtml = () => S.config.signatures.map((s) => s.html || '').concat((S.config.blocks || []).map((b) => b.html || ''), (S.config.campaigns || []).flatMap((c) => [c.html || ''].concat(Object.values(c.langs || {}).map((v) => v.html || ''), (c.more || []).map((v) => v.html || ''))));
  const widthOf = (tag) => Number((/\swidth\s*=\s*["']?(\d+)/i.exec(tag) || [])[1]) || 0;
  // The width an image is shown at in all of that, the widest (0: none set).
  function shownWidth(name) { let w = 0; for (const h of usedHtml()) for (const m of h.matchAll(/<img\b[^>]*>/gi)) if (m[0].includes('assets/' + name + '"')) w = Math.max(w, widthOf(m[0])); return w; }

  // Pixels: an image loaded, drawn at a width (its shape kept), and what is in it.
  function loadImage(src) { return new Promise((ok, no) => { const img = new Image(); img.onload = () => ok(img); img.onerror = no; img.src = src; }); }
  async function imageOf(blob) { const url = URL.createObjectURL(blob); try { return await loadImage(url); } finally { URL.revokeObjectURL(url); } }
  function drawn(img, w) {
    const cv = document.createElement('canvas'), g = cv.getContext('2d');
    cv.width = Math.max(1, Math.round(w)); cv.height = Math.max(1, Math.round(img.naturalHeight * cv.width / (img.naturalWidth || 1)));
    g.imageSmoothingQuality = 'high'; g.drawImage(img, 0, 0, cv.width, cv.height);
    return { cv, data: g.getImageData(0, 0, cv.width, cv.height) };
  }
  const hasClear = (data) => { const d = data.data; for (let i = 3; i < d.length; i += 4) if (d[i] < 250) return true; return false; };
  const jpegOf = (cv, quality) => new Promise((r) => cv.toBlob(r, 'image/jpeg', quality));
  // A GIF that moves: more than one frame (a graphic control block each). It stays as it is.
  function animatedGif(bytes) { let frames = 0; for (let i = 0; i + 2 < bytes.length; i++) if (bytes[i] === 0x21 && bytes[i + 1] === 0xf9 && bytes[i + 2] === 0x04 && ++frames > 1) return true; return false; }

  // The images of the draft nothing uses (signatures, building blocks, campaigns, their forms).
  function unusedImages() {
    const c = S.config, text = c.signatures.map((s) => sigFull(s)).concat((c.blocks || []).map((b) => b.html || ''),
      (c.campaigns || []).flatMap((x) => [x].concat(x.more || [], Object.values(x.langs || {})).flatMap((p) => [p.html || '', 'assets/' + ((p.design || {}).img || '') + '"'])));
    const all = text.join('\n');
    return (S.assets || []).map((a) => a.name).filter((n) => !all.includes('assets/' + n + '"') && !all.includes('assets/' + n + "'") && !all.includes('assets/' + n + '?'));
  }
  // Those images deleted on the server, asked first: what came of it in words (null: not asked
  // for). The undo steps start anew after it: the files they name are gone.
  async function removeUnused() {
    const names = unusedImages(); let n = 0;
    if (!names.length) return null;
    if (!(await ask(t('Remove these images from the draft? Nothing uses them: {0}', names.join(', ')), t('Remove'), true))) return null;
    for (const name of names) { try { await api('assets/delete', { name }); S.assets = S.assets.filter((a) => a.name !== name); n++; } catch (e) { /* that one stays */ } }
    changed({ list: false });
    ghReset();
    return n === 1 ? t('1 unused image removed from the draft.') : t('{0} unused images removed from the draft.', n);
  }
  $('imgUnused').addEventListener('click', async () => { const said = await removeUnused(); if (said) toast(said); renderImageList(); });
  // After the signatures' HTML changed here: the HTML tab shows the new one (else reopening would
  // take the old text in it for a change), saved, the signature or building block open again.
  function reopenChanged() {
    if (S.mode === 'html' && sig()) showSource(sig().html);
    changed();
    if (S.view !== 'sig') return;
    if (S.blockEdit) openBlock(S.blockEdit);
    else if (S.sigId) openSig(S.sigId);
  }
  // The name an image is saved again under: without the -2, -3 an earlier save gave it (the server
  // never writes over a file, so each save would add one), when that name is free. Only the
  // server's endings (one or two digits): xmas-2026.png stays as it is.
  function freshName(name) {
    const base = name.replace(/(?:-\d{1,2})+(\.[A-Za-z0-9]+)$/, '$1');
    return (S.assets || []).some((a) => a.name === base) ? name : base;
  }
  // An image put in another's place: uploaded under a new name (one is never written over), named
  // so wherever it is used, the old one gone when nothing uses it any more. The new one.
  async function swapAsset(name, blob, ext) {
    const r = await api('assets/upload', { name: freshName(name.replace(/\.[a-z0-9]+$/i, '') + ext), data: await b64Of(blob) });
    S.assets = S.assets.filter((x) => x.name !== r.name).concat([{ name: r.name, size: r.size }]);
    renameAsset(name, r.name);
    if (r.name !== name && unusedImages().includes(name)) { try { await api('assets/delete', { name }); S.assets = S.assets.filter((x) => x.name !== name); } catch (e) { /* it stays */ } }
    return r;
  }
  // An image named anew wherever it is used: signatures, building blocks, campaign banners (made
  // from their form, or from their own HTML); what was measured of a file under the new name
  // before (a name can come back once its file is gone) forgotten.
  function renameAsset(name, to) {
    forgetAsset(to);
    if (to === name) return;
    const re = new RegExp('assets/' + name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=["\'\\s>?#])', 'g');
    S.config.signatures.forEach((x) => { x.html = (x.html || '').replace(re, 'assets/' + to); });
    (S.config.blocks || []).forEach((x) => { x.html = (x.html || '').replace(re, 'assets/' + to); });
    (S.config.campaigns || []).forEach((c) => {
      [c].concat(Object.values(c.langs || {}), c.more || []).forEach((v) => {
        if (!v.design) return;
        if (v.design.img === name) v.design.img = to;
        if (typeof v.design.html === 'string') v.design.html = v.design.html.replace(re, 'assets/' + to);
      });
      campBuild(c);
    });
  }
  function forgetAsset(name) {
    for (const k of Object.keys(classicImg)) if (k.startsWith(name + '@')) delete classicImg[k];
    delete dark[name]; delete darkWait[name];
  }

  // The images for the web, in one go (quiet; prepareClassic says what came of it): each shown as
  // small as it is shown (twice that for sharp screens; strong: as wide as shown, more compressed),
  // a still GIF as a PNG, a PNG with a palette where it can be, one without transparency as a JPEG.
  // only: these images only (by name).
  async function optimizeImages(opts) {
    const strong = !!(opts && opts.strong);
    const only = opts && opts.only ? new Set(opts.only) : null;
    const a = await shrinkUsed({ strong, only });
    // Made smaller under a new name: still one of these.
    if (only) for (const n of Object.values(a.renamed)) only.add(n);
    const used = (re) => [...new Set(usedHtml().flatMap((h) => [...h.matchAll(re)].map((m) => m[1])))].filter((n) => !only || only.has(n));
    const c = await gifToPng(used(/src="assets\/([^"?#]+\.gif)"/gi), strong);
    const b = await pngToJpeg(used(/src="assets\/([^"?#]+\.png)"/gi), strong ? 0.72 : 0.86);
    return a.n + b.n + c.n;
  }
  // The images with a width in the HTML, as small as they are shown: twice that (sharp on
  // high-resolution screens; strong: as wide as shown, JPEGs at a lower quality). One without a
  // width stays as it is.
  async function shrinkUsed(opts) {
    const only = opts && opts.only, strong = !!(opts && opts.strong);
    const shown = new Map();
    for (const h of usedHtml()) for (const m of h.matchAll(/<img\b[^>]*>/gi)) {
      const src = (/\ssrc\s*=\s*["']?assets\/([^"'\s>]+)/i.exec(m[0]) || [])[1], w = widthOf(m[0]);
      if (src && w && (!only || (only instanceof Set ? only.has(src) : only === src))) shown.set(src, Math.max(shown.get(src) || 0, w));
    }
    let n = 0, from = 0, to = 0;
    const renamed = {}, done = [];
    for (const [name, w] of shown) {
      try {
        const blob = await fetchAsset(name);
        const f = new File([blob], name, { type: blob.type || (/\.png$/i.test(name) ? 'image/png' : 'image/jpeg') });
        const small = await shrinkUpload(f, Math.max(48, (strong ? 1 : 2) * w), strong ? { quality: 0.72, always: true, palette: 'lossy' } : { palette: 'exact' });
        if (!small.shrunk || small.size >= f.size) continue;
        const r = await swapAsset(name, small, (/\.[a-z0-9]+$/i.exec(name) || [''])[0]);
        if (r.name !== name) renamed[name] = r.name;
        n++; from += f.size; to += small.size; done.push({ name: r.name, from: f.size, to: small.size });
      } catch (e) { /* that one stays */ }
    }
    if (n) reopenChanged();
    return { n, from, to, renamed, done };
  }
  // Still GIFs (an icon, a logo) as PNGs with a palette, as wide as they are shown (twice that,
  // unless strong): a GIF is not made smaller otherwise, here or by the add-in. One that moves stays.
  async function gifToPng(gifs, strong) {
    let n = 0;
    for (const name of gifs) {
      try {
        const blob = await fetchAsset(name), bytes = new Uint8Array(await blob.arrayBuffer());
        if (animatedGif(bytes)) continue;
        const img = await imageOf(new Blob([bytes], { type: 'image/gif' }));
        const shown = shownWidth(name);
        const png = await palettePng(drawn(img, Math.min(img.naturalWidth, shown ? (strong ? 1 : 2) * shown : img.naturalWidth)).data, true);
        if (!png || png.size >= bytes.length) continue;
        await swapAsset(name, png, '.png');
        n++;
      } catch (e) { /* that one stays */ }
    }
    if (n) reopenChanged();
    return { n };
  }
  // These PNGs as JPEGs where they have no transparent pixel (a photo, a banner is much smaller
  // so); one with transparency (a logo) stays: a JPEG would give it a box.
  async function pngToJpeg(pngs, quality) {
    let n = 0, from = 0, to = 0;
    const done = [];
    for (const name of pngs) {
      try {
        const img = await loadImage(assetUrl(name)), { cv, data } = drawn(img, img.naturalWidth);
        if (hasClear(data)) continue;
        const blob = await jpegOf(cv, quality || 0.86), was = assetSize(name);
        if (!blob || (was && blob.size >= was)) continue;
        const r = await swapAsset(name, blob, '.jpg');
        n++; from += was; to += blob.size; done.push({ name: r.name, from: was, to: blob.size });
      } catch (e) { /* that one stays */ }
    }
    if (n) reopenChanged();
    return { n, from, to, done };
  }
  // One image made at most this many bytes, or as near as it goes: drawn narrower step by step,
  // without a transparent pixel as a JPEG at a lower quality, with one as a PNG with fewer colours.
  // Without a width in the HTML it shows as wide as its file: then only made lighter, not
  // narrower. The smallest made takes its place when it is smaller than the file: the new one, or null.
  async function squeezeTo(name, bytes) {
    try {
      const was = assetSize(name);
      if (was && was <= bytes) return null;
      const blob = await fetchAsset(name);
      if ((/\.gif$/i.test(name) || blob.type === 'image/gif') && animatedGif(new Uint8Array(await blob.arrayBuffer()))) return null;
      const img = await imageOf(blob);
      const set = shownWidth(name), shown = set || img.naturalWidth;
      const clear = hasClear(drawn(img, Math.min(img.naturalWidth, shown)).data);
      let best = null;
      for (const scale of set ? [1, 0.85, 0.7, 0.55, 0.45, 0.35] : [1]) {
        const { cv, data } = drawn(img, Math.max(16, Math.min(img.naturalWidth, shown * scale)));
        for (const step of clear ? [256, 128, 64, 32, 16] : [0.6, 0.5, 0.4, 0.32, 0.25]) {
          const b = clear ? await palettePng(data, true, step) : await jpegOf(cv, step);
          if (b && (!best || b.size < best.size)) best = b;
          if (best && best.size <= bytes) break;
        }
        if (best && best.size <= bytes) break;
      }
      if (!best || best.size >= (was || blob.size)) return null;
      return await swapAsset(name, best, clear ? '.png' : '.jpg');
    } catch (e) { return null; }
  }
  // Whether an image is dark lines on nothing (a logo that disappears on a dark background):
  // looked at once, in the background; the publish checks say it the next time they are shown
  // (Publish… waits for the ones used). Done when it is known (at once when it was already).
  const darkWait = {};
  function darkCheck(name) {
    if (name in dark) return darkWait[name] || Promise.resolve();
    dark[name] = null;
    const p = darkWait[name] = loadImage(assetUrl(name)).then((img) => {
      try {
        const d = drawn(img, Math.min(120, img.naturalWidth)).data.data;
        let clear = 0, seen = 0, lum = 0;
        for (let i = 0; i < d.length; i += 4) { if (d[i + 3] < 40) { clear++; continue; } seen++; lum += 0.2126 * d[i] + 0.7152 * d[i + 1] + 0.0722 * d[i + 2]; }
        dark[name] = clear > d.length / 4 * 0.2 && seen > 0 && lum / seen < 70;
      } catch (e) { dark[name] = false; }
    }, () => { dark[name] = false; }).finally(() => { if (darkWait[name] === p) delete darkWait[name]; });
    return p;
  }
