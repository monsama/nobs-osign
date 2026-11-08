  // ---------------------------------------------------------------------------------------------
  // Classic Outlook: what a signature takes in the mailbox there (about 30 KB), and the two ways to
  // make it fit: preparing (the HTML lighter, the images optimized) and the images small enough.
  // ---------------------------------------------------------------------------------------------
  // Classic Outlook keeps the signatures that go in by themselves in the mailbox's settings
  // (about 30 KB, keepForClassic in the add-in): roughly what one takes there, with the banner
  // running now and its images as files (made a bit smaller there, so it is an upper bound).
  const CLASSIC_KB = 30;
  // What an image takes there: as the pane makes it (taskpane.js shrink: 1.5 times as wide as it
  // is shown, JPEG/PNG again at 0.82, as base64), measured once in the background; until then
  // the file as it is.
  const classicImg = {};
  // The steps the add-in makes an image smaller by when it does not fit there (common.js, the same).
  const CLASSIC_SQUEEZE = [{ scale: 1, quality: 0.7 }, { scale: 1, quality: 0.55 }, { scale: 0.8, quality: 0.5 }, { scale: 0.65, quality: 0.45 }];
  const classicPending = new Set();
  function classicMeasure(name, shown, step) {
    const k = name + '@' + shown + (step ? '@' + step.scale + '/' + step.quality : '');
    if (k in classicImg) return;
    classicImg[k] = null;
    let done; const wait = new Promise((r) => { done = r; }); classicPending.add(wait);
    const settle = () => { classicPending.delete(wait); done(); };
    const type = /\.png$/i.test(name) ? 'image/png' : /\.jpe?g$/i.test(name) ? 'image/jpeg' : '';
    const img = new Image();
    img.onload = () => {
      try {
        const file = Math.ceil(assetSize(name) / 3) * 4;
        if (!type) { classicImg[k] = file; settle(); return; }
        const want = Math.min(img.naturalWidth, Math.max(1, (shown || img.naturalWidth) * (step ? step.scale : 1.5)));
        const b64 = (drawn(img, want).cv.toDataURL(type, step ? step.quality : 0.82).split(',')[1] || '').length;
        classicImg[k] = file && file < b64 ? file : b64;
      } catch (e) { classicImg[k] = assetSize(name) * 4 / 3; }
      settle(); schedulePreview();
    };
    img.onerror = () => { classicImg[k] = 0; settle(); };
    img.src = assetUrl(name);
  }
  // Roughly what one signature takes there (its HTML, its images as there, a photo), and its images.
  function classicParts(s) {
    const c = core.campaignFor(S.config, s, { anyone: true, kind: 'new' });
    const banner = c ? [c].concat(c.more || []).map((x) => x.html || '').join('') : '';
    const html = sigFull(s) + banner, seen = new Map();
    for (const m of html.matchAll(/<img\b[^>]*>/gi)) {
      const n = (/src="assets\/([^"?#]+)"/i.exec(m[0]) || [])[1]; if (!n || seen.has(n)) continue;
      const w = widthOf(m[0]);
      classicMeasure(n, w);
      const last = CLASSIC_SQUEEZE[CLASSIC_SQUEEZE.length - 1];
      classicMeasure(n, w, last);
      const v = classicImg[n + '@' + w], vs = classicImg[n + '@' + w + '@' + last.scale + '/' + last.quality];
      const bytes = typeof v === 'number' ? v : assetSize(n) * 4 / 3;
      seen.set(n, { name: n, bytes, least: typeof vs === 'number' ? Math.min(vs, bytes) : bytes, exact: typeof v === 'number' });
    }
    const photos = /src="photos\//.test(html) ? 8 * 1024 : 0;
    const images = [...seen.values()].sort((x, y) => y.bytes - x.bytes);
    // As the add-in fills the room (keepForClassic): the text first, then the images, smallest
    // first, as many as fit; the ones that do not are left out there.
    const text = html.length * 1.3, room = CLASSIC_KB * 1024;
    let used = text + photos;
    const missing = [], squeezed = [];
    for (const x of images.slice().reverse()) {
      if (used + x.bytes <= room) used += x.bytes;
      // Made smaller there, step by step, till it fits (the last step's size here).
      else if (used + x.least <= room) { used += x.least; squeezed.push(x.name); }
      else missing.push(x.name);
    }
    return { kb: Math.round((text + images.reduce((n, x) => n + x.bytes, 0) + photos) / 1024), images, missing, squeezed, textFits: text <= room, text, photos, textKb: Math.round(text / 1024), photoKb: Math.round(photos / 1024) };
  }
  function classicKb(s) { return classicParts(s).kb; }
  // As classicParts, once every image is measured.
  async function classicPartsSettled(s) { classicParts(s); while (classicPending.size) await Promise.all([...classicPending]); return classicParts(s); }
  // The image work (preparing, fitting) one at a time, the editing held meanwhile (it would put
  // the old names back), the message kept up till it is done. The undo steps from before start
  // anew after it: the files they name are gone.
  let preparing = false;
  async function imageWork(what, fn) {
    if (preparing) { toast(t('Still working on the images…')); return; }
    preparing = true;
    document.body.classList.add('image-work'); $('source').readOnly = true;
    try { if (ddoc && ddoc.activeElement) ddoc.activeElement.blur(); } catch (e) { /* no design view */ }
    try {
      flushSource();
      toast(what, false, null, true);
      await fn();
    } finally {
      preparing = false;
      document.body.classList.remove('image-work'); $('source').readOnly = false;
      ghReset();
      schedulePreview();
    }
  }
  const defaultSigs = () => S.config.signatures.filter((x) => x.defaultNew || x.defaultReply);
  // The default signatures classic Outlook would put in without some of it.
  async function classicLeft() { const out = []; for (const x of defaultSigs()) { const p = await classicPartsSettled(x); if (p.missing.length || !p.textFits) out.push({ s: x, p }); } return out; }
  // What the used images weigh.
  function usedWeight() { const used = new Set(usedHtml().flatMap((h) => [...h.matchAll(/assets\/([^"'?#\s>]+)/g)].map((m) => m[1]))); return (S.assets || []).filter((x) => used.has(x.name)).reduce((n, x) => n + (x.size || 0), 0); }
  // The one way to make the signatures ready for classic Outlook (and lighter for every mail): the
  // HTML made lighter, the images optimized; still too large, the images of those signatures
  // smaller still; then said whether it fits now, or which image still does not.
  function prepareClassic() {
    return imageWork(t('Preparing the signatures for classic Outlook…'), async () => {
      const before = usedWeight();
      const saved = slimAll();
      await optimizeImages();
      let left = await classicLeft();
      if (left.some((x) => x.p.textFits)) { await optimizeImages({ strong: true, only: new Set(left.flatMap((x) => x.p.images.map((i) => i.name))) }); left = await classicLeft(); }
      const after = usedWeight();
      const did = [saved > 0 ? t('HTML {0} lighter', kb2(saved)) : '', after < before ? t('images {0} → {1}', kb2(before), kb2(after)) : ''].filter(Boolean).join(', ');
      if (!defaultSigs().length) toast(did ? t('Done: {0}. No signature is a default, so classic Outlook puts none in by itself.', did) : t('Nothing to make lighter. No signature is a default, so classic Outlook puts none in by itself.'));
      else if (!left.length) toast(did ? t('Ready for classic Outlook: {0}.', did) : t('Ready for classic Outlook: nothing needed changing.'));
      else {
        const textOnly = left.every((x) => !x.p.textFits);
        toast(t('Still too large for classic Outlook: {0}. Missing there: {1}. Show these images narrower, or use fewer; Insert in the pane puts in everything.', left.map((x) => x.s.name).join(', '), [...new Set(left.flatMap((x) => x.p.missing))].join(', ') || t('the text')), true,
          textOnly ? null : { label: t('Images small enough to fit'), run: fitClassic });
      }
    });
  }
  // The further way, when preparing is not enough: the images of the default signatures that
  // classic Outlook would still go without made as small in KB as it takes for all of it to fit
  // (narrower drawn, a JPEG at a lower quality, a PNG with fewer colours), at a cost in
  // sharpness that shows. Every mail gets them so, as with optimizing.
  function fitClassic() {
    return imageWork(t('Making the images small enough for classic Outlook…'), async () => {
      const room = CLASSIC_KB * 1024 * 0.95;
      let todo = await classicLeft();
      // The images made smaller, by their name now (one made smaller twice is one).
      const done = new Set();
      // Each round a share of what is left for the images (the text and a photo first), each image
      // its part of it by its size now; still not enough, the next round asks less.
      for (let round = 0; round < 4 && todo.some((x) => x.p.textFits); round++) {
        const goal = new Map();
        for (const { p } of todo) {
          if (!p.textFits) continue;
          const avail = room - p.text - p.photos, total = p.images.reduce((n, x) => n + x.bytes, 0);
          if (!total) continue;
          const share = Math.max(0.02, avail / total) * Math.pow(0.8, round);
          for (const x of p.images) goal.set(x.name, Math.min(goal.has(x.name) ? goal.get(x.name) : Infinity, Math.max(600, x.bytes * share * 3 / 4)));
        }
        let any = false;
        for (const [name, bytes] of goal) {
          const r = await squeezeTo(name, bytes);
          if (r) { done.delete(name); done.add(r.name); any = true; }
        }
        if (any) reopenChanged();
        todo = await classicLeft();
        if (!any) break;
      }
      if (!todo.length) toast(done.size ? t('Ready for classic Outlook: {0} images made smaller to fit.', done.size) : t('It fits classic Outlook already: nothing made smaller.'));
      else if (todo.every((x) => !x.p.textFits)) toast(t('Still too large for classic Outlook: {0}. Missing there: {1}. Show these images narrower, or use fewer; Insert in the pane puts in everything.', todo.map((x) => x.s.name).join(', '), t('the text')), true);
      else toast(t('Even as small as they go, still too large for classic Outlook: {0}. Use fewer images there; Insert in the pane puts in everything.', todo.map((x) => x.s.name).join(', ')), true);
    });
  }
  // The defaults (what goes in by itself) too large for that.
  function classicHeavy() {
    return defaultSigs().map((s) => ({ s, kb: classicKb(s), p: classicParts(s) })).filter((x) => x.kb > CLASSIC_KB && (x.p.missing.length || !x.p.textFits));
  }
  // A signature's HTML made lighter, nothing that shows changed: what Word, Outlook and browsers
  // leave in it when copied (mso-…, -webkit-… and other default styles, lang, empty spans), and
  // runs of spaces and line breaks. Answers the HTML and how many bytes it saved.
  const DEAD_STYLE = /^(mso-|-webkit-|orphans$|widows$|font-variant-ligatures$|font-variant-caps$|text-decoration-thickness$|text-decoration-style$|text-decoration-color$|caret-color$|font-kerning$)/i;
  function slimHtml(html) {
    const tpl = document.createElement('template'); tpl.innerHTML = String(html || '');
    const all = [...tpl.content.querySelectorAll('*')];
    for (const n of all) {
      n.removeAttribute('lang');
      if (n.hasAttribute('style')) {
        const keep = String(n.getAttribute('style')).split(';').map((d) => d.trim()).filter(Boolean).filter((d) => {
          const p = d.split(':')[0].trim();
          return !DEAD_STYLE.test(p) && !/^\s*(initial|unset)\s*$/i.test(d.slice(d.indexOf(':') + 1));
        });
        if (keep.length) n.setAttribute('style', keep.join(';')); else n.removeAttribute('style');
      }
    }
    // A span with nothing on it says nothing: its content in its place.
    for (const n of [...tpl.content.querySelectorAll('span')]) if (!n.attributes.length) n.replaceWith(...n.childNodes);
    let out = tpl.innerHTML;
    // Whitespace: between block tags it shows nothing; elsewhere one space is what shows.
    out = out.replace(/>\s+<(\/?(?:table|tbody|thead|tr|td|th|p|div|ul|ol|li|br)\b)/gi, '><$1').replace(/(<\/?(?:table|tbody|thead|tr|td|th|p|div|ul|ol|li)\b[^>]*>)\s+/gi, '$1').replace(/[ \t\r\n]{2,}/g, ' ');
    return { html: out, saved: String(html || '').length - out.length };
  }
  // Every signature and building block made lighter (one step to undo).
  function slimAll() {
    // What was typed in the HTML tab first; after, the tab shows the lighter HTML (else reopening
    // the signature would take the old text in the tab for a change).
    flushSource();
    let saved = 0;
    for (const s of S.config.signatures) { const r = slimHtml(s.html); if (r.saved > 0) { s.html = r.html; saved += r.saved; } }
    for (const b of S.config.blocks || []) { const r = slimHtml(b.html); if (r.saved > 0) { b.html = r.html; saved += r.saved; } }
    if (saved > 0) reopenChanged();
    return saved;
  }
