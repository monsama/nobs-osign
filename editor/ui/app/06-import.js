  // ---------------------------------------------------------------------------------------------
  // Import: templates (.htm) with their images, from a folder or picked files. CI-Sign keeps them
  // as UTF-16 files that load "images/<name>"; classic Outlook as "<name>_files/<name>".
  // ---------------------------------------------------------------------------------------------
  let impItems = [], impImages = new Map(), impBlocks = new Map(), impPhotos = [], impPhotosOn = false;
  const IMG_EXT = /\.(png|jpe?g|gif)$/i;
  // CI-Sign's text blocks: {##Name}…{end ##Name} in a file of their own (__intern.htm), each with
  // the days it is shown ({dats 01.12.2026}{date 31.12.2026}), used in a template as {##Name}.
  const BLOCK_DEF = /\{##([^}\s]+)\}([\s\S]*?)\{end ##\1\}/g;
  const BLOCK_REF = /\{##([^}\s]+)\}/g;
  const ciDay = (v) => { const m = /^(\d{1,2})\.(\d{1,2})\.(\d{4})$/.exec(String(v || '').trim()); return m ? m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2) : ''; };
  function readBlocks(text) {
    for (const m of text.matchAll(BLOCK_DEF)) {
      let html = m[2], from = '', until = '';
      html = html.replace(/\{dats\s+([^}]*)\}/i, (a, d) => { from = ciDay(d); return ''; }).replace(/\{date\s+([^}]*)\}/i, (a, d) => { until = ciDay(d); return ''; });
      // A block whose last day is past is shown no more: it stays out (CI-Sign keeps its samples).
      impBlocks.set(m[1].toLowerCase(), { name: m[1], html: html.trim(), from, until, expired: !!until && until < core.today() });
    }
  }

  function relativeSrcs(doc) {
    return [...doc.querySelectorAll('img[src]')].map((i) => i.getAttribute('src')).filter((src) => !/^(https?:|cid:|data:)/i.test(src));
  }

  async function readImportFiles(files) {
    impItems = []; impImages = new Map(); impBlocks = new Map();
    const list = [...files];
    // The same image name in several places (CI-Sign keeps copies in ci-sign\, in each
    // AutoBackup-… and in its update folder): the one nearest the top, outside the copies, wins.
    const rank = (f) => { const p = (f.webkitRelativePath || f.name).replace(/\\/g, '/'); return (/\/(AutoBackup-[^/]*|ci-sign-update|editor|Templates)\//i.test('/' + p) ? 100 : 0) + p.split('/').length; };
    list.filter((f) => IMG_EXT.test(f.name)).sort((a, b) => rank(a) - rank(b)).forEach((f) => { if (!impImages.has(f.name.toLowerCase())) impImages.set(f.name.toLowerCase(), f); });
    for (const f of list.filter((x) => /\.html?$/i.test(x.name))) {
      // A folder holds more than templates (CI-Sign's own editor, for one): those stay unticked.
      const path = f.webkitRelativePath || f.name;
      const text = core.decodeHtmlFile(new Uint8Array(await f.arrayBuffer()));
      const d = new DOMParser().parseFromString(text, 'text/html');
      const body = d.body ? d.body.innerHTML : text;
      const name = f.name.replace(/\.html?$/i, '');
      const fields = core.placeholdersIn(body);
      const blocks = /\{end ##[^}]+\}/.test(text);
      if (blocks) readBlocks(body);
      const refs = blocks ? [] : [...new Set([...body.matchAll(BLOCK_REF)].map((m) => m[1]))];
      const srcs = relativeSrcs(d);
      const perPerson = srcs.filter((x) => x.indexOf('@@') >= 0);
      const imgs = srcs.filter((x) => x.indexOf('@@') < 0).map((x) => decodeURIComponent(x.split(/[\\/]/).pop()));
      const missing = imgs.filter((x) => !impImages.has(x.toLowerCase()));
      const looksLikeSig = fields.length > 0 && !blocks && !/\\(editor|ci-sign-update|AutoBackup-[^\\/]*|Templates)[\\/]/i.test(path.replace(/\//g, '\\'));
      impItems.push({ file: f, path, name, body, fields, blocks, refs, perPerson, imgs, missing, checked: looksLikeSig });
    }
    impItems.sort((a, b) => (b.checked - a.checked) || a.name.localeCompare(b.name));
    // Photos per person: with a template that uses one (images/@@sAMAccountName.jpg), the images
    // no template names, of that type, named like a user (mustera.jpg, not Newsletter-2.jpg).
    const per = impItems.filter((x) => x.checked && x.perPerson.length).map((x) => x.perPerson[0]);
    impPhotos = [];
    if (per.length) {
      const ext = (per[0].split('.').pop() || 'jpg').toLowerCase();
      const used = new Set(impItems.flatMap((x) => x.imgs.map((n) => n.toLowerCase())));
      impPhotos = [...impImages.entries()].filter(([k]) => !used.has(k) && k.endsWith('.' + ext) && /^[a-z][a-z0-9._]{1,40}$/i.test(k.slice(0, -ext.length - 1))).map(([, f]) => f);
    }
    // Off unless ticked: photos in the folder are used at once, without publishing.
    impPhotosOn = false;
    // The same template twice (CI-Sign keeps a copy next to the program): the first one is ticked.
    const seen = new Set();
    impItems.forEach((it) => { const k = it.name.toLowerCase(); if (it.checked && seen.has(k)) it.checked = false; if (it.checked) seen.add(k); });
    renderImportList();
  }

  function renderImportList() {
    const ul = $('impList'); ul.textContent = '';
    impItems.forEach((it) => {
      const box = el('input', { type: 'checkbox' }); box.checked = it.checked;
      box.addEventListener('change', () => { it.checked = box.checked; $('impGo').disabled = !impItems.some((x) => x.checked); });
      const exists = S.config.signatures.some((x) => x.name.toLowerCase() === it.name.toLowerCase());
      const meta = [];
      if (it.blocks) meta.push(el('span', { class: 'warn' }, t('CI-Sign text blocks ({##…}), not a signature. ')));
      meta.push(it.fields.length ? t('{0} fields', it.fields.length) : t('no fields'));
      meta.push(t(' · {0} image(s)', it.imgs.length));
      if (it.missing.length) meta.push(el('span', { class: 'warn' }, t(' · not picked: {0}', [...new Set(it.missing)].join(', '))));
      if (it.refs.length) {
        const known = it.refs.filter((r) => impBlocks.has(r.toLowerCase())), unknown = it.refs.filter((r) => !impBlocks.has(r.toLowerCase()));
        const timed = known.filter((r) => { const b = impBlocks.get(r.toLowerCase()); return (b.from || b.until) && !b.expired; });
        const old = known.filter((r) => impBlocks.get(r.toLowerCase()).expired);
        if (old.length) meta.push(t(' · past text blocks left out: {0}', old.join(', ')));
        if (known.length) meta.push(t(' · text blocks {0}', known.join(', ')) + (timed.length ? t(' (with days: a campaign each)') : ''));
        if (unknown.length) meta.push(el('span', { class: 'warn' }, t(' · text blocks not picked: {0} (pick the file with them, e.g. __intern.htm)', unknown.join(', '))));
      }
      if (it.perPerson.length) meta.push(t(' · photo per person {0} → photos\\{1}', it.perPerson[0], it.perPerson[0].split(/[\\/]/).pop()));
      if (exists) meta.push(t(' · replaces the signature of the same name'));
      ul.append(el('li', { class: it.checked ? '' : 'skip' }, box, el('div', null, el('div', { class: 'imp-name' }, it.name), el('div', { class: 'imp-meta' }, el('span', null, it.path), el('br'), ...meta))));
    });
    if (impPhotos.length) {
      const box = el('input', { type: 'checkbox' }); box.checked = impPhotosOn;
      box.addEventListener('change', () => { impPhotosOn = box.checked; });
      const names = impPhotos.map((f) => f.name);
      ul.append(el('li', { class: impPhotosOn ? '' : 'skip' }, box, el('div', null,
        el('div', { class: 'imp-name' }, t('Photos per person: {0} into photos\\ in the OSign folder', impPhotos.length)),
        el('div', { class: 'imp-meta' }, names.slice(0, 12).join(', ') + (names.length > 12 ? ' …' : '') + t(' · used at once, without publishing; a photo already there of the same name is replaced')))));
    }
    const n = impItems.length;
    $('impNote').textContent = n ? t('{0} template file(s) found, {1} image(s) picked.', n, impImages.size) : t('No .htm files among what was picked.');
    $('impGo').disabled = !impItems.some((x) => x.checked);
  }

  async function uploadImage(file) {
    file = await shrinkUpload(file);
    const data = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = rej; r.readAsDataURL(file); });
    const r = await api('assets/upload', { name: file.name, data });
    S.assets = S.assets.filter((a) => a.name !== r.name).concat([{ name: r.name, size: r.size }]).sort((a, b) => a.name.localeCompare(b.name));
    return r.name;
  }

  async function runImport() {
    $('impErr').textContent = ''; $('impGo').disabled = true;
    const uploaded = new Map(); // picked file name (lower case) -> asset name
    let first = null, count = 0;
    const campaigns = [];
    try {
      const imported = [];
      for (const it of impItems.filter((x) => x.checked)) {
        const t = inertTemplate(sanitize(it.body));
        for (const img of t.content.querySelectorAll('img[src]')) {
          const src = img.getAttribute('src');
          if (/^(https?:|cid:|data:)/i.test(src)) continue;
          // One image per person ("images/@@sAMAccountName.jpg"): from the photos folder.
          if (src.indexOf('@@') >= 0) { img.setAttribute('src', 'photos/' + src.split(/[\\/]/).pop()); continue; }
          const base = decodeURIComponent(src.split(/[\\/]/).pop()), key = base.toLowerCase();
          if (impImages.has(key) && !uploaded.has(key)) uploaded.set(key, await uploadImage(impImages.get(key)));
          // Not picked: it points where the image will be once uploaded under that name.
          img.setAttribute('src', 'assets/' + (uploaded.get(key) || core.assetName(base)));
        }
        // Text blocks: one without days goes in as it is; one with days becomes a campaign of its
        // own (the template with that block, on those days), and the template itself goes without.
        const html = t.innerHTML.trim();
        const timed = [];
        const plain = html.replace(BLOCK_REF, (all, name) => {
          const b = impBlocks.get(name.toLowerCase());
          if (!b) return all;
          if (b.expired) return '';
          if (b.from || b.until) { timed.push(b); return ''; }
          return sanitize(b.html);
        });
        const put = (name, body, from, until) => {
          let s = S.config.signatures.find((x) => x.name.toLowerCase() === name.toLowerCase());
          if (!s) { s = core.newSignature(newId(name), name); S.config.signatures.push(s); }
          s.html = body; s.validFrom = from || ''; s.validUntil = until || '';
          imported.push(s);
          return s;
        };
        const s = put(it.name, plain);
        for (const b of timed) {
          const c = put(it.name + ' – ' + b.name, html.replace(BLOCK_REF, (all, name) => {
            const x = impBlocks.get(name.toLowerCase());
            return !x ? all : x === b || !(x.from || x.until || x.expired) ? sanitize(x.html) : '';
          }), b.from, b.until);
          c.audience = JSON.parse(JSON.stringify(s.audience));
          campaigns.push([c, s]);
        }
        first = first || s; count++;
      }
      // With no default yet, the one called Standard (or the first) becomes it.
      if (first && !S.config.signatures.some((x) => x.defaultNew)) {
        const std = S.config.signatures.find((x) => /^standard$/i.test(x.name)) || first;
        std.defaultNew = true; std.defaultReply = true;
      }
      // A campaign stands in for its template while it runs: the same defaults.
      for (const [c, base] of campaigns) { c.defaultNew = !!base.defaultNew; c.defaultReply = !!base.defaultReply; }
      // Placeholders of CI-Sign beyond the fields there are (@@telephoneNumber, @@company …): fields.
      const newFields = addMissingFields(imported);
      changed();
      $('dlgImport').close();
      if (first) openSig(first.id);
      let photosDone = 0;
      if (impPhotosOn) {
        for (let f of impPhotos) {
          f = await shrinkUpload(f, PHOTO_MAX_W);
          const data = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = rej; r.readAsDataURL(f); });
          const r = await api('photos/upload', { name: f.name, data }).catch((e) => ({ error: e.message }));
          if (!r.error) photosDone++;
        }
      }
      toast((uploaded.size ? t('{0} template(s) imported, with {1} image(s).', count, uploaded.size) : t('{0} template(s) imported.', count))
        + (impPhotosOn ? ' ' + t('{0} of {1} photo(s) uploaded.', photosDone, impPhotos.length) : '')
        + (newFields.length ? ' ' + t('Fields added: {0}. Check their AD attribute and Graph source under Fields.', newFields.map((k) => '@@' + k).join(', ')) : '')
        + ' ' + t('The signatures go live only when you publish.'));
    } catch (e) { $('impErr').textContent = t(e.message); $('impGo').disabled = false; }
  }

  $('btnImportOpen').addEventListener('click', () => {
    impItems = []; impImages = new Map(); impPhotos = []; impPhotosOn = false; renderImportList(); $('impNote').textContent = ''; $('impErr').textContent = '';
    $('dlgImport').showModal();
  });
  ['impFiles', 'impFolder'].forEach((id) => $(id).addEventListener('change', async () => {
    const files = $(id).files; if (!files || !files.length) return;
    try { await readImportFiles(files); } catch (e) { $('impErr').textContent = t(e.message); }
    $(id).value = '';
  }));
  $('impGo').addEventListener('click', (e) => { e.preventDefault(); runImport(); });

