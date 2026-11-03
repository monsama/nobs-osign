  // ---------------------------------------------------------------------------------------------
  // Building blocks: a part several signatures share (a disclaimer, a banner, the address), kept
  // once in config.blocks and put into a signature as {{key}} (core.expandBlocks).
  // ---------------------------------------------------------------------------------------------
  const BLOCK_STARTER = {
    disclaimer: '<p style="margin:0;font-family:Arial,sans-serif;font-size:7.5pt;color:#888888">Diese E-Mail kann vertrauliche Informationen enthalten. Wenn Sie nicht der richtige Empfänger sind, informieren Sie bitte den Absender und löschen Sie diese E-Mail.</p>',
    address: '<p style="margin:0;font-family:Arial,sans-serif;font-size:9pt;color:#555555">@@company · @@streetAddress · @@postalCode @@l</p>'
  };
  const blockUses = (key) => S.config.signatures.filter((s) => core.blocksIn(s.html).some((k) => k.toLowerCase() === key.toLowerCase())).map((s) => s.name);
  function blockKey(name) {
    const base = String(name || 'block').normalize('NFKD').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-+|-+$/g, '').toLowerCase().slice(0, 30) || 'block';
    let k = /^[a-z]/.test(base) ? base : 'b-' + base, i = 2;
    const taken = (x) => S.config.blocks.some((b) => b.key.toLowerCase() === x);
    const first = k; while (taken(k)) k = first + '-' + i++;
    return k;
  }
  function blockPreview(frame, b) {
    const values = {};
    S.config.fields.forEach((f) => { values[f.key] = f.sample || ''; });
    const html = core.render(core.compile(sanitize(b.html), scratchDoc), values, { assetBase: '/' });
    frame.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:10px 12px;font-family:Arial,sans-serif;font-size:10pt;color:#222}p{margin:0}table{border-collapse:collapse}</style></head><body>' + html + '</body></html>';
  }
  function renderBlocks() {
    const list = $('blockList');
    // Empty: the button in the empty state is the one; below the list otherwise.
    $('btnAddBlock').parentNode.hidden = !S.config.blocks.length;
    if (!S.config.blocks.length) {
      list.replaceChildren(emptyState('blocks', t('No building blocks yet.'), t('A disclaimer or a banner in many signatures: made once here, changed in one place.'), [t('+ Add building block'), () => $('btnAddBlock').click()]));
      return;
    }
    list.replaceChildren(...S.config.blocks.map((b, i) => {
      const frame = el('iframe', { class: 'block-pv', sandbox: 'allow-same-origin', tabindex: '-1', title: t('Preview') });
      const name = el('input', { value: b.name, 'aria-label': t('Name') });
      const key = el('input', { value: b.key, class: 'mono', 'aria-label': t('Its name in a signature') });
      const src = el('textarea', { class: 'mono', rows: '5', spellcheck: 'false', 'aria-label': t('HTML') });
      src.value = b.html;
      const uses = blockUses(b.key);
      const used = el('small', { class: 'hint' }, uses.length ? t('In: {0}', uses.join(', ')) : t('Not in any signature yet.'));
      name.addEventListener('input', () => { b.name = name.value; changed({ list: false, preview: false }); });
      key.addEventListener('change', () => {
        const nk = key.value.trim().replace(/^\{\{|\}\}$/g, '');
        if (!/^[A-Za-z][A-Za-z0-9_-]{0,39}$/.test(nk)) { toast(t('A building block\'s name is letters, digits, - and _, starting with a letter.'), true); key.value = b.key; return; }
        if (S.config.blocks.some((x) => x !== b && x.key.toLowerCase() === nk.toLowerCase())) { toast(t('{{{0}}} is already a building block.', nk), true); key.value = b.key; return; }
        // The signatures that use it go along.
        const re = new RegExp('\\{\\{' + b.key.replace(/[-]/g, '\\-') + '\\}\\}', 'gi');
        S.config.signatures.forEach((s) => { s.html = s.html.replace(re, '{{' + nk + '}}'); });
        b.key = nk; changed(); renderBlocks();
      });
      // Edited as a signature is: the design view (fields as chips, bold, a link …) or its HTML.
      const ed = blockEditor(b, () => { src.value = b.html; blockPreview(frame, b); });
      src.hidden = true;
      src.addEventListener('input', () => { b.html = src.value; changed({ list: false }); blockPreview(frame, b); });
      const modeBtn = el('button', { type: 'button', class: 'ghost small block-mode', onclick: () => {
        const toHtml = src.hidden;
        if (toHtml) { src.value = b.html; } else { ed.load(b.html); }
        src.hidden = !toHtml; ed.box.hidden = toHtml; modeBtn.textContent = toHtml ? t('Design') : t('HTML');
      } }, t('HTML'));
      const del = el('button', { type: 'button', class: 'ghost small danger', 'aria-label': t('Remove the building block {0}', b.name), title: t('Remove'), onclick: async () => {
        const u = blockUses(b.key);
        if (u.length && !(await ask(t('{{{0}}} is in: {1}. Remove it anyway? It stays there as typed until you take it out.', b.key, u.join(', ')), t('Remove'), true))) return;
        S.config.blocks.splice(i, 1); changed(); renderBlocks();
        toast(t('Building block {0} removed.', b.name), false, { label: t('Undo'), run: () => { if (!S.config.blocks.some((x) => x.key === b.key)) { S.config.blocks.splice(Math.min(i, S.config.blocks.length), 0, b); changed(); renderBlocks(); } } });
      } }, '✕');
      const card = el('div', { class: 'block-card', 'data-block': b.key },
        el('div', { class: 'block-head' }, name, el('span', { class: 'block-key' }, '{{', key, '}}'), el('span', { class: 'spacer' }), del),
        el('div', { class: 'block-body' }, el('div', { class: 'block-edit' }, el('div', { class: 'block-tools' }, ed.tools, el('span', { class: 'spacer' }), modeBtn), ed.box, src), frame), used);
      blockPreview(frame, b);
      return card;
    }));
  }
  // A small design view of its own for a building block: the signature editor's look and chips.
  function blockEditor(b, onChange) {
    const frame = el('iframe', { class: 'block-ed', title: t('Building block') });
    const box = el('div', { class: 'block-ed-box' }, frame);
    let doc = null;
    const htmlOf = () => {
      const copy = doc.body.cloneNode(true);
      copy.querySelectorAll('span.osf[data-f]').forEach((c) => c.replaceWith(doc.createTextNode('@@' + c.dataset.f)));
      copy.querySelectorAll('span.osb[data-b]').forEach((c) => c.replaceWith(doc.createTextNode('{{' + c.dataset.b + '}}')));
      return copy.innerHTML;
    };
    const input = () => { b.html = sanitize(htmlOf()); changed({ list: false }); onChange(); };
    const load = (html) => { if (!doc) return; doc.body.innerHTML = sanitize(html || ''); chipify(doc.body); };
    frame.addEventListener('load', () => {
      doc = frame.contentDocument;
      doc.body.contentEditable = 'true';
      try { doc.execCommand('styleWithCSS', false, true); doc.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) { /* older browsers */ }
      load(b.html);
      doc.addEventListener('input', input);
      doc.addEventListener('paste', (e) => {
        const h = e.clipboardData && e.clipboardData.getData('text/html');
        if (!h) return;
        e.preventDefault();
        const body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(h);
        doc.execCommand('insertHTML', false, sanitize((body ? body[1] : h).replace(/<!--StartFragment-->|<!--EndFragment-->/g, '')));
        chipify(doc.body); input();
      });
    }, { once: true });
    frame.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><base href="/"><style>' + EDITOR_CSS + 'body{margin:10px 12px;min-height:60px}</style></head><body></body></html>';
    const run = (cmd, val) => { if (!doc) return; frame.contentWindow.focus(); doc.execCommand(cmd, false, val === undefined ? null : val); input(); };
    const tbtn = (icon, title, fn) => { const x = el('button', { type: 'button', class: 'ghost small', title: t(title), 'aria-label': t(title) }); x.innerHTML = svg(icon); x.addEventListener('mousedown', (e) => e.preventDefault()); x.addEventListener('click', fn); return x; };
    const fields = el('select', { class: 'field-pick', 'aria-label': t('Insert a field') });
    fillFieldSelect(fields);
    [...fields.querySelectorAll('optgroup')].filter((g) => g.label === t('Building blocks')).forEach((g) => g.remove());
    fields.addEventListener('focus', () => { fillFieldSelect(fields); [...fields.querySelectorAll('optgroup')].filter((g) => g.label === t('Building blocks')).forEach((g) => g.remove()); });
    fields.addEventListener('change', () => { const k = fields.value; fields.value = ''; if (k) run('insertHTML', fieldChip(doc, k).outerHTML); });
    const images = el('select', { 'aria-label': t('Image'), title: t('Image') });
    const fillImages = () => images.replaceChildren(el('option', { value: '' }, t('Image')), ...S.assets.map((a) => el('option', { value: a.name }, a.name)));
    fillImages(); images.addEventListener('focus', fillImages);
    images.addEventListener('change', () => { const n = images.value; images.value = ''; if (n) run('insertHTML', '<img src="assets/' + encodeURIComponent(n).replace(/%2F/gi, '/') + '" alt="" style="border:0">'); });
    const link = async () => {
      const sel = doc && doc.getSelection(), range = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
      const url = await askFor(t('Link address (https://…, mailto:…, tel:…). Fields work here too: mailto:@@mail'), 'https://', t('Link'));
      if (!url || url === 'https://') return;
      if (!/^(https?:|mailto:|tel:)/i.test(url.trim())) { toast(t('A link starts with https://, mailto: or tel:.'), true); return; }
      frame.contentWindow.focus();
      if (range) { const s2 = doc.getSelection(); s2.removeAllRanges(); s2.addRange(range); }
      run('createLink', url.trim());
    };
    const tools = el('div', { class: 'block-btns' }, fields, images,
      tbtn('bold', 'Bold (Ctrl+B)', () => run('bold')), tbtn('italic', 'Italic (Ctrl+I)', () => run('italic')), tbtn('underline', 'Underline (Ctrl+U)', () => run('underline')),
      tbtn('link', 'Link', link), tbtn('unlink', 'Remove link', () => run('unlink')), tbtn('clear', 'Clear formatting', () => run('removeFormat')));
    return { box, tools, load };
  }
  $('btnAddBlock').addEventListener('click', () => {
    const first = !S.config.blocks.length;
    const name = first ? 'Disclaimer' : t('New building block');
    const b = { key: blockKey(name), name, html: first ? BLOCK_STARTER.disclaimer : BLOCK_STARTER.address };
    S.config.blocks.push(b);
    if (!first) addMissingFields([{ html: b.html }]);
    changed({ list: false }); renderBlocks();
    const cards = $('blockList').querySelectorAll('.block-card');
    cards[cards.length - 1].querySelector('input').select();
  });
