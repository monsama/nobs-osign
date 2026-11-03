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
      src.addEventListener('input', () => { b.html = src.value; changed({ list: false }); blockPreview(frame, b); });
      const del = el('button', { type: 'button', class: 'ghost small danger', 'aria-label': t('Remove the building block {0}', b.name), title: t('Remove'), onclick: async () => {
        const u = blockUses(b.key);
        if (u.length && !(await ask(t('{{{0}}} is in: {1}. Remove it anyway? It stays there as typed until you take it out.', b.key, u.join(', ')), t('Remove'), true))) return;
        S.config.blocks.splice(i, 1); changed(); renderBlocks();
        toast(t('Building block {0} removed.', b.name), false, { label: t('Undo'), run: () => { if (!S.config.blocks.some((x) => x.key === b.key)) { S.config.blocks.splice(Math.min(i, S.config.blocks.length), 0, b); changed(); renderBlocks(); } } });
      } }, '✕');
      const card = el('div', { class: 'block-card', 'data-block': b.key },
        el('div', { class: 'block-head' }, name, el('span', { class: 'block-key' }, '{{', key, '}}'), el('span', { class: 'spacer' }), del),
        el('div', { class: 'block-body' }, src, frame), used);
      blockPreview(frame, b);
      return card;
    }));
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
