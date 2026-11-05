  // ---------------------------------------------------------------------------------------------
  // Editor tools: the house colours, spacing, and a bar for what was clicked in the design (a link,
  // a table cell; the image's bar is in 02-design). One bar at a time.
  // ---------------------------------------------------------------------------------------------

  // A colour as #rrggbb (from #rgb, #rrggbb or rgb(r, g, b)), or ''.
  function hexColour(c) {
    const v = String(c || '').trim().toLowerCase();
    let m = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/.exec(v);
    if (m) return '#' + m[1] + m[1] + m[2] + m[2] + m[3] + m[3];
    if (/^#[0-9a-f]{6}$/.test(v)) return v;
    m = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(v);
    if (m && (m[4] === undefined || Number(m[4]) > 0)) return '#' + [m[1], m[2], m[3]].map((n) => Math.min(255, Number(n)).toString(16).padStart(2, '0')).join('');
    return '';
  }
  // The text colours a signature uses (color: in its styles, <font color>), most used first.
  function coloursIn(html) {
    const n = new Map(), re = /(?:[^-]color\s*:\s*|\bcolor\s*=\s*["']?)(#[0-9a-f]{3,6}\b|rgba?\([^)]*\))/gi;
    let m;
    while ((m = re.exec(' ' + String(html || '')))) { const h = hexColour(m[1]); if (h) n.set(h, (n.get(h) || 0) + 1); }
    return [...n.entries()].sort((a, b) => b[1] - a[1]).map((e) => e[0]);
  }
  const houseColours = () => (S.config && Array.isArray(S.config.colors) ? S.config.colors : []);

  // The house colours next to the text colour: the firm's colours one click away (kept in the
  // draft for every signature), with the ones this signature uses below them.
  function paletteTool(colourInput) {
    const pop = el('div', { class: 'pal', hidden: true, role: 'dialog', 'aria-label': t('House colours') });
    const btn = el('button', { type: 'button', class: 'pal-btn', title: t('House colours'), 'aria-label': t('House colours'), 'aria-haspopup': 'true' });
    btn.innerHTML = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 9l6 6 6-6"/></svg>';
    btn.addEventListener('mousedown', (e) => e.preventDefault());
    const use = (hex) => { colourInput.value = hex; exec('foreColor', hex); close(); };
    const swatch = (hex, removable) => {
      const b = el('button', { type: 'button', class: 'sw', title: hex + (removable ? ' - ' + t('× removes it from the house colours') : ''), 'aria-label': hex, style: 'background:' + hex });
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', () => use(hex));
      if (removable) {
        const x = el('span', { class: 'x', title: t('Remove from the house colours'), 'aria-hidden': 'true' }, '×');
        x.addEventListener('click', (e) => { e.stopPropagation(); S.config.colors = houseColours().filter((c) => c !== hex); changed({ list: false, preview: false }); fill(); });
        b.append(x);
      }
      return b;
    };
    function fill() {
      const house = houseColours(), s = sig(), used = s ? coloursIn(s.html).filter((c) => !house.includes(c)) : [];
      const add = el('button', { type: 'button', class: 'ghost small pal-add', title: t('The colour now in the colour box, for all signatures') }, t('+ Add {0} to the house colours', colourInput.value));
      add.addEventListener('mousedown', (e) => e.preventDefault());
      add.disabled = house.includes(colourInput.value) || house.length >= 16;
      add.addEventListener('click', () => { S.config.colors = house.concat([colourInput.value]); changed({ list: false, preview: false }); fill(); });
      pop.replaceChildren(
        el('h4', null, t('House colours')),
        house.length ? el('div', { class: 'pal-row' }, house.map((c) => swatch(c, true))) : el('p', { class: 'pal-none' }, t('None yet: pick one in the colour box, then add it here.')),
        used.length ? el('h4', null, t('In this signature')) : null,
        el('div', { class: 'pal-row' }, used.slice(0, 12).map((c) => swatch(c, false))),
        add);
    }
    function close() { pop.hidden = true; btn.setAttribute('aria-expanded', 'false'); }
    btn.addEventListener('click', () => { if (!pop.hidden) { close(); return; } fill(); pop.hidden = false; btn.setAttribute('aria-expanded', 'true'); });
    document.addEventListener('mousedown', (e) => { if (!pop.hidden && !wrap.contains(e.target)) close(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !pop.hidden) close(); });
    // A click or Escape in the signature closes it too (the editor is a frame of its own).
    if (designReady) designReady.then(() => { ddoc.addEventListener('mousedown', () => { if (!pop.hidden) close(); }); ddoc.addEventListener('keydown', (e) => { if (e.key === 'Escape' && !pop.hidden) close(); }); });
    const wrap = el('span', { class: 'pal-wrap' }, btn, pop);
    return wrap;
  }

  // Spacing for the lines selected: the line height, the space below a paragraph (in a cell: below
  // the cell's content).
  function spacingTool() {
    const sel = el('select', { title: t('Spacing'), 'aria-label': t('Spacing') }, el('option', { value: '' }, t('Spacing')),
      el('optgroup', { label: t('Line spacing') }, ['1', '1.15', '1.3', '1.5', '2'].map((v) => el('option', { value: 'lh:' + v }, Number(v).toLocaleString(document.documentElement.lang || undefined, { minimumFractionDigits: v.includes('.') ? v.split('.')[1].length : 0 })))),
      el('optgroup', { label: t('Space below') }, ['0', '4', '8', '12', '20'].map((v) => el('option', { value: 'mb:' + v }, v + ' px'))));
    sel.addEventListener('change', () => { const v = sel.value; sel.value = ''; if (v) setSpacing(v.slice(0, 2), v.slice(3)); });
    return sel;
  }
  const BLOCK_SEL = 'p, div, li, td, th, h1, h2, h3, h4, h5, h6, blockquote';
  // The paragraphs (or cells) the selection touches; text right in the body gets a paragraph first.
  function blocksInSelection() {
    const sel = ddoc.getSelection();
    if (!sel.rangeCount || !ddoc.body.contains(sel.getRangeAt(0).commonAncestorContainer)) return [];
    const r = sel.getRangeAt(0);
    const blockOf = (n) => { const e = n.nodeType === 1 ? n : n.parentNode; const b = e && e.closest(BLOCK_SEL); return b && ddoc.body.contains(b) && b !== ddoc.body ? b : null; };
    const out = new Set();
    const add = (n) => { const b = blockOf(n); if (b) out.add(b); };
    add(r.startContainer); add(r.endContainer);
    if (!r.collapsed) {
      const w = ddoc.createTreeWalker(r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentNode, NodeFilter.SHOW_TEXT);
      while (w.nextNode()) if (r.intersectsNode(w.currentNode) && w.currentNode.nodeValue.trim()) add(w.currentNode);
    }
    if (!out.size) {
      // Text right in the body: a paragraph made of it, then that one.
      ddoc.execCommand('formatBlock', false, 'div');
      const s2 = ddoc.getSelection();
      if (s2.rangeCount) add(s2.getRangeAt(0).startContainer);
    }
    // The innermost only: a paragraph in a cell, not the cell as well.
    const list = [...out];
    return list.filter((b) => !list.some((o) => o !== b && b.contains(o)));
  }
  // Line spacing (lh) or the space below (mb) on the paragraphs selected; one step to undo.
  function setSpacing(kind, v) {
    if (S.mode !== 'design' || !ddoc) return;
    design.contentWindow.focus();
    const blocks = blocksInSelection();
    if (!blocks.length) { toast(t('Put the cursor in the text first.'), true); return; }
    blocks.forEach((b) => {
      const cell = /^T[DH]$/.test(b.nodeName);
      if (kind === 'lh') b.style.lineHeight = v === '1' ? '' : v;
      else if (cell) b.style.paddingBottom = v === '0' ? '' : v + 'px';
      else b.style.marginBottom = v === '0' ? '' : v + 'px';
      if (!b.getAttribute('style')) b.removeAttribute('style');
    });
    onDesignInput();
  }

  // What a click in the design shows: the image's bar, a link's, a table cell's, or none.
  let selLink = null, selCell = null;
  function hideToolBars() {
    selLink = selCell = null;
    $('linkBar').hidden = true; $('cellBar').hidden = true;
  }
  function onToolClick(target) {
    hideToolBars();
    if (!target || target.nodeType !== 1 || S.mode !== 'design') return;
    const a = target.closest('a[href]');
    if (a && ddoc.body.contains(a)) {
      selLink = a; $('linkHref').value = a.getAttribute('href') || ''; $('linkBar').hidden = false; return;
    }
    const td = target.closest('td, th');
    if (td && ddoc.body.contains(td)) {
      selCell = td;
      $('cellV').value = (td.style.verticalAlign || td.getAttribute('valign') || 'top').toLowerCase().replace(/^center$/, 'middle');
      if (!['top', 'middle', 'bottom'].includes($('cellV').value)) $('cellV').value = 'top';
      $('cellW').value = parseInt(td.getAttribute('width') || td.style.width, 10) || '';
      $('cellGap').value = parseInt(td.style.paddingRight, 10) || (td.style.padding ? parseInt(td.style.padding.split(/\s+/)[1] || td.style.padding, 10) || '' : '');
      $('cellLines').checked = !!td.style.border && !/none|^0/.test(td.style.border);
      $('cellBar').hidden = false;
    }
  }
  // A link's address changed in its bar: https, mailto: or tel: only.
  $('linkHref').addEventListener('change', () => {
    if (!selLink) return;
    const v = $('linkHref').value.trim();
    if (!/^(https?:|mailto:|tel:)/i.test(v)) { toast(t('A link starts with https://, mailto: or tel:.'), true); $('linkHref').value = selLink.getAttribute('href') || ''; return; }
    selLink.setAttribute('href', v); onDesignInput();
  });
  $('linkOpen').addEventListener('click', () => {
    const v = selLink ? selLink.getAttribute('href') || '' : '';
    // Only a web address opens here; a field in it (mailto:@@mail) has no value yet.
    if (!/^https?:\/\//i.test(v) || /@@/.test(v)) { toast(t('Link: {0}', v)); return; }
    window.open(v, '_blank', 'noopener,noreferrer');
  });
  $('linkRemove').addEventListener('click', () => { if (!selLink) return; selLink.replaceWith(...selLink.childNodes); hideToolBars(); onDesignInput(); });
  $('linkDone').addEventListener('click', hideToolBars);

  // The cell's bar onto the cell (and its lines onto every cell of its table).
  function applyCell() {
    const td = selCell; if (!td) return;
    td.style.verticalAlign = $('cellV').value; td.removeAttribute('valign');
    const w = parseInt($('cellW').value, 10);
    if (w > 0) { td.setAttribute('width', w); td.style.width = w + 'px'; } else { td.removeAttribute('width'); td.style.width = ''; }
    const gap = parseInt($('cellGap').value, 10);
    // padding: written out per side, so one changed leaves the others as they were.
    if (td.style.padding) { const p = td.style.padding; td.style.padding = ''; td.style.paddingTop = td.style.paddingTop || p.split(/\s+/)[0]; }
    td.style.paddingRight = gap > 0 ? gap + 'px' : '';
    const table = td.closest('table');
    if (table) table.querySelectorAll(':scope > tbody > tr > td, :scope > tr > td, :scope > tbody > tr > th').forEach((c) => {
      c.style.border = $('cellLines').checked ? '1px solid #d0d5dd' : '';
      if (!c.getAttribute('style')) c.removeAttribute('style');
    });
    if (!td.getAttribute('style')) td.removeAttribute('style');
    onDesignInput();
  }
  ['cellV', 'cellW', 'cellGap', 'cellLines'].forEach((id) => $(id).addEventListener('change', applyCell));
  $('cellDone').addEventListener('click', hideToolBars);

  // The image's corners (rounded, round) and its place (left, middle, right of its line).
  function imgToolsShow(img) {
    const r = img.style.borderRadius || '';
    $('imgRound').value = r === '50%' ? '50%' : parseInt(r, 10) > 0 ? '6' : '';
    const b = img.closest(BLOCK_SEL), al = b && ddoc.body.contains(b) ? (b.style.textAlign || b.getAttribute('align') || '') : '';
    document.querySelectorAll('[data-ialign]').forEach((x) => x.classList.toggle('on', x.dataset.ialign === (al || 'left')));
  }
  $('imgRound').addEventListener('change', () => {
    if (!selImg) return;
    const v = $('imgRound').value;
    selImg.style.borderRadius = v === '50%' ? '50%' : v ? v + 'px' : '';
    // Round: as wide as high, the picture cut to fit (a person's photo in a circle).
    if (v === '50%') { const w = parseInt(selImg.getAttribute('width'), 10) || selImg.offsetWidth; selImg.setAttribute('height', w); selImg.style.objectFit = 'cover'; }
    else selImg.style.objectFit = '';
    if (!selImg.getAttribute('style')) selImg.removeAttribute('style');
    onDesignInput(); gripPlace();
  });
  // The place as the toolbar's align icons, the word as its tooltip.
  document.querySelectorAll('[data-ialign]').forEach((b) => { b.setAttribute('aria-label', b.textContent); b.innerHTML = svg(b.dataset.ialign); });
  document.querySelectorAll('[data-ialign]').forEach((b) => b.addEventListener('click', () => {
    if (!selImg) return;
    const img = selImg, top = img.closest('a') || img;
    let block = img.closest(BLOCK_SEL);
    // An image right in the body: a line of its own made for it.
    if (!block || !ddoc.body.contains(block)) { block = ddoc.createElement('div'); top.replaceWith(block); block.append(top); }
    block.style.textAlign = b.dataset.ialign === 'left' ? '' : b.dataset.ialign; block.removeAttribute('align');
    if (!block.getAttribute('style')) block.removeAttribute('style');
    imgToolsShow(img); onDesignInput(); gripPlace();
  }));

  // The preview as narrow as a phone; kept for this browser.
  (function phonePreview() {
    const kept = () => { try { return localStorage.getItem('osign.phone') === '1'; } catch (e) { return false; } };
    $('pvPhone').checked = kept();
    const apply = () => { document.querySelector('.mail').classList.toggle('phone', $('pvPhone').checked); };
    apply();
    $('pvPhone').addEventListener('change', () => { try { if ($('pvPhone').checked) localStorage.setItem('osign.phone', '1'); else localStorage.removeItem('osign.phone'); } catch (e) { /* no storage */ } apply(); });
  })();
