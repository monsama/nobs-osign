  // ---------------------------------------------------------------------------------------------
  // The design editor: an iframe of its own, so the app's styles never reach the signature
  // ---------------------------------------------------------------------------------------------
  const design = $('design');
  let ddoc = null, designReady = null;
  const EDITOR_CSS = 'html{background:#fff}body{margin:14px 16px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222;outline:none;word-wrap:break-word}' +
    'img{max-width:100%;cursor:pointer}img.sel{outline:2px solid #2f6fde}table{border-collapse:collapse}td{vertical-align:top}' +
    // A person's photo (photos/@@sAMAccountName.jpg) has no picture here: a grey figure in its
    // place and size, framed, so it is seen where it is and how large.
    'img[src*="@@"]{content:var(--osign-ph,url("data:image/svg+xml,<svg xmlns=\'http://www.w3.org/2000/svg\' viewBox=\'0 0 100 80\'><rect width=\'100\' height=\'80\' fill=\'%23e9ecef\'/><circle cx=\'50\' cy=\'31\' r=\'14\' fill=\'%23b8bec6\'/><ellipse cx=\'50\' cy=\'82\' rx=\'30\' ry=\'26\' fill=\'%23b8bec6\'/></svg>"));background:#e9ecef;outline:1px dashed #8a94a3;outline-offset:-1px;object-fit:cover}' +
    'table:not([border]) td{outline:1px dashed #d6dbe3}p{margin:0}a{color:#2f6fde}' +
    // A field, shown by its name in a plain token; in the signature it is the placeholder @@key.
    'span.osf{display:inline-block;padding:0 .35em;margin:0 .05em;border-radius:3px;background:#f2f4f7;color:#344054;' +
    'line-height:1.35;white-space:nowrap;cursor:default;user-select:all;box-shadow:inset 0 0 0 1px #d0d5dd}' +
    'span.osf.unknown{background:#fff;color:#b42318;box-shadow:inset 0 0 0 1px #f0b4ad}' +
    // A building block: its name in a band of its own, as wide as the text.
    'span.osb{display:inline-block;padding:2px .6em;margin:1px 0;border-radius:4px;background:#eaf1fb;color:#1d4ed8;font:600 9pt Arial,sans-serif;cursor:default;user-select:all;box-shadow:inset 0 0 0 1px #b6cdf5}' +
    'span.osb::before{content:"\\25A4  "}span.osb.unknown{background:#fff;color:#b42318;box-shadow:inset 0 0 0 1px #f0b4ad}';

  function initDesign() {
    designReady = new Promise((resolve) => {
      design.addEventListener('load', () => {
        ddoc = design.contentDocument;
        zoomWheel(ddoc, 'design'); zoomApply('design');
        phRefresh();
        ddoc.addEventListener('keydown', (e) => onKeys(e));
        ddoc.body.contentEditable = 'true';
        try { ddoc.execCommand('styleWithCSS', false, true); ddoc.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) { }
        ddoc.addEventListener('input', onDesignInput);
        ddoc.addEventListener('keyup', updateToolbarState);
        ddoc.addEventListener('mouseup', updateToolbarState);
        ddoc.addEventListener('click', onDesignClick);
        ddoc.addEventListener('paste', onPaste);
        resolve();
      }, { once: true });
    });
    // The base makes "assets/logo.png" load from this server, as it will from the published one.
    design.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><base href="/"><style>' + EDITOR_CSS + '</style></head><body></body></html>';
  }

  async function loadDesign(html) {
    await designReady;
    ddoc.body.innerHTML = html || '';
    chipify(ddoc.body);
    hideImgBar();
  }

  // Fields in the design view: each @@key in the text becomes a chip with the field's name
  // (<span class="osf" data-f="key">Vorname</span>), moved and deleted as one; in the signature
  // it is @@key again (designHtml). Placeholders in attributes (mailto:@@mail, a photo's address)
  // stay as they are.
  // A field as people read it: its name (Vorname), or @@key for what is not a field.
  const fieldName = (k) => { const f = S.config.fields.find((x) => x.key === k); return f ? (f.label || f.key) : '@@' + k; };
  const FIELD_RE = /@@([A-Za-z0-9_]+)/g;
  function fieldChip(doc, key) {
    const f = S.config.fields.find((x) => x.key === key);
    const c = doc.createElement('span');
    c.className = 'osf' + (f ? '' : ' unknown');
    c.dataset.f = key;
    c.contentEditable = 'false';
    c.title = '@@' + key + (f ? '' : ' - ' + t('not a field'));
    c.textContent = f ? (f.label || f.key) : '@@' + key;
    return c;
  }
  function blockChip(doc, key) {
    const b = S.config.blocks.find((x) => x.key.toLowerCase() === key.toLowerCase());
    const c = doc.createElement('span');
    c.className = 'osb' + (b ? '' : ' unknown');
    c.dataset.b = b ? b.key : key;
    c.contentEditable = 'false';
    c.title = '{{' + key + '}}' + (b ? ' - ' + t('building block, changed under Building blocks') : ' - ' + t('not a building block'));
    c.textContent = b ? b.name : '{{' + key + '}}';
    return c;
  }
  const TOKEN_RE = /@@([A-Za-z0-9_]+)|\{\{([A-Za-z][A-Za-z0-9_-]{0,39})\}\}/g;
  // Makes chips of the placeholders in the text under root (and re-arms chips pasted or copied in);
  // answers the last chip made, for the caret to go after it.
  function chipify(root) {
    const doc = root.ownerDocument;
    root.querySelectorAll('span.osf[data-f]').forEach((c) => { c.replaceWith(fieldChip(doc, c.dataset.f)); });
    root.querySelectorAll('span.osb[data-b]').forEach((c) => { c.replaceWith(blockChip(doc, c.dataset.b)); });
    const walker = doc.createTreeWalker(root, NodeFilter.SHOW_TEXT), texts = [];
    while (walker.nextNode()) if (/@@|\{\{/.test(walker.currentNode.nodeValue) && !walker.currentNode.parentNode.closest('span.osf, span.osb')) texts.push(walker.currentNode);
    let last = null;
    texts.forEach((n) => {
      const v = n.nodeValue, frag = doc.createDocumentFragment();
      let at = 0, m, made = false;
      TOKEN_RE.lastIndex = 0;
      while ((m = TOKEN_RE.exec(v))) {
        if (m.index > at) frag.append(v.slice(at, m.index));
        last = m[1] ? fieldChip(doc, m[1]) : blockChip(doc, m[2]); frag.append(last); made = true;
        at = m.index + m[0].length;
      }
      if (made) { if (at < v.length) frag.append(v.slice(at)); n.replaceWith(frag); }
    });
    return last;
  }
  // The design view's HTML as it goes into the signature: chips back to @@key, the editor's own
  // marks (a selected image's outline) gone.
  function designHtml() {
    const copy = ddoc.body.cloneNode(true);
    copy.querySelectorAll('span.osf[data-f]').forEach((c) => {
      const extra = [...c.childNodes].filter((n) => n.nodeType === 1);
      c.replaceWith(ddoc.createTextNode('@@' + c.dataset.f), ...extra);
    });
    copy.querySelectorAll('span.osb[data-b]').forEach((c) => c.replaceWith(ddoc.createTextNode('{{' + c.dataset.b + '}}')));
    copy.querySelectorAll('img.sel').forEach((i) => { i.classList.remove('sel'); if (!i.className) i.removeAttribute('class'); });
    return copy.innerHTML;
  }

  function onDesignInput() {
    const s = sig(); if (!s) return;
    // A placeholder typed by hand becomes a chip once it is complete (a space or punctuation after it).
    const sel = ddoc.getSelection(), node = sel.rangeCount ? sel.anchorNode : null;
    if (node && node.nodeType === 3 && /@@[A-Za-z0-9_]+[^A-Za-z0-9_]|\{\{[A-Za-z][A-Za-z0-9_-]*\}\}/.test(node.nodeValue) && !node.parentNode.closest('span.osf, span.osb')) {
      const holder = node.parentNode;
      const tail = node.nodeValue.slice(sel.anchorOffset);
      const last = chipify(holder);
      if (last && !tail) { const r = ddoc.createRange(); r.setStartAfter(last.nextSibling || last); r.collapse(true); sel.removeAllRanges(); sel.addRange(r); }
    }
    s.html = designHtml();
    changed({ list: false });
  }

  function onPaste(e) {
    const html = e.clipboardData && e.clipboardData.getData('text/html');
    if (!html) return;
    e.preventDefault();
    const body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html);
    const clean = sanitize((body ? body[1] : html).replace(/<!--StartFragment-->|<!--EndFragment-->/g, ''));
    ddoc.execCommand('insertHTML', false, clean);
    chipify(ddoc.body);
    onDesignInput();
  }

  function exec(cmd, val) {
    if (cmd === 'undo' || cmd === 'redo') { histStep(cmd === 'undo' ? -1 : 1); return; }
    if (S.mode !== 'design') return;
    design.contentWindow.focus();
    ddoc.execCommand(cmd, false, val === undefined ? null : val);
    onDesignInput(); updateToolbarState();
  }

  // execCommand only knows sizes 1-7: size 7 is set, and what it made gets the real size.
  function setFontSize(pt) {
    exec('fontSize', '7');
    ddoc.querySelectorAll('font[size="7"]').forEach((f) => { const sp = ddoc.createElement('span'); sp.style.fontSize = pt + 'pt'; sp.append(...f.childNodes); f.replaceWith(sp); });
    ddoc.querySelectorAll('[style*="xxx-large"]').forEach((n) => { n.style.fontSize = pt + 'pt'; });
    onDesignInput();
  }

  // Inserted as it is written. execCommand('insertHTML') would copy the editor's own font onto
  // every element it inserts (an image with a font-size), which then goes out in the signature.
  function insertHtml(html) {
    if (S.mode !== 'design') return;
    design.contentWindow.focus();
    const frag = ddoc.createRange().createContextualFragment(html);
    const last = frag.lastChild;
    const sel = ddoc.getSelection();
    if (!sel.rangeCount || !ddoc.body.contains(sel.getRangeAt(0).commonAncestorContainer)) ddoc.body.append(frag);
    else {
      const r = sel.getRangeAt(0);
      // Never inside a field's chip (the caret can end up there, e.g. with Ctrl+End): after it.
      const inChip = (n) => (n.nodeType === 1 ? n : n.parentNode).closest('span.osf, span.osb');
      const chip = inChip(r.startContainer) || inChip(r.endContainer);
      if (chip) { r.setStartAfter(chip); r.collapse(true); }
      r.deleteContents(); r.insertNode(frag);
      if (last) { r.setStartAfter(last); r.collapse(true); sel.removeAllRanges(); sel.addRange(r); }
    }
    onDesignInput();
  }

  const ICON = {
    bold: '<path d="M7 5h6a3.5 3.5 0 0 1 0 7H7zM7 12h7a3.5 3.5 0 0 1 0 7H7z"/>',
    italic: '<path d="M14 5h-4M14 19h-4M14 5l-4 14"/>',
    underline: '<path d="M7 4v7a5 5 0 0 0 10 0V4M5 20h14"/>',
    left: '<path d="M4 6h16M4 10h10M4 14h16M4 18h10"/>',
    center: '<path d="M4 6h16M7 10h10M4 14h16M7 18h10"/>',
    right: '<path d="M4 6h16M10 10h10M4 14h16M10 18h10"/>',
    link: '<path d="M10 14a4 4 0 0 0 6 0l3-3a4 4 0 0 0-6-6l-1 1M14 10a4 4 0 0 0-6 0l-3 3a4 4 0 0 0 6 6l1-1"/>',
    unlink: '<path d="M10 14a4 4 0 0 0 6 0l3-3a4 4 0 0 0-6-6l-1 1M14 10a4 4 0 0 0-6 0l-3 3a4 4 0 0 0 6 6l1-1M4 4l16 16"/>',
    image: '<rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-8 8"/>',
    table: '<rect x="3" y="5" width="18" height="14" rx="1"/><path d="M3 12h18M12 5v14"/>',
    hr: '<path d="M4 12h16"/>',
    clear: '<path d="M6 5h12M12 5l-3 14M4 20l16-16"/>',
    undo: '<path d="M9 14L4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 0 10h-3"/>',
    redo: '<path d="M15 14l5-5-5-5"/><path d="M20 9H9a5 5 0 0 0 0 10h3"/>'
  };
  const svg = (name) => '<svg viewBox="0 0 24 24" aria-hidden="true">' + ICON[name] + '</svg>';

  function buildToolbar() {
    const tb = $('toolbar');
    const btn = (name, title, fn, cmd) => { title = t(title); const b = el('button', { type: 'button', title, 'aria-label': title, 'data-cmd': cmd || '' }); b.innerHTML = svg(name); b.addEventListener('mousedown', (e) => e.preventDefault()); b.addEventListener('click', fn); return b; };
    const sep = () => el('span', { class: 'sep' });
    const fonts = ['Arial', 'Aptos', 'Calibri', 'Segoe UI', 'Verdana', 'Tahoma', 'Helvetica', 'Georgia', 'Times New Roman'];
    const fontSel = el('select', { title: t('Font'), 'aria-label': t('Font') }, el('option', { value: '' }, t('Font')), fonts.map((f) => el('option', { value: f }, f)));
    fontSel.addEventListener('change', () => { if (fontSel.value) exec('fontName', fontSel.value + (fontSel.value === 'Georgia' || fontSel.value === 'Times New Roman' ? ', serif' : ', sans-serif')); fontSel.value = ''; });
    const sizeSel = el('select', { title: t('Size'), 'aria-label': t('Size') }, el('option', { value: '' }, t('Size')), [8, 9, 10, 11, 12, 14, 16, 18, 24].map((n) => el('option', { value: n }, n + ' pt')));
    sizeSel.addEventListener('change', () => { if (sizeSel.value) setFontSize(sizeSel.value); sizeSel.value = ''; });
    const color = el('input', { type: 'color', title: t('Text colour'), 'aria-label': t('Text colour'), value: '#222222' });
    color.addEventListener('input', () => exec('foreColor', color.value));
    const fieldSel = el('select', { class: 'field-pick', title: t('Insert a field'), 'aria-label': t('Insert a field') });
    fieldSel.addEventListener('focus', () => fillFieldSelect(fieldSel));
    fieldSel.addEventListener('change', () => {
      const k = fieldSel.value; fieldSel.value = '';
      if (!k) return;
      if (k.startsWith('block:')) {
        // A building block: on a line of its own.
        const b = k.slice(6);
        if (S.mode === 'html') insertAtCursor($('source'), '<div>{{' + b + '}}</div>');
        else insertHtml('<div>' + blockChip(ddoc, b).outerHTML + '</div>');
        return;
      }
      if (S.mode === 'html') insertAtCursor($('source'), '@@' + k);
      else { const c = fieldChip(ddoc, k); insertHtml(c.outerHTML); }
    });
    fillFieldSelect(fieldSel);
    // The HTML tab: indent it again after typing in it.
    const fmt = el('button', { type: 'button', class: 'html-only', title: t('Indent the HTML again (Alt+Shift+F)') }, t('Format'));
    fmt.addEventListener('click', formatSource);
    fmt.style.display = S.mode === 'html' ? '' : 'none';
    // Two lines, at any width: what to insert, the font and Undo/Redo (at the right) on the first,
    // the formatting on the second. In groups that wrap as a whole on a narrow screen.
    const g = (cls, ...kids) => el('span', { class: 'tb-g ' + cls }, ...kids);
    const hist = g('tb-hist', btn('undo', 'Undo (Ctrl+Z)', () => histStep(-1), 'undo'), btn('redo', 'Redo (Ctrl+Y)', () => histStep(1), 'redo'));
    tb.append(
      g('tb-keep', fieldSel, fmt), sep(), g('tb-font', fontSel, sizeSel, color), hist, el('span', { class: 'tb-break', 'aria-hidden': 'true' }),
      g('tb-style', btn('bold', 'Bold (Ctrl+B)', () => exec('bold'), 'bold'), btn('italic', 'Italic (Ctrl+I)', () => exec('italic'), 'italic'), btn('underline', 'Underline (Ctrl+U)', () => exec('underline'), 'underline')), sep(),
      g('tb-align', btn('left', 'Align left', () => exec('justifyLeft'), 'justifyLeft'), btn('center', 'Centre', () => exec('justifyCenter'), 'justifyCenter'), btn('right', 'Align right', () => exec('justifyRight'), 'justifyRight')), sep(),
      g('tb-insert', btn('link', 'Link', makeLink), btn('unlink', 'Remove link', () => exec('unlink')), btn('image', 'Image', openImageDialog), btn('table', 'Two columns (logo beside the text)', insertTable), btn('hr', 'Line', () => exec('insertHorizontalRule')), btn('clear', 'Clear formatting', () => exec('removeFormat')))
    );
  }

  // Placeholders the signatures use that no field has: fields for them (as CI-Sign read them from
  // AD, with the Graph property where there is one). Answers the keys added.
  function addMissingFields(sigs) {
    const have = new Set(S.config.fields.map((f) => f.key));
    const added = [];
    for (const s of sigs || S.config.signatures) {
      for (const k of core.placeholdersIn(sigFull(s))) if (!have.has(k)) { have.add(k); S.config.fields.push(core.fieldFor(k)); added.push(k); }
    }
    return added;
  }
  const FIELD_GROUPS = {
    person: ['givenname', 'sn', 'displayname', 'initials', 'description', 'title', 'department', 'samaccountname', 'userprincipalname', 'photo', 'thumbnailphoto'],
    contact: ['mail', 'mobile', 'homephone', 'telephonenumber', 'othertelephone', 'ipphone', 'facsimiletelephonenumber', 'pager', 'wwwhomepage', 'othermobile'],
    company: ['company', 'physicaldeliveryofficename', 'streetaddress', 'postofficebox', 'postalcode', 'l', 'st', 'co', 'c']
  };
  // A separator only between two groups on the same line: none at the start or end of a line
  // where the toolbar wraps.
  function tidySeps(bar) {
    if (!bar) return;
    const seps = [...bar.querySelectorAll(':scope > .sep')];
    seps.forEach((s) => { s.classList.remove('edge'); });
    const shown = (n) => n && n.offsetParent !== null && getComputedStyle(n).display !== 'none';
    const near = (s, dir) => { let n = s[dir]; while (n && (!shown(n) || n.classList.contains('sep'))) n = n[dir]; return n; };
    seps.forEach((s) => {
      const a = near(s, 'previousElementSibling'), b = near(s, 'nextElementSibling');
      if (!a || !b || a.offsetTop !== b.offsetTop || s.offsetTop !== b.offsetTop) s.classList.add('edge');
    });
  }
  if (window.ResizeObserver) new ResizeObserver(() => tidySeps($('toolbar'))).observe($('toolbar'));
  function fillFieldSelect(sel) {
    sel.textContent = '';
    sel.append(el('option', { value: '' }, t('+ Insert field')));
    if (!S.config) return;
    // In building blocks, by what the field is read from: the person, how to reach them, the company.
    const groups = [['Person', FIELD_GROUPS.person], ['Contact', FIELD_GROUPS.contact], ['Company', FIELD_GROUPS.company], ['More', null]];
    const of = (f) => String(f.ad || f.key).toLowerCase();
    const placed = new Set();
    for (const [name, keys] of groups) {
      const fs = S.config.fields.filter((f) => !placed.has(f) && (!keys || keys.includes(of(f))));
      if (!fs.length) continue;
      const g = el('optgroup', { label: t(name) });
      fs.forEach((f) => { placed.add(f); g.append(el('option', { value: f.key }, f.label || f.key)); });
      sel.append(g);
    }
    // Building blocks: not into a building block (one inside another is not put in).
    if (S.config.blocks.length && !S.blockEdit) {
      const g = el('optgroup', { label: t('Building blocks') });
      S.config.blocks.forEach((b) => g.append(el('option', { value: 'block:' + b.key }, b.name)));
      sel.append(g);
    }
  }

  function insertAtCursor(ta, text) {
    const a = ta.selectionStart, b = ta.selectionEnd;
    ta.setRangeText(text, a, b, 'end'); ta.focus(); ta.dispatchEvent(new Event('input'));
  }

  function updateToolbarState() {
    if (!ddoc) return;
    document.querySelectorAll('#toolbar button[data-cmd]').forEach((b) => {
      const c = b.dataset.cmd; if (!c || c === 'undo' || c === 'redo') return;
      let on = false; try { on = ddoc.queryCommandState(c); } catch (e) { }
      b.classList.toggle('on', on);
    });
  }

  async function makeLink() {
    // Where the link goes: the selection in the design view, kept while the question is open.
    const sel = ddoc && ddoc.getSelection(), range = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    const url = await askFor(t('Link address (https://…, mailto:…, tel:…). Fields work here too: mailto:@@mail'), 'https://', t('Link'));
    if (!url || url === 'https://') return;
    if (!/^(https?:|mailto:|tel:)/i.test(url.trim())) { toast(t('A link starts with https://, mailto: or tel:.'), true); return; }
    design.contentWindow.focus();
    if (range) { const s2 = ddoc.getSelection(); s2.removeAllRanges(); s2.addRange(range); }
    exec('createLink', url.trim());
  }

  function insertTable() {
    insertHtml('<table cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tr>' +
      '<td style="padding:0 14px 0 0;vertical-align:top">Logo</td>' +
      '<td style="vertical-align:top;font-family:Arial,sans-serif;font-size:10pt"><b>@@givenName @@sn</b><br>@@title<br>Mobile @@mobile<br>@@mail</td>' +
      '</tr></table><p><br></p>');
  }

  // Selected image: width, alt text and a link, in a bar above the editor.
  let selImg = null;
  function onDesignClick(e) {
    if (e.target && e.target.nodeName === 'IMG') {
      hideImgBar(); selImg = e.target; selImg.classList.add('sel');
      // A person's photo: so named, with the placeholder for whoever has none (all signatures).
      const person = /@@([A-Za-z0-9_]+)/.exec(selImg.getAttribute('src') || '');
      $('imgKind').textContent = person ? t('Photo per person (@@{0})', person[1]) : t('Image');
      $('imgPh').hidden = !person;
      if (person) phRefresh();
      $('imgW').value = selImg.getAttribute('width') || selImg.naturalWidth || '';
      $('imgAlt').value = selImg.getAttribute('alt') || '';
      const a = selImg.closest('a'); $('imgHref').value = a ? a.getAttribute('href') : '';
      $('imgBar').hidden = false;
    } else hideImgBar();
  }
  function hideImgBar() {
    if (selImg) { selImg.classList.remove('sel'); if (!selImg.className) selImg.removeAttribute('class'); }
    selImg = null; $('imgBar').hidden = true;
  }
  function applyImg() {
    if (!selImg) return;
    const w = parseInt($('imgW').value, 10);
    // A height set too (a person's photo in its shape): it goes along, so nothing gets stretched.
    const ow = parseInt(selImg.getAttribute('width'), 10), oh = parseInt(selImg.getAttribute('height'), 10);
    if (w > 0 && ow > 0 && oh > 0) { selImg.setAttribute('width', w); selImg.setAttribute('height', Math.round(oh * w / ow)); selImg.style.width = ''; selImg.style.height = ''; if (!selImg.getAttribute('style')) selImg.removeAttribute('style'); }
    else if (w > 0) { selImg.setAttribute('width', w); selImg.removeAttribute('height'); selImg.style.width = ''; selImg.style.height = ''; if (!selImg.getAttribute('style')) selImg.removeAttribute('style'); }
    selImg.setAttribute('alt', $('imgAlt').value);
    const href = $('imgHref').value.trim(), a = selImg.closest('a');
    if (href && /^(https?:|mailto:|tel:)/i.test(href)) {
      if (a) a.setAttribute('href', href); else { const na = ddoc.createElement('a'); na.href = href; selImg.replaceWith(na); na.append(selImg); }
    } else if (!href && a) a.replaceWith(...a.childNodes);
    onDesignInput();
  }

  // Images: pick one already in the draft, or upload one.
  function openImageDialog() {
    $('imgErr').textContent = '';
    renderImageList();
    $('dlgImage').showModal();
  }
  function renderImageList() {
    const box = $('imgList'); box.textContent = '';
    if (!S.assets.length) box.append(el('p', { class: 'lead' }, t('No images yet.')));
    S.assets.forEach((a) => {
      const del = el('span', { class: 'del', title: t('Remove from the draft'), role: 'button', onclick: (e) => { e.stopPropagation(); deleteAsset(a.name); } }, t('remove'));
      box.append(el('button', { type: 'button', onclick: () => { $('dlgImage').close(); insertImage('assets/' + a.name); } },
        el('img', { src: '/assets/' + encodeURIComponent(a.name), alt: '' }), el('span', { class: 'nm', title: a.name }, a.name), del));
    });
  }
  // Each person's own photo: photos/@@<login>.jpg, the photos kept under Photos.
  $('imgPerson').addEventListener('click', () => {
    $('dlgImage').close();
    const rule = photoRule(), key = rule ? rule.key : (fieldKeys().includes('sAMAccountName') ? 'sAMAccountName' : 'mail');
    insertHtml('<img src="photos/@@' + key + '.' + (rule ? rule.ext : 'jpg') + '" width="100" alt="" style="border:0">');
  });
  function insertImage(src) {
    const i = new Image();
    i.onload = () => { const w = Math.min(i.naturalWidth, 240); insertHtml('<img src="' + core.escapeHtml(src) + '" width="' + w + '" alt="" style="border:0">'); };
    i.onerror = () => insertHtml('<img src="' + core.escapeHtml(src) + '" alt="" style="border:0">');
    i.src = '/' + src;
  }
  async function deleteAsset(name) {
    const used = S.config.signatures.filter((s) => sigFull(s).includes('assets/' + name)).map((s) => s.name);
    if (used.length && !(await ask(t('{0} is used in: {1}. Remove it anyway?', name, used.join(', ')), t('Remove'), true))) return;
    try { await api('assets/delete', { name }); S.assets = S.assets.filter((a) => a.name !== name); renderImageList(); changed({ list: false }); } catch (e) { $('imgErr').textContent = t(e.message); }
  }
  // An image as small as a signature needs, before it is uploaded: at most 1200 pixels wide
  // (twice a wide signature, so it stays sharp on high-resolution screens), a JPEG saved again
  // at good quality. A PNG stays a PNG (it may be transparent), a GIF as it is (it may move).
  // Smaller files go faster into every mail, and fit into what classic Outlook keeps.
  const IMG_MAX_W = 1200;
  // A person's photo is shown small (about 100 pixels): 400 wide is plenty.
  const PHOTO_MAX_W = 400;
  async function shrinkUpload(f, maxW) {
    if (!/^image\/(png|jpeg)$/.test(f.type) || typeof createImageBitmap !== 'function') return f;
    let bmp;
    try { bmp = await createImageBitmap(f); } catch (e) { return f; }
    const scale = Math.min(1, (maxW || IMG_MAX_W) / bmp.width);
    // A PNG of the right size stays as it is; a JPEG is saved again only when it is large.
    if (scale === 1 && (f.type === 'image/png' || f.size <= 150 * 1024)) return f;
    const c = document.createElement('canvas');
    c.width = Math.round(bmp.width * scale); c.height = Math.round(bmp.height * scale);
    const g = c.getContext('2d'); g.imageSmoothingQuality = 'high'; g.drawImage(bmp, 0, 0, c.width, c.height);
    const blob = await new Promise((res) => c.toBlob(res, f.type, 0.86));
    if (!blob || blob.size >= f.size) return f;
    const out = new File([blob], f.name, { type: f.type });
    out.shrunk = { from: f.size, to: blob.size, w: bmp.width, w2: c.width };
    return out;
  }
  const kb = (n) => (n >= 1024 * 1024 ? (n / 1024 / 1024).toFixed(1) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB');
  $('imgFile').addEventListener('change', async () => {
    let f = $('imgFile').files[0]; $('imgFile').value = '';
    if (!f) return;
    $('imgErr').textContent = '';
    f = await shrinkUpload(f);
    if (f.shrunk) toast(t('Image made smaller: {0} → {1} ({2} → {3} pixels wide).', kb(f.shrunk.from), kb(f.shrunk.to), f.shrunk.w, f.shrunk.w2));
    if (f.size > 1024 * 1024) { $('imgErr').textContent = t('The image is larger than 1 MB.'); return; }
    const data = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = rej; r.readAsDataURL(f); });
    try {
      const r = await api('assets/upload', { name: f.name, data });
      S.assets = S.assets.filter((a) => a.name !== r.name).concat([{ name: r.name, size: r.size }]).sort((a, b) => a.name.localeCompare(b.name));
      $('dlgImage').close();
      insertImage(r.path);
    } catch (e) { $('imgErr').textContent = t(e.message); }
  });

  // Undo and redo for the signature open: its HTML after each change (typing a little bit at a
  // time), up to 100 steps back. Everything counts: typing, a field put in, an image's width, the
  // HTML tab; the browser's own undo knows only some of it, and nothing once the focus left.
  const hist = { id: null, stack: [], pos: 0, timer: null };
  function histReset(s) { clearTimeout(hist.timer); hist.id = s ? s.id : null; hist.stack = s ? [s.html || ''] : []; hist.pos = 0; histButtons(); }
  function histRecord() {
    clearTimeout(hist.timer); hist.timer = null;
    const s = sig(); if (!s || s.id !== hist.id) return;
    const h = s.html || '';
    if (hist.stack[hist.pos] === h) return;
    hist.stack = hist.stack.slice(0, hist.pos + 1);
    hist.stack.push(h);
    if (hist.stack.length > 101) hist.stack.shift();
    hist.pos = hist.stack.length - 1;
    histButtons();
  }
  // Typing goes in as one step once it pauses.
  function histNote() { if (S.view === 'sig' && hist.id) { clearTimeout(hist.timer); hist.timer = setTimeout(histRecord, 500); } }
  function histStep(dir) {
    if (hist.timer) histRecord();
    const s = sig(); if (!s || s.id !== hist.id) return;
    if (S.mode === 'html') flushSource();
    histRecord();
    const to = hist.pos + dir;
    if (to < 0 || to >= hist.stack.length) return;
    hist.pos = to;
    s.html = hist.stack[to];
    loadDesign(s.html).then(() => { if (S.mode === 'design') { design.contentWindow.focus(); const r = ddoc.createRange(); r.selectNodeContents(ddoc.body); r.collapse(false); const sel = ddoc.getSelection(); sel.removeAllRanges(); sel.addRange(r); } });
    $('source').value = prettyHtml(s.html);
    changed({ list: false });
    clearTimeout(hist.timer); hist.timer = null;
    histButtons();
  }
  function histButtons() {
    const u = document.querySelector('#toolbar [data-cmd=undo]'), r = document.querySelector('#toolbar [data-cmd=redo]');
    if (u) u.disabled = hist.pos <= 0 && !hist.timer;
    if (r) r.disabled = hist.pos >= hist.stack.length - 1;
  }

  // Zoom: the design view and the preview each their own, 50 to 200 %, kept for this browser.
  // The buttons above each, or Ctrl + mouse wheel over it; the percentage puts it back to 100 %.
  // At first 80 %: the signature as large as Outlook shows it on most screens, the wide ones whole.
  const ZOOM_MIN = 50, ZOOM_MAX = 200, ZOOM_STEP = 10, ZOOM_START = 80;
  const zoom = { design: ZOOM_START, preview: ZOOM_START };
  const zoomKept = (k, v) => { try { if (v === undefined) return localStorage.getItem('osign.zoom.' + k); if (v === null) localStorage.removeItem('osign.zoom.' + k); else localStorage.setItem('osign.zoom.' + k, v); } catch (e) { /* no storage */ } return null; };
  for (const k of Object.keys(zoom)) { const v = Number(zoomKept(k)); if (v >= ZOOM_MIN && v <= ZOOM_MAX) zoom[k] = v; }
  function zoomApply(k) {
    const doc = k === 'design' ? ddoc : $('preview').contentDocument;
    if (doc && doc.documentElement) doc.documentElement.style.zoom = String(zoom[k] / 100);
    // The HTML tab follows the design view's zoom with its text size.
    if (k === 'design') $('source').style.fontSize = zoom[k] === 100 ? '' : (12.5 * zoom[k] / 100).toFixed(1) + 'px';
    const box = document.querySelector('.zoom[data-zoom="' + k + '"]');
    if (box) {
      box.querySelector('.zpct').textContent = zoom[k] + ' %';
      box.querySelector('[data-z="-1"]').disabled = zoom[k] <= ZOOM_MIN;
      box.querySelector('[data-z="1"]').disabled = zoom[k] >= ZOOM_MAX;
    }
  }
  function zoomBy(k, dir) {
    zoom[k] = dir === 0 ? ZOOM_START : Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round((zoom[k] + dir * ZOOM_STEP) / ZOOM_STEP) * ZOOM_STEP));
    zoomKept(k, zoom[k] === ZOOM_START ? null : String(zoom[k]));
    zoomApply(k);
  }
  document.querySelectorAll('.zoom[data-zoom]').forEach((box) => {
    box.addEventListener('click', (e) => { const b = e.target.closest('button[data-z]'); if (b) zoomBy(box.dataset.zoom, Number(b.dataset.z)); });
  });
  // Ctrl + wheel over a view zooms that view, not the page.
  function zoomWheel(doc, k) {
    if (!doc || doc.__osignZoom) return;
    doc.__osignZoom = true;
    doc.addEventListener('wheel', (e) => { if (!e.ctrlKey) return; e.preventDefault(); zoomBy(k, e.deltaY < 0 ? 1 : -1); }, { passive: false });
  }
  $('preview').addEventListener('load', () => { zoomWheel($('preview').contentDocument, 'preview'); zoomApply('preview'); });
  zoomApply('preview');

  // The HTML tab: the signature indented, a block (table, row, cell, paragraph) and a line (<br>)
  // each on a line of its own; what that adds (a line break and the indent next to a block or a
  // <br>) goes again when it is kept, so the signature is the same as before.
  const BLOCKS = 'table|thead|tbody|tfoot|tr|td|th|p|div|ul|ol|li|h[1-6]|blockquote|center|hr';
  // Line breaks (and the indent after them) next to a block or a <br>: they show nothing.
  const NLWS = '[ \\t]*(?:\\r?\\n[ \\t]*)+';
  const unwrap = (h) => String(h || '').replace(new RegExp(NLWS + '(?=</?(?:' + BLOCKS + '|br)\\b)', 'gi'), '').replace(new RegExp('(</?(?:' + BLOCKS + '|br)\\b[^>]*>)' + NLWS, 'gi'), '$1');
  function prettyHtml(html) {
    const src = unwrap(html);
    const parts = src.split(/(<!--[\s\S]*?-->|<[^>]+>)/);
    const isBlock = new RegExp('^<(/?)(' + BLOCKS + ')\\b', 'i');
    let out = '', depth = 0, atLineStart = true;
    const nl = () => { out = out.replace(/[ \t]+$/, ''); if (out && !out.endsWith('\n')) out += '\n'; out += '  '.repeat(depth); atLineStart = true; };
    for (const p of parts) {
      if (!p) continue;
      const m = isBlock.exec(p);
      if (m && !m[1]) {
        nl(); out += p; atLineStart = false;
        // What the block holds starts on a line of its own, indented.
        if (!/^<hr\b/i.test(p) && !/\/>$/.test(p)) { depth++; nl(); }
      } else if (m && m[1]) {
        depth = Math.max(0, depth - 1);
        // An empty block stays on one line: <td></td>.
        if (new RegExp('<' + m[2] + '\\b[^>]*>\\s*$', 'i').test(out)) out = out.replace(/\s+$/, '');
        else nl();
        out += p; atLineStart = false;
      } else if (/^<br\b/i.test(p)) {
        out += p; nl();
      } else {
        out += atLineStart ? p.replace(/^[ \t]+/, '') : p;
        if (p.trim()) atLineStart = false;
      }
    }
    return out.replace(/[ \t]+$/gm, '').replace(/^\n+/, '').replace(/\n+$/, '');
  }
  // The HTML tab indented again, the cursor kept near where it was (by line).
  function formatSource() {
    const src = $('source'), line = src.value.slice(0, src.selectionStart).split('\n').length;
    src.value = prettyHtml(plainHtml(src.value));
    const lines = src.value.split('\n'), at = lines.slice(0, Math.min(line, lines.length) - 1).join('\n').length + (line > 1 ? 1 : 0);
    src.focus(); src.setSelectionRange(at, at);
  }
  $('source').addEventListener('keydown', (e) => { if (e.altKey && e.shiftKey && (e.key === 'F' || e.key === 'f')) { e.preventDefault(); formatSource(); } });
  function plainHtml(text) {
    return unwrap(text).replace(/^\s+|\s+$/g, '');
  }

  // Design / HTML
  function setMode(mode) {
    if (mode === S.mode) return;
    const s = sig();
    if (mode === 'html') { $('source').value = s ? prettyHtml(s.html) : ''; }
    else { flushSource(); loadDesign(s ? s.html : ''); }
    // What was done in the other tab is a step of its own.
    histRecord();
    S.mode = mode;
    document.querySelectorAll('.tab[data-mode]').forEach((t) => { t.classList.toggle('active', t.dataset.mode === mode); t.setAttribute('aria-selected', t.dataset.mode === mode); });
    design.hidden = mode !== 'design';
    $('source').hidden = mode !== 'html';
    // HTML: only inserting a field and Format; the design tools in Design.
    document.querySelectorAll('#toolbar > *').forEach((n) => { n.style.display = mode === 'html' && !n.classList.contains('tb-keep') ? 'none' : ''; });
    document.querySelectorAll('#toolbar .html-only').forEach((n) => { n.style.display = mode === 'html' ? '' : 'none'; });
    tidySeps($('toolbar'));
    hideImgBar();
  }
  // What was typed in the HTML tab, made safe, once the tab is left.
  function flushSource() {
    if (S.mode !== 'html') return;
    const s = sig(); if (!s) return;
    const clean = sanitize(plainHtml($('source').value));
    if (clean !== s.html) { s.html = clean; changed({ list: false }); }
  }
  $('source').addEventListener('input', () => { const s = sig(); if (!s) return; s.html = plainHtml($('source').value); changed({ list: false }); });
  $('source').addEventListener('blur', flushSource);

