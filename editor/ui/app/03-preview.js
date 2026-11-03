  // ---------------------------------------------------------------------------------------------
  // Preview
  // ---------------------------------------------------------------------------------------------
  let pvTimer = null;
  function schedulePreview() { clearTimeout(pvTimer); pvTimer = setTimeout(renderPreview, 150); }

  function previewValues() {
    const src = $('pvSource').value, v = {};
    S.config.fields.forEach((f) => { v[f.key] = src === 'sample' ? (f.sample || '') : (src === 'ad' || src === 'entra') && S.adValues && S.adFrom === src ? (S.adValues[f.key] || '') : ''; });
    return v;
  }

  // The signature as the preview shows it (its values, the building blocks put in), with its images
  // inside (data: addresses), as one file: to try it in a mail program without publishing.
  async function exportHtml() {
    const s = sig(); if (!s) return null;
    const html = core.render(core.compile(sanitize(sigFull(s)), scratchDoc), previewValues(), { assetBase: '/', removeEmptyLines: s.removeEmptyLines !== false });
    const doc = new DOMParser().parseFromString('<!doctype html><body>' + html + '</body>', 'text/html');
    await Promise.all([...doc.querySelectorAll('img[src]')].map(async (img) => {
      // Only the signatures' own images and photos, and only what is an image.
      if (!/^\/(assets|photos)\/[^/?#]+(\?[^#]*)?$/.test(img.getAttribute('src'))) { if (/^\//.test(img.getAttribute('src'))) img.remove(); return; }
      try {
        const r = await fetch(img.getAttribute('src'), { credentials: 'same-origin' });
        if (!r.ok) { img.remove(); return; }
        const b = await r.blob();
        if (!/^image\//.test(b.type)) { img.remove(); return; }
        img.src = await new Promise((ok) => { const fr = new FileReader(); fr.onload = () => ok(fr.result); fr.readAsDataURL(b); });
      } catch (e) { img.remove(); }
    }));
    return { body: doc.body.innerHTML, name: (s.name || 'signature').replace(/[^\w\- ]+/g, '').trim() || 'signature' };
  }
  $('pvCopy').addEventListener('click', async () => {
    const x = await exportHtml(); if (!x) return;
    try {
      await navigator.clipboard.write([new ClipboardItem({ 'text/html': new Blob([x.body], { type: 'text/html' }), 'text/plain': new Blob([core.htmlToText(x.body)], { type: 'text/plain' }) })]);
      toast(t('Copied, with these values: paste it into a mail or into Outlook\'s signature settings.'));
    } catch (e) { toast(t('The browser did not let it be copied: use Download instead.'), true); }
  });
  $('pvSave').addEventListener('click', async () => {
    const x = await exportHtml(); if (!x) return;
    const file = '<!doctype html><html><head><meta charset="utf-8"><title>' + core.escapeHtml(x.name) + '</title></head><body>' + x.body + '</body></html>';
    const a = el('a', { href: URL.createObjectURL(new Blob([file], { type: 'text/html' })), download: x.name + '.htm' });
    document.body.append(a); a.click(); a.remove();
    toast(t('Downloaded: {0}.htm, with these values.', x.name));
  });

  function renderPreview() {
    const s = sig(); if (!s || S.view !== 'sig') return;
    const html = core.render(core.compile(sanitize(sigFull(s)), scratchDoc), previewValues(), { assetBase: '/', removeEmptyLines: s.removeEmptyLines !== false });
    const asText = $('pvText').checked;
    const greet = asText ? '' : '<p>' + core.escapeHtml(t('Hello,')) + '</p><p>' + core.escapeHtml(t('this is how the signature looks under a message.')) + '</p><p>&nbsp;</p>';
    const body = asText ? '<pre style="font:11pt/1.4 Consolas,monospace;white-space:pre-wrap">' + core.escapeHtml(core.htmlToText(html)) + '</pre>' : greet + html;
    // Dark mode, roughly as Outlook shows it to recipients who use it: the colours turned around,
    // the images as they are (a dark logo on a transparent background gets hard to see).
    const dark = $('pvDark').checked ? 'html{background:#fff;filter:invert(1) hue-rotate(180deg)}img{filter:invert(1) hue-rotate(180deg)}' : '';
    $('preview').srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:14px 16px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222}p{margin:0}table{border-collapse:collapse}' + dark + '</style></head><body>' + body + '</body></html>';
    // A person's photo that is not there goes, as the add-in leaves it out; the note says which.
    $('preview').onload = () => {
      const d = $('preview').contentDocument; if (!d) return;
      // A link in the signature does not open here (the site would refuse to be shown in the
      // preview): where it goes is said instead.
      d.addEventListener('click', (e) => {
        const a = e.target.closest && e.target.closest('a[href]');
        if (!a) return;
        e.preventDefault();
        toast(t('Link: {0}', a.getAttribute('href')));
      });
      d.querySelectorAll('a[href]').forEach((a) => { a.title = a.getAttribute('href'); });
      const gone = (img) => {
        const name = decodeURIComponent(img.getAttribute('src').split('/').pop());
        img.remove();
        if (!$('pvNote').querySelector('[data-photo="' + CSS.escape(name) + '"]')) {
          $('pvNote').append(el('div', { class: 'warn', 'data-photo': name }, t('No photo {0} in the photos folder: the signature goes without it.', name) + ' ',
            el('a', { href: '#', onclick: (e) => { e.preventDefault(); photoAddFor(name); } }, t('Add the photo'))));
        }
      };
      d.querySelectorAll('img[src*="/photos/"]').forEach((img) => { if (img.complete && !img.naturalWidth) gone(img); else img.addEventListener('error', () => gone(img)); });
    };

    // What to know about this signature before it goes out.
    const notes = [], keys = fieldKeys(), used = core.placeholdersIn(sigFull(s));
    const unknown = used.filter((k) => keys.indexOf(k) < 0);
    const noGraph = used.filter((k) => { const f = S.config.fields.find((x) => x.key === k); return f && core.fieldSource(S.config, f) === 'none'; });
    if (unknown.length) notes.push(el('div', { class: 'warn' }, t('Not a field: {0} - left as typed.', unknown.map((k) => '@@' + k).join(', ')) + ' ',
      el('a', { href: '#', onclick: (e) => {
        e.preventDefault();
        const added = addMissingFields([s]);
        changed({ list: false });
        toast(t('Fields added: {0}. Check their AD attribute and Graph source under Fields.', added.map((k) => '@@' + k).join(', ')));
      } }, t('Add them as fields'))));
    if ($('pvDark').checked && /<img\b/i.test(sigFull(s))) notes.push(el('div', null, t('Dark mode, roughly: Outlook turns the colours itself. Check that the logo still shows on a dark background (a white edge or a light version helps).')));
    if (core.isCampaign(s) && !core.inPeriod(s)) notes.push(el('div', { class: 'warn' }, t('A campaign, not running today ({0}): nobody gets it now.', campaignText(s))));
    if (noGraph.length) notes.push(el('div', { class: 'warn' }, core.valueSource(S.config) === 'graph'
      ? t('No Graph source, so empty in Outlook (see Fields): {0}', noGraph.map(fieldName).join(', '))
      : t('No source (neither Graph nor AD attribute), so empty in Outlook (see Fields): {0}', noGraph.map(fieldName).join(', '))));
    if (($('pvSource').value === 'ad' || $('pvSource').value === 'entra') && S.adName && S.adFrom === $('pvSource').value) {
      // What this user would get, as the add-in decides it (the From address being their own).
      const v = S.adValues || {}, from = v.mail || '', g = S.adGroups || [];
      const nm = (x) => (x ? x.name : t('none'));
      const open = core.isAvailable(s, v, from, g);
      notes.push(el('div', { class: 'pv-groups' }, ($('pvSource').value === 'entra' ? t('Values of {0} from Entra ID.', S.adName) : t('Values of {0} from Active Directory.', S.adName)) + ' ' +
        (open ? t('This signature is open to them.') : t('This signature is NOT open to them.')) + ' ' +
        t('New mail: {0}, replies: {1}.', nm(core.pick(S.config, v, from, 'new', null, g)), nm(core.pick(S.config, v, from, 'reply', null, g)))));
    }
    $('pvNote').replaceChildren(...notes);
  }

  $('pvSource').addEventListener('change', () => {
    const entra = $('pvSource').value === 'entra';
    const ad = $('pvSource').value === 'ad';
    $('pvAd').hidden = !ad && !entra;
    if (entra && !S.adValues) $('pvUser').focus();
    if (ad && !(S.info && S.info.features && S.info.features.adLookup)) toast(onServer() ? t('Active Directory is not set up on the server: {0}', t(S.info.features.adProblem)) : t('Active Directory is not reachable from here: this needs Windows on a machine in the domain.'), true);
    if (ad && !S.adValues) $('pvUser').focus();
    renderPreview();
  });
  $('pvText').addEventListener('change', renderPreview);
  $('pvDark').addEventListener('change', renderPreview);
  // A real person's values and groups: from AD (the service account) or from Entra ID (the
  // server's app), whichever the preview is set to.
  async function adLookup() {
    const from = $('pvSource').value === 'entra' ? 'entra' : 'ad';
    try {
      const r = from === 'entra'
        ? await api('entra/lookup', { query: $('pvUser').value, fields: S.config.fields.map((f) => ({ key: f.key, graph: f.graph })) })
        : await api('ad/lookup', { query: $('pvUser').value, fields: S.config.fields.map((f) => ({ key: f.key, ad: f.ad })) });
      S.adValues = r.values; S.adName = r.displayName; S.adGroups = r.groups || []; S.adFrom = from; renderPreview();
    } catch (e) {
      // Several match: the choice instead of the error.
      if (/More than one/.test(e.message) || /Mehr als ein/.test(e.message)) { $('pvUser').focus(); pvSuggest.open(); return; }
      toast(t(e.message), true);
    }
  }
  $('pvLookup').addEventListener('click', adLookup);
  // Enter: the suggestion picked (↓), or else what is typed.
  $('pvUser').addEventListener('keydown', (e) => { if (e.key === 'Enter' && !pvSuggest.picking()) adLookup(); });

  // Suggestions under a field while typing (from AD or Entra ID): ↑/↓ to move, Enter or a click to
  // take one, Esc to close. multi: a list with commas, the suggestion takes the part being typed.
  function suggest(input, { fetch: get, pick, multi, min = 2 }) {
    const box = el('div', { class: 'suggest', role: 'listbox', id: input.id + 'Suggest', hidden: true });
    document.body.append(box);
    input.setAttribute('role', 'combobox'); input.setAttribute('aria-autocomplete', 'list');
    input.setAttribute('aria-controls', box.id); input.setAttribute('aria-expanded', 'false');
    let items = [], at = -1, timer = null, seq = 0;
    const term = () => (multi ? input.value.split(/[,;]/).pop() : input.value).trim();
    const close = () => { box.hidden = true; items = []; at = -1; input.setAttribute('aria-expanded', 'false'); };
    function place() { const r = input.getBoundingClientRect(); box.style.left = r.left + 'px'; box.style.top = (r.bottom + 2) + 'px'; box.style.minWidth = r.width + 'px'; }
    function render() {
      box.replaceChildren(...items.map((it, i) => el('div', { class: 'sg-item' + (i === at ? ' on' : ''), role: 'option', 'aria-selected': i === at ? 'true' : 'false',
        onmousedown: (e) => { e.preventDefault(); take(i); } }, el('span', { class: 'sg-main' }, it.label), it.sub ? el('small', null, it.sub) : null)));
      box.hidden = !items.length; input.setAttribute('aria-expanded', String(!box.hidden));
      if (!box.hidden) place();
    }
    function take(i) {
      const it = items[i]; if (!it) return;
      if (multi) {
        const before = input.value.replace(/[^,;]*$/, '');
        input.value = (before ? before.trim() + ' ' : '') + it.value + ', ';
      } else input.value = it.value;
      close();
      input.dispatchEvent(new Event('input', { bubbles: true, cancelable: true }));
      close();
      if (pick) pick(it);
    }
    async function load() {
      const q = term(), mine = ++seq;
      if (q.length < min) { close(); return; }
      try { const list = await get(q); if (mine !== seq || document.activeElement !== input) return; items = list.slice(0, 15); at = -1; render(); } catch (e) { close(); }
    }
    input.addEventListener('input', (e) => { if (!e.isTrusted) return; clearTimeout(timer); timer = setTimeout(load, 200); });
    input.addEventListener('focus', () => { if (min === 0 && !input.value.trim()) load(); });
    input.addEventListener('keydown', (e) => {
      if (box.hidden) { if (e.key === 'ArrowDown') { e.preventDefault(); load(); } return; }
      if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); at = (at + (e.key === 'ArrowDown' ? 1 : -1) + items.length) % items.length; render(); }
      else if (e.key === 'Enter' && at >= 0) { e.preventDefault(); take(at); }
      else if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(); }
    });
    input.addEventListener('blur', () => setTimeout(close, 120));
    window.addEventListener('resize', () => { if (!box.hidden) place(); });
    document.addEventListener('scroll', () => { if (!box.hidden) place(); }, true);
    return { open: load, close, picking: () => !box.hidden && at >= 0 };
  }
  const adOn = () => !!(S.info && S.info.features && S.info.features.adLookup);
  const person = (u) => ({ value: u.login, label: u.name || u.login, sub: [u.login, u.mail, u.department].filter((x, i, a) => x && a.indexOf(x) === i).join(' · ') });
  const pvSuggest = suggest($('pvUser'), {
    fetch: async (q) => ($('pvSource').value === 'entra' ? (await api('entra/users', { query: q })).users : adOn() ? (await api('ad/users', { query: q })).users : []).map(person),
    pick: () => adLookup()
  });

