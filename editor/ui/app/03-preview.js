  // ---------------------------------------------------------------------------------------------
  // Preview
  // ---------------------------------------------------------------------------------------------
  let pvTimer = null;
  function schedulePreview() { clearTimeout(pvTimer); pvTimer = setTimeout(renderPreview, 150); }

  function previewValues() {
    // "Sample values, one field empty" (without:key): the sample values, that one left empty.
    const pick = $('pvSource').value, none = pick.startsWith('without:') ? pick.slice(8) : '', src = none ? 'sample' : pick, v = {};
    S.config.fields.forEach((f) => { v[f.key] = core.formatValue(f, src === 'sample' ? (f.sample || '') : (src === 'ad' || src === 'entra') && S.adValues && S.adFrom === src ? (S.adValues[f.key] || '') : ''); });
    // One field empty, the others as they are: which line goes without it.
    if (none && Object.prototype.hasOwnProperty.call(v, none)) v[none] = '';
    return v;
  }

  // The day the preview shows (today, or the one set with the calendar), and its campaign.
  const pvNow = () => ($('pvDayOn').checked && $('pvDay').value ? new Date($('pvDay').value + 'T12:00:00') : new Date());
  function pvCampaign(s) {
    const src = $('pvSource').value, real = (src === 'ad' || src === 'entra') && S.adValues && S.adFrom === src;
    return core.campaignFor(S.config, s, { now: pvNow(), kind: 'new', anyone: !real, values: real ? S.adValues : previewValues(), from: real ? S.adValues.mail || '' : '', groups: real ? S.adGroups || [] : null });
  }
  // A signature compiled for the preview, with the banner of the campaign running that day.
  function pvCompiled(s) {
    const c = pvCampaign(s), v = c ? (c.langs && c.langs[core.langOf(previewValues())]) || c : null;
    return core.fillCampaign(core.compile(sanitize(sigFull(s)), scratchDoc), v ? core.compile(sanitize(v.html || ''), scratchDoc) : '', c && c.position);
  }
  $('pvDayOn').addEventListener('change', () => { $('pvDayBox').hidden = !$('pvDayOn').checked; if ($('pvDayOn').checked && !$('pvDay').value) $('pvDay').value = core.today(); if ($('pvDayOn').checked) $('pvDay').focus(); renderPreview(); });
  $('pvDay').addEventListener('input', renderPreview);

  // The signature as the preview shows it (its values, the building blocks put in), with its images
  // inside (data: addresses), as one file: to try it in a mail program without publishing.
  async function exportHtml() {
    const s = sig(); if (!s) return null;
    const html = core.render(pvCompiled(s), previewValues(), { assetBase: '/', removeEmptyLines: s.removeEmptyLines !== false });
    const doc = new DOMParser().parseFromString('<!doctype html><body>' + html + '</body>', 'text/html');
    await Promise.all([...doc.querySelectorAll('img[src]')].map(async (img) => {
      // Only the signatures' own images and photos, and only what is an image.
      if (!/^\/(assets|photos)\/[^/?#]+(\?[^#]*)?$/.test(img.getAttribute('src'))) { if (/^\//.test(img.getAttribute('src'))) img.remove(); return; }
      try {
        const r = await fetch(img.getAttribute('src'), { credentials: 'same-origin' });
        if (!r.ok) { img.remove(); return; }
        const b = await r.blob();
        if (!/^image\//.test(b.type)) { img.remove(); return; }
        img.src = await dataUrlOf(b);
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

  // All signatures side by side, each as the preview shows it (the same values): to see that they
  // look alike before publishing. A click opens one.
  function openAll() {
    flushSource();
    const src = $('pvSource');
    $('allLead').textContent = t('With the values: {0}. Change them in the preview (above the preview, left).', src.options[src.selectedIndex].textContent + (src.value === 'ad' || src.value === 'entra' ? ' ' + ($('pvUser').value || '') : ''));
    const values = previewValues();
    $('allList').replaceChildren(...(S.config.signatures.length ? S.config.signatures.map((s) => {
      const html = core.render(pvCompiled(s), values, { assetBase: '/', removeEmptyLines: s.removeEmptyLines !== false });
      // Its images from this server (signed in): same origin, still no scripts. A photo that is not
      // there goes, as in the preview.
      const frame = el('iframe', { tabindex: '-1', 'aria-hidden': 'true', sandbox: 'allow-same-origin' });
      frame.addEventListener('load', () => { const d = frame.contentDocument; if (d) d.querySelectorAll('img').forEach((i) => { if (i.complete && !i.naturalWidth) i.remove(); else i.addEventListener('error', () => i.remove()); }); });
      frame.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:12px 14px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222;zoom:.8}p{margin:0}table{border-collapse:collapse}</style></head><body>' + html + '</body></html>';
      const b = el('button', { type: 'button', class: 'tpl-card', 'data-sig': s.id }, frame,
        el('span', { class: 'tpl-name' }, s.name || t('(no name)')),
        el('span', { class: 'tpl-text' }, [s.defaultNew ? t('New') : '', s.defaultReply ? t('Reply') : '', s.language ? s.language.toUpperCase() : '', audienceText(s)].filter(Boolean).join(' · ')));
      b.addEventListener('click', () => { $('dlgAll').close(); openSig(s.id); });
      return b;
    }) : [emptyState('blocks', t('No signatures yet'), '')]));
    $('dlgAll').showModal();
  }
  $('btnAll').addEventListener('click', openAll);

  function renderPreview() {
    const s = sig(); if (!s || S.view !== 'sig') return;
    const html = core.render(pvCompiled(s), previewValues(), { assetBase: '/', removeEmptyLines: s.removeEmptyLines !== false });
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
            el('button', { type: 'button', class: 'linklike', onclick: () => photoAddFor(name) }, t('Add the photo'))));
        }
      };
      d.querySelectorAll('img[src*="/photos/"]').forEach((img) => { if (img.complete && !img.naturalWidth) gone(img); else img.addEventListener('error', () => gone(img)); });
    };

    // What to know about this signature before it goes out.
    const notes = [], keys = fieldKeys(), used = core.placeholdersIn(sigFull(s));
    const unknown = used.filter((k) => keys.indexOf(k) < 0);
    const noGraph = used.filter((k) => { const f = S.config.fields.find((x) => x.key === k); return f && core.fieldSource(S.config, f) === 'none'; });
    if (unknown.length) notes.push(el('div', { class: 'warn' }, t('Not a field: {0} - left as typed.', unknown.map((k) => '@@' + k).join(', ')) + ' ',
      el('button', { type: 'button', class: 'linklike', onclick: () => {
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
    // The campaign shown: its banner, or why there is none (on another day: whether the signature is used then).
    const camp = pvCampaign(s), loc = I18N.lang === 'de' ? 'de-CH' : 'en-GB', onDay = $('pvDayOn').checked && $('pvDay').value;
    const dayText = onDay ? pvNow().toLocaleDateString(loc) : '';
    if (onDay && core.isCampaign(s) && !core.inPeriod(s, pvNow())) notes.push(el('div', { class: 'warn' }, t('On {0} this signature is not used (its own days: {1}).', dayText, campaignText(s))));
    if (camp) notes.push(el('div', null, (onDay ? t('On {0}: the banner of «{1}».', dayText, camp.name) : t('Banner of the campaign «{0}», running now.', camp.name)) + ' ',
      el('button', { type: 'button', class: 'linklike', onclick: () => { campSel = camp.id; goView('campaigns'); } }, t('Open'))));
    else if (onDay) notes.push(el('div', null, t('On {0}: no campaign banner.', dayText)));
    // Classic Outlook: what it keeps in the mailbox, when it gets tight.
    if (s.defaultNew || s.defaultReply) {
      const cp = classicParts(s), kb = cp.kb;
      // Said as it will be there: without which images (the largest first), or not at all.
      // What the room goes to, and what helps with each.
      const share = ' ' + t('Of it: text ~{0} KB, photo ~{1} KB, images ~{2} KB.', cp.textKb, cp.photoKb, Math.round(cp.images.reduce((n, x) => n + x.bytes, 0) / 1024));
      const act = el('span', null, share, ' ', el('button', { type: 'button', class: 'linklike classic-prep', onclick: () => prepareClassic() }, t('Prepare for classic Outlook')),
        // The further way, when the text fits and images would be left out: as small as it takes.
        ...(cp.textFits && cp.missing.length ? [' · ', el('button', { type: 'button', class: 'linklike classic-fit', title: t('Makes the images as small in KB as it takes for all of it to fit; they look less sharp, in every mail.'), onclick: () => fitClassic() }, t('Images small enough to fit'))] : []));
      if (kb > CLASSIC_KB && !cp.textFits) notes.push(el('div', { class: 'warn' }, t('Classic Outlook: even the text of this signature is more than the {0} KB its mailbox keeps: there it is not put in by itself, only with Insert in the pane.', CLASSIC_KB), act));
      else if (kb > CLASSIC_KB && cp.missing.length) notes.push(el('div', { class: 'warn' }, t('Classic Outlook: about {0} KB, more than the {1} KB its mailbox keeps. When it puts the signature in by itself, it goes without: {2}. The rest fits; Insert in the pane puts in everything.', kb, CLASSIC_KB, cp.missing.map((x) => x + ' (~' + Math.round((cp.images.find((y) => y.name === x) || {}).bytes / 1024) + ' KB)').join(', ')), act));
      // The images in full from next to the add-in's files: said once it would not fit otherwise.
      else if (cp.hosted && kb > CLASSIC_KB) notes.push(el('div', null, t('Classic Outlook: the images come in as they are from the add-in\'s address (Cloudflare Pages), not from the mailbox: all of it goes in, sharp. Without a network there, it takes the smaller copies kept in the mailbox.')));
      // It fits with those images made smaller there: said, nothing to do.
      else if (kb > CLASSIC_KB) notes.push(el('div', null, t('Classic Outlook: about {0} KB; so that all of it fits into the {1} KB its mailbox keeps, the add-in makes {2} smaller there (only there: every other Outlook gets them as they are).', kb, CLASSIC_KB, cp.squeezed.join(', '))));
      else if (kb > CLASSIC_KB * 0.7) notes.push(el('div', null, t('Classic Outlook: about {0} of {1} KB, nearly full.', kb, CLASSIC_KB)));
    }
    const none = $('pvSource').value.startsWith('without:') ? $('pvSource').value.slice(8) : '';
    if (none && used.includes(none)) notes.push(emptyFieldNote(s, none));
    $('pvNote').replaceChildren(...notes);
    fillEmptyPick(s, used);
  }
  // "Sample values, one field empty": the fields this signature uses, to leave one of them out.
  // One gone from the signature: back to the sample values.
  function fillEmptyPick(s, used) {
    const src = $('pvSource'), group = $('pvWithout'), keep = src.value, keys = used.filter((k) => fieldKeys().includes(k));
    const names = keys.map((k) => k + '=' + fieldName(k)).join();
    if (group.dataset.keys === names) return;
    group.dataset.keys = names;
    group.label = t('Sample values, one field empty');
    group.replaceChildren(...keys.map((k) => el('option', { value: 'without:' + k }, t('Without {0}', fieldName(k)))));
    group.hidden = !keys.length;
    if (keep.startsWith('without:')) { src.value = keys.includes(keep.slice(8)) ? keep : 'sample'; if (src.value !== keep) schedulePreview(); }
  }
  // What goes when one field is empty: the lines the preview loses, in words.
  function emptyFieldNote(s, key) {
    const html = (vals) => core.render(pvCompiled(s), vals, { assetBase: '/', removeEmptyLines: s.removeEmptyLines !== false });
    const without = previewValues(), withIt = Object.assign({}, without, { [key]: (S.config.fields.find((f) => f.key === key) || {}).sample || 'x' });
    const lines = (h) => core.htmlToText(h).split('\n').map((l) => l.trim()).filter(Boolean);
    const a = lines(html(withIt)), b = lines(html(without));
    // Fewer lines without it: those are what goes (a line kept just lacks the value).
    const gone = b.length < a.length ? a.filter((l) => !b.includes(l)).slice(0, a.length - b.length) : [];
    if (s.removeEmptyLines === false) return el('div', null, t('Without {0}: its lines stay, empty (this signature keeps empty lines).', fieldName(key)));
    if (!gone.length) return el('div', null, t('Without {0}: no line goes, only the value is missing.', fieldName(key)));
    return el('div', null, t('Without {0}, this goes: {1}', fieldName(key), gone.map((l) => '«' + l + '»').join(', ')));
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

