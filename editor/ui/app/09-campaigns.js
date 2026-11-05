  // ---------------------------------------------------------------------------------------------
  // Campaigns: a banner for some days, put into the signatures that are there (core.withCampaign),
  // made here without HTML: an image with a link, a line of text with a button, an image beside a
  // text, or HTML of one's own; in other languages too. With their days on a timeline, what a
  // signature gets on a day, a note when one ends with none after it, and a copy for next year.
  // config.campaigns: see shared/signature-core.js; design is what this form holds, html is made of it.
  // ---------------------------------------------------------------------------------------------
  const CAMP_LANGS = { de: 'Deutsch', en: 'English', fr: 'Français', it: 'Italiano' };
  const CAMP_TPL = ['image', 'text', 'imagetext', 'html'];
  // campAlt: which of the banners taking turns the form shows (0: the first).
  let campSel = null, campLang = '', campAlt = 0, campPast = false;
  const WEEKDAYS = [1, 2, 3, 4, 5, 6, 0];
  const isOver = (c) => !!c.validUntil && c.validUntil < core.today();
  const campaigns = () => (S.config.campaigns = Array.isArray(S.config.campaigns) ? S.config.campaigns : []);
  const campById = (id) => campaigns().find((c) => c.id === id) || null;
  const okLink = (u) => /^(https?:|mailto:|tel:)/i.test(String(u || '').trim());

  // A campaign as the form starts it: from today, for two weeks, in all signatures, for everyone.
  function newCampaign(name) {
    const d = new Date(), from = core.today(d); d.setDate(d.getDate() + 13);
    return { id: newId('c-' + name), name, validFrom: from, validUntil: core.today(d), design: { tpl: 'image', img: '', width: 400, alt: '', link: 'https://', text: '', button: '', color: houseColours()[0] || '#1d4ed8', html: '' },
      html: '', utm: true, signatures: [], position: 'below', newOnly: false, audience: { departments: [], domains: [], groups: [], everyone: true }, langs: {} };
  }
  // The banner's HTML from the form: tables and inline styles, as Outlook (Word) shows them.
  function bannerHtml(d, c) {
    const e = core.escapeHtml, link = okLink(d.link) && d.link.trim() !== 'https://' ? d.link.trim() : '';
    const w = Math.max(40, Math.min(800, parseInt(d.width, 10) || 400)), color = /^#[0-9a-f]{6}$/i.test(d.color || '') ? d.color : '#1d4ed8';
    const img = d.img ? '<img src="assets/' + e(d.img) + '" width="' + w + '" alt="' + e(d.alt || '') + '" style="border:0;display:block">' : '';
    const a = (inner, style) => (link ? '<a href="' + e(link) + '"' + (style ? ' style="' + style + '"' : '') + '>' + inner + '</a>' : inner);
    const text = e(d.text || '').replace(/\n/g, '<br>');
    const btn = d.button ? '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tr><td bgcolor="' + color + '" style="background:' + color + ';border-radius:4px;padding:6px 14px">' +
      a(e(d.button), 'color:#ffffff;font-family:Arial,sans-serif;font-size:10pt;font-weight:bold;text-decoration:none') + '</td></tr></table>' : '';
    let html = '';
    if (d.tpl === 'html') html = sanitize(d.html || '');
    else if (d.tpl === 'text') html = '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:10px"><tr>' +
      '<td style="padding:0 14px 0 0;vertical-align:middle;font-family:Arial,sans-serif;font-size:10pt;color:' + color + ';font-weight:bold">' + text + '</td>' +
      (btn ? '<td style="vertical-align:middle">' + btn + '</td>' : '') + '</tr></table>';
    else if (d.tpl === 'imagetext') html = '<table cellpadding="0" cellspacing="0" style="border-collapse:collapse;margin-top:10px"><tr>' +
      (img ? '<td style="padding:0 14px 0 0;vertical-align:middle">' + a(img) + '</td>' : '') +
      '<td style="vertical-align:middle;font-family:Arial,sans-serif;font-size:10pt;color:#222222"><span style="color:' + color + ';font-weight:bold">' + text + '</span>' + (btn ? '<div style="margin-top:6px">' + btn + '</div>' : '') + '</td></tr></table>';
    else html = img ? '<div style="margin-top:10px">' + a(img) + '</div>' : '';
    return c.utm ? core.withUtm(html, c.name) : html;
  }
  // The banner made again from the form, in every language.
  function campBuild(c) {
    c.html = bannerHtml(c.design || {}, c);
    Object.values(c.langs || {}).forEach((v) => { v.html = bannerHtml(v.design || {}, c); });
    (c.more || []).forEach((v) => { v.html = bannerHtml(v.design || {}, c); });
  }
  const pickCamp = (id) => { campSel = id; campLang = ''; campAlt = 0; };
  // A change: kept at once (made into HTML, saved a moment later); what shows it drawn again now
  // (a click) or once typing pauses (campSoon: the list, timeline, preview and day not at every key).
  let campTimer = null;
  const campRedraw = () => { clearTimeout(campTimer); campTimer = null; campList(); campTimeline(); campPreview(); campDayTable(); renderStatusBar(); };
  const campKeep = () => { const c = campById(campSel); if (c) campBuild(c); changed({ list: false, preview: false }); };
  const campChanged = () => { campKeep(); campRedraw(); };
  const campSoon = () => { campKeep(); clearTimeout(campTimer); campTimer = setTimeout(campRedraw, 250); };

  // ---- the view --------------------------------------------------------------------------------
  function renderCampaignsView() {
    const list = campaigns();
    if (campSel && !campById(campSel)) campSel = null;
    if (!campSel && list.length) campSel = list[0].id;
    campTimeline(); campDayTable(); campList(); campEditor();
  }
  function campList() {
    const list = campaigns(), box = $('campList');
    $('btnAddCampaign').parentNode.hidden = !list.length;
    if (!list.length) {
      box.replaceChildren(emptyState('blocks', t('No banner campaigns yet.'), t('A banner for some days (the holidays, a fair, a webinar) in the signatures there are: no copy of a signature.'), [t('+ New campaign'), () => $('btnAddCampaign').click()]));
      return;
    }
    const item = (c) => el('button', { type: 'button', class: 'camp-item' + (c.id === campSel ? ' on' : ''), 'aria-pressed': String(c.id === campSel), 'data-camp': c.id, onclick: () => { pickCamp(c.id); campList(); campEditor(); campTimeline(); } },
      el('b', null, c.name || t('(no name)')), el('small', null, [campaignText(c) || t('always'), core.inPeriod(c) ? t('running') : isOver(c) ? t('over') : '', (c.more || []).length ? t('{0} banners in turns', c.more.length + 1) : ''].filter(Boolean).join(' · ')));
    // The ones over, folded away (the one selected stays in view).
    const now = list.filter((c) => !isOver(c)), past = list.filter(isOver);
    const open = campPast || past.some((c) => c.id === campSel);
    const fold = past.length ? el('button', { type: 'button', class: 'camp-fold', 'aria-expanded': String(open), onclick: () => { campPast = !open; campList(); campTimeline(); } }, (open ? '▾ ' : '▸ ') + t('Over ({0})', past.length)) : null;
    box.replaceChildren(...now.map(item), ...(fold ? [fold] : []), ...(open ? past.map(item) : []));
  }
  // Up and down in the list: the next campaign (as in the list of signatures).
  $('campList').addEventListener('keydown', (e) => {
    if (e.key !== 'ArrowDown' && e.key !== 'ArrowUp') return;
    const items = [...$('campList').querySelectorAll('.camp-item, .camp-fold')], i = items.indexOf(document.activeElement);
    if (i < 0) return;
    e.preventDefault();
    const n = items[Math.max(0, Math.min(items.length - 1, i + (e.key === 'ArrowDown' ? 1 : -1)))];
    if (n) n.focus();
  });

  // The form of the campaign selected: its head, then a card each for when, the banner, where and
  // who (campWhen … campWho), the preview beside. f: the helpers the cards share.
  function campEditor() {
    const box = $('campEdit'), c = campById(campSel);
    box.hidden = !c;
    if (!c) { box.replaceChildren(); return; }
    c.design = c.design || { tpl: 'image' }; c.langs = c.langs || {}; c.audience = c.audience || { everyone: true };
    if (campLang && !c.langs[campLang]) campLang = '';
    c.more = Array.isArray(c.more) ? c.more : [];
    if (campLang || campAlt > c.more.length) campAlt = 0;
    const d = campLang ? c.langs[campLang].design : campAlt ? c.more[campAlt - 1].design : c.design;
    const f = campHelpers();
    box.replaceChildren(campHead(c, f),
      el('div', { class: 'camp-body' },
        el('div', { class: 'camp-form' }, campWhen(c, f), campBanner(c, d, f), campWhere(c, f), campWho(c, f)),
        el('aside', { class: 'camp-side' }, el('h3', null, t('Preview')),
          el('iframe', { id: 'campPv', class: 'camp-pv', sandbox: 'allow-same-origin', title: t('Preview'), tabindex: '-1' }),
          el('p', { class: 'hint' }, t('Under the first signature it goes into, with the sample values.')))));
    campPreview();
  }
  // What the cards are made of: a field (typing redraws once it pauses), a label beside or above
  // it, a card.
  function campHelpers() {
    const input = (attrs, get, set) => {
      const i = el('input', attrs); i.value = get() == null ? '' : get();
      const box = attrs.type === 'checkbox';
      i.addEventListener(box ? 'change' : 'input', () => { set(box ? i.checked : i.value); if (box) campChanged(); else campSoon(); });
      if (box) i.checked = !!get();
      return i;
    };
    const row = (label, ...kids) => el('label', { class: 'camp-f' }, el('span', null, label), ...kids);
    const stack = (label, ...kids) => el('label', { class: 'camp-stack' }, el('span', null, label), ...kids);
    const sect = (title, ...kids) => el('section', { class: 'camp-card' }, el('h3', null, title), ...kids.filter((k) => k !== null && k !== undefined));
    return { input, row, stack, sect };
  }
  // The head: the name, whether it runs, and to copy or remove it.
  function campHead(c, f) {
    const name = f.input({ 'aria-label': t('Name'), class: 'camp-title', placeholder: t('Name of the campaign') }, () => c.name, (v) => { c.name = v; });
    const copy = el('button', { type: 'button', class: 'ghost', title: t('The same campaign one year later (the year in its name too)'), onclick: () => {
      const n = JSON.parse(JSON.stringify(c)), y = (x) => { if (!x) return x; const dd = new Date(x + 'T12:00:00'); dd.setFullYear(dd.getFullYear() + 1); return core.today(dd); };
      n.name = /\b(19|20)\d\d\b/.test(c.name) ? c.name.replace(/\b((?:19|20)\d\d)\b/g, (m) => String(Number(m) + 1)) : c.name + ' ' + t('(next year)');
      n.id = newId('c-' + n.name); n.validFrom = y(c.validFrom); n.validUntil = y(c.validUntil);
      campaigns().push(n); pickCamp(n.id); campBuild(n); changed({ list: false }); renderCampaignsView();
      toast(t('Copied for next year: {0}, {1}.', n.name, campaignText(n) || t('always')));
    } }, t('Copy for next year'));
    const del = el('button', { type: 'button', class: 'ghost danger', onclick: () => {
      const i = campaigns().indexOf(c); campaigns().splice(i, 1); campSel = null; changed({ list: false }); renderCampaignsView();
      toast(t('Campaign {0} removed.', c.name), false, { label: t('Undo'), run: () => { campaigns().splice(Math.min(i, campaigns().length), 0, c); campSel = c.id; changed({ list: false }); renderCampaignsView(); } });
    } }, t('Remove'));

    const state = core.inPeriod(c) && (!c.days || !c.days.length || c.days.includes(new Date().getDay())) ? ['now', t('running')] : isOver(c) ? ['over', t('over')] : ['soon', t('planned')];
    return el('div', { class: 'camp-head' }, name, el('span', { class: 'camp-state ' + state[0] }, state[1]), el('span', { class: 'spacer' }), copy, del);
  }
  // When: its first and last day, and the weekdays.
  function campWhen(c, f) {
    const { input } = f;
    // When
    const from = input({ type: 'date', 'aria-label': t('From') }, () => c.validFrom, (v) => { c.validFrom = v; campList(); });
    const until = input({ type: 'date', 'aria-label': t('Until') }, () => c.validUntil, (v) => { c.validUntil = v; campList(); });
    // Only on some weekdays (none ticked: every day).
    const dayName = (n) => new Date(2026, 0, 4 + n).toLocaleDateString(I18N.lang === 'de' ? 'de-CH' : 'en-GB', { weekday: 'short' });
    const days = el('div', { class: 'camp-days', role: 'group', 'aria-label': t('Only on these weekdays') },
      ...WEEKDAYS.map((n) => {
        const b = el('input', { type: 'checkbox', 'aria-label': dayName(n) }); b.checked = (c.days || []).includes(n);
        b.addEventListener('change', () => { const set = new Set(c.days || []); if (b.checked) set.add(n); else set.delete(n); c.days = WEEKDAYS.filter((x) => set.has(x)); if (!c.days.length) delete c.days; campChanged(); });
        return el('label', { class: 'chip' }, b, dayName(n));
      }));

    return f.sect(t('When'),
      el('div', { class: 'camp-grid2' }, f.stack(t('From'), from), f.stack(t('Until'), until)),
      el('div', { class: 'camp-stack' }, el('span', null, t('Weekdays'), ' ', el('small', { class: 'hint' }, '(' + t('none ticked: every day') + ')')), days));
  }
  // The banner: its language and the ones in turns, its kind, its parts.
  function campBanner(c, d, f) {
    const { input, row } = f;
    // The banner: its kind, then its parts.
    const tplNames = { image: t('Image with link'), text: t('Text with button'), imagetext: t('Image and text'), html: t('HTML of your own') };
    const tpls = el('div', { class: 'seg camp-tpl', role: 'radiogroup', 'aria-label': t('Kind of banner') }, CAMP_TPL.map((k) => {
      const r = el('input', { type: 'radio', name: 'campTpl', value: k }); r.checked = (d.tpl || 'image') === k;
      r.addEventListener('change', () => { d.tpl = k; campChanged(); campEditor(); });
      return el('label', null, r, tplNames[k]);
    }));
    const parts = [];
    if (d.tpl !== 'text' && d.tpl !== 'html') {
      const pick = el('select', { 'aria-label': t('Image') }, el('option', { value: '' }, t('- choose an image -')), (S.assets || []).map((a) => el('option', { value: a.name }, a.name)));
      pick.value = d.img || '';
      pick.addEventListener('change', () => { d.img = pick.value; campChanged(); });
      const up = el('input', { type: 'file', accept: 'image/png,image/jpeg,image/gif', hidden: true });
      up.addEventListener('change', async () => {
        let f = up.files[0]; up.value = ''; if (!f) return;
        f = await shrinkUpload(f);
        if (f.size > 1024 * 1024) { toast(t('The image is larger than 1 MB.'), true); return; }
        const data = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(',')[1] || ''); r.onerror = rej; r.readAsDataURL(f); });
        try {
          const r = await api('assets/upload', { name: f.name, data });
          S.assets = S.assets.filter((a) => a.name !== r.name).concat([{ name: r.name, size: r.size }]).sort((a, b) => a.name.localeCompare(b.name));
          d.img = r.name;
          // As wide as it is (at most the width set), so it is not blown up.
          const i = new Image(); i.onload = () => { if (i.naturalWidth && i.naturalWidth < (parseInt(d.width, 10) || 400)) d.width = i.naturalWidth; campChanged(); campEditor(); }; i.onerror = () => { campChanged(); campEditor(); }; i.src = '/assets/' + encodeURIComponent(r.name);
        } catch (e) { toast(t(e.message), true); }
      });
      parts.push(row(t('Image'), pick, el('button', { type: 'button', class: 'ghost small', onclick: () => up.click() }, t('Upload…')), up),
        row(t('Width'), input({ type: 'number', min: 40, max: 800, class: 'num' }, () => d.width, (v) => { d.width = v; }), el('span', { class: 'unit' }, 'px')),
        row(t('Alt text'), input({ placeholder: t('what the image says, for those who see no images') }, () => d.alt, (v) => { d.alt = v; })));
    }
    if (d.tpl === 'text' || d.tpl === 'imagetext') {
      const ta = el('textarea', { rows: 2, 'aria-label': t('Text') }); ta.value = d.text || '';
      ta.addEventListener('input', () => { d.text = ta.value; campSoon(); });
      const col = el('select', { 'aria-label': t('Colour') }, [...new Set(houseColours().concat([d.color || '#1d4ed8']))].map((x) => el('option', { value: x, style: 'color:' + x }, '■ ' + x)));
      col.value = d.color || '#1d4ed8';
      const own = el('input', { type: 'color', 'aria-label': t('Another colour'), value: /^#[0-9a-f]{6}$/i.test(d.color || '') ? d.color : '#1d4ed8' });
      col.addEventListener('change', () => { d.color = col.value; own.value = col.value; campChanged(); });
      own.addEventListener('input', () => { d.color = own.value; campSoon(); });
      own.addEventListener('change', () => campEditor());
      parts.push(row(t('Text'), ta), row(t('Button'), input({ placeholder: t('e.g. Sign up now (empty: none)') }, () => d.button, (v) => { d.button = v; })),
        row(t('Colour'), col, own, houseColours().length ? null : el('small', { class: 'hint' }, t('House colours: the arrow beside the text colour in the editor.'))));
    }
    if (d.tpl === 'html') {
      const ta = el('textarea', { rows: 6, class: 'mono', 'aria-label': t('HTML') }); ta.value = d.html || '';
      ta.addEventListener('input', () => { d.html = ta.value; campSoon(); });
      parts.push(row(t('HTML'), ta));
    }
    if (d.tpl !== 'html') {
      const li = input({ placeholder: 'https://' }, () => d.link, (v) => { d.link = v; });
      li.addEventListener('change', () => { if (li.value.trim() && li.value.trim() !== 'https://' && !okLink(li.value)) toast(t('A link starts with https://, mailto: or tel:.'), true); });
      parts.push(row(t('Link'), li));
    }
    parts.push(el('label', { class: 'camp-chk', title: t('Adds utm_source=email, utm_medium=signature and utm_campaign=<name> to the web addresses: the web statistics then show the visits that came from the signatures.') },
      input({ type: 'checkbox' }, () => c.utm, (v) => { c.utm = v; }), ' ', t('Link tracking (utm_campaign={0})', core.campaignSlug(c.name))));

    // Languages: the banner for all, and one per language for those who have it in Entra ID / AD.
    const langTabs = el('div', { class: 'camp-langs', role: 'tablist', 'aria-label': t('Language') },
      el('button', { type: 'button', role: 'tab', class: 'tab' + (!campLang ? ' active' : ''), 'aria-selected': String(!campLang), onclick: () => { campLang = ''; campAlt = 0; campEditor(); campPreview(); } }, t('All languages')),
      ...Object.keys(c.langs).map((k) => el('span', { class: 'camp-lang' },
        el('button', { type: 'button', role: 'tab', class: 'tab' + (campLang === k ? ' active' : ''), 'aria-selected': String(campLang === k), onclick: () => { campLang = k; campEditor(); campPreview(); } }, CAMP_LANGS[k] || k),
        el('button', { type: 'button', class: 'ghost small x', title: t('Remove this language'), 'aria-label': t('Remove this language'), onclick: () => { delete c.langs[k]; campLang = ''; campChanged(); campEditor(); } }, '×'))));
    // Banners that take turns, mail by mail (for all languages; one in a language is that one).
    const altTabs = campLang ? null : el('div', { class: 'camp-langs camp-alts', role: 'tablist', 'aria-label': t('Banners in turns') },
      ...[c.design].concat(c.more.map((x) => x.design)).map((x, i) => el('span', { class: 'camp-lang' },
        el('button', { type: 'button', role: 'tab', class: 'tab' + (campAlt === i ? ' active' : ''), 'aria-selected': String(campAlt === i), onclick: () => { campAlt = i; campEditor(); } }, t('Banner {0}', i + 1)),
        i ? el('button', { type: 'button', class: 'ghost small x', title: t('Remove this banner'), 'aria-label': t('Remove this banner'), onclick: () => { c.more.splice(i - 1, 1); campAlt = 0; campChanged(); campEditor(); } }, '×') : null)),
      c.more.length < 9 ? el('button', { type: 'button', class: 'ghost small', title: t('Another banner: they take turns, mail by mail'), onclick: () => { c.more.push({ design: JSON.parse(JSON.stringify(d)), html: '' }); campAlt = c.more.length; campChanged(); campEditor(); } }, t('+ Banner in turns')) : null);
    const missing = Object.keys(CAMP_LANGS).filter((k) => !c.langs[k]);
    if (missing.length) {
      const add = el('select', { 'aria-label': t('+ Language'), class: 'camp-addlang' }, el('option', { value: '' }, t('+ Language')), missing.map((k) => el('option', { value: k }, CAMP_LANGS[k])));
      add.addEventListener('change', () => { const k = add.value; if (!k) return; c.langs[k] = { design: JSON.parse(JSON.stringify(c.design)), html: '' }; campLang = k; campChanged(); campEditor(); });
      langTabs.append(add);
    }

    return f.sect(t('Banner'),
      el('div', { class: 'camp-variant' }, el('span', { class: 'camp-lbl' }, t('Language')), langTabs),
      altTabs ? el('div', { class: 'camp-variant' }, el('span', { class: 'camp-lbl' }, t('In turns')), altTabs) : null,
      el('div', { class: 'camp-variant' }, el('span', { class: 'camp-lbl' }, t('Kind')), tpls),
      el('div', { class: 'camp-fields' }, ...parts));
  }
  // Where: in which signatures, where in them, in new mail only.
  function campWhere(c, f) {
    const { input } = f;
    // Where: which signatures, where in them, only in new mail.
    const sigs = S.config.signatures, targets = c.signatures || [];
    const all = el('input', { type: 'radio', name: 'campWhere' }), some = el('input', { type: 'radio', name: 'campWhere' });
    all.checked = !targets.length; some.checked = !!targets.length;
    all.addEventListener('change', () => { c.signatures = []; campChanged(); campEditor(); });
    some.addEventListener('change', () => { c.signatures = sigs.slice(0, 1).map((s) => s.id); campChanged(); campEditor(); });
    const sigBoxes = targets.length ? el('div', { class: 'camp-sigs' }, sigs.map((s) => {
      const b = el('input', { type: 'checkbox' }); b.checked = targets.includes(s.id);
      b.addEventListener('change', () => { c.signatures = b.checked ? targets.concat([s.id]) : targets.filter((x) => x !== s.id); if (!c.signatures.length) c.signatures = [s.id]; campChanged(); campEditor(); });
      return el('label', { class: 'chip' }, b, s.name || t('(no name)'));
    })) : null;
    const meant = targets.length ? sigs.filter((s) => targets.includes(s.id)) : sigs;
    const noSlot = meant.filter((s) => !core.hasSlot(s.html));
    const pos = el('select', { 'aria-label': t('Where in the signature') }, el('option', { value: 'below' }, t('below the signature')), el('option', { value: 'above' }, t('above the signature')));
    pos.value = c.position === 'above' ? 'above' : 'below';
    pos.addEventListener('change', () => { c.position = pos.value; campChanged(); });
    const role = !!(S.info && S.info.user && S.info.user.campaigns);
    // Where in them: in the place {{campaign}}; the ones without it, below or above (said in one
    // line, the names folded).
    const addPlace = role ? el('small', { class: 'hint' }, t('An editor can put the place into them.')) :
      el('button', { type: 'button', class: 'ghost small', title: t('Puts {{campaign}} at the end of these signatures: then you can move it where it should go'), onclick: () => {
        noSlot.forEach((s) => { s.html = String(s.html || '').replace(/\s*$/, '') + '\n<div>{{campaign}}</div>'; });
        changed(); campEditor(); campPreview();
        toast(t('The place {{campaign}} is in {0} now: move it in the design where the banner should go.', noSlot.map((s) => s.name).join(', ')));
      } }, t('Add the place to them'));
    const slotNote = noSlot.length
      ? el('div', { class: 'camp-slot' },
        el('details', null, el('summary', null, noSlot.length === meant.length ? t('None of them has the place {{campaign}}') : t('{0} of {1} without the place {{campaign}}', noSlot.length, meant.length)),
          el('p', { class: 'hint' }, noSlot.map((x) => x.name).join(', '))),
        el('span', { class: 'camp-slot-pos' }, t('There the banner goes'), pos), addPlace)
      : el('div', { class: 'camp-slot ok' }, '✓ ' + t('Each of them has the place {{campaign}}: the banner goes there.'));

    return f.sect(t('Where'),
      el('div', { class: 'seg camp-where', role: 'radiogroup', 'aria-label': t('Where') }, el('label', null, all, t('In all signatures')), el('label', null, some, t('Only in these'))),
      sigBoxes, slotNote,
      el('label', { class: 'camp-chk' }, input({ type: 'checkbox' }, () => c.newOnly, (v) => { c.newOnly = v; }), t('Only in new mail, not in replies and forwards')));
  }
  // Who: groups and departments, or everyone.
  function campWho(c, f) {
    const { input } = f;
    // Who: everyone, or groups and departments (as with signatures).
    const a = c.audience;
    const list = (k) => input({ placeholder: k === 'groups' ? t('e.g. Sales, Marketing') : t('e.g. Sales') }, () => (a[k] || []).join(', '), (v) => { a[k] = v.split(',').map((x) => x.trim()).filter(Boolean); a.everyone = !(a.groups || []).length && !(a.departments || []).length; });

    return f.sect(t('Who'),
      el('div', { class: 'camp-grid2' }, f.stack(t('Groups'), list('groups')), f.stack(t('Departments'), list('departments'))),
      el('p', { class: 'hint' }, t('Both empty: everyone who gets one of these signatures.')));
  }

  // The banner under the first signature it is meant for, with the sample values.
  function campPreview() {
    const frame = $('campPv'), c = campById(campSel); if (!frame || !c) return;
    const meant = (c.signatures || []).length ? S.config.signatures.filter((s) => c.signatures.includes(s.id)) : S.config.signatures;
    const s = meant[0], values = {};
    S.config.fields.forEach((f) => { values[f.key] = core.formatValue(f, f.sample || ''); });
    const shown = campLang && c.langs[campLang] ? c.langs[campLang] : campAlt && (c.more || [])[campAlt - 1] ? c.more[campAlt - 1] : c;
    const banner = core.compile(sanitize(shown.html || ''), scratchDoc);
    const base = s ? core.compile(sanitize(sigFull(s)), scratchDoc) : '';
    const html = core.render(core.fillCampaign(base, banner, c.position), values, { assetBase: '/', removeEmptyLines: !s || s.removeEmptyLines !== false });
    // A person's photo that is not there goes, as in the preview.
    frame.onload = () => { const dd = frame.contentDocument; if (dd) dd.querySelectorAll('img').forEach((i) => { if (i.complete && !i.naturalWidth) i.remove(); else i.addEventListener('error', () => i.remove()); }); };
    frame.srcdoc = '<!doctype html><html><head><meta charset="utf-8"><style>body{margin:12px 14px;font-family:Aptos,Calibri,Arial,sans-serif;font-size:11pt;color:#222;zoom:.85}p{margin:0}table{border-collapse:collapse}</style></head><body>' + (html || '<p style="color:#888">' + core.escapeHtml(t('No signature yet.')) + '</p>') + '</body></html>';
  }

  // A test run for one person (from AD or Entra ID): on that day, which signature they get for new
  // mail and for replies, with which banner, and whether it fits classic Outlook.
  let campPerson = null;
  async function campWhoLookup() {
    const q = $('campWho').value.trim(); if (!q) return;
    const entra = !(S.info && S.info.features && S.info.features.adLookup) && onServer() && S.info.features.publishTarget === 'sharepoint';
    try {
      const r = entra ? await api('entra/lookup', { query: q, fields: S.config.fields.map((f) => ({ key: f.key, graph: f.graph })) })
        : await api('ad/lookup', { query: q, fields: S.config.fields.map((f) => ({ key: f.key, ad: f.ad })) });
      campPerson = { name: r.displayName, values: r.values || {}, groups: r.groups || [] };
      $('campWhoOff').hidden = false; campDayTable();
    } catch (e) { toast(t(e.message), true); }
  }
  suggest($('campWho'), {
    fetch: async (q) => ((S.info && S.info.features && S.info.features.adLookup) ? (await api('ad/users', { query: q })).users : onServer() && S.info.features.publishTarget === 'sharepoint' ? (await api('entra/users', { query: q })).users : []).map(person),
    pick: () => campWhoLookup()
  });
  $('campWhoGo').addEventListener('click', campWhoLookup);
  $('campWho').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); campWhoLookup(); } });
  $('campWhoOff').addEventListener('click', () => { campPerson = null; $('campWho').value = ''; $('campWhoOff').hidden = true; campDayTable(); });
  function campPersonRows(now) {
    const v = campPerson.values, from = v.mail || '', rows = [el('li', { class: 'who' }, el('b', null, campPerson.name || t('(no name)')))];
    for (const kind of ['new', 'reply']) {
      const s = core.pick(S.config, v, from, kind, null, campPerson.groups, now);
      const c = s ? core.campaignFor(S.config, s, { now, kind, values: v, from, groups: campPerson.groups }) : null;
      const kb = s ? classicKb(s) : 0;
      rows.push(el('li', { class: s ? '' : 'off' }, el('span', { class: 'k' }, kind === 'new' ? t('New mail') : t('Replies')),
        el('span', null, !s ? t('no signature') : s.name + (c ? ' · ' + t('banner «{0}»', c.name) : ' · ' + t('no banner')) + (kb > CLASSIC_KB ? ' · ' + t('too large for classic Outlook (~{0} KB)', kb) : ''))));
    }
    return rows;
  }

  // What each signature gets on a day: its banner, or none.
  function campDayTable() {
    const box = $('campDayList'); if (!box) return;
    const day = $('campDay').value || core.today(), now = new Date(day + 'T12:00:00');
    if (campPerson) { box.replaceChildren(...campPersonRows(now)); return; }
    const rows = S.config.signatures.map((s) => {
      const on = core.inPeriod(s, now), c = on ? core.campaignFor(S.config, s, { now, anyone: true, kind: 'new' }) : null;
      return el('li', { class: on ? '' : 'off' }, el('b', null, s.name || t('(no name)')), ' ',
        el('span', null, !on ? t('not used that day (its own days)') : c ? t('banner «{0}»', c.name) + (c.newOnly ? ' ' + t('(new mail only)') : '') : t('no banner')));
    });
    box.replaceChildren(...(rows.length ? rows : [el('li', { class: 'off' }, t('No signature yet.'))]));
  }

  // Signature campaigns (a whole signature for some days) and banner campaigns on one timeline.
  function campTimeline() {
    // The ones over only when unfolded (or selected): the timeline is about what comes.
    const camps = S.config.signatures.filter((s) => core.isCampaign(s) && (campPast || !isOver(s))).map((s) => ({ kind: 'sig', o: s }))
      .concat(campaigns().filter((c) => campPast || !isOver(c) || c.id === campSel).map((c) => ({ kind: 'banner', o: c })));
    const DAY = 86400000, day = (x) => Date.parse(x + 'T00:00:00');
    const now = new Date(), today = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
    const froms = camps.map((x) => (x.o.validFrom ? day(x.o.validFrom) : null)).filter((x) => x !== null);
    const untils = camps.map((x) => (x.o.validUntil ? day(x.o.validUntil) + DAY : null)).filter((x) => x !== null);
    const start = Math.min(today - 7 * DAY, ...froms), end = Math.max(today + 60 * DAY, ...untils);
    const span = end - start, pct = (x) => (100 * (x - start) / span).toFixed(2) + '%';
    const loc = I18N.lang === 'de' ? 'de-CH' : 'en-GB';
    const range = (s) => [s.validFrom ? day(s.validFrom) : start, s.validUntil ? day(s.validUntil) + DAY : end];
    const sameWho = (a, b) => a.kind === b.kind && (a.kind === 'sig' ? audienceText(a.o) === audienceText(b.o) && (a.o.language || '') === (b.o.language || '')
      : audienceText(a.o) === audienceText(b.o) && JSON.stringify((a.o.signatures || []).slice().sort()) === JSON.stringify((b.o.signatures || []).slice().sort()));
    const clash = (a, b) => a !== b && sameWho(a, b) && range(a.o)[0] < range(b.o)[1] && range(b.o)[0] < range(a.o)[1];
    const ticks = [];
    for (let d = new Date(start); d.getTime() <= end; d = new Date(d.getFullYear(), d.getMonth() + 1, 1)) {
      const m = new Date(d.getFullYear(), d.getMonth(), 1).getTime();
      if (m >= start) ticks.push(el('span', { class: 'camp-tick', style: 'left:' + pct(m) }, new Date(m).toLocaleDateString(loc, { month: 'short', year: '2-digit' })));
    }
    const fmt = (x) => (x ? new Date(day(x)).toLocaleDateString(loc) : '…');
    $('campTimeline').replaceChildren(
      el('div', { class: 'camp-axis' }, ...ticks, el('span', { class: 'camp-today', style: 'left:' + pct(today), title: t('Today') })),
      ...camps.map((x) => {
        const s = x.o, [a, b] = range(s), over = camps.some((o) => clash(x, o));
        const bar = el('button', { type: 'button', class: 'camp-bar' + (x.kind === 'banner' ? ' banner' : '') + (core.inPeriod(s) ? ' now' : b <= today ? ' past' : '') + (over ? ' clash' : '') + (x.kind === 'banner' && s.id === campSel ? ' sel' : ''),
          style: 'left:' + pct(a) + ';width:calc(' + pct(b) + ' - ' + pct(a) + ')',
          title: fmt(s.validFrom) + ' - ' + fmt(s.validUntil) + ' · ' + (x.kind === 'banner' ? t('banner') : t('whole signature')) + ' · ' + audienceText(s) + (over ? ' · ' + t('at the same time as another one for the same people') : ''),
          'aria-label': (s.name || t('(no name)')) + ': ' + fmt(s.validFrom) + ' - ' + fmt(s.validUntil),
          onclick: () => { if (x.kind === 'sig') openSig(s.id); else { pickCamp(s.id); campList(); campEditor(); campTimeline(); } } });
        return el('div', { class: 'camp-row' }, el('span', { class: 'camp-name', title: s.name }, (x.kind === 'banner' ? '▭ ' : '✉ ') + (s.name || t('(no name)'))), el('div', { class: 'camp-track' }, bar, el('span', { class: 'camp-today', style: 'left:' + pct(today) })));
      }),
      ...(camps.length ? [] : [el('p', { class: 'muted' }, t('No campaigns yet.'))]));
  }

  // Ending in 3 days with none after it: said in the status bar (a click opens it).
  function campEndingNote() {
    const pill = $('stCampaign'); if (!pill || !S.config) return;
    const e = core.campaignsEnding(S.config, 3)[0];
    pill.hidden = !e;
    if (e) { pill.textContent = e.left ? t('Campaign «{0}» ends in {1} day(s), none follows', e.c.name, e.left) : t('Campaign «{0}» ends today, none follows', e.c.name); pill.dataset.id = e.c.id; }
  }
  $('stCampaign').addEventListener('click', () => { campSel = $('stCampaign').dataset.id || campSel; flushSource(); goView('campaigns'); });

  $('btnAddCampaign').addEventListener('click', () => {
    const c = newCampaign(campaigns().length ? t('New campaign') : t('Holidays {0}', new Date().getFullYear()));
    campaigns().push(c); pickCamp(c.id); campBuild(c); changed({ list: false }); renderCampaignsView();
    const n = $('campEdit').querySelector('input'); if (n) { n.focus(); n.select(); }
  });
  // A new one from an earlier one: its banner, signatures and rules, its days from today on (as long).
  $('campFrom').addEventListener('focus', () => {
    const sel = $('campFrom');
    sel.replaceChildren(el('option', { value: '' }, t('From an earlier one…')), ...campaigns().slice().reverse().map((c) => el('option', { value: c.id }, c.name || t('(no name)'))));
  });
  $('campFrom').addEventListener('change', () => {
    const o = campById($('campFrom').value); $('campFrom').value = ''; if (!o) return;
    const n = JSON.parse(JSON.stringify(o)), len = o.validFrom && o.validUntil ? Math.round((Date.parse(o.validUntil) - Date.parse(o.validFrom)) / 86400000) : 13;
    const d = new Date(); n.validFrom = core.today(d); d.setDate(d.getDate() + len); n.validUntil = core.today(d);
    n.name = /\b(19|20)\d\d\b/.test(o.name || '') ? o.name.replace(/\b(19|20)\d\d\b/, String(new Date().getFullYear())) : (o.name || '') + ' ' + t('(copy)');
    if (n.name === o.name) n.name += ' ' + t('(copy)');
    n.id = newId('c-' + n.name);
    campaigns().push(n); pickCamp(n.id); campBuild(n); changed({ list: false }); renderCampaignsView();
    toast(t('Made from «{0}»: {1}.', o.name, campaignText(n)));
  });
  $('campDay').addEventListener('input', campDayTable);
