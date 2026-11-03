  // ---------------------------------------------------------------------------------------------
  // Replace in all signatures (Ctrl+H): a text, in what people read (text, link targets, alt and
  // title), never in the HTML's tags. Shown first per signature and building block; undone in one go.
  // ---------------------------------------------------------------------------------------------
  function replaceIn(html, find, by, caseToo) {
    const tpl = document.createElement('template'); tpl.innerHTML = html;
    const re = new RegExp(find.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), caseToo ? 'g' : 'gi');
    let n = 0;
    // Never inside a field (@@mail) or a building block ({{name}}): those stay as they are.
    const sub = (s) => s.split(/(@@[A-Za-z][A-Za-z0-9_]*|\{\{[A-Za-z][A-Za-z0-9_-]*\}\})/).map((part, i) => (i % 2 ? part : part.replace(re, () => { n++; return by; }))).join('');
    const walk = (node) => {
      for (const c of node.childNodes) {
        if (c.nodeType === 3) { const v = sub(c.nodeValue); if (v !== c.nodeValue) c.nodeValue = v; }
        else if (c.nodeType === 1) {
          for (const a of ['href', 'alt', 'title']) if (c.hasAttribute(a)) { const v = sub(c.getAttribute(a)); if (v !== c.getAttribute(a)) c.setAttribute(a, v); }
          walk(c);
        }
      }
    };
    walk(tpl.content);
    // Made safe as anything typed is (a link target replaced with javascript: does not stay).
    return { n, html: n ? sanitize(tpl.innerHTML) : html };
  }
  const repTargets = () => S.config.signatures.map((s) => ({ kind: 'sig', obj: s, name: s.name || t('(no name)') }))
    .concat((S.config.blocks || []).map((b) => ({ kind: 'block', obj: b, name: '{{' + b.key + '}} ' + (b.name || '') })));
  function repPreview() {
    const find = $('repFind').value, by = $('repBy').value, cs = $('repCase').checked;
    $('repErr').textContent = '';
    if (!find) { $('repList').replaceChildren(); $('repGo').disabled = true; $('repSum').textContent = ''; return; }
    const hits = repTargets().map((x) => Object.assign(x, replaceIn(x.obj.html || '', find, by, cs))).filter((x) => x.n);
    $('repList').replaceChildren(...hits.map((x, i) => el('li', null, el('label', { class: 'check' },
      el('input', { type: 'checkbox', checked: '', 'data-i': String(i) }), el('span', null, (x.kind === 'block' ? t('Building block') + ' ' : '') + x.name),
      el('small', null, ' ' + (x.n === 1 ? t('1 place') : t('{0} places', x.n)))))));
    $('repList')._hits = hits;
    $('repSum').textContent = hits.length ? t('Found in {0} of {1}.', hits.length, repTargets().length) : t('Found nowhere.');
    $('repGo').disabled = !hits.length;
  }
  function openReplace() {
    flushSource();
    $('repFind').value = ''; $('repBy').value = ''; repPreview();
    $('dlgReplace').showModal(); $('repFind').focus();
  }
  ['repFind', 'repBy'].forEach((id) => $(id).addEventListener('input', repPreview));
  $('repCase').addEventListener('change', repPreview);
  $('btnReplace').addEventListener('click', openReplace);
  $('repGo').addEventListener('click', (e) => {
    e.preventDefault();
    const hits = $('repList')._hits || [], picked = [...$('repList').querySelectorAll('input:checked')].map((c) => hits[Number(c.dataset.i)]);
    if (!picked.length) { $('repErr').textContent = t('Tick at least one.'); return; }
    const before = picked.map((x) => [x.obj, x.obj.html]);
    picked.forEach((x) => { x.obj.html = x.html; });
    const n = picked.reduce((a, x) => a + x.n, 0);
    const refresh = () => { changed(); if (S.view === 'sig') { if (S.blockEdit) openBlock(S.blockEdit); else if (S.sigId) openSig(S.sigId); } if (S.view === 'blocks') renderBlocks(); };
    refresh();
    $('dlgReplace').close();
    toast(t('Replaced {0} times in {1}.', n, picked.length), false, { label: t('Undo'), run: () => { before.forEach(([o, h]) => { o.html = h; }); refresh(); } });
  });
