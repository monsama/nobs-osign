  // ---------------------------------------------------------------------------------------------
  // Undo for the whole draft (Ctrl+Z / Ctrl+Y outside text fields, ↶ ↷ in the top bar): every
  // change to signatures, fields, building blocks and options, typing taken together, up to 100
  // steps in this browser tab. Undone, the page goes where the change was, and says what it was.
  // Publishing, photos, images and the server's settings are not part of the draft: not in here.
  // ---------------------------------------------------------------------------------------------
  const gh = { stack: [], pos: -1, timer: null, applying: false };
  const GH_MAX = 100;
  const ghWhere = () => ({ view: S.view, sigId: S.sigId, block: S.blockEdit });
  function ghReset() {
    clearTimeout(gh.timer); gh.timer = null;
    gh.stack = S.config ? [Object.assign({ json: JSON.stringify(S.config), what: '' }, ghWhere())] : []; gh.pos = gh.stack.length - 1;
    ghButtons();
  }
  // A change noted: kept as a step half a second after the last one (typing taken together).
  function ghNote() { if (gh.applying || !gh.stack.length) return; clearTimeout(gh.timer); gh.timer = setTimeout(ghRecord, 500); }
  // A step kept when the draft is not as at the last step; the oldest go past 100.
  function ghRecord() {
    clearTimeout(gh.timer); gh.timer = null;
    if (!gh.stack.length) return;
    const json = JSON.stringify(S.config), prev = gh.stack[gh.pos];
    if (json === prev.json) return;
    gh.stack = gh.stack.slice(0, gh.pos + 1);
    const where = ghWhere(), sigOf = ghSigOf(prev.json, json);
    // The signature the change was in (also one added or removed), so undo opens that one.
    if (sigOf && !where.block) { where.view = 'sig'; where.sigId = sigOf; }
    gh.stack.push(Object.assign({ json, what: ghDescribe(prev.json, json) }, where));
    if (gh.stack.length > GH_MAX) gh.stack.shift();
    gh.pos = gh.stack.length - 1;
    ghButtons();
  }
  // The one signature added or removed between two states, if it is one.
  function ghSigOf(a, b) {
    let x, y; try { x = JSON.parse(a).signatures || []; y = JSON.parse(b).signatures || []; } catch (e) { return ''; }
    const xi = new Map(x.map((s) => [s.id, JSON.stringify(s)])), yi = new Map(y.map((s) => [s.id, JSON.stringify(s)]));
    // Only one added or removed: a change in a page (Photos, Fields) stays on that page.
    const ids = [...new Set([...xi.keys(), ...yi.keys()])].filter((id) => xi.has(id) !== yi.has(id));
    return ids.length === 1 ? ids[0] : '';
  }
  // What changed between two states, in a few words.
  function ghDescribe(a, b) {
    let x, y; try { x = JSON.parse(a); y = JSON.parse(b); } catch (e) { return t('a change'); }
    const nm = (s) => s.name || t('(no name)');
    const xs = x.signatures || [], ys = y.signatures || [];
    const xi = new Map(xs.map((s) => [s.id, s])), yi = new Map(ys.map((s) => [s.id, s]));
    const added = ys.filter((s) => !xi.has(s.id)), gone = xs.filter((s) => !yi.has(s.id));
    if (added.length) return added.length === 1 ? t('signature «{0}» added', nm(added[0])) : t('{0} signatures added', added.length);
    if (gone.length) return gone.length === 1 ? t('signature «{0}» removed', nm(gone[0])) : t('{0} signatures removed', gone.length);
    if (xs.map((s) => s.id).join() !== ys.map((s) => s.id).join()) return t('order of the signatures');
    const diff = ys.filter((s) => JSON.stringify(s) !== JSON.stringify(xi.get(s.id)));
    if (diff.length === 1) {
      const s = diff[0], o = xi.get(s.id);
      if (s.name !== o.name) return t('signature renamed to «{0}»', nm(s));
      if (s.html !== o.html || s.source !== o.source) return t('text of «{0}»', nm(s));
      return t('settings of «{0}»', nm(s));
    }
    if (diff.length > 1) return t('{0} signatures', diff.length);
    const xf = x.fields || [], yf = y.fields || [];
    if (yf.length > xf.length) return t('field «{0}» added', (yf.find((f) => !xf.some((g) => g.key === f.key)) || {}).label || '');
    if (yf.length < xf.length) return t('field «{0}» removed', (xf.find((f) => !yf.some((g) => g.key === f.key)) || {}).label || '');
    if (JSON.stringify(xf) !== JSON.stringify(yf)) return t('fields');
    if (JSON.stringify(x.blocks || []) !== JSON.stringify(y.blocks || [])) return t('building blocks');
    if (JSON.stringify(x.campaigns || []) !== JSON.stringify(y.campaigns || [])) return t('campaigns');
    if (JSON.stringify(x.colors || []) !== JSON.stringify(y.colors || [])) return t('house colours');
    return t('options of the add-in');
  }
  // ↶ and ↷: on or off, their tooltip naming the step.
  function ghButtons() {
    const u = $('ghUndo'), r = $('ghRedo'); if (!u) return;
    // While the images are worked on, no step: one from before would bring back files now gone.
    if (preparing) { offFor(u, t('Still working on the images…')); offFor(r, t('Still working on the images…')); return; }
    const back = gh.pos > 0 ? gh.stack[gh.pos] : null, fwd = gh.pos < gh.stack.length - 1 ? gh.stack[gh.pos + 1] : null;
    offFor(u, back ? '' : t('Nothing to undo yet.'));
    offFor(r, fwd ? '' : t('Nothing to redo.'));
    if (back) u.setAttribute('title', t('Undo: {0} (Ctrl+Z)', back.what));
    if (fwd) r.setAttribute('title', t('Redo: {0} (Ctrl+Y)', fwd.what));
  }
  // One step back or forward: the draft as it was then, the page where the change was.
  function ghStep(dir) {
    if (imagesBusy()) return false;
    flushSource();
    if (gh.timer) ghRecord();
    const to = gh.pos + dir;
    if (to < 0 || to >= gh.stack.length) return false;
    const entry = dir < 0 ? gh.stack[gh.pos] : gh.stack[to];
    gh.applying = true;
    try {
      S.config = JSON.parse(gh.stack[to].json);
      gh.pos = to;
      // Where the change was: that signature, building block or page.
      if (entry.view === 'sig' && entry.block && S.config.blocks.some((b) => b.key === entry.block)) openBlock(entry.block);
      else if (entry.view === 'sig' && entry.sigId && S.config.signatures.some((s) => s.id === entry.sigId)) openSig(entry.sigId);
      else if (entry.view === 'sig') { if (S.config.signatures.length) openSig(S.config.signatures[0].id); else show('empty'); }
      else goView(entry.view);
      changed();
    } finally { gh.applying = false; }
    clearTimeout(gh.timer); gh.timer = null;
    ghButtons();
    toast((dir < 0 ? t('Undone: {0}', entry.what) : t('Redone: {0}', entry.what)), false, { label: dir < 0 ? t('Redo') : t('Undo'), run: () => ghStep(-dir) });
    return true;
  }
  $('ghUndo').addEventListener('click', () => ghStep(-1));
  $('ghRedo').addEventListener('click', () => ghStep(1));

  // Load again (⟳ in the top bar): what was typed saved first, then the draft, what is published,
  // the state and the server's info read anew, without reloading the page. The page stays where it
  // is; the undo steps too, as long as the draft on the server is the one here.
  async function refreshAll() {
    const b = $('btnRefresh');
    if (b.classList.contains('busy')) return;
    b.classList.add('busy'); b.disabled = true;
    try {
      flushSource();
      if (saveTimer) await save();
      await saving;
      const where = ghWhere(), before = stableJson(prepared()), keep = { stack: gh.stack, pos: gh.pos };
      const [info, st, draft] = await Promise.all([api('info'), api('settings/get'), api('draft/get')]);
      S.info = info; S.settings = st.settings; lastSettings = st;
      applyDraft(draft);
      // The users as last read (and whether they were written since), where values come from AD.
      if (core.usesAdExport(S.config)) await loadSnapshot();
      const same = stableJson(prepared()) === before;
      if (same) { gh.stack = keep.stack; gh.pos = keep.pos; gh.stack[gh.pos].json = JSON.stringify(S.config); ghButtons(); }
      if (where.view === 'sig' && where.block && S.config.blocks.some((x) => x.key === where.block)) openBlock(where.block);
      else if (where.view === 'sig' && where.sigId && S.config.signatures.some((s) => s.id === where.sigId)) openSig(where.sigId);
      else if (where.view !== 'sig' && where.view !== 'empty') goView(where.view);
      renderStatusBar();
      loadApproval().catch(() => {});
      toast(same ? t('Loaded again: as it was.') : t('Loaded again: the draft on the server was changed (by someone else, or in another window).'));
    } catch (e) { toast(t(e.message), true); }
    finally { b.classList.remove('busy'); b.disabled = false; }
  }
  $('btnRefresh').addEventListener('click', refreshAll);
