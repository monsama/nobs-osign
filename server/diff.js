/*
 * What changed between two configurations, in words: signatures added, removed, renamed, their
 * HTML (the lines that differ), who gets them, and the fields and options. For the history
 * ("Compare") and for an approval (what would go out).
 */
'use strict';

const MAX_LINES = 600;
const parse = (t) => { try { return JSON.parse(t || '{}') || {}; } catch (e) { return {}; } };
const same = (a, b) => JSON.stringify(a === undefined ? null : a) === JSON.stringify(b === undefined ? null : b);
// HTML as lines: one tag or text run each, so a changed word shows as one line.
const lines = (html) => String(html || '').replace(/>\s*</g, '>\n<').split('\n').map((l) => l.trim()).filter(Boolean).slice(0, MAX_LINES);

// The lines removed (-) and added (+), by the longest common subsequence.
function lineDiff(a, b) {
  const n = a.length, m = b.length;
  const L = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
  for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) L[i][j] = a[i] === b[j] ? L[i + 1][j + 1] + 1 : Math.max(L[i + 1][j], L[i][j + 1]);
  const out = [];
  let i = 0, j = 0;
  while (i < n && j < m) {
    if (a[i] === b[j]) { i++; j++; } else if (L[i + 1][j] >= L[i][j + 1]) out.push({ op: '-', text: a[i++] }); else out.push({ op: '+', text: b[j++] });
  }
  while (i < n) out.push({ op: '-', text: a[i++] });
  while (j < m) out.push({ op: '+', text: b[j++] });
  return out;
}

// What a signature's "who gets it" is made of: everything but its name and HTML.
const rules = (s) => { const r = Object.assign({}, s); delete r.id; delete r.name; delete r.html; return r; };

// → [{ kind: 'added'|'removed'|'renamed'|'html'|'rules'|'fields'|'options', name, detail?, lines? }]
function diffConfigs(beforeJson, afterJson) {
  const a = parse(beforeJson), b = parse(afterJson);
  const out = [];
  const before = new Map((a.signatures || []).map((s) => [s.id, s]));
  const after = new Map((b.signatures || []).map((s) => [s.id, s]));
  for (const [id, s] of after) {
    const o = before.get(id);
    if (!o) { out.push({ kind: 'added', name: s.name || id }); continue; }
    if ((o.name || '') !== (s.name || '')) out.push({ kind: 'renamed', name: s.name || id, detail: o.name || id });
    if ((o.html || '') !== (s.html || '')) out.push({ kind: 'html', name: s.name || id, lines: lineDiff(lines(o.html), lines(s.html)) });
    const ro = rules(o), rs = rules(s);
    const keys = [...new Set([...Object.keys(ro), ...Object.keys(rs)])].filter((k) => !same(ro[k], rs[k])).sort();
    if (keys.length) out.push({ kind: 'rules', name: s.name || id, detail: keys.join(', ') });
  }
  for (const [id, s] of before) if (!after.has(id)) out.push({ kind: 'removed', name: s.name || id });
  if (!same(a.fields || [], b.fields || [])) {
    const fa = new Map((a.fields || []).map((f) => [f.key, f])), fb = new Map((b.fields || []).map((f) => [f.key, f]));
    const changed = [...new Set([...fa.keys(), ...fb.keys()])].filter((k) => !same(fa.get(k), fb.get(k))).sort();
    out.push({ kind: 'fields', name: '', detail: changed.join(', ') });
  }
  const oa = a.options || {}, ob = b.options || {};
  const opts = [...new Set([...Object.keys(oa), ...Object.keys(ob)])].filter((k) => !same(oa[k], ob[k])).sort();
  if (opts.length) out.push({ kind: 'options', name: '', detail: opts.join(', ') });
  return out;
}

module.exports = { diffConfigs, lineDiff };
