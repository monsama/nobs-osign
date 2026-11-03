// node --test tests/i18n.test.js
//
// The editor's German: every text the code shows with t('…') and every text of the page has its
// German (editor/ui/i18n.js), and no text is there twice (the second would win, unseen).
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { JSDOM } = require('jsdom');

const ROOT = path.join(__dirname, '..', 'editor', 'ui');
const src = fs.readFileSync(path.join(ROOT, 'i18n.js'), 'utf8');
// The dictionary as the page has it: i18n.js run with German chosen, DE taken from its scope.
function dictionary() {
  const body = src.replace(/document\.documentElement\.lang = lang;[\s\S]*$/, 'globalThis.__DE = DE; })();');
  const ctx = { navigator: { language: 'de-CH' }, localStorage: { getItem: () => 'de' }, document: {}, globalThis: null };
  ctx.globalThis = ctx;
  vm.runInNewContext(body, ctx);
  return ctx.__DE;
}
const norm = (s) => String(s).replace(/\s+/g, ' ').trim();
// What needs no German: the same in both, names, codes, symbols.
const SAME = (s) => !/[a-z]{3}/i.test(s) || /^[\w.+-]+@[\w.-]+$|^[\w.-]+\.(xml|json|png|jpg|htm)$|^· /.test(s) || /^(NOBS OSign|OSign|Outlook|HTML|OK|Info|Server|Name|Status|Graph|Entra ID|Active Directory|SharePoint|Cloudflare Pages|Azure Static Web Apps|IIS|StartTLS|Microsoft Graph|Add-in|admin|Deutsch|English|Français|Italiano|Disclaimer|Mail|Logo|Backup|Audit-Log|Tenant|Buy me a coffee|Manifest:|root)$/i.test(s);

test('German: no text twice in the dictionary', () => {
  const keys = [...src.slice(0, src.indexOf('  DE.html = {')).matchAll(/^ {4}'((?:[^'\\]|\\.)*)':/gm)].map((m) => m[1]);
  const seen = new Set(), twice = [];
  keys.forEach((k) => { if (seen.has(k)) twice.push(k); seen.add(k); });
  assert.deepEqual(twice, []);
});

test('German: every t(\'…\') of the editor\'s code has its German', () => {
  const DE = dictionary();
  const has = (s) => Object.prototype.hasOwnProperty.call(DE.text, s) || (s.indexOf(': ') > 0 && Object.prototype.hasOwnProperty.call(DE.text, s.slice(0, s.indexOf(': '))));
  const dir = path.join(ROOT, 'app');
  const missing = new Set();
  for (const f of fs.readdirSync(dir).filter((x) => x.endsWith('.js'))) {
    const code = fs.readFileSync(path.join(dir, f), 'utf8');
    for (const m of code.matchAll(/\bt\('((?:[^'\\]|\\.)*)'/g)) {
      const s = m[1].replace(/\\'/g, "'").replace(/\\\\/g, '\\');
      if (!SAME(s) && !has(s)) missing.add(f + ': ' + s);
    }
  }
  assert.deepEqual([...missing], []);
});

test('German: every text of the page has its German', () => {
  const DE = dictionary();
  const dom = new JSDOM(fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8'));
  const doc = dom.window.document, missing = new Set();
  const walker = doc.createTreeWalker(doc.body, 4);
  while (walker.nextNode()) {
    const n = walker.currentNode, key = norm(n.nodeValue);
    if (!key || SAME(key) || n.parentNode.closest('[data-th], code, script, style, textarea, kbd, option[value=""] ~ option[lang]')) continue;
    if (!Object.prototype.hasOwnProperty.call(DE.text, key)) missing.add(key);
  }
  doc.querySelectorAll('[placeholder], [title], [aria-label]').forEach((el) => ['placeholder', 'title', 'aria-label'].forEach((a) => {
    const v = el.getAttribute(a);
    if (v && !SAME(v) && !/^(https?:|[A-Z]:\\|\\\\|ldaps?:|00000000|CN=|OU=|DC=|\+41|e\.g\.)/.test(v) && !Object.prototype.hasOwnProperty.call(DE.text, v)) missing.add('[' + a + '] ' + v);
  }));
  doc.querySelectorAll('[data-th]').forEach((el) => { if (!DE.html[el.getAttribute('data-th')]) missing.add('data-th=' + el.getAttribute('data-th')); });
  assert.deepEqual([...missing], []);
});

test('German in the add-in: every t(\'…\') of the pane has its German, none twice', () => {
  const pane = fs.readFileSync(path.join(__dirname, '..', 'addin', 'src', 'taskpane.js'), 'utf8');
  const block = pane.slice(pane.indexOf('const DE = {'), pane.indexOf('};', pane.indexOf('const DE = {')));
  const keys = [...block.matchAll(/^ {2}(['"])((?:(?!\1)[^\\]|\\.)*)\1:/gm)].map((m) => m[2].replace(/\\(['"])/g, '$1'));
  const seen = new Set(), twice = [];
  keys.forEach((k) => { if (seen.has(k)) twice.push(k); seen.add(k); });
  assert.deepEqual(twice, []);
  const missing = new Set();
  for (const m of pane.matchAll(/\bt\((['"])((?:(?!\1)[^\\]|\\.)*)\1/g)) {
    const s = m[2].replace(/\\(['"])/g, '$1');
    if (!SAME(s) && !seen.has(s)) missing.add(s);
  }
  const html = fs.readFileSync(path.join(__dirname, '..', 'addin', 'taskpane.html'), 'utf8');
  const dom = new JSDOM(html).window.document;
  dom.querySelectorAll('h1 .sub, label, #refresh, #help summary').forEach((el) => el.childNodes.forEach((n) => {
    const v = n.nodeType === 3 ? n.nodeValue.trim() : '';
    if (v && !SAME(v) && !seen.has(v)) missing.add(v);
  }));
  assert.deepEqual([...missing], []);
});

test('French and Italian in the add-in: every text the German has, and every message', async () => {
  const pane = fs.readFileSync(path.join(__dirname, '..', 'addin', 'src', 'taskpane.js'), 'utf8');
  const block = pane.slice(pane.indexOf('const DE = {'), pane.indexOf('};', pane.indexOf('const DE = {')));
  const keys = [...block.matchAll(/^ {2}(['"])((?:(?!\1)[^\\]|\\.)*)\1:/gm)].map((m) => m[2].replace(/\\(['"])/g, '$1'));
  const { LANG } = await import(path.join(__dirname, '..', 'addin', 'src', 'lang.js'));
  const common = fs.readFileSync(path.join(__dirname, '..', 'addin', 'src', 'common.js'), 'utf8');
  const msgs = (common.slice(common.indexOf('const MSG_DE = ['), common.indexOf('];', common.indexOf('const MSG_DE = ['))).match(/^\s+\[\//gm) || []).length;
  for (const l of ['fr', 'it']) {
    assert.deepEqual(keys.filter((k) => !LANG[l].ui[k]), [], l + ': missing');
    assert.deepEqual(Object.keys(LANG[l].ui).filter((k) => !keys.includes(k)), [], l + ': not in the German');
    assert.equal(LANG[l].msg.length, msgs, l + ': one per message');
  }
});
