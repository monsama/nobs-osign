/*
 * Whether a newer version of NOBS OSign is out: twice a day the server asks GitHub for the latest
 * release (one request out, nothing comes in; no data of yours goes with it). The editor then
 * says so in its status bar. Off with OSIGN_UPDATE_CHECK=0.
 */
'use strict';

const LATEST = 'https://api.github.com/repos/monsama/nobs-osign/releases/latest';
const EVERY_MS = 12 * 60 * 60 * 1000;

// 0.10.1 > 0.9.3: by number, part by part; a pre-release (0.9.0-rc.1) is never "newer".
function newer(a, b) {
  const p = (v) => (/^v?(\d+)\.(\d+)\.(\d+)$/.exec(String(v || '').trim()) || []).slice(1).map(Number);
  const x = p(a), y = p(b);
  if (x.length !== 3 || y.length !== 3) return false;
  for (let i = 0; i < 3; i++) if (x[i] !== y[i]) return x[i] > y[i];
  return false;
}

function createUpdates({ env, version, fetchFn, log, now = Date.now }) {
  const on = String(env.OSIGN_UPDATE_CHECK || '0') !== '0';
  const url = env.OSIGN_UPDATE_URL || LATEST;
  let found = null, checkedAt = 0, running = null;

  async function check() {
    checkedAt = now();
    try {
      const r = await fetchFn(url, { headers: { Accept: 'application/vnd.github+json', 'User-Agent': 'nobs-osign/' + version }, signal: AbortSignal.timeout(10000) });
      if (!r.ok) return;
      const j = await r.json();
      const latest = String(j.tag_name || '').replace(/^v/, '');
      found = newer(latest, version) ? { version: latest, url: /^https:\/\/github\.com\//.test(j.html_url || '') ? j.html_url : '', at: String(j.published_at || '') } : null;
      if (found) log.info('A newer version of NOBS OSign is out: ' + latest + ' (this is ' + version + ').');
    } catch (e) { /* asked again later */ }
  }
  // What is known now; a check starts when the last one is older than EVERY_MS (never waited for).
  function status() {
    if (!on) return null;
    if (!running && now() - checkedAt > EVERY_MS) running = check().finally(() => { running = null; });
    return found;
  }
  return { status, check: () => (on ? check() : Promise.resolve()), on };
}

module.exports = { createUpdates, newer };
