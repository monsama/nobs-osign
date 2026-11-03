/*
 * Publishing, and four eyes before it (settings: requireApproval): one person asks, another one
 * looks at what changed and publishes it. Routes for server/app.js; live() gives what changes with
 * the connections (adOn).
 */
'use strict';
const fs = require('node:fs');
const { writeAtomic } = require('./store');
const { diffConfigs } = require('./diff');

function createPublishRoutes({ store, history, settings, users, approvalFile, approval, needsApproval, needsApprovalFor, publisher, publishNow, compiledProblem, live }) {
  const same = (a, b) => String(a || '').toLowerCase() === String(b || '').toLowerCase();
  // What waits for approval, for the editor: who asked, when, and whether it is the one asking now.
  function approvalAnswer(s) {
    const a = approval();
    if (!a) return { pending: null, required: needsApprovalFor(s), publisher: publisher(s) };
    const current = store.draftHash() === a.hash;
    return {
      required: needsApprovalFor(s), publisher: publisher(s),
      pending: { by: a.by, at: a.at, message: a.message, current, mine: same(a.login, s.user.login), mayApprove: !same(a.login, s.user.login) && current && publisher(s),
        hash: a.hash, changes: diffConfigs(history.latest(), a.configJson).concat(current ? imageChanges() : []) }
    };
  }
  // The draft's images against the last published ones: added, removed, replaced (same name).
  function imageChanges() {
    const was = history.latestImages(), now = store.assetHashes(), out = [];
    for (const n of Object.keys(now).sort()) if (!(n in was)) out.push({ kind: 'imgadd', name: n }); else if (was[n] !== now[n]) out.push({ kind: 'imgchg', name: n });
    for (const n of Object.keys(was).sort()) if (!(n in now)) out.push({ kind: 'imgdel', name: n });
    return out;
  }

  const routes = {
    publish: async (d, s) => {
      if (needsApproval()) return { error: 'Publishing needs a second person\'s approval here: ask for it (Publish… → Ask for approval).' };
      if (!publisher(s)) return { error: 'You edit the draft; publishing is for the publishing group: ask for approval (Publish… → Ask for approval).' };
      return publishNow(d.configJson, s.user.name || s.user.login, d.message);
    },
    // Four eyes (settings: requireApproval).
    'approval/get': (d, s) => approvalAnswer(s),
    'approval/request': (d, s) => {
      if (!needsApprovalFor(s)) return { error: 'Approval before publishing is not on.' };
      const saved = store.saveDraft(typeof d.configJson === 'string' ? d.configJson : '');
      if (saved.error) return saved;
      const bad = compiledProblem(typeof d.configJson === 'string' ? d.configJson : '');
      if (bad) return { error: bad };
      writeAtomic(approvalFile, JSON.stringify({ by: s.user.name || s.user.login, login: s.user.login, at: Date.now(), message: String(d.message || '').slice(0, 200), configJson: d.configJson, hash: store.draftHash() }));
      return approvalAnswer(s);
    },
    'approval/approve': async (d, s) => {
      const a = approval();
      if (!a) return { error: 'Nothing is waiting for approval.' };
      if (same(a.login, s.user.login)) return { error: 'Someone else has to approve what you asked for.' };
      if (!publisher(s)) return { error: 'Only the publishing group approves.' };
      if (store.draftHash() !== a.hash) return { error: 'The draft changed after approval was asked for: it has to be asked for again.' };
      // Only what the approver was shown: asked for again since, it is to be looked at again.
      if (String(d.hash || '') !== a.hash) return { error: 'The request changed while you looked at it: look at it again, then approve.' };
      const r = await publishNow(a.configJson, a.by + ' (' + (s.user.name || s.user.login) + ' approved)', a.message);
      if (r.error) return r;
      fs.rmSync(approvalFile, { force: true });
      // The users' values from AD, as the approved signatures ask for them (not as anyone asked
      // before approval): read and written now.
      if (live().adOn) {
        try { const u = await users.exportPublished(settings.target()); if (!u.skipped) r.users = u; } catch (e) { r.usersError = 'Writing the users: ' + e.message; }
      }
      return r;
    },
    'approval/cancel': (d, s) => {
      const a = approval();
      if (a && !same(a.login, s.user.login) && !s.user.local) return { error: 'Only who asked for it, or the local account admin, can withdraw it.' };
      fs.rmSync(approvalFile, { force: true });
      return {};
    }
  };
  return { routes };
}

module.exports = { createPublishRoutes };
