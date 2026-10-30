/*
 * The event handler: Outlook runs it when a new message, reply, forward or meeting is started, and
 * when the From address changes. It puts the right signature in, or none. Only when it could not
 * (the signatures or the user's details could not be read) does it show a short notice on the message.
 */
import { core, localize, loadConfig, loadUserValues, getPrefs, composeInfo, fromAddress, insertSignature, isMessage } from './common.js';

/* global Office */

// The notice above the message, under a key of its own: a later attempt that works takes it away.
const NOTICE = 'osign';
function notices(item) { return item && item.notificationMessages; }
function showNotice(item, reason) {
  const n = notices(item); if (!n || !n.replaceAsync) return Promise.resolve();
  const german = /^de/i.test(Office.context.displayLanguage || '');
  // What to do first (the pane shows the whole story, and signs in where that is needed), then why.
  let text = (german ? 'NOBS OSign hat keine Signatur eingefügt; mehr unter Signaturen. ' : 'NOBS OSign inserted no signature; see Signatures. ') + localize(reason, german);
  // Outlook takes at most 150 characters.
  if (text.length > 150) text = text.slice(0, 149) + '\u2026';
  return new Promise((resolve) => {
    try {
      n.replaceAsync(NOTICE, { type: (Office.MailboxEnums.ItemNotificationMessageType || {}).InformationalMessage || 'informationalMessage', message: text, icon: 'Icon.16', persistent: false }, () => resolve());
    } catch (e) { resolve(); }
  });
}
function clearNotice(item) {
  const n = notices(item); if (!n || !n.removeAsync) return Promise.resolve();
  return new Promise((resolve) => { try { n.removeAsync(NOTICE, () => resolve()); } catch (e) { resolve(); } });
}

async function sign(event, what) {
  const item = Office.context.mailbox.item;
  try {
    const cfg = await loadConfig({ interactive: false });
    if (!isMessage(item) && !(cfg.options && cfg.options.meetings)) return;
    const info = await composeInfo(item);
    const user = await loadUserValues(cfg, { interactive: false });
    if (!user.values) {
      const why = user.error ? user.error.message : 'your details could not be read.';
      console.log('OSign: no user details, no signature. ' + why);
      await showNotice(item, why);
      return;
    }
    const from = await fromAddress(item);
    const sig = core.pick(cfg, user.values, from, info.kind, getPrefs(), user.groups);
    // A changed sender may have no signature where the first one had: then the old one goes.
    if (!sig && what !== 'fromChanged') return;
    await insertSignature(item, cfg, sig, user.values, info.html);
    await clearNotice(item);
  } catch (e) {
    const why = e && e.message ? e.message : String(e);
    console.log('OSign: ' + why);
    await showNotice(item, why);
  } finally {
    event.completed();
  }
}

function onNewMessageCompose(event) { sign(event, 'new'); }
function onNewAppointmentOrganizer(event) { sign(event, 'new'); }
function onMessageFromChanged(event) { sign(event, 'fromChanged'); }

// The names are the ones in the manifest's LaunchEvent entries.
Office.actions.associate('onNewMessageComposeHandler', onNewMessageCompose);
Office.actions.associate('onNewAppointmentOrganizerHandler', onNewAppointmentOrganizer);
Office.actions.associate('onMessageFromChangedHandler', onMessageFromChanged);
