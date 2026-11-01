/*
 * The event handler: Outlook runs it when a new message, reply, forward or meeting is started, and
 * when the From address changes. It puts the right signature in, or none. Only when it could not
 * (the signatures or the user's details could not be read) does it show a short notice on the message.
 */
import { core, localize, loadConfig, loadUserValues, getPrefs, getVariants, isClassicOutlook, composeInfo, fromAddress, insertSignature, isMessage } from './common.js';

/* global Office */

// The notice above the message, under a key of its own: a later attempt that works takes it away.
const NOTICE = 'osign';
function notices(item) { return item && item.notificationMessages; }
function showNotice(item, reason) {
  const n = notices(item); if (!n || !n.replaceAsync) return Promise.resolve();
  const german = /^de/i.test(Office.context.displayLanguage || '');
  // What to do first (the pane shows the whole story, and signs in where that is needed), then why.
  let text = (german ? 'OSign hat keine Signatur eingefügt; mehr unter Signaturen. ' : 'OSign inserted no signature; see Signatures. ') + localize(reason, german);
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

/* global OSIGN_VERSION */
const VERSION = typeof OSIGN_VERSION !== 'undefined' ? OSIGN_VERSION : '';

// Each step goes to the console (F12, filter "OSign"): when nothing is put in, it shows where it stopped.
function log(m) { try { console.log('OSign ' + VERSION + ': ' + m); } catch (e) { /* none */ } }

const CLASSIC_HINT = 'In classic Outlook, open OSign → Signatures once; from then on it goes in by itself.';

async function sign(event, what) {
  let item = null;
  try {
    item = Office.context.mailbox.item;
    const req = Office.context.requirements;
    log(what + ' started (sign-in through Outlook: ' + (req && req.isSetSupported && req.isSetSupported('NestedAppAuth', '1.1') ? 'yes' : 'no') + ').');
    const cfg = core.withVariants(await loadConfig({ interactive: false }), getVariants());
    log('signatures loaded.');
    if (!isMessage(item) && !(cfg.options && cfg.options.meetings)) { log('a meeting, and meetings are off.'); return; }
    const info = await composeInfo(item);
    const user = await loadUserValues(cfg, { interactive: false });
    if (!user.values) {
      const why = user.error ? user.error.message : 'your details could not be read.';
      log('no user details, no signature. ' + why);
      await showNotice(item, why);
      return;
    }
    const from = await fromAddress(item);
    const sig = core.pick(cfg, user.values, from, info.kind, getPrefs(), user.groups);
    // A changed sender may have no signature where the first one had: then the old one goes.
    if (!sig && what !== 'fromChanged') { log('no signature is meant for ' + (from || 'this sender') + ' (' + info.kind + ').'); return; }
    await insertSignature(item, cfg, sig, user.values, info.html);
    await clearNotice(item);
    log(sig ? 'signature "' + sig.name + '" put in.' : 'signature taken out.');
  } catch (e) {
    let why = e && e.message ? e.message : String(e);
    log(why);
    // Classic Outlook refuses the sign-in here, and no copy is kept yet: what to do, not the code.
    if (isClassicOutlook() && /untrusted_origin|Nested App Auth|sign-in|Signing in/i.test(why)) why = CLASSIC_HINT;
    await showNotice(item, why);
  } finally {
    event.completed();
  }
}

function onNewMessageCompose(event) { log('Outlook called onNewMessageCompose.'); sign(event, 'new'); }
function onNewAppointmentOrganizer(event) { log('Outlook called onNewAppointmentOrganizer.'); sign(event, 'new'); }
function onMessageFromChanged(event) { log('Outlook called onMessageFromChanged.'); sign(event, 'fromChanged'); }

// Errors nobody caught land in the log too, under the same name.
try {
  if (typeof addEventListener === 'function') {
    addEventListener('error', (e) => log('error: ' + (e && e.message)));
    addEventListener('unhandledrejection', (e) => log('error: ' + (e && e.reason && e.reason.message ? e.reason.message : e && e.reason)));
  }
} catch (e) { /* no window in classic Outlook */ }

// Outlook waits for the add-in to say it is ready before it hands over an event.
if (typeof Office.onReady === 'function') {
  Office.onReady((info) => log('Office ready (' + (info && info.host) + ', ' + (info && info.platform) + ').'));
}

// The names are the ones in the manifest's LaunchEvent entries.
Office.actions.associate('onNewMessageComposeHandler', onNewMessageCompose);
Office.actions.associate('onNewAppointmentOrganizerHandler', onNewAppointmentOrganizer);
Office.actions.associate('onMessageFromChangedHandler', onMessageFromChanged);
log('event handler ready.');
