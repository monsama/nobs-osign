/*
 * The event handler: Outlook runs it when a new message, reply, forward or meeting is started, and
 * when the From address changes. It shows nothing; it puts the right signature in, or none.
 */
import { core, loadConfig, loadUserValues, getPrefs, composeInfo, fromAddress, insertSignature, isMessage } from './common.js';

/* global Office */

async function sign(event, what) {
  try {
    const item = Office.context.mailbox.item;
    const cfg = await loadConfig({ interactive: false });
    if (!isMessage(item) && !(cfg.options && cfg.options.meetings)) return;
    const info = await composeInfo(item);
    const user = await loadUserValues(cfg, { interactive: false });
    if (!user.values) { console.log('OSign: no user details, no signature. ' + (user.error ? user.error.message : '')); return; }
    const from = await fromAddress(item);
    const sig = core.pick(cfg, user.values, from, info.kind, getPrefs(), user.groups);
    // A changed sender may have no signature where the first one had: then the old one goes.
    if (!sig && what !== 'fromChanged') return;
    await insertSignature(item, cfg, sig, user.values, info.html);
  } catch (e) {
    console.log('OSign: ' + (e && e.message ? e.message : e));
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
