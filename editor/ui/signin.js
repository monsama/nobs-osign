/*
 * Microsoft's sign-in comes back here (the setup assistant's "Sign in with Microsoft"): the answer
 * is in the fragment (#code=…&state=… or #error=…), never sent to the server. The editor that
 * opened this window reads it from here and closes it; it is also handed over by message, on this
 * site only, where the window still knows its opener.
 */
'use strict';
(function () {
  var p = new URLSearchParams(location.hash.slice(1));
  var answer = { osignSignIn: true, code: p.get('code') || '', state: p.get('state') || '', error: p.get('error') || '', errorText: p.get('error_description') || '' };
  try { if (window.opener) window.opener.postMessage(answer, location.origin); } catch (e) { /* read from here instead */ }
  document.getElementById('msg').textContent = 'You can close this window.';
})();
