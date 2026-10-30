// Stands in for addin/src/auth.js in the tests: Outlook's sign-in (nested app authentication) is
// not there outside Outlook. The token is shaped as Entra's (its claims carry the scopes, scp),
// and names its scopes after it too, so the stand-in Graph can check them. No btoa here: classic
// Outlook's runtime has none.
const B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_';
function b64url(text) {
  const bytes = unescape(encodeURIComponent(text)).split('').map((c) => c.charCodeAt(0));
  let out = '', bits = 0, n = 0;
  for (const b of bytes) { n = (n << 8) | b; bits += 8; while (bits >= 6) { bits -= 6; out += B64[(n >> bits) & 63]; } }
  if (bits) out += B64[(n << (6 - bits)) & 63];
  return out;
}
export async function getGraphToken(interactive, scopes) {
  if (globalThis.__osignSignInFails) throw new Error('sign-in failed');
  return b64url('{"alg":"none"}') + '.' + b64url(JSON.stringify({ scp: (scopes || []).join(' ') })) + '.token:' + (scopes || []).join(' ');
}
