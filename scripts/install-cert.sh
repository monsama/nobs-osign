#!/bin/bash
# Installs the web certificate for the editor's address (OSIGN_HOST) in the shared proxy, in place
# of the proxy's own one: issued by your CA (in AD: the template WebServer), so the PCs trust it, the
# browser does not warn and the setup script can send its result.
#   sudo ./scripts/install-cert.sh osign.pfx           a PFX/P12 with the key (asks for its password)
#   sudo ./scripts/install-cert.sh bundle.pem          the key and the certificate(s) in one PEM file
#   sudo ./scripts/install-cert.sh --remove            back to the proxy's own
# Checked before it goes in: the key belongs to the certificate, it holds now and for OSIGN_HOST.
# Kept in the proxy's folder (certificates/sites/nobs-osign.crt and .key); updates keep it.
set -Eeuo pipefail
cd "$(cd "$(dirname "$0")/.." && pwd)"
# shellcheck source=scripts/osign-lib.sh
. scripts/osign-lib.sh

[ "$(id -u)" = 0 ] || fail "Run it as root: sudo ./scripts/install-cert.sh <file>"
command -v openssl >/dev/null || fail "openssl is missing."
host="$(env_get OSIGN_HOST)"
[ -n "$host" ] || fail "OSIGN_HOST is not set in .env."
proxy_find || fail "The shared proxy (nobs-proxy) is not running here, or OSIGN_PROXY_DIR does not name its folder."
dir="$PROXY_DIR/certificates/sites"
crt="$dir/nobs-osign.crt"; key="$dir/nobs-osign.key"

reload() {
    proxy_site
    docker restart nobs-proxy >/dev/null
    sleep 3
}

if [ "${1:-}" = "--remove" ]; then
    rm -f "$crt" "$key"
    reload
    info "Removed: the proxy serves its own certificate for $host again."
    exit 0
fi

src="${1:-}"
[ -n "$src" ] && [ -f "$src" ] || fail "Name the certificate file: sudo ./scripts/install-cert.sh osign.pfx (or bundle.pem)"
work="$(mktemp -d)"; chmod 700 "$work"
trap 'rm -rf "$work"' EXIT

# Into PEM: a PFX with its password (OpenSSL 3 needs -legacy for older ones), a PEM as it is.
case "${src,,}" in
    *.pfx|*.p12)
        pass="${OSIGN_PFX_PASSWORD:-}"
        if [ -z "$pass" ]; then read -r -s -p "Password of $src: " pass; echo; fi
        export OSIGN_PFX_PASS="$pass"
        openssl pkcs12 -in "$src" -nodes -passin env:OSIGN_PFX_PASS -out "$work/all.pem" 2>/dev/null \
            || openssl pkcs12 -legacy -in "$src" -nodes -passin env:OSIGN_PFX_PASS -out "$work/all.pem" 2>/dev/null \
            || fail "The PFX could not be opened: wrong password, or not a PFX."
        unset OSIGN_PFX_PASS ;;
    *) cp "$src" "$work/all.pem" ;;
esac

# The key, and each certificate in the file.
openssl pkey -in "$work/all.pem" -out "$work/key.pem" 2>/dev/null || fail "There is no private key in $src (export it with the key: PFX, or PEM with the key)."
awk -v d="$work" '/-----BEGIN CERTIFICATE-----/{n++; f=d"/x509-"n".pem"} f{print > f} /-----END CERTIFICATE-----/{f=""}' "$work/all.pem"
ls "$work"/x509-*.pem >/dev/null 2>&1 || fail "There is no certificate in $src."

# The one for the key comes first, the CA's after it.
want="$(openssl pkey -in "$work/key.pem" -pubout 2>/dev/null | openssl sha256 | awk '{print $NF}')"
leaf=""
for c in "$work"/x509-*.pem; do
    [ "$(openssl x509 -in "$c" -noout -pubkey 2>/dev/null | openssl sha256 | awk '{print $NF}')" = "$want" ] && { leaf="$c"; break; }
done
[ -n "$leaf" ] || fail "None of the certificates in $src belongs to its key."

openssl x509 -in "$leaf" -noout -checkend 0 >/dev/null || fail "The certificate has run out ($(openssl x509 -in "$leaf" -noout -enddate | cut -d= -f2))."
names="$(openssl x509 -in "$leaf" -noout -ext subjectAltName 2>/dev/null | grep -oE 'DNS:[^,]+' | cut -d: -f2 | tr '\n' ' ')"
covers() { local n; for n in $names; do [ "${n,,}" = "${1,,}" ] && return 0; [[ "$n" == \*.* ]] && [[ "${1,,}" == *."${n#\*.}" ]] && return 0; done; return 1; }
covers "$host" || fail "The certificate is not for $host (it is for: ${names:-no names}). Issue one with $host as a DNS name."
short="${host%%.*}"
[ "$short" != "$host" ] && ! covers "$short" && warn "It is not for the short name $short: https://$short/ will warn. Add DNS:$short when you issue the next one."

{ cat "$leaf"; for c in "$work"/x509-*.pem; do [ "$c" != "$leaf" ] && cat "$c"; done; } > "$work/chain.pem"
mkdir -p "$dir"
install -m 644 "$work/chain.pem" "$crt"
install -m 600 "$work/key.pem" "$key"
reload

info "Installed for $host: $(openssl x509 -in "$leaf" -noout -subject | sed 's/^subject=//'), issued by $(openssl x509 -in "$leaf" -noout -issuer | sed 's/^issuer=//'), valid until $(openssl x509 -in "$leaf" -noout -enddate | cut -d= -f2)."
served="$(echo | openssl s_client -connect 127.0.0.1:443 -servername "$host" 2>/dev/null | openssl x509 -noout -fingerprint -sha256 2>/dev/null | cut -d= -f2 || true)"
mine="$(openssl x509 -in "$leaf" -noout -fingerprint -sha256 | cut -d= -f2)"
if [ "$served" = "$mine" ]; then info "The proxy serves it: https://$host/"; else warn "The proxy does not serve it yet. See: docker logs nobs-proxy"; fi
info "Delete $src now: it holds the private key."
