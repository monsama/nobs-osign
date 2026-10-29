#!/bin/bash
# What deploy-nobs-osign.sh and scripts/update.sh, backup.sh, restore.sh share. Sourced from the install folder.
PROJECT=nobs-osign
PROXY_NET=nobs-proxy

info() { echo "[INFO]  $*"; }
warn() { echo "[WARN]  $*"; }
fail() { echo "[ERROR] $*" >&2; exit 1; }

# A value of .env, without the quotes it may be written in.
env_get() { grep -E "^$1=" .env 2>/dev/null | tail -1 | cut -d= -f2- | sed -E "s/^([\"'])(.*)\\1$/\\2/" || true; }
# Sets a value in .env, in place of what was there.
env_set() { sed -i -E "/^$1=/d" .env; echo "$1=$2" >> .env; }

compose() { docker compose -p "$PROJECT" "$@"; }

# The key the passwords set in the editor are kept with (secrets/app_key): made once, never changed;
# without it they cannot be read. Readable by the app's user (UID 1000) only.
app_key_ensure() {
    [ -s secrets/app_key ] && return 0
    mkdir -p secrets; chmod 700 secrets
    ( umask 077; openssl rand -base64 48 | tr -d '\n' > secrets/app_key )
    chown 1000:1000 secrets/app_key; chmod 400 secrets/app_key
    info "secrets/app_key generated"
}

# The image of the release in this folder: the version in docker-compose.yml.
release_image() { grep -oE 'ghcr\.io/monsama/nobs-osign:[0-9]+\.[0-9]+\.[0-9]+' docker-compose.yml | head -1; }

# The network the shared proxy reaches the app on (made by NOBS Inventory; made here if it is not there yet).
proxy_network() { docker network inspect "$PROXY_NET" >/dev/null 2>&1 || docker network create "$PROXY_NET" >/dev/null; }

wait_healthy() {
    local status
    for _ in $(seq 1 30); do
        status="$(docker inspect -f '{{.State.Health.Status}}' nobs-osign 2>/dev/null || echo starting)"
        [ "$status" = healthy ] && return 0
        [ "$status" = unhealthy ] && { docker logs --tail 20 nobs-osign 2>&1 | sed 's/^/        /'; return 1; }
        sleep 2
    done
    return 1
}

# The shared proxy's folder: OSIGN_PROXY_DIR, or where the running proxy has its sites from.
proxy_find() {
    PROXY_DIR="${OSIGN_PROXY_DIR:-}"
    if [ -z "$PROXY_DIR" ] && docker inspect nobs-proxy >/dev/null 2>&1; then
        PROXY_DIR="$(docker inspect -f '{{range .Mounts}}{{if eq .Destination "/etc/caddy/sites"}}{{.Source}}{{end}}{{end}}' nobs-proxy)"
        PROXY_DIR="${PROXY_DIR%/sites}"
    fi
    [ -n "$PROXY_DIR" ] && [ -d "$PROXY_DIR/sites" ]
}

# The site: the name people type, and its short form for an internal name. Written each time, so a
# changed OSIGN_HOST takes effect. A certificate of its own (NOBS Inventory's
# scripts/install-cert.sh nobs-osign <bundle.pem>) is kept; else the proxy's default one.
proxy_site() {
    local host short names
    host="$(env_get OSIGN_HOST)"
    # A name and nothing else, as the deploy script checks it: it goes into the proxy's site file.
    [[ "$host" =~ ^[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?)+$ ]] || fail "OSIGN_HOST in .env is not a host name: $host"
    names="https://$host"; short="${host%%.*}"; [ "$short" != "$host" ] && names="$names, https://$short"
    { echo "# NOBS OSign. Written from OSIGN_HOST by deploy-nobs-osign.sh and scripts/update.sh of NOBS OSign: a change here is overwritten."
      echo "$names {"
      if [ -s "$PROXY_DIR/certificates/sites/nobs-osign.crt" ] && [ -s "$PROXY_DIR/certificates/sites/nobs-osign.key" ]; then
          printf '\ttls /certs/sites/nobs-osign.crt /certs/sites/nobs-osign.key\n\timport headers\n'
      else
          printf '\timport tls_and_headers\n'
      fi
      printf '\treverse_proxy nobs-osign:8080\n}\n'
    } > "$PROXY_DIR/sites/nobs-osign.caddy"
}
