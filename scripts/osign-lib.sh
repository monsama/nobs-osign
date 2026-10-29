#!/bin/bash
# What deploy-osign.sh and scripts/update.sh, backup.sh, restore.sh share. Sourced from the install folder.
PROJECT=osign
PROXY_NET=nobs-proxy

info() { echo "[INFO]  $*"; }
warn() { echo "[WARN]  $*"; }
fail() { echo "[ERROR] $*" >&2; exit 1; }

# A value of .env, without the quotes it may be written in.
env_get() { grep -E "^$1=" .env 2>/dev/null | tail -1 | cut -d= -f2- | sed -E "s/^([\"'])(.*)\\1$/\\2/" || true; }
# Sets a value in .env, in place of what was there.
env_set() { sed -i -E "/^$1=/d" .env; echo "$1=$2" >> .env; }

compose() { docker compose -p "$PROJECT" "$@"; }

# The image of the release in this folder: the version in docker-compose.yml.
release_image() { grep -oE 'ghcr\.io/monsama/osign:[0-9]+\.[0-9]+\.[0-9]+' docker-compose.yml | head -1; }

# The network the shared proxy reaches the app on (made by NOBS Inventory; made here if it is not there yet).
proxy_network() { docker network inspect "$PROXY_NET" >/dev/null 2>&1 || docker network create "$PROXY_NET" >/dev/null; }

wait_healthy() {
    local status
    for _ in $(seq 1 30); do
        status="$(docker inspect -f '{{.State.Health.Status}}' osign 2>/dev/null || echo starting)"
        [ "$status" = healthy ] && return 0
        [ "$status" = unhealthy ] && { docker logs --tail 20 osign 2>&1 | sed 's/^/        /'; return 1; }
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
# scripts/install-cert.sh osign <bundle.pem>) is kept; else the proxy's default one.
proxy_site() {
    local host short names
    host="$(env_get OSIGN_HOST)"
    names="https://$host"; short="${host%%.*}"; [ "$short" != "$host" ] && names="$names, https://$short"
    { echo "# OSign. Written from OSIGN_HOST by deploy-osign.sh and scripts/update.sh of OSign: a change here is overwritten."
      echo "$names {"
      if [ -s "$PROXY_DIR/certificates/sites/osign.crt" ] && [ -s "$PROXY_DIR/certificates/sites/osign.key" ]; then
          printf '\ttls /certs/sites/osign.crt /certs/sites/osign.key\n\timport headers\n'
      else
          printf '\timport tls_and_headers\n'
      fi
      printf '\treverse_proxy osign:8080\n}\n'
    } > "$PROXY_DIR/sites/osign.caddy"
}
