#!/bin/bash
# ============================================================
# deploy-osign.sh - the OSign editor with Docker Compose, behind the shared proxy of NOBS Inventory
#
#   sudo ./deploy-osign.sh        first install, or re-run after changes (safe to repeat)
#
# Creates (never overwrites existing values):
#   .env                         non-secret settings (from .env.example)
#   secrets/admin_password       the local account "admin" (generated, shown once at the end)
#   secrets/ldap_service_password, secrets/graph_client_secret, secrets/cf_api_token
#                                empty until you fill them in (see .env.example)
#
# Then starts the app, waits until it is healthy, and adds its site to the shared proxy (nobs-proxy,
# installed with NOBS Inventory): osign.caddy in the proxy's sites, reloaded. Nothing is opened to
# the internet.
#
# Environment overrides (optional):
#   OSIGN_HOST=osign.example.ch      the name people type (asked for on the first run)
#   OSIGN_PROXY_DIR=/opt/nobs-inventory/proxy   the shared proxy's folder (found from the running proxy)
#   INSTALL_BACKUP_CRON=true         daily backup at 03:25 (default: true)
# ============================================================
set -Eeuo pipefail
trap 'echo "[ERROR] Failed at line $LINENO: $BASH_COMMAND" >&2' ERR
cd "$(dirname "$0")"
BASE_DIR="$(pwd)"
INSTALL_BACKUP_CRON="${INSTALL_BACKUP_CRON:-true}"
# shellcheck source=scripts/osign-lib.sh
. "$BASE_DIR/scripts/osign-lib.sh"

step_preflight() {
    echo; echo "Step 1: Preflight"
    if [ "$(id -u)" -ne 0 ] && ! id -nG | grep -qw docker; then fail "Run with sudo (or as a member of the docker group)."; fi
    command -v docker >/dev/null || fail "Docker is not installed. Install Docker Engine and the compose plugin first."
    docker compose version >/dev/null 2>&1 || fail "The Docker Compose plugin is missing."
    command -v openssl >/dev/null || fail "openssl is missing."
    docker info >/dev/null 2>&1 || fail "The Docker daemon is not running."
    info "OK"
}

step_env() {
    echo; echo "Step 2: Settings (.env)"
    local host="${OSIGN_HOST:-}"
    [ -f .env ] || { cp .env.example .env; info ".env created from .env.example"; }
    [ -n "$host" ] || host="$(env_get OSIGN_HOST)"
    if [ -z "$host" ] || [ "$host" = osign.example.ch ]; then
        if [ -t 0 ]; then read -r -p "Name people will type (e.g. osign.example.ch): " host; fi
        [ -n "$host" ] && [ "$host" != osign.example.ch ] || fail "Set OSIGN_HOST, e.g. OSIGN_HOST=osign.example.ch sudo ./deploy-osign.sh"
    fi
    env_set OSIGN_HOST "$host"
    # The image follows the release that was installed. A custom image name is kept as it is.
    local img; img="$(env_get OSIGN_IMAGE)"
    if [ -z "$img" ] || [[ "$img" == ghcr.io/monsama/osign:* ]]; then env_set OSIGN_IMAGE "$(release_image)"; fi
    chmod 640 .env
    info ".env: OSIGN_HOST=$host, OSIGN_IMAGE=$(env_get OSIGN_IMAGE)"
}

step_secrets() {
    echo; echo "Step 3: Secrets (./secrets, never committed)"
    mkdir -p secrets certificates; chmod 700 secrets
    FIRST_PASSWORD=""
    if [ ! -s secrets/admin_password ]; then
        FIRST_PASSWORD="$(openssl rand -base64 48 | tr -dc 'A-Za-z0-9' | head -c 24)"
        ( umask 077; printf '%s' "$FIRST_PASSWORD" > secrets/admin_password )
        info "secrets/admin_password generated"
    else
        info "secrets/admin_password exists - kept"
    fi
    for s in ldap_service_password graph_client_secret cf_api_token; do
        [ -f "secrets/$s" ] || { ( umask 077; : > "secrets/$s" ); info "secrets/$s created (empty until you fill it in)"; }
    done
    # The app runs as the 'node' user (UID 1000) in the container: it must own the files to read them.
    chown 1000:1000 secrets/*
    chmod 400 secrets/*
}

step_start() {
    echo; echo "Step 4: Image and start"
    local image; image="$(env_get OSIGN_IMAGE)"
    if docker pull "$image" >/dev/null 2>&1; then info "Pulled $image"
    else warn "Could not pull $image - building it from this folder instead."; docker build -t "$image" . >/dev/null; fi
    proxy_network
    compose up -d --no-build
    wait_healthy || fail "The app did not become healthy. See: docker logs osign"
    info "App healthy"
}

step_proxy() {
    echo; echo "Step 5: Shared proxy (HTTPS)"
    proxy_find || { warn "The shared proxy (nobs-proxy) is not running here: install NOBS Inventory first, or put your own proxy in front of osign:8080 on the Docker network nobs-proxy."; return 0; }
    proxy_site
    docker restart nobs-proxy >/dev/null
    info "Site added: $PROXY_DIR/sites/osign.caddy (proxy reloaded)"
    local host code; host="$(env_get OSIGN_HOST)"; sleep 3
    code="$(curl -sk -o /dev/null -w '%{http_code}' --resolve "$host:443:127.0.0.1" "https://$host/health" || true)"
    if [ "$code" = 200 ]; then info "HTTPS OK: https://$host/"; else warn "HTTPS check returned ${code:-none}. Check: docker logs nobs-proxy"; fi
}

step_cron() {
    echo; echo "Step 6: Backups"
    if [ "$INSTALL_BACKUP_CRON" != true ]; then info "Daily backup not installed (INSTALL_BACKUP_CRON=false)"; return; fi
    if [ "$(id -u)" -ne 0 ]; then warn "Not root - daily backup cron not installed."; return; fi
    cat > /etc/cron.d/osign-backup <<CRON
# OSign backup - installed by deploy-osign.sh
SHELL=/bin/bash
PATH=/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin
25 3 * * * root $BASE_DIR/scripts/backup.sh >> /var/log/osign-backup.log 2>&1
CRON
    chmod 644 /etc/cron.d/osign-backup
    info "Daily backup at 03:25 -> $BASE_DIR/backup/ (log: /var/log/osign-backup.log)"
}

summary() {
    echo; echo "=========================================="
    echo "  OSign is running"
    echo "=========================================="
    echo "  URL:     https://$(env_get OSIGN_HOST)/"
    if [ -n "$FIRST_PASSWORD" ]; then
        echo "  Admin:   login 'admin', password below - shown ONCE:"
        echo "           $FIRST_PASSWORD"
        echo "           It is also in $BASE_DIR/secrets/admin_password"
    else
        echo "  Admin:   login 'admin'. Password: $BASE_DIR/secrets/admin_password (sudo cat to read it)"
    fi
    echo "  Next:    fill in .env and ./secrets (Active Directory, Graph, Cloudflare; docs/SETUP.md),"
    echo "           then sudo ./deploy-osign.sh again. The setup assistant in the editor does the rest."
    echo "  Update:  sudo ./scripts/update.sh      Backup: sudo ./scripts/backup.sh"
    echo "  Logs:    docker logs -f osign"
    echo "=========================================="
}

step_preflight
step_env
step_secrets
step_start
step_proxy
step_cron
summary
