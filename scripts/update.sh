#!/bin/bash
# Updates the NOBS OSign server to the latest release, or to the one given: downloads its files (the
# archive and its checksum from GitHub), backs up, replaces the files of this folder (never .env,
# secrets, certificates or backup) and starts the release image. Switches back when the new one does
# not become healthy.
#   sudo ./scripts/update.sh            the latest release
#   sudo ./scripts/update.sh 0.5.0      that release
set -Eeuo pipefail
# The update replaces this very file, and bash reads a script as it goes: it runs from a copy.
BASE_DIR="${OSIGN_DIR:-$(cd "$(dirname "$0")/.." && pwd)}"
if [ -z "${OSIGN_UPDATE_COPY:-}" ]; then
    copy="$(mktemp)"; cp "$0" "$copy"
    OSIGN_UPDATE_COPY="$copy" OSIGN_DIR="$BASE_DIR" exec bash "$copy" "$@"
fi
trap 'rm -f "$OSIGN_UPDATE_COPY"' EXIT
cd "$BASE_DIR"
# shellcheck source=scripts/osign-lib.sh
. scripts/osign-lib.sh
# The repository as it was first named: GitHub forwards it once it is renamed.
REPO=monsama/osign

command -v curl >/dev/null || fail "curl is missing."
want="${1:-}"; want="${want#v}"
if [ -z "$want" ]; then
    want="$(curl -fsSL "https://api.github.com/repos/$REPO/releases/latest" | grep -oE '"tag_name": *"v[0-9]+\.[0-9]+\.[0-9]+"' | grep -oE '[0-9]+\.[0-9]+\.[0-9]+')" || fail "The latest release could not be found on GitHub."
fi
[[ "$want" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] || fail "Not a version: $want"
current="$(env_get OSIGN_IMAGE)"; current="${current##*:}"
[ "$current" = "$want" ] && { info "NOBS OSign $want is installed already."; exit 0; }
info "Updating from ${current:-?} to $want"

work="$(mktemp -d)"; trap 'rm -rf "$work" "$OSIGN_UPDATE_COPY"' EXIT
name="nobs-osign-server-$want.tar.gz"
curl -fsSL -o "$work/$name" "https://github.com/$REPO/releases/download/v$want/$name" || fail "The release $want has no $name."
curl -fsSL -o "$work/$name.sha256" "https://github.com/$REPO/releases/download/v$want/$name.sha256" || fail "The release $want has no checksum."
( cd "$work" && sha256sum -c --quiet "$name.sha256" ) || fail "The checksum does not match: the download is damaged."

./scripts/backup.sh
# The files before the update, to switch back to.
tar -czf "backup/code-before-update-$(date +%Y%m%d-%H%M%S).tar.gz" --exclude=./backup --exclude=./secrets --exclude=./.env --exclude=./certificates .
old_image="$(env_get OSIGN_IMAGE)"
tar -xzf "$work/$name" -C "$work"
src="$(find "$work" -mindepth 1 -maxdepth 1 -type d -name 'nobs-osign-server-*' | head -1)"
[ -n "$src" ] || fail "The archive does not hold the server's files."
cp -a "$src"/. "$BASE_DIR"/
env_set OSIGN_IMAGE "ghcr.io/monsama/nobs-osign:$want"
docker pull "$(env_get OSIGN_IMAGE)" >/dev/null || fail "The image of $want could not be pulled."
compose up -d --no-build >/dev/null
if wait_healthy; then
    proxy_find && { proxy_site; docker restart nobs-proxy >/dev/null; }
    info "NOBS OSign $want is running."
else
    warn "NOBS OSign $want did not become healthy: switching back to $old_image."
    docker logs --tail 20 nobs-osign 2>&1 | sed 's/^/        /'
    env_set OSIGN_IMAGE "$old_image"
    compose up -d --no-build >/dev/null
    if wait_healthy; then fail "Switched back to $old_image; it runs. The update to $want did not go through."; fi
    fail "Switched back to $old_image, but it is not healthy either. See: docker logs nobs-osign"
fi
