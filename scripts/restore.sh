#!/bin/bash
# Restores OSign's data from a backup of scripts/backup.sh. .env and the secrets only with
# RESTORE_SECRETS=yes.
#   sudo ./scripts/restore.sh backup/osign-<time>.tar.gz[.enc]
set -Eeuo pipefail
cd "$(dirname "$0")/.."
# shellcheck source=scripts/osign-lib.sh
. scripts/osign-lib.sh

file="${1:-}"
[ -f "$file" ] || fail "Give the backup file: sudo ./scripts/restore.sh backup/osign-<time>.tar.gz"
if [ -f "$file.sha256" ]; then ( cd "$(dirname "$file")" && sha256sum -c --quiet "$(basename "$file").sha256" ) || fail "The checksum does not match: the backup is damaged."; fi
read -r -p "This replaces OSign's draft and settings with the backup. Type RESTORE to go on: " ok
[ "$ok" = RESTORE ] || fail "Not restored."
work="$(mktemp -d)"; trap 'rm -rf "$work"' EXIT
if [[ "$file" == *.enc ]]; then
    pass="$(env_get BACKUP_PASSPHRASE_FILE)"
    if [ -n "$pass" ] && [ -s "$pass" ]; then openssl enc -d -aes-256-cbc -pbkdf2 -iter 200000 -in "$file" -out "$work/b.tar.gz" -pass "file:$pass"
    else openssl enc -d -aes-256-cbc -pbkdf2 -iter 200000 -in "$file" -out "$work/b.tar.gz"; fi
    file="$work/b.tar.gz"
fi
tar -xzf "$file" -C "$work"
[ -f "$work/data.tar" ] || fail "This is not a backup of OSign (no data.tar in it)."
image="$(env_get OSIGN_IMAGE)"
compose stop app >/dev/null
# Into the volume, as the app's user, with the app's own image (it has tar).
docker run --rm --user 0 -v osign-data:/data -v "$work:/r:ro" --entrypoint sh "$image" -c 'rm -rf /data/* && tar -xf /r/data.tar -C / && chown -R 1000:1000 /data'
if [ "${RESTORE_SECRETS:-}" = yes ]; then
    cp -a "$work/.env" .env
    [ -d "$work/secrets" ] && { rm -rf secrets; cp -a "$work/secrets" secrets; }
    info ".env and secrets restored"
fi
compose up -d --no-build >/dev/null
wait_healthy || fail "The app did not become healthy after the restore. See: docker logs osign"
info "Restored from $1"
