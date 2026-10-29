#!/bin/bash
# Backs up NOBS OSign: its data (the draft, its images, the settings, the users as last read) with .env and
# the secrets, into backup/nobs-osign-<time>.tar.gz (checksummed, kept 14 days). The signatures themselves
# are in SharePoint. With BACKUP_PASSPHRASE_FILE in .env the backup is encrypted (.tar.gz.enc).
#   sudo ./scripts/backup.sh
set -Eeuo pipefail
cd "$(dirname "$0")/.."
# shellcheck source=scripts/osign-lib.sh
. scripts/osign-lib.sh

mkdir -p backup; chmod 700 backup
stamp="$(date +%Y%m%d-%H%M%S)"
work="$(mktemp -d)"; trap 'rm -rf "$work"' EXIT
docker cp nobs-osign:/data - > "$work/data.tar" || fail "The data could not be read: is the container nobs-osign there?"
cp -a .env docker-compose.yml "$work/"
[ -d secrets ] && cp -a secrets "$work/"
out="backup/nobs-osign-$stamp.tar.gz"
tar -czf "$out" -C "$work" .
pass="$(env_get BACKUP_PASSPHRASE_FILE)"
if [ -n "$pass" ]; then
    [ -s "$pass" ] || fail "BACKUP_PASSPHRASE_FILE ($pass) is empty or missing."
    openssl enc -aes-256-cbc -pbkdf2 -iter 200000 -salt -in "$out" -out "$out.enc" -pass "file:$pass"
    rm -f "$out"; out="$out.enc"
fi
chmod 600 "$out"
( cd backup && sha256sum "$(basename "$out")" > "$(basename "$out").sha256" )
find backup -name 'nobs-osign-*' -mtime +14 -delete
info "Backup: $out (it holds the secrets: keep copies as private as they are)"
