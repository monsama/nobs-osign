#!/usr/bin/env bash
# Makes a release from main: checks the tree, bumps the version if needed, runs the tests, and pushes the tag.
# The release workflow then builds and publishes it. Usage: scripts/release.sh 0.5.0
set -Eeuo pipefail
cd "$(dirname "$0")/.."

V="${1:-}"; V="${V#v}"
[[ "$V" =~ ^[0-9]+\.[0-9]+\.[0-9]+$ ]] || { echo "Give the version as x.y.z, for example 0.5.0." >&2; exit 1; }
[ "$(git rev-parse --abbrev-ref HEAD)" = main ] || { echo "Release from main." >&2; exit 1; }
[ -z "$(git status --porcelain)" ] || { echo "The tree has changes. Commit or stash them first." >&2; exit 1; }
git fetch -q origin main
[ "$(git rev-parse HEAD)" = "$(git rev-parse origin/main)" ] || { echo "main is not origin/main. Pull or push first." >&2; exit 1; }
git rev-parse -q --verify "refs/tags/v$V" >/dev/null && { echo "The tag v$V already exists here." >&2; exit 1; }
git ls-remote --tags origin "v$V" | grep -q . && { echo "The tag v$V already exists on origin." >&2; exit 1; }

npm ci --silent
npm test
npm audit --omit=dev --audit-level=high

if [ "$(node -p 'require("./package.json").version')" != "$V" ]; then
  node scripts/bump-version.mjs "$V"
  git add -A
  git commit -q -m "Release $V"
  git push -q origin main
  echo "Version $V committed and pushed to main."
fi

git tag -a "v$V" -m "v$V"
git push -q origin "refs/tags/v$V"
echo "Tag v$V pushed. The release workflow now builds and publishes it."
