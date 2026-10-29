#!/usr/bin/env node
/* Sets the version everywhere it is kept: package.json, package-lock.json, the editor
   (editor/OSign.ps1, $script:AppVersion) and the server's image (docker-compose.yml). The add-in's manifest takes its version from package.json.
   Usage: node scripts/bump-version.mjs 0.5.0 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const version = process.argv[2];
if (!/^\d+\.\d+\.\d+$/.test(version || '')) { console.error('Give the new version as x.y.z, for example 0.5.0.'); process.exit(1); }

/* Replaces the pattern in the file. Fails when the pattern is not there at all, so a changed file is not missed. */
const edit = (file, pattern, replacement) => {
  const p = path.join(root, file);
  const before = fs.readFileSync(p, 'utf8');
  if (!pattern.test(before)) throw new Error(`${file}: the version was not found`);
  fs.writeFileSync(p, before.replace(pattern, replacement));
};

edit('package.json', /("version": ")[^"]+(")/, `$1${version}$2`);
/* package-lock.json names the app's version twice, at the top and for its root package. Every other "version"
   in it belongs to a dependency, so only the first two are changed. */
{
  const p = path.join(root, 'package-lock.json');
  let n = 0;
  const text = fs.readFileSync(p, 'utf8').replace(/("version": ")[^"]+(")/g, (m, a, b) => (n++ < 2 ? `${a}${version}${b}` : m));
  if (n < 2) throw new Error('package-lock.json: the version was not found');
  fs.writeFileSync(p, text);
}
edit('editor/OSign.ps1', /(\$script:AppVersion = ')[^']+(')/, `$1${version}$2`);
edit('docker-compose.yml', /(ghcr\.io\/monsama\/osign:)\d+\.\d+\.\d+/, `$1${version}`);
console.log(`Version set to ${version} in package.json, package-lock.json, editor/OSign.ps1 and docker-compose.yml.`);
