import { cp, mkdir, readFile, rm } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';

const root = new URL('../', import.meta.url);
const manifest = JSON.parse(await readFile(new URL('src/vendor/integrity.json', root), 'utf8'));
for (const { files } of Object.values(manifest)) {
  for (const [name, expected] of Object.entries(files)) {
    const bytes = await readFile(new URL('src/vendor/' + name, root));
    const actual = createHash('sha256').update(bytes).digest('hex');
    if (actual !== expected) throw new Error('Vendored file integrity mismatch: ' + name);
  }
}
const output = new URL('dist/', root);
await rm(output, { recursive: true, force: true });
await mkdir(output, { recursive: true });
await cp(new URL('src/', root), output, { recursive: true });
console.log('Built static app: ' + fileURLToPath(output));
