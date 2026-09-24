import { readFile, readdir, mkdir, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
await import('./build.mjs');
const root = new URL('../', import.meta.url);
const { version } = JSON.parse(await readFile(new URL('package.json', root), 'utf8'));
const zip = new JSZip();
async function add(directory, prefix) {
  const entries = (await readdir(directory, { withFileTypes: true })).sort((a, b) => a.name.localeCompare(b.name));
  for (const entry of entries) {
    if (entry.isDirectory()) await add(new URL(entry.name + '/', directory), prefix + entry.name + '/');
    else zip.file(prefix + entry.name, await readFile(new URL(entry.name, directory)), { date: new Date('2026-01-01T00:00:00Z') });
  }
}
await add(new URL('dist/', root), 'dist/');
for (const file of ['LICENSE', 'THIRD_PARTY_NOTICES.md', 'scripts/serve.mjs']) {
  let content = await readFile(new URL(file, root));
  if (file === 'THIRD_PARTY_NOTICES.md') content = content.toString().replaceAll('src/vendor/', 'dist/vendor/');
  zip.file(file, content, { date: new Date('2026-01-01T00:00:00Z') });
}
zip.file('README.txt', 'Table Workshop ' + version + '\n\nInstall Node.js 22 or newer.\nRun: node scripts/serve.mjs\nOpen http://127.0.0.1:4173/\n\nAlternatively: python3 -m http.server 4173 --bind 127.0.0.1 --directory dist\n\nThe dist folder can be deployed to any static HTTPS host.\nAll table content is processed locally in the browser.\nSource: https://github.com/peta-webster/table-workshop\n', { date: new Date('2026-01-01T00:00:00Z') });
const bytes = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE', compressionOptions: { level: 9 } });
const name = 'table-workshop-v' + version + '-web.zip';
await mkdir(new URL('release/', root), { recursive: true });
await writeFile(new URL('release/' + name, root), bytes);
await writeFile(new URL('release/SHA256SUMS', root), createHash('sha256').update(bytes).digest('hex') + '  ' + name + '\n');
console.log('Packaged release/' + name);
