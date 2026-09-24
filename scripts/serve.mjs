import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { resolve, relative, extname, sep, isAbsolute } from 'node:path';
import { fileURLToPath } from 'node:url';

// Bind locally by default. Deploy dist/ on a static host for public access.
const root = fileURLToPath(new URL('../dist/', import.meta.url));
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png' };
const port = Number(process.env.PORT || 4173);
if (!Number.isInteger(port) || port < 0 || port > 65535) throw new Error('PORT must be between 0 and 65535.');
const server = createServer(async (request, response) => {
  if (!['GET', 'HEAD'].includes(request.method)) { response.writeHead(405); response.end(); return; }
  try {
    const pathname = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
    let path = resolve(root, '.' + pathname);
    const within = relative(root, path);
    if (within === '..' || within.startsWith('..' + sep) || isAbsolute(within)) { response.writeHead(403); response.end(); return; }
    if ((await stat(path)).isDirectory()) path = resolve(path, 'index.html');
    const data = await readFile(path);
    response.writeHead(200, { 'Content-Type': types[extname(path)] || 'application/octet-stream', 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
    response.end(request.method === 'HEAD' ? undefined : data);
  } catch { response.writeHead(404); response.end('Not found'); }
});
server.on('error', error => {
  console.error(error.code === 'EADDRINUSE' ? 'Port ' + port + ' is busy. Set PORT to another port and try again.' : error.message);
  process.exitCode = 1;
});
server.listen(port, '127.0.0.1', () => console.log('Table Workshop: http://127.0.0.1:' + server.address().port + '/\nKeep this terminal open. Press Ctrl+C to stop.'));
