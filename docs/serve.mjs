import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { resolve, extname } from 'node:path';
import { fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('../build/docs/', import.meta.url));
try { await stat(resolve(root, 'index.html')); }
catch { console.error('Build the documentation first: npm run docs:build'); process.exit(1); }
const types = { '.html': 'text/html; charset=utf-8', '.txt': 'text/plain; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'text/javascript', '.mjs': 'text/javascript', '.json': 'application/json', '.png': 'image/png', '.woff2': 'font/woff2' };
const server = createServer(async (request, response) => {
  if (!['GET', 'HEAD'].includes(request.method)) { response.writeHead(405); response.end(); return; }
  try {
    const url = new URL(request.url, 'http://localhost');
    const pathname = decodeURIComponent(url.pathname);
    if (pathname === '/' || pathname === '/minkexcel') {
      response.writeHead(302, { Location: '/minkexcel/' }); response.end(); return;
    }
    if (!pathname.startsWith('/minkexcel/')) { response.writeHead(404); response.end('Not found'); return; }
    const path = resolve(root, pathname.slice('/minkexcel/'.length) || 'index.html');
    if (!path.startsWith(root)) { response.writeHead(404); response.end('Not found'); return; }
    const content = await readFile(path);
    const extension = extname(path);
    const cache = ['.woff2', '.png'].includes(extension) ? 'public, max-age=3600' : 'no-store';
    response.writeHead(200, { 'Content-Type': types[extension] || 'application/octet-stream', 'Cache-Control': cache });
    response.end(request.method === 'HEAD' ? undefined : content);
  } catch { response.writeHead(404); response.end('Not found'); }
});
server.listen(Number(process.env.DOCS_PORT || 4173), '127.0.0.1', () => {
  console.log(`Documentation preview: http://127.0.0.1:${server.address().port}/minkexcel/`);
});
server.on('error', error => { console.error(error.message); process.exitCode = 1; });
