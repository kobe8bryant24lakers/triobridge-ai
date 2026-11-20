import { createServer } from 'node:https';
import { readFile, stat } from 'node:fs/promises';
import { resolve, extname, sep } from 'node:path';
import { tlsOptions, createProxy } from './server-utils.mjs';

try {
  process.loadEnvFile('.env');
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}
const root = resolve('dist');
try {
  await stat(resolve(root, 'index.html'));
} catch {
  throw new Error('请先运行 npm run build。');
}
const port = 3000;
const proxy = createProxy(process.env.MODEL_PROXY_TARGET, `https://localhost:${port}`);
const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'application/javascript',
  '.css': 'text/css',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.json': 'application/json',
};
const server = createServer(tlsOptions(), (req, res) => {
  proxy(req, res, async () => {
    try {
      if (!['localhost:3000', '127.0.0.1:3000'].includes(req.headers.host)) {
        res.writeHead(403);
        res.end();
        return;
      }
      if (req.method !== 'GET' && req.method !== 'HEAD') {
        res.writeHead(405);
        res.end();
        return;
      }
      const path = decodeURIComponent(new URL(req.url, 'https://localhost').pathname);
      const file = resolve(root, `.${path === '/' ? '/index.html' : path}`);
      if (!file.startsWith(root + sep)) {
        res.writeHead(403);
        res.end();
        return;
      }
      const content = await readFile(file);
      res.writeHead(200, {
        'Content-Type': types[extname(file)] || 'application/octet-stream',
        'X-Content-Type-Options': 'nosniff',
        'Cache-Control': file.endsWith('.html') ? 'no-cache' : 'public, max-age=3600',
      });
      res.end(req.method === 'HEAD' ? undefined : content);
    } catch {
      res.writeHead(404);
      res.end('Not found');
    }
  }).catch(() => {
    if (!res.headersSent) res.writeHead(500);
    res.end();
  });
});
server.listen(port, '127.0.0.1', () =>
  console.log(`TrioBridge AI: https://localhost:${port} (loopback only)`),
);
