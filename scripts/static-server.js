import http from 'node:http';
import fs from 'node:fs/promises';
import path from 'node:path';

const mime = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml'
};

function within(root, target) {
  const relative = path.relative(root, target);
  return relative !== '..' && !relative.startsWith('..' + path.sep) && !path.isAbsolute(relative);
}

export function createStaticServer({ root, assets = new Map() }) {
  return http.createServer(async (req, res) => {
    if (!['GET', 'HEAD'].includes(req.method)) {
      res.writeHead(405, { Allow: 'GET, HEAD' }).end('Method not allowed');
      return;
    }
    let pathname;
    try {
      pathname = decodeURIComponent(new URL(req.url, 'http://localhost').pathname);
      if (/[\\\0]/.test(pathname)) throw new Error('Invalid path');
    } catch {
      res.writeHead(400).end('Bad request');
      return;
    }
    const asset = assets.get(pathname);
    const target = asset || path.resolve(root, '.' + (pathname === '/' ? '/index.html' : pathname));
    if (!asset && !within(root, target)) {
      res.writeHead(403).end('Forbidden');
      return;
    }
    try {
      // Also prevent symbolic links inside public/ from exposing project files.
      if (!asset && !within(await fs.realpath(root), await fs.realpath(target))) {
        res.writeHead(403).end('Forbidden');
        return;
      }
      const data = await fs.readFile(target);
      res.writeHead(200, {
        'Content-Type': mime[path.extname(target)] || 'application/octet-stream',
        'Content-Length': data.length,
        'Cache-Control': 'no-cache',
        'X-Content-Type-Options': 'nosniff'
      });
      res.end(req.method === 'HEAD' ? undefined : data);
    } catch (error) {
      const missing = ['ENOENT', 'ENOTDIR', 'EISDIR'].includes(error.code);
      if (!missing) console.error('静态资源读取失败', error);
      res.writeHead(missing ? 404 : 500).end(missing ? 'Not found' : 'Internal server error');
    }
  });
}
