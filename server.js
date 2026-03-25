// Minimal static file server for Railway deployment
// Serves the Vite-built dist/ folder with correct MIME types and compression support

import http from 'http';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import zlib from 'zlib';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DIST = path.join(__dirname, 'dist');
const PORT = parseInt(process.env.PORT || '3000', 10);

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'application/javascript; charset=utf-8',
  '.css':  'text/css; charset=utf-8',
  '.json': 'application/json',
  '.png':  'image/png',
  '.ico':  'image/x-icon',
  '.svg':  'image/svg+xml',
  '.woff2':'font/woff2',
  '.woff': 'font/woff',
};

const server = http.createServer((req, res) => {
  // Strip query string
  let urlPath = req.url.split('?')[0];

  // Map / to index.html, SPA fallback for everything else
  let filePath = path.join(DIST, urlPath === '/' ? 'index.html' : urlPath);

  // Security: prevent path traversal
  if (!filePath.startsWith(DIST)) {
    res.writeHead(403); res.end('Forbidden'); return;
  }

  // If file doesn't exist, serve index.html (SPA fallback)
  if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
    filePath = path.join(DIST, 'index.html');
  }

  const ext  = path.extname(filePath);
  const mime = MIME[ext] || 'application/octet-stream';
  const data = fs.readFileSync(filePath);

  // Gzip if client supports it and file is compressible
  const compressible = ['.html','.js','.css','.json','.svg'].includes(ext);
  const acceptsGzip  = (req.headers['accept-encoding'] || '').includes('gzip');

  const headers = {
    'Content-Type': mime,
    'Cache-Control': ext === '.html' ? 'no-cache' : 'public, max-age=31536000, immutable',
    'X-Content-Type-Options': 'nosniff',
  };

  if (compressible && acceptsGzip) {
    headers['Content-Encoding'] = 'gzip';
    zlib.gzip(data, (err, compressed) => {
      if (err) { res.writeHead(500); res.end(); return; }
      res.writeHead(200, headers);
      res.end(compressed);
    });
  } else {
    res.writeHead(200, headers);
    res.end(data);
  }
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`LocalLens running on http://0.0.0.0:${PORT}`);
});
