// Serves preview/index.html on http://127.0.0.1:5199 for local checks, with a
// fake POST /exec standing in for the Apps Script web app (nothing is saved).
import { readFileSync } from 'node:fs';
import { createServer } from 'node:http';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..', 'preview', 'index.html');
createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/exec') {
    let body = '';
    req.on('data', (c) => (body += c));
    req.on('end', () => {
      const data = JSON.parse(body || '{}');
      console.log('[mock] booking', req.headers['content-type'], 'cookie:', req.headers.cookie ? 'yes' : 'no', 'ageMonths:', data.ageMonths);
      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(/^01\d{9}$/.test(data.phone) ? { ok: true, ref: 'BK-0001' } : { ok: false, error: 'phone' }));
    });
    return;
  }
  res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8', 'Cache-Control': 'no-store' });
  res.end(readFileSync(root));
}).listen(5199, '127.0.0.1', () => console.log('preview on http://127.0.0.1:5199'));
