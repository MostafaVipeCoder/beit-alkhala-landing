// Builds the static copy of the page for GitHub Pages: fills the template the
// way HtmlService does (logo + CONFIG from Code.gs) and writes docs/index.html.
// Bookings go to CONFIG.bookingUrl (the Apps Script web app's doPost).
//   node tools/build.mjs            ->  docs/index.html
//   node tools/build.mjs --preview  ->  preview/index.html, posting to the local mock in serve.mjs

import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import vm from 'node:vm';

const preview = process.argv.includes('--preview');
const code = readFileSync('Code.gs', 'utf8');
const sandbox = {};
vm.runInNewContext(`${code.slice(0, code.indexOf('const SHEET_NAME'))}; this.CONFIG = CONFIG;`, sandbox);
const config = { ...sandbox.CONFIG };
if (preview) config.bookingUrl = '/exec';
if (!/^(https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec|\/exec)$/.test(config.bookingUrl || '')) {
  throw new Error('CONFIG.bookingUrl in Code.gs must be the web app /exec link');
}

const html = readFileSync('Index.html', 'utf8')
  .replace("<?!= include('Logo'); ?>", readFileSync('Logo.html', 'utf8'))
  .replace('<?!= JSON.stringify(config) ?>', JSON.stringify(config).replace(/</g, '\\u003c'));
if (html.includes('<?')) throw new Error('unfilled template tag in Index.html');

const out = preview ? 'preview' : 'docs';
mkdirSync(out, { recursive: true });
writeFileSync(`${out}/index.html`, html);
console.log(`${out}/index.html written (bookings -> ${config.bookingUrl})`);
