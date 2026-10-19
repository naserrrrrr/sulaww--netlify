// Copy the single-file build to the site folder served by Netlify: /galaxy/index.html
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const here = path.dirname(fileURLToPath(import.meta.url));
const src = path.resolve(here, '..', 'dist-single', 'index.html');
const dst = path.resolve(here, '..', '..', 'galaxy', 'index.html');
fs.mkdirSync(path.dirname(dst), { recursive: true });
fs.copyFileSync(src, dst);
console.log(`galaxy/index.html ${(fs.statSync(dst).size / 1024 / 1024).toFixed(2)} MB`);
