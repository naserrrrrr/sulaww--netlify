// Render the official rocket Lottie into spritesheets (one per hull colour).
// Usage: node scripts/render-lottie.mjs   (run from game/ via `npm run render:rocket`)
// Output: public/assets/brand/rocket-<colour>.png + rocket-sheet.json (frame meta)
import { chromium } from 'playwright';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(here, '..', '..');
const assets = path.join(root, 'public', 'assets', 'brand');
const lottieLib = fs.readFileSync(path.join(root, 'game', 'node_modules', 'lottie-web', 'build', 'player', 'lottie.min.js'), 'utf8');
const src = JSON.parse(fs.readFileSync(path.join(assets, 'lottie-rocket.json'), 'utf8'));

const FRAMES = 12;
const FRAME_W = 180; // px per frame (1920 Lottie units -> 180 px, rocket is centered)
const FRAME_H = 180;

// Hull palette swaps: official #5143DE hull + its two shades -> brand colours.
const VARIANTS = {
  blue: null,
  orange: { '#5143de': '#fa500e', '#3a299b': '#b83a0a', '#1e1259': '#7a2607', '#271466': '#7a2607' },
  green: { '#5143de': '#00df6a', '#3a299b': '#0a9a4c', '#1e1259': '#06603a', '#271466': '#06603a' },
  rose: { '#5143de': '#eb459f', '#3a299b': '#a62e70', '#1e1259': '#6b1d48', '#271466': '#6b1d48' },
};

const hex = (k) => '#' + k.slice(0, 3).map((v) => Math.round(v * 255).toString(16).padStart(2, '0')).join('');
const rgb = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);

function recolor(data, map) {
  const walk = (o) => {
    if (Array.isArray(o)) return o.forEach(walk);
    if (!o || typeof o !== 'object') return;
    if ((o.ty === 'fl' || o.ty === 'st') && o.c && Array.isArray(o.c.k) && typeof o.c.k[0] === 'number') {
      const to = map[hex(o.c.k)];
      if (to) o.c.k = [...rgb(to), o.c.k[3] ?? 1];
    }
    Object.values(o).forEach(walk);
  };
  walk(data);
  return data;
}

const exe = fs.existsSync('/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell')
  ? '/opt/pw-browsers/chromium_headless_shell-1194/chrome-linux/headless_shell'
  : undefined;
const browser = await chromium.launch({ executablePath: exe });
const page = await browser.newPage({ viewport: { width: FRAME_W * FRAMES, height: FRAME_H }, deviceScaleFactor: 1 });

for (const [name, map] of Object.entries(VARIANTS)) {
  const data = map ? recolor(structuredClone(src), map) : src;
  const cells = Array.from({ length: FRAMES }, (_, i) => `<div class="c" id="c${i}"></div>`).join('');
  await page.setContent(`<style>body{margin:0;background:transparent;display:flex}.c{width:${FRAME_W}px;height:${FRAME_H}px}</style>${cells}<script>${lottieLib}</script><script>
    const d=${JSON.stringify(data)};const n=${FRAMES};
    for(let i=0;i<n;i++){const a=lottie.loadAnimation({container:document.getElementById('c'+i),renderer:'svg',loop:false,autoplay:false,animationData:JSON.parse(JSON.stringify(d))});
      a.goToAndStop(Math.floor(d.ip+(d.op-d.ip)*i/n),true);}
  </script>`);
  await page.waitForTimeout(300);
  await page.screenshot({ path: path.join(assets, `rocket-${name}.png`), omitBackground: true });
  console.log('rocket-' + name + '.png');
}
await browser.close();
fs.writeFileSync(path.join(assets, 'rocket-sheet.json'), JSON.stringify({ frameWidth: FRAME_W, frameHeight: FRAME_H, frames: FRAMES, variants: Object.keys(VARIANTS) }));
