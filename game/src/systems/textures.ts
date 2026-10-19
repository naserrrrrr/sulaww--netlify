// Procedural textures for everything the brand guide does not provide
// (meteors, shots, gates, ice, boss parts). Flat shapes, rounded, brand colours only.
import Phaser from 'phaser';
import { hex } from '../config/palette';

const ROCK = 0x3a299b; // official hull shade from the rocket animation
const ROCK_DARK = hex('deep');

function blob(g: Phaser.GameObjects.Graphics, cx: number, cy: number, r: number, seed: number, fill: number): void {
  const pts: Phaser.Math.Vector2[] = [];
  const n = 11;
  let s = seed;
  const rnd = () => ((s = (s * 9301 + 49297) % 233280) / 233280);
  for (let i = 0; i < n; i++) {
    const a = (i / n) * Math.PI * 2;
    const rr = r * (0.78 + rnd() * 0.3);
    pts.push(new Phaser.Math.Vector2(cx + Math.cos(a) * rr, cy + Math.sin(a) * rr));
  }
  g.fillStyle(fill, 1);
  g.fillPoints(pts, true);
  // craters
  g.fillStyle(ROCK_DARK, 0.55);
  for (let i = 0; i < 3; i++) {
    const a = rnd() * Math.PI * 2;
    const d = r * (0.15 + rnd() * 0.45);
    g.fillCircle(cx + Math.cos(a) * d, cy + Math.sin(a) * d, r * (0.12 + rnd() * 0.14));
  }
  // rim light
  g.lineStyle(Math.max(2, r * 0.08), hex('sky'), 0.35);
  g.strokePoints(pts, true);
}

function star(g: Phaser.GameObjects.Graphics, cx: number, cy: number, r: number, color: number, pts = 5): void {
  const v: Phaser.Math.Vector2[] = [];
  for (let i = 0; i < pts * 2; i++) {
    const a = -Math.PI / 2 + (i * Math.PI) / pts;
    const rr = i % 2 === 0 ? r : r * 0.45;
    v.push(new Phaser.Math.Vector2(cx + Math.cos(a) * rr, cy + Math.sin(a) * rr));
  }
  g.fillStyle(color, 1);
  g.fillPoints(v, true);
}

export function buildTextures(scene: Phaser.Scene): void {
  const mk = (key: string, w: number, h: number, draw: (g: Phaser.GameObjects.Graphics) => void) => {
    if (scene.textures.exists(key)) return;
    const g = scene.make.graphics({ x: 0, y: 0 }, false);
    draw(g);
    g.generateTexture(key, w, h);
    g.destroy();
  };

  for (let i = 0; i < 3; i++) mk(`meteor-s${i}`, 80, 80, (g) => blob(g, 40, 40, 34, 11 + i * 7, ROCK));
  for (let i = 0; i < 2; i++) mk(`meteor-b${i}`, 170, 170, (g) => blob(g, 85, 85, 74, 31 + i * 5, ROCK));
  mk('fragment', 36, 36, (g) => blob(g, 18, 18, 15, 77, hex('orange')));
  mk('junk', 40, 40, (g) => {
    g.fillStyle(ROCK_DARK, 1);
    g.fillRoundedRect(4, 4, 32, 32, 8);
    g.lineStyle(3, hex('sky'), 0.8);
    g.strokeRoundedRect(4, 4, 32, 32, 8);
    g.fillStyle(hex('orange'), 1);
    g.fillCircle(20, 20, 6);
  });
  mk('chaser', 52, 52, (g) => {
    g.fillStyle(hex('orange'), 1);
    g.fillCircle(26, 26, 24);
    g.fillStyle(ROCK_DARK, 1);
    g.fillCircle(26, 26, 12);
    g.fillStyle(hex('yellow'), 1);
    g.fillCircle(26, 26, 5);
  });
  mk('bullet', 12, 34, (g) => {
    g.fillStyle(hex('sky'), 1);
    g.fillRoundedRect(0, 0, 12, 34, 6);
    g.fillStyle(0xffffff, 1);
    g.fillRoundedRect(3, 4, 6, 22, 3);
  });
  mk('glow', 128, 128, (g) => {
    for (let r = 64; r > 0; r -= 4) {
      g.fillStyle(0xffffff, 0.035);
      g.fillCircle(64, 64, r);
    }
  });
  mk('particle', 16, 16, (g) => {
    g.fillStyle(0xffffff, 1);
    g.fillCircle(8, 8, 8);
  });
  mk('spark', 6, 6, (g) => {
    g.fillStyle(0xffffff, 1);
    g.fillRect(0, 0, 6, 6);
  });
  mk('speedline', 4, 90, (g) => {
    g.fillStyle(0xffffff, 1);
    g.fillRoundedRect(0, 0, 4, 90, 2);
  });
  mk('shield', 150, 150, (g) => {
    g.lineStyle(6, hex('green'), 0.9);
    g.strokeCircle(75, 75, 68);
    g.fillStyle(hex('green'), 0.12);
    g.fillCircle(75, 75, 68);
  });
  mk('ring', 200, 200, (g) => {
    g.lineStyle(6, hex('orange'), 1);
    g.strokeCircle(100, 100, 96);
  });
  mk('gate', 320, 90, (g) => {
    for (const x of [0, 236]) {
      g.fillStyle(hex('sky'), 1);
      g.fillRoundedRect(x, 10, 84, 70, 18);
      g.fillStyle(hex('water'), 1);
      g.fillRoundedRect(x + 14, 24, 56, 42, 12);
      g.fillStyle(hex('primary'), 1);
      g.fillCircle(x + 42, 45, 9);
    }
    g.lineStyle(4, hex('flamingo'), 0.6);
    g.beginPath();
    for (let i = 0; i < 10; i++) {
      g.moveTo(96 + i * 14, 45);
      g.lineTo(102 + i * 14, 45);
    }
    g.strokePath();
  });
  mk('ice', 130, 56, (g) => {
    g.fillStyle(hex('sky'), 1);
    g.fillRoundedRect(0, 0, 130, 56, 16);
    g.fillStyle(hex('water'), 1);
    g.fillRoundedRect(10, 8, 70, 14, 7);
    g.fillStyle(0xffffff, 0.5);
    g.fillRoundedRect(90, 30, 28, 10, 5);
  });
  mk('boss-rock', 340, 340, (g) => blob(g, 170, 170, 152, 5, ROCK));
  mk('crack', 70, 70, (g) => {
    g.lineStyle(9, hex('orange'), 1);
    g.beginPath();
    g.moveTo(12, 8);
    g.lineTo(30, 30);
    g.lineTo(22, 42);
    g.lineTo(44, 50);
    g.lineTo(58, 64);
    g.strokePath();
    g.lineStyle(4, hex('yellow'), 1);
    g.beginPath();
    g.moveTo(12, 8);
    g.lineTo(30, 30);
    g.lineTo(22, 42);
    g.lineTo(44, 50);
    g.lineTo(58, 64);
    g.strokePath();
  });
  mk('core', 60, 60, (g) => {
    g.fillStyle(hex('flamingo'), 1);
    g.fillCircle(30, 30, 27);
    g.fillStyle(0xffffff, 1);
    g.fillCircle(30, 30, 13);
    g.fillStyle(hex('rose'), 1);
    g.fillCircle(30, 30, 6);
  });
  mk('beam', 70, 1400, (g) => {
    g.fillStyle(hex('orange'), 0.55);
    g.fillRect(0, 0, 70, 1400);
    g.fillStyle(hex('yellow'), 0.9);
    g.fillRect(24, 0, 22, 1400);
  });
  mk('warnline', 8, 1400, (g) => {
    g.fillStyle(hex('orange'), 0.9);
    g.fillRect(0, 0, 8, 1400);
  });
  mk('edge-hazard', 44, 260, (g) => {
    g.fillStyle(hex('orange'), 1);
    for (let i = 0; i < 6; i++) {
      g.fillTriangle(0, i * 44, 44, i * 44 + 22, 0, i * 44 + 44);
    }
  });
  mk('gold-star', 72, 72, (g) => {
    star(g, 36, 36, 34, hex('yellow'));
    star(g, 36, 36, 18, 0xfff4a6);
  });
  mk('heart', 44, 44, (g) => {
    g.fillStyle(hex('green'), 1);
    g.fillCircle(14, 15, 11);
    g.fillCircle(30, 15, 11);
    g.fillTriangle(4, 20, 40, 20, 22, 40);
  });
  mk('pw-shield', 84, 84, (g) => {
    g.fillStyle(hex('green'), 1);
    g.fillCircle(42, 42, 40);
    g.lineStyle(6, 0xffffff, 1);
    g.strokeCircle(42, 42, 24);
  });
  mk('pw-magnet', 84, 84, (g) => {
    g.fillStyle(hex('rose'), 1);
    g.fillCircle(42, 42, 40);
    g.lineStyle(10, 0xffffff, 1);
    g.beginPath();
    g.arc(42, 40, 18, Math.PI, 0, false);
    g.strokePath();
    g.fillStyle(0xffffff, 1);
    g.fillRect(19, 40, 10, 16);
    g.fillRect(55, 40, 10, 16);
  });
  mk('pw-triple', 84, 84, (g) => {
    g.fillStyle(hex('flamingo'), 1);
    g.fillCircle(42, 42, 40);
    g.fillStyle(0xffffff, 1);
    for (const x of [26, 42, 58]) g.fillRoundedRect(x - 4, 20, 8, 44, 4);
  });
  mk('wall', 110, 110, (g) => blob(g, 55, 55, 50, 99, hex('primary')));
}
