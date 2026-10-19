// DOM overlay helpers: screens live in #ui, in-game HUD in #hud, Bayan's radio in #radio.
import lottie, { AnimationItem } from 'lottie-web/build/player/lottie_svg';
import { sfx } from '../systems/audio';
import { GAME_W } from '../config/balance';

export const $ = <T extends HTMLElement = HTMLElement>(sel: string, root: ParentNode = document): T => {
  const el = root.querySelector<T>(sel);
  if (!el) throw new Error(`missing ${sel}`);
  return el;
};

const ui = () => $('#ui');
let liveAnims: AnimationItem[] = [];

export function showScreen(html: string, opts: { clear?: boolean } = {}): HTMLElement {
  const root = ui();
  killLotties();
  root.innerHTML = html;
  root.classList.toggle('clear', !!opts.clear);
  root.hidden = false;
  root.querySelectorAll('button').forEach((b) =>
    b.addEventListener('pointerdown', () => sfx.unlock(), { once: true }),
  );
  root.querySelectorAll('button:not(.silent)').forEach((b) => b.addEventListener('click', () => sfx.play('click')));
  return root;
}

export function hideScreen(): void {
  killLotties();
  const root = ui();
  root.hidden = true;
  root.innerHTML = '';
}

export function killLotties(): void {
  liveAnims.forEach((a) => a.destroy());
  liveAnims = [];
}

export function mountLottie(container: HTMLElement, data: unknown, opts: { loop?: boolean; speed?: number } = {}): AnimationItem {
  const anim = lottie.loadAnimation({
    container,
    renderer: 'svg',
    loop: opts.loop ?? true,
    autoplay: true,
    animationData: structuredClone(data),
    rendererSettings: { preserveAspectRatio: 'xMidYMid meet' },
  });
  if (opts.speed) anim.setSpeed(opts.speed);
  liveAnims.push(anim);
  return anim;
}

export function on(root: ParentNode, sel: string, fn: (el: HTMLElement, ev: Event) => void): void {
  root.querySelectorAll<HTMLElement>(sel).forEach((el) => el.addEventListener('click', (ev) => fn(el, ev)));
}

export function toast(msg: string): void {
  const root = ui();
  const t = document.createElement('div');
  t.className = 'toast';
  t.textContent = msg;
  root.appendChild(t);
  setTimeout(() => t.remove(), 1900);
}

/** Keep the DOM overlay exactly on top of the Phaser canvas. */
export function fitOverlay(canvas: HTMLCanvasElement): void {
  const stage = $('#stage');
  const r = canvas.getBoundingClientRect();
  const s = stage.getBoundingClientRect();
  const o = $('#overlay');
  o.style.left = `${r.left - s.left}px`;
  o.style.top = `${r.top - s.top}px`;
  o.style.width = `${r.width}px`;
  o.style.height = `${r.height}px`;
  document.documentElement.style.setProperty('--canvas-scale', String(r.width / GAME_W));
}

export const fmt = (n: number): string => Math.round(n).toLocaleString('en-US');
