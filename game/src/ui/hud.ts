// In-game HUD (DOM). Hearts + combo top-right, coins + score top-left, turbo bottom.
import { $, fmt } from './dom';
import { images } from '../assets';
import { S } from '../config/strings.ar';
import { player as P } from '../config/balance';
import { save } from '../systems/save';

const heartSvg = `<svg class="heart" viewBox="0 0 24 24"><path fill="#00DF6A" d="M12 21s-7.5-4.7-9.6-9.2C.9 8.3 3 4.5 6.7 4.5c2 0 3.4 1.1 4.3 2.4.9-1.3 2.3-2.4 4.3-2.4 3.7 0 5.8 3.8 4.3 7.3C19.5 16.3 12 21 12 21z"/></svg>`;

export class Hud {
  root: HTMLElement;
  private hearts: HTMLElement[] = [];
  private comboEl!: HTMLElement;
  private comboBar!: HTMLElement;
  private scoreEl!: HTMLElement;
  private coinsEl!: HTMLElement;
  private turboBtn!: HTMLButtonElement;
  private turboFill!: HTMLElement;
  private gravityEl!: HTMLElement;
  private bossBar!: HTMLElement;
  private bossFill!: HTMLElement;
  private bossLbl!: HTMLElement;
  private vignette!: HTMLElement;
  private lastHearts = P.hearts;

  constructor(onPause: () => void, onTurbo: () => void) {
    this.root = $('#hud');
    this.root.innerHTML = `
      <div class="vignette"></div>
      <div class="hud-top">
        <div class="hud-right">
          <div class="hearts">${heartSvg.repeat(P.hearts)}</div>
          <div class="combo"></div>
          <div class="combo-bar"><i></i></div>
        </div>
        <div class="hud-left">
          <div class="score">0</div>
          <div class="coins"><img src="${images['icon-coin']}" alt=""><span>0</span></div>
        </div>
      </div>
      <button id="pauseBtn" aria-label="${S.paused}"></button>
      <div class="boss-bar" hidden><div class="lbl"></div><div class="bar"><i></i></div></div>
      <div class="hud-bottom">
        <button id="turboBtn" aria-label="${S.hud.turbo}"><div class="fill"></div><img src="${images['icon-rocket']}" alt=""><span>${S.hud.turbo}</span></button>
        <div class="gravity" hidden><img src="${images['icon-arrow']}" alt=""><span></span></div>
      </div>
      ${matchMedia('(pointer: fine)').matches ? `<div class="keys-hint">${S.tutorialKeys}</div>` : ''}`;
    this.hearts = Array.from(this.root.querySelectorAll<HTMLElement>('.heart'));
    this.comboEl = $('.combo', this.root);
    this.comboBar = $('.combo-bar i', this.root);
    this.scoreEl = $('.score', this.root);
    this.coinsEl = $('.coins span', this.root);
    this.turboBtn = $<HTMLButtonElement>('#turboBtn', this.root);
    this.turboFill = $('#turboBtn .fill', this.root);
    this.gravityEl = $('.gravity', this.root);
    this.bossBar = $('.boss-bar', this.root);
    this.bossFill = $('.boss-bar i', this.root);
    this.bossLbl = $('.boss-bar .lbl', this.root);
    this.vignette = $('.vignette', this.root);
    $('#pauseBtn', this.root).addEventListener('click', onPause);
    this.turboBtn.addEventListener('pointerdown', (e) => {
      e.preventDefault();
      e.stopPropagation();
      onTurbo();
    });
    this.root.hidden = false;
  }

  setHearts(n: number): void {
    this.hearts.forEach((h, i) => {
      h.classList.toggle('off', i >= n);
      if (i === n && n < this.lastHearts) {
        h.classList.add('pop');
        setTimeout(() => h.classList.remove('pop'), 320);
      }
    });
    this.lastHearts = n;
  }

  setScore(n: number): void {
    this.scoreEl.textContent = fmt(n);
  }

  setCoins(n: number, pop = false): void {
    this.coinsEl.textContent = fmt(n);
    if (pop) {
      const c = this.coinsEl.parentElement!;
      c.classList.remove('pop');
      void c.offsetWidth;
      c.classList.add('pop');
    }
  }

  setCombo(mult: number, progress: number, bump = false): void {
    this.comboEl.textContent = mult > 1 ? `${S.hud.combo} ×${mult}` : '';
    this.comboEl.classList.toggle('x5', mult >= 5);
    this.comboBar.style.width = `${Math.round(progress * 100)}%`;
    if (bump) {
      this.comboEl.classList.remove('bump');
      void this.comboEl.offsetWidth;
      this.comboEl.classList.add('bump');
    }
  }

  setTurbo(charge: number, active: boolean): void {
    this.turboFill.style.height = `${Math.round(charge * 100)}%`;
    this.turboBtn.classList.toggle('ready', charge >= 1 && !active);
    this.turboBtn.classList.toggle('active', active);
  }

  setGravity(dir: -1 | 0 | 1, warn: boolean): void {
    this.gravityEl.hidden = dir === 0;
    this.gravityEl.classList.toggle('left', dir === -1);
    this.gravityEl.classList.toggle('warn', warn);
    $('span', this.gravityEl).textContent = dir === 1 ? 'يمين' : 'يسار';
  }

  showBoss(name: string): void {
    this.bossLbl.textContent = name;
    this.bossBar.hidden = false;
    this.bossFill.style.width = '100%';
  }

  setBoss(frac: number): void {
    this.bossFill.style.width = `${Math.max(0, Math.round(frac * 100))}%`;
  }

  hideBoss(): void {
    this.bossBar.hidden = true;
  }

  banner(text: string, good = false): void {
    const b = document.createElement('div');
    b.className = 'banner' + (good ? ' good' : '');
    b.textContent = text;
    this.root.appendChild(b);
    setTimeout(() => b.remove(), 2300);
  }

  storm(on: boolean): void {
    if (save.data.settings.reduceMotion) return;
    this.vignette.classList.toggle('storm', on);
  }

  hitFlash(): void {
    if (save.data.settings.reduceMotion) return;
    this.vignette.classList.add('hit');
    setTimeout(() => this.vignette.classList.remove('hit'), 120);
  }

  destroy(): void {
    this.root.hidden = true;
    this.root.innerHTML = '';
  }
}
