// Bayan's radio message box: avatar + typewriter text, auto-hides after 2.5 s.
import { $ } from './dom';
import { images } from '../assets';

export class Radio {
  private root = $('#radio');
  private txt: HTMLElement;
  private timer = 0;
  private typeTimer = 0;
  private queue: string[] = [];
  private lastMsg = '';

  constructor() {
    this.root.innerHTML = `<img src="${images['avatar-bayan']}" alt=""><div class="txt"><span class="name">بيان</span><span class="body"></span></div>`;
    this.txt = $('.body', this.root);
    this.root.hidden = false;
  }

  say(msg: string, force = false): void {
    if (!force && msg === this.lastMsg && this.root.classList.contains('show')) return;
    if (this.root.classList.contains('show') && !force) {
      if (this.queue.length < 2) this.queue.push(msg);
      return;
    }
    this.lastMsg = msg;
    clearTimeout(this.timer);
    clearInterval(this.typeTimer);
    this.root.classList.add('show');
    let i = 0;
    this.txt.textContent = '';
    this.typeTimer = window.setInterval(() => {
      i++;
      this.txt.textContent = msg.slice(0, i);
      if (i >= msg.length) clearInterval(this.typeTimer);
    }, 28);
    this.timer = window.setTimeout(() => {
      this.root.classList.remove('show');
      const next = this.queue.shift();
      if (next) setTimeout(() => this.say(next, true), 250);
    }, 2500);
  }

  destroy(): void {
    clearTimeout(this.timer);
    clearInterval(this.typeTimer);
    this.root.classList.remove('show');
    this.root.hidden = true;
  }
}
