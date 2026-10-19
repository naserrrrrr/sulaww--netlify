// Procedural sound effects (Web Audio API). No audio files, no music.
import { save } from './save';

export type Sfx =
  | 'laser'
  | 'coin'
  | 'diamond'
  | 'hit'
  | 'crash'
  | 'power'
  | 'nearMiss'
  | 'gate'
  | 'alarm'
  | 'bossHit'
  | 'click'
  | 'victory'
  | 'turbo'
  | 'star'
  | 'crack';

class AudioSystem {
  private ctx: AudioContext | null = null;
  private master: GainNode | null = null;
  private noiseBuf: AudioBuffer | null = null;
  private lastLaser = 0;

  /** Must be called from a user gesture; safe to call repeatedly. */
  unlock(): void {
    if (this.ctx) {
      if (this.ctx.state === 'suspended') void this.ctx.resume();
      return;
    }
    const Ctx = window.AudioContext || (window as unknown as { webkitAudioContext: typeof AudioContext }).webkitAudioContext;
    if (!Ctx) return;
    this.ctx = new Ctx();
    this.master = this.ctx.createGain();
    this.master.connect(this.ctx.destination);
    this.applySettings();
    const len = this.ctx.sampleRate;
    this.noiseBuf = this.ctx.createBuffer(1, len, this.ctx.sampleRate);
    const d = this.noiseBuf.getChannelData(0);
    for (let i = 0; i < len; i++) d[i] = Math.random() * 2 - 1;
  }

  applySettings(): void {
    if (!this.master) return;
    const s = save.data.settings;
    this.master.gain.value = s.sound ? s.volume * 0.6 : 0;
  }

  play(name: Sfx): void {
    if (!this.ctx || !this.master || !save.data.settings.sound) return;
    const t = this.ctx.currentTime;
    switch (name) {
      case 'laser':
        if (t - this.lastLaser < 0.05) return;
        this.lastLaser = t;
        this.tone('square', 880, 220, 0.08, 0.12, t);
        break;
      case 'coin':
        this.tone('sine', 1320, 1320, 0.07, 0.18, t);
        this.tone('sine', 1760, 1760, 0.12, 0.14, t + 0.06);
        break;
      case 'diamond':
        [1046, 1318, 1568, 2093].forEach((f, i) => this.tone('triangle', f, f, 0.12, 0.16, t + i * 0.06));
        break;
      case 'star':
        [1568, 1976, 2349, 2637].forEach((f, i) => this.tone('sine', f, f * 1.01, 0.2, 0.14, t + i * 0.07));
        break;
      case 'hit':
        this.noise(0.25, 0.5, t, 900);
        this.tone('sawtooth', 200, 60, 0.3, 0.3, t);
        break;
      case 'crash':
        this.noise(0.18, 0.35, t, 2500);
        this.tone('triangle', 300, 90, 0.15, 0.2, t);
        break;
      case 'crack':
        this.noise(0.12, 0.3, t, 4000);
        this.tone('square', 500, 180, 0.12, 0.15, t);
        break;
      case 'power':
        [523, 659, 784, 1046].forEach((f, i) => this.tone('triangle', f, f, 0.15, 0.2, t + i * 0.07));
        break;
      case 'nearMiss':
        this.tone('sine', 600, 1400, 0.12, 0.18, t);
        break;
      case 'gate':
        this.tone('triangle', 700, 1400, 0.18, 0.2, t);
        this.tone('sine', 1400, 2100, 0.18, 0.1, t + 0.05);
        break;
      case 'alarm':
        for (let i = 0; i < 3; i++) this.tone('square', 520, 520, 0.12, 0.12, t + i * 0.18);
        break;
      case 'bossHit':
        this.tone('sawtooth', 160, 70, 0.22, 0.28, t);
        this.noise(0.1, 0.25, t, 1500);
        break;
      case 'click':
        this.tone('sine', 900, 700, 0.05, 0.1, t);
        break;
      case 'turbo':
        this.tone('sawtooth', 200, 900, 0.5, 0.18, t);
        this.noise(0.4, 0.15, t, 3000);
        break;
      case 'victory':
        [523, 659, 784, 1046, 1318, 1568].forEach((f, i) => this.tone('triangle', f, f, 0.35, 0.2, t + i * 0.11));
        break;
    }
  }

  private tone(type: OscillatorType, f0: number, f1: number, dur: number, vol: number, at: number): void {
    const ctx = this.ctx!;
    const o = ctx.createOscillator();
    const g = ctx.createGain();
    o.type = type;
    o.frequency.setValueAtTime(f0, at);
    o.frequency.exponentialRampToValueAtTime(Math.max(20, f1), at + dur);
    g.gain.setValueAtTime(0.0001, at);
    g.gain.exponentialRampToValueAtTime(vol, at + 0.01);
    g.gain.exponentialRampToValueAtTime(0.0001, at + dur);
    o.connect(g).connect(this.master!);
    o.start(at);
    o.stop(at + dur + 0.02);
  }

  private noise(dur: number, vol: number, at: number, cutoff: number): void {
    const ctx = this.ctx!;
    const src = ctx.createBufferSource();
    src.buffer = this.noiseBuf;
    const f = ctx.createBiquadFilter();
    f.type = 'lowpass';
    f.frequency.value = cutoff;
    const g = ctx.createGain();
    g.gain.setValueAtTime(vol, at);
    g.gain.exponentialRampToValueAtTime(0.0001, at + dur);
    src.connect(f).connect(g).connect(this.master!);
    src.start(at);
    src.stop(at + dur);
  }
}

export const sfx = new AudioSystem();

export function vibrate(pattern: number | number[]): void {
  const s = save.data.settings;
  if (!s.vibrate || s.reduceMotion) return;
  try {
    navigator.vibrate?.(pattern);
  } catch {
    /* unsupported */
  }
}
