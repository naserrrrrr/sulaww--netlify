import Phaser from 'phaser';
import { GAME_H } from '../config/balance';

export type PickupKind = 'coin' | 'diamond' | 'shield' | 'magnet' | 'triple' | 'star' | 'heart';

const TEX: Record<PickupKind, string> = {
  coin: 'icon-coin',
  diamond: 'icon-diamond',
  shield: 'pw-shield',
  magnet: 'pw-magnet',
  triple: 'pw-triple',
  star: 'gold-star',
  heart: 'heart',
};
const SIZE: Record<PickupKind, number> = { coin: 52, diamond: 60, shield: 76, magnet: 76, triple: 76, star: 72, heart: 56 };

export class Pickup extends Phaser.Physics.Arcade.Sprite {
  kind: PickupKind = 'coin';
  baseVy = 0;
  bob = 0;
  baseScale = 1;
  pulled = false;

  constructor(scene: Phaser.Scene, x: number, y: number) {
    super(scene, x, y, 'icon-coin');
  }

  spawn(kind: PickupKind, x: number, y: number, vy: number, vx = 0): this {
    this.kind = kind;
    this.setTexture(TEX[kind]);
    const s = SIZE[kind] / Math.max(this.width, this.height);
    this.baseScale = s;
    this.setScale(s);
    this.setActive(true).setVisible(true);
    this.setPosition(x, y);
    this.setAlpha(1);
    this.setRotation(0);
    this.baseVy = vy;
    this.bob = Math.random() * Math.PI * 2;
    this.pulled = false;
    const body = this.body as Phaser.Physics.Arcade.Body;
    body.enable = true;
    const r = SIZE[kind] / 2 / s;
    body.setCircle(r, this.width / 2 - r, this.height / 2 - r);
    body.reset(x, y);
    body.setVelocity(vx, vy);
    return this;
  }

  preUpdate(t: number, d: number): void {
    super.preUpdate(t, d);
    if (!this.pulled) {
      this.bob += d * 0.004;
      this.setScale(this.baseScale * (1 + Math.sin(this.bob) * 0.06));
      if (this.kind === 'star' || this.kind === 'diamond') this.rotation += d * 0.002;
    }
    if (this.y > GAME_H + 80) this.kill();
  }

  kill(): void {
    this.setActive(false).setVisible(false);
    (this.body as Phaser.Physics.Arcade.Body).enable = false;
  }
}
