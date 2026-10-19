import Phaser from 'phaser';
import { GAME_H, GAME_W } from '../config/balance';

export class Gate extends Phaser.Physics.Arcade.Sprite {
  passed = false;
  sway = 0;
  swayAmp = 0;
  baseX = 0;

  constructor(scene: Phaser.Scene, x: number, y: number) {
    super(scene, x, y, 'gate');
  }

  spawn(x: number, y: number, vy: number, swayAmp = 0): this {
    this.passed = false;
    this.baseX = x;
    this.swayAmp = swayAmp;
    this.sway = Math.random() * Math.PI * 2;
    this.setActive(true).setVisible(true);
    this.setPosition(x, y);
    this.setAlpha(1);
    const body = this.body as Phaser.Physics.Arcade.Body;
    body.enable = true;
    body.setSize(140, 60, true); // only the gap counts
    body.reset(x, y);
    body.setVelocity(0, vy);
    return this;
  }

  preUpdate(t: number, d: number): void {
    super.preUpdate(t, d);
    if (this.swayAmp > 0) {
      this.sway += d * 0.0022;
      const nx = Phaser.Math.Clamp(this.baseX + Math.sin(this.sway) * this.swayAmp, 170, GAME_W - 170);
      (this.body as Phaser.Physics.Arcade.Body).velocity.x = ((nx - this.x) * 1000) / Math.max(1, d);
    }
    if (this.y > GAME_H + 80) this.kill();
  }

  kill(): void {
    this.setActive(false).setVisible(false);
    (this.body as Phaser.Physics.Arcade.Body).enable = false;
  }
}
