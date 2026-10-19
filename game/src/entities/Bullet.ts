import Phaser from 'phaser';
import { laser } from '../config/balance';

export class Bullet extends Phaser.Physics.Arcade.Sprite {
  damage = 1;

  constructor(scene: Phaser.Scene, x: number, y: number) {
    super(scene, x, y, 'bullet');
  }

  fire(x: number, y: number, angleDeg: number, damage: number): this {
    this.damage = damage;
    this.setActive(true).setVisible(true);
    this.setPosition(x, y);
    this.setRotation(Phaser.Math.DegToRad(angleDeg));
    const body = this.body as Phaser.Physics.Arcade.Body;
    body.enable = true;
    body.setSize(12, 34);
    body.reset(x, y);
    const a = Phaser.Math.DegToRad(angleDeg - 90);
    body.setVelocity(Math.cos(a) * laser.bulletSpeed, Math.sin(a) * laser.bulletSpeed);
    return this;
  }

  preUpdate(t: number, d: number): void {
    super.preUpdate(t, d);
    if (this.y < -60 || this.x < -40 || this.x > 760) this.kill();
  }

  kill(): void {
    this.setActive(false).setVisible(false);
    (this.body as Phaser.Physics.Arcade.Body).enable = false;
  }
}
