import Phaser from 'phaser';
import { GAME_H, GAME_W } from '../config/balance';

export class Meteor extends Phaser.Physics.Arcade.Sprite {
  big = false;
  wall = false;
  hp = 1;
  radius = 30;
  spin = 0;
  passedPlayer = false;

  constructor(scene: Phaser.Scene, x: number, y: number) {
    super(scene, x, y, 'meteor-s0');
  }

  spawn(x: number, y: number, vx: number, vy: number, big: boolean, hp = big ? 3 : 1, wall = false): this {
    const key = wall ? 'wall' : big ? `meteor-b${Phaser.Math.Between(0, 1)}` : `meteor-s${Phaser.Math.Between(0, 2)}`;
    this.setTexture(key);
    this.big = big;
    this.wall = wall;
    this.hp = hp;
    this.radius = wall ? 42 : big ? 62 : 28;
    this.spin = wall ? 0 : Phaser.Math.FloatBetween(-90, 90);
    this.passedPlayer = false;
    this.setActive(true).setVisible(true);
    this.setPosition(x, y);
    this.setRotation(Phaser.Math.FloatBetween(0, Math.PI * 2));
    this.setScale(1);
    this.setAlpha(1);
    const body = this.body as Phaser.Physics.Arcade.Body;
    body.enable = true;
    body.setCircle(this.radius, this.width / 2 - this.radius, this.height / 2 - this.radius);
    body.reset(x, y);
    body.setVelocity(vx, vy);
    body.setAngularVelocity(this.spin);
    return this;
  }

  preUpdate(t: number, d: number): void {
    super.preUpdate(t, d);
    if (this.y > GAME_H + 140 || this.x < -160 || this.x > GAME_W + 160) this.kill();
  }

  kill(): void {
    this.setActive(false).setVisible(false);
    (this.body as Phaser.Physics.Arcade.Body).enable = false;
  }
}
