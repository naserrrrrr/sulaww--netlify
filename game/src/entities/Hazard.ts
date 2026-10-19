// Shared pooled hazards: boss fragments, junk, ice blocks, chasers.
import Phaser from 'phaser';
import { GAME_H, GAME_W, bossRobot } from '../config/balance';

export type HazardKind = 'fragment' | 'junk' | 'ice' | 'chaser' | 'edge';

export class Hazard extends Phaser.Physics.Arcade.Sprite {
  kind: HazardKind = 'fragment';
  hp = 1;
  life = 0;
  target: Phaser.GameObjects.Components.Transform | null = null;
  radius = 14;

  constructor(scene: Phaser.Scene, x: number, y: number) {
    super(scene, x, y, 'fragment');
  }

  spawn(kind: HazardKind, x: number, y: number, vx: number, vy: number): this {
    this.kind = kind;
    this.setTexture(kind === 'edge' ? 'edge-hazard' : kind);
    this.hp = kind === 'chaser' ? bossRobot.chaserHp : 1;
    this.life = 0;
    this.target = null;
    this.setActive(true).setVisible(true);
    this.setPosition(x, y);
    this.setAlpha(1);
    this.setScale(1);
    this.setRotation(0);
    const body = this.body as Phaser.Physics.Arcade.Body;
    body.enable = true;
    if (kind === 'ice') {
      body.setSize(118, 44, true);
      body.setCollideWorldBounds(false);
    } else if (kind === 'edge') {
      body.setSize(30, 240, true);
      body.setCollideWorldBounds(false);
    } else {
      this.radius = kind === 'chaser' ? 22 : kind === 'junk' ? 16 : 13;
      body.setCircle(this.radius, this.width / 2 - this.radius, this.height / 2 - this.radius);
      body.setCollideWorldBounds(false);
    }
    body.reset(x, y);
    body.setVelocity(vx, vy);
    body.setAngularVelocity(kind === 'ice' || kind === 'edge' ? 0 : Phaser.Math.Between(-120, 120));
    return this;
  }

  preUpdate(t: number, d: number): void {
    super.preUpdate(t, d);
    this.life += d / 1000;
    const body = this.body as Phaser.Physics.Arcade.Body;
    if (this.kind === 'ice') {
      if (this.x < 66) body.velocity.x = Math.abs(body.velocity.x);
      else if (this.x > GAME_W - 66) body.velocity.x = -Math.abs(body.velocity.x);
    }
    if (this.kind === 'chaser' && this.target) {
      const a = Phaser.Math.Angle.Between(this.x, this.y, this.target.x, this.target.y);
      const sp = bossRobot.chaserSpeed;
      body.velocity.x = Phaser.Math.Linear(body.velocity.x, Math.cos(a) * sp, 0.04);
      body.velocity.y = Phaser.Math.Linear(body.velocity.y, Math.sin(a) * sp, 0.04);
      if (this.life > 9) this.kill();
    }
    if (this.y > GAME_H + 140 || this.y < -300 || this.x < -200 || this.x > GAME_W + 200) this.kill();
  }

  kill(): void {
    this.setActive(false).setVisible(false);
    (this.body as Phaser.Physics.Arcade.Body).enable = false;
  }
}
