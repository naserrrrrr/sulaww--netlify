import Phaser from 'phaser';
import { GAME_H, GAME_W, player as P, turbo as T } from '../config/balance';
import { RocketSkin } from '../config/palette';
import { ROCKET_FRAMES } from '../assets';

export interface PlayerInput {
  keys: { left: boolean; right: boolean; up: boolean; down: boolean };
}

export class Player extends Phaser.Physics.Arcade.Sprite {
  target = new Phaser.Math.Vector2();
  invulnUntil = 0;
  shieldUntil = 0;
  tripleUntil = 0;
  magnetUntil = 0;
  turboCharge = 0;
  turboUntil = 0;
  gravityOffset = 0;
  dead = false;
  private dragStart: { px: number; py: number; x: number; y: number } | null = null;
  private prevX = 0;
  private shieldSprite: Phaser.GameObjects.Image;
  private flame: Phaser.GameObjects.Particles.ParticleEmitter;
  private speedLines: Phaser.GameObjects.Particles.ParticleEmitter;

  constructor(scene: Phaser.Scene, skin: RocketSkin, lowQuality: boolean) {
    super(scene, GAME_W / 2, GAME_H - 260, `rocket-${skin}`, 0);
    scene.add.existing(this);
    scene.physics.add.existing(this);
    const key = `rocket-${skin}-fly`;
    if (!scene.anims.exists(key)) {
      scene.anims.create({ key, frames: scene.anims.generateFrameNumbers(`rocket-${skin}`, { start: 0, end: ROCKET_FRAMES - 1 }), frameRate: 24, repeat: -1 });
    }
    this.play(key);
    this.setDepth(20);
    this.target.set(this.x, this.y);
    this.prevX = this.x;
    const body = this.body as Phaser.Physics.Arcade.Body;
    body.setCircle(P.hitRadius, this.width / 2 - P.hitRadius, this.height / 2 - P.hitRadius + 6);
    body.setAllowGravity(false);

    this.shieldSprite = scene.add.image(this.x, this.y, 'shield').setDepth(21).setVisible(false).setBlendMode(Phaser.BlendModes.ADD);
    this.flame = scene.add.particles(0, 0, 'particle', {
      follow: this,
      followOffset: { x: 0, y: 58 },
      speedY: { min: 260, max: 420 },
      speedX: { min: -40, max: 40 },
      scale: { start: 1.1, end: 0 },
      alpha: { start: 0.9, end: 0 },
      lifespan: { min: 180, max: 320 },
      quantity: lowQuality ? 1 : 2,
      frequency: lowQuality ? 40 : 22,
      tint: [0xffec00, 0xfa500e, 0xfff4a6],
      blendMode: Phaser.BlendModes.ADD,
    }).setDepth(19);
    this.speedLines = scene.add.particles(0, 0, 'speedline', {
      x: { min: 20, max: GAME_W - 20 },
      y: -60,
      speedY: { min: 1500, max: 2200 },
      alpha: { start: 0.5, end: 0 },
      lifespan: 700,
      quantity: 2,
      frequency: 30,
      emitting: false,
    }).setDepth(5);
  }

  /** Relative drag: the rocket moves by the same delta as the finger. */
  pointerDown(p: Phaser.Input.Pointer): void {
    this.dragStart = { px: p.worldX, py: p.worldY, x: this.target.x, y: this.target.y };
    this.gravityOffset = 0;
  }

  pointerMove(p: Phaser.Input.Pointer): void {
    if (!this.dragStart || !p.isDown) return;
    this.target.x = this.dragStart.x + (p.worldX - this.dragStart.px) + this.gravityOffset;
    this.target.y = this.dragStart.y + (p.worldY - this.dragStart.py);
  }

  pointerUp(): void {
    this.dragStart = null;
  }

  get turboActive(): boolean {
    return this.scene.time.now < this.turboUntil;
  }
  get shielded(): boolean {
    return this.scene.time.now < this.shieldUntil;
  }
  get invulnerable(): boolean {
    return this.scene.time.now < this.invulnUntil || this.shielded || this.turboActive;
  }

  tryTurbo(): boolean {
    if (this.turboCharge < 1 || this.turboActive || this.dead) return false;
    this.turboCharge = 0;
    this.turboUntil = this.scene.time.now + T.durationSeconds * 1000;
    this.speedLines.start();
    this.scene.time.delayedCall(T.durationSeconds * 1000, () => this.speedLines.stop());
    this.scene.tweens.add({ targets: this, scaleX: 0.8, scaleY: 1.25, duration: 120, yoyo: true, ease: 'Quad.easeOut' });
    return true;
  }

  hitSquash(): void {
    this.scene.tweens.add({ targets: this, scaleX: 1.3, scaleY: 0.75, duration: 90, yoyo: true, ease: 'Quad.easeOut' });
  }

  update(dt: number, input: PlayerInput, gravityDir: number, gravityDrift: number): void {
    if (this.dead) return;
    const sp = P.maxSpeed * dt * (this.turboActive ? T.speedBoost : 1);
    const k = input.keys;
    if (k.left || k.right || k.up || k.down) {
      this.target.x += ((k.right ? 1 : 0) - (k.left ? 1 : 0)) * sp;
      this.target.y += ((k.down ? 1 : 0) - (k.up ? 1 : 0)) * sp;
    }
    if (gravityDir !== 0) {
      const g = gravityDir * gravityDrift * dt;
      this.target.x += g;
      this.gravityOffset += g;
    }
    this.target.x = Phaser.Math.Clamp(this.target.x, P.marginX, GAME_W - P.marginX);
    this.target.y = Phaser.Math.Clamp(this.target.y, P.marginTop, GAME_H - P.marginBottom);

    const f = 1 - Math.pow(1 - P.smoothing, dt * 60);
    let nx = this.x + (this.target.x - this.x) * f;
    let ny = this.y + (this.target.y - this.y) * f;
    const maxStep = P.maxSpeed * dt * (this.turboActive ? T.speedBoost : 1);
    const dx = nx - this.x;
    const dy = ny - this.y;
    const dist = Math.hypot(dx, dy);
    if (dist > maxStep) {
      nx = this.x + (dx / dist) * maxStep;
      ny = this.y + (dy / dist) * maxStep;
    }
    this.setPosition(nx, ny);
    (this.body as Phaser.Physics.Arcade.Body).reset(nx, ny);

    const vx = (nx - this.prevX) / Math.max(dt, 1e-4);
    this.prevX = nx;
    this.setRotation(Phaser.Math.Linear(this.rotation, Phaser.Math.Clamp(vx / P.maxSpeed, -1, 1) * 0.28, 0.2));

    // invulnerability blink (<= 3 Hz)
    const now = this.scene.time.now;
    if (now < this.invulnUntil && !this.shielded && !this.turboActive) {
      this.setAlpha(Math.floor(now / (1000 / P.blinkHz / 2)) % 2 === 0 ? 1 : 0.35);
    } else this.setAlpha(1);

    this.shieldSprite.setVisible(this.shielded).setPosition(nx, ny).setRotation(now * 0.001);
    this.shieldSprite.setAlpha(this.shieldUntil - now < 1500 ? 0.4 + 0.6 * Math.abs(Math.sin(now * 0.012)) : 1);
    this.flame.setParticleScale(this.turboActive ? 1.8 : 1);
  }

  die(): void {
    this.dead = true;
    this.flame.stop();
    this.speedLines.stop();
    this.shieldSprite.setVisible(false);
    (this.body as Phaser.Physics.Arcade.Body).enable = false;
  }

  destroy(fromScene?: boolean): void {
    this.flame.destroy();
    this.speedLines.destroy();
    this.shieldSprite.destroy();
    super.destroy(fromScene);
  }
}
