// Final boss: "Khurda", the broken cleaning robot. Three phases, every attack telegraphed. Repaired, never destroyed.
import Phaser from 'phaser';
import { GAME_H, GAME_W, bossRobot as B } from '../config/balance';
import type { PlayScene } from '../scenes/Play';

type Attack = 'beam' | 'junk' | 'chasers' | 'edges';

export class BossRobot {
  robot: Phaser.Physics.Arcade.Image;
  cores: Phaser.Physics.Arcade.Image[] = [];
  coreHp: number[] = [];
  alive = true;
  entered = false;
  phase = 1;
  private t = 0;
  private attackTimer = 2.2;
  private exposedUntil = 0;
  private busy = false;
  private dir = 1;
  private sparks: Phaser.GameObjects.Particles.ParticleEmitter;
  private smoke: Phaser.GameObjects.Particles.ParticleEmitter;
  private warnLine: Phaser.GameObjects.Image;
  private beam: Phaser.Physics.Arcade.Image;
  private edgeWarn: Phaser.GameObjects.Rectangle[];
  private attackIndex = 0;
  private speedMul = 1;

  constructor(private scene: PlayScene) {
    this.robot = scene.physics.add.image(GAME_W / 2, -360, 'robot-yellow').setDepth(10).setScale(0.62);
    const body = this.robot.body as Phaser.Physics.Arcade.Body;
    body.setSize(230, 440, true).setAllowGravity(false).setImmovable(true);
    scene.tweens.add({ targets: this.robot, angle: { from: -3, to: 3 }, duration: 380, yoyo: true, repeat: -1, ease: 'Sine.easeInOut' });
    this.sparks = scene.add.particles(0, 0, 'spark', {
      follow: this.robot,
      followOffset: { x: 0, y: -110 },
      speed: { min: 120, max: 360 },
      angle: { min: 200, max: 340 },
      scale: { start: 1.4, end: 0 },
      lifespan: 400,
      quantity: 2,
      frequency: 90,
      tint: [0xffec00, 0xfa500e, 0xffffff],
      gravityY: 500,
    }).setDepth(11);
    this.smoke = scene.add.particles(0, 0, 'particle', {
      follow: this.robot,
      followOffset: { x: 30, y: -140 },
      speedY: { min: -80, max: -40 },
      speedX: { min: -20, max: 20 },
      scale: { start: 1.5, end: 3.5 },
      alpha: { start: 0.35, end: 0 },
      lifespan: 1400,
      frequency: 120,
      tint: 0x3a299b,
    }).setDepth(9);
    for (let i = 0; i < B.cores; i++) {
      const c = scene.physics.add.image(0, 0, 'core').setDepth(12).setData('off', { x: i === 0 ? -52 : 52, y: 30 });
      (c.body as Phaser.Physics.Arcade.Body).setCircle(27, 3, 3).setAllowGravity(false);
      this.cores.push(c);
      this.coreHp.push(B.coreHp);
    }
    this.warnLine = scene.add.image(0, GAME_H / 2, 'warnline').setDepth(8).setVisible(false);
    this.beam = scene.physics.add.image(0, GAME_H / 2, 'beam').setDepth(8).setVisible(false).setBlendMode(Phaser.BlendModes.ADD);
    (this.beam.body as Phaser.Physics.Arcade.Body).setAllowGravity(false).setSize(44, 1400, true);
    (this.beam.body as Phaser.Physics.Arcade.Body).enable = false;
    this.edgeWarn = [0, GAME_W].map((x) => scene.add.rectangle(x, GAME_H / 2, 60, GAME_H, 0xfa500e, 0).setDepth(8));
    this.setExposed(false);
  }

  get hpFrac(): number {
    return this.coreHp.reduce((a, b) => a + b, 0) / (B.coreHp * B.cores);
  }
  get exposed(): boolean {
    return this.scene.time.now < this.exposedUntil;
  }

  setExposed(on: boolean, seconds = B.coreExposeSeconds): void {
    this.exposedUntil = on ? this.scene.time.now + seconds * 1000 : 0;
    this.cores.forEach((c, i) => {
      const live = this.coreHp[i] > 0;
      (c.body as Phaser.Physics.Arcade.Body).enable = on && live;
      c.setVisible(live).setAlpha(on ? 1 : 0.35).setScale(on ? 1.15 : 0.85);
    });
  }

  damageCore(core: Phaser.Physics.Arcade.Image, dmg: number): boolean {
    const i = this.cores.indexOf(core);
    if (i < 0 || this.coreHp[i] <= 0 || !this.exposed) return false;
    this.coreHp[i] -= dmg;
    this.scene.tweens.add({ targets: core, scale: 1.35, duration: 60, yoyo: true });
    this.robot.setTint(0xffffff);
    this.scene.time.delayedCall(60, () => this.robot.clearTint());
    if (this.hpFrac <= B.phase3At && this.phase < 3) this.enterPhase3();
    else if (this.hpFrac <= 0.65 && this.phase < 2) this.phase = 2;
    if (this.coreHp[i] <= 0) {
      this.coreHp[i] = 0;
      core.setVisible(false);
      (core.body as Phaser.Physics.Arcade.Body).enable = false;
      this.scene.explode(core.x, core.y, 0xf876ff, 30);
      if (this.coreHp.every((h) => h <= 0)) this.repair();
      return true;
    }
    return false;
  }

  private enterPhase3(): void {
    this.phase = 3;
    this.speedMul = B.phase3SpeedMul;
    this.scene.radioSay(this.scene.strings.radio.robotPhase3);
    this.scene.hud.banner('المرحلة الأخيرة!');
  }

  private repair(): void {
    this.alive = false;
    this.busy = true;
    this.hideBeam();
    this.sparks.stop();
    this.smoke.stop();
    this.scene.tweens.killTweensOf(this.robot);
    this.robot.setAngle(0);
    this.scene.tweens.add({ targets: this.robot, y: 360, scale: 0.68, duration: 900, ease: 'Sine.easeOut' });
    this.scene.time.addEvent({ delay: 120, repeat: 14, callback: () => this.scene.explode(this.robot.x + Phaser.Math.Between(-110, 110), this.robot.y + Phaser.Math.Between(-180, 160), 0x00df6a, 14, 'spark') });
    this.scene.tweens.add({ targets: this.robot, angle: { from: -6, to: 6 }, duration: 260, yoyo: true, repeat: 5, delay: 1000 });
  }

  update(dt: number): void {
    if (!this.alive) return;
    this.t += dt;
    const targetY = 330;
    if (!this.entered) {
      this.robot.y = Phaser.Math.Linear(this.robot.y, targetY, 1 - Math.pow(0.97, dt * 60));
      if (Math.abs(this.robot.y - targetY) < 2) {
        this.entered = true;
        this.setExposed(true, 2);
      }
    } else {
      if (!this.busy) {
        this.robot.x += this.dir * 120 * this.speedMul * dt;
        if (this.robot.x > GAME_W - 170) this.dir = -1;
        if (this.robot.x < 170) this.dir = 1;
      }
      this.attackTimer -= dt * this.speedMul;
      if (this.attackTimer <= 0 && !this.busy) this.nextAttack();
      if (!this.exposed && this.cores.some((c, i) => this.coreHp[i] > 0 && (c.body as Phaser.Physics.Arcade.Body).enable)) this.setExposed(false);
    }
    (this.robot.body as Phaser.Physics.Arcade.Body).reset(this.robot.x, this.robot.y);
    this.cores.forEach((c) => {
      const o = c.getData('off') as { x: number; y: number };
      c.setPosition(this.robot.x + o.x, this.robot.y + o.y);
      (c.body as Phaser.Physics.Arcade.Body).reset(c.x, c.y);
      if (this.exposed) c.setScale(1.05 + 0.1 * Math.sin(this.t * 10));
    });
  }

  private nextAttack(): void {
    const seq: Attack[] = this.phase === 1 ? ['beam', 'beam', 'junk'] : this.phase === 2 ? ['beam', 'junk', 'chasers'] : ['edges', 'beam', 'chasers', 'junk'];
    const atk = seq[this.attackIndex++ % seq.length];
    this.busy = true;
    const done = (delay: number) => {
      this.scene.time.delayedCall(delay, () => {
        this.busy = false;
        this.attackTimer = (this.phase === 3 ? 1.2 : 1.8) / this.speedMul;
        this.setExposed(true, this.phase === 3 ? B.coreExposeSeconds : B.coreExposeSeconds + 0.6);
      });
    };
    if (atk === 'beam') {
      const fromLeft = Math.random() < 0.5;
      const x0 = fromLeft ? 60 : GAME_W - 60;
      const x1 = fromLeft ? GAME_W - 60 : 60;
      this.warnLine.setVisible(true).setX(x0).setAlpha(1);
      this.scene.tweens.add({ targets: this.warnLine, alpha: 0.2, duration: 120, yoyo: true, repeat: 3 });
      this.scene.radioSay(this.scene.strings.radio.beam);
      this.scene.sfx('alarm');
      this.scene.time.delayedCall(B.beamWarnSeconds * 1000, () => {
        this.warnLine.setVisible(false);
        this.beam.setVisible(true).setX(x0);
        (this.beam.body as Phaser.Physics.Arcade.Body).enable = true;
        this.scene.tweens.add({
          targets: this.beam,
          x: x1,
          duration: (B.beamSweepSeconds / this.speedMul) * 1000,
          ease: 'Sine.easeInOut',
          onUpdate: () => (this.beam.body as Phaser.Physics.Arcade.Body).reset(this.beam.x, this.beam.y),
          onComplete: () => this.hideBeam(),
        });
      });
      done((B.beamWarnSeconds + B.beamSweepSeconds / this.speedMul) * 1000 + 100);
    } else if (atk === 'junk') {
      const rows = this.phase === 3 ? 4 : 3;
      for (let r = 0; r < rows; r++) {
        this.scene.time.delayedCall(r * 420, () => {
          const cx = Phaser.Math.Between(140, GAME_W - 140);
          for (let i = -2; i <= 2; i++) {
            if (i === 0 && r % 2 === 0) continue; // V formation with a gap in the middle
            this.scene.spawnHazard('junk', cx + i * 70, this.robot.y + 120 + Math.abs(i) * 30, i * 30, 300 * this.speedMul);
          }
        });
      }
      done(rows * 420 + 300);
    } else if (atk === 'chasers') {
      for (let i = 0; i < (this.phase === 3 ? 3 : 2); i++) {
        this.scene.time.delayedCall(i * 350, () => {
          const h = this.scene.spawnHazard('chaser', this.robot.x + (i ? 80 : -80), this.robot.y + 150, 0, 120);
          if (h) h.target = this.scene.player;
        });
      }
      done(1200);
    } else {
      // edges: both screen edges become hazardous after a glow warning
      this.edgeWarn.forEach((r) => this.scene.tweens.add({ targets: r, fillAlpha: 0.45, duration: 200, yoyo: true, repeat: 3 }));
      this.scene.sfx('alarm');
      this.scene.time.delayedCall(900, () => {
        for (let y = 100; y < GAME_H; y += 260) {
          this.scene.spawnHazard('edge', 16, y, 0, 320);
          const r = this.scene.spawnHazard('edge', GAME_W - 16, y, 0, 320);
          if (r) r.setFlipX(true);
        }
      });
      done(2600);
    }
  }

  private hideBeam(): void {
    this.beam.setVisible(false);
    (this.beam.body as Phaser.Physics.Arcade.Body).enable = false;
    this.warnLine.setVisible(false);
  }

  destroy(): void {
    this.robot.destroy();
    this.cores.forEach((c) => c.destroy());
    this.beam.destroy();
    this.warnLine.destroy();
    this.sparks.destroy();
    this.smoke.destroy();
    this.edgeWarn.forEach((r) => r.destroy());
  }
}
