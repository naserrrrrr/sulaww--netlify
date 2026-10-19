// Stage 2: Saturn's rings. Side gravity, ring gates, ice blocks, debris passages, a fork, then Khurda.
import Phaser from 'phaser';
import { GAME_W, saturn as SB, saturnBossAt, saturnSections } from '../config/balance';
import { S } from '../config/strings.ar';
import { PlayScene } from './Play';
import { BossRobot } from '../entities/BossRobot';
import { Bullet } from '../entities/Bullet';

export class StageSaturn extends PlayScene {
  readonly stageId = 'saturn' as const;
  readonly bossAt = saturnBossAt;
  readonly bossName = 'خُردة';
  readonly planetKey = 'planet-saturn';
  private boss: BossRobot | null = null;
  private gravityTimer = 4;
  private gravityWarned = false;
  private gateTimer = 2;
  private wallTimer = 8;
  private forkTold = false;
  private narrowTold = false;
  private gatesTold = false;

  constructor() {
    super('StageSaturn');
  }

  protected stageCreate(): void {
    this.boss = null;
    this.gravityDir = Math.random() < 0.5 ? -1 : 1;
    this.gravityDrift = SB.gravityAccel;
    this.gravityTimer = Phaser.Math.FloatBetween(SB.gravityFlipMin, SB.gravityFlipMax);
    this.gravityWarned = false;
    this.gateTimer = 2;
    this.wallTimer = 8;
    this.forkTold = this.narrowTold = this.gatesTold = false;
    this.hud.setGravity(this.gravityDir, false);
    this.time.delayedCall(2500, () => this.radioSay(this.gravityDir === 1 ? S.radio.gravityRight : S.radio.gravityLeft, true));
  }

  protected segmentKey(): string {
    if (this.bossStarted) return 'boss';
    return this.elapsed < saturnSections[0].to ? 's1' : 's2';
  }

  private section() {
    return saturnSections.find((s) => this.elapsed < s.to) ?? saturnSections[1];
  }

  protected currentSpeed(): number {
    const s = this.section();
    return (s.speed[0] + s.speed[1]) / 2;
  }

  private get forkActive(): boolean {
    return this.elapsed > 78 && this.elapsed < 96;
  }

  protected stageUpdate(dt: number): void {
    const s = this.section();
    // gravity flips with a warning
    this.gravityTimer -= dt;
    if (this.gravityTimer <= SB.gravityWarnSeconds && !this.gravityWarned) {
      this.gravityWarned = true;
      this.hud.setGravity(this.gravityDir, true);
    }
    if (this.gravityTimer <= 0) {
      this.gravityDir = this.gravityDir === 1 ? -1 : 1;
      this.gravityWarned = false;
      this.gravityTimer = Phaser.Math.FloatBetween(SB.gravityFlipMin, SB.gravityFlipMax);
      this.hud.setGravity(this.gravityDir, false);
      this.radioSay(this.gravityDir === 1 ? S.radio.gravityRight : S.radio.gravityLeft, true);
      this.sfx('click');
    }
    if (!this.gatesTold && this.elapsed > 4) {
      this.gatesTold = true;
      this.radioSay(S.radio.gates);
    }
    if (s.narrow && !this.narrowTold) {
      this.narrowTold = true;
      this.hud.banner(S.radio.narrow);
      this.radioSay(S.radio.narrow, true);
    }
    if (this.forkActive && !this.forkTold) {
      this.forkTold = true;
      this.hud.banner(S.radio.forkSafe, true);
      this.radioSay(S.radio.forkSafe, true);
    }

    // meteors
    this.spawnAcc += s.rate * this.assistFactor * dt;
    while (this.spawnAcc >= 1) {
      this.spawnAcc -= 1;
      const vy = Phaser.Math.Between(s.speed[0], s.speed[1]);
      if (Math.random() < s.ice) {
        const fromLeft = Math.random() < 0.5;
        this.spawnHazard('ice', fromLeft ? 70 : GAME_W - 70, -40, (fromLeft ? 1 : -1) * SB.iceSpeed, vy * 0.35);
      } else {
        const xMin = this.forkActive ? GAME_W / 2 + 40 : 50;
        this.spawnMeteor(Phaser.Math.Between(xMin, GAME_W - 50), -50, Phaser.Math.Between(-30, 30), vy, Math.random() < 0.18);
      }
    }
    if (this.forkActive && Math.random() < dt * 0.9) {
      this.spawnPickup('coin', Phaser.Math.Between(GAME_W / 2 + 60, GAME_W - 60), -40, vyOf(s) * 0.8);
    }
    // gates
    this.gateTimer -= dt;
    if (this.gateTimer <= 0 && s.gates) {
      this.gateTimer = Phaser.Math.FloatBetween(4, 5.5);
      this.spawnGate(Phaser.Math.Between(180, GAME_W - 180), -60, vyOf(s) * 0.8, s.narrow ? 130 : 0);
    }
    // debris walls with one gap
    if (s.narrow) {
      this.wallTimer -= dt;
      if (this.wallTimer <= 0) {
        this.wallTimer = Phaser.Math.FloatBetween(5.5, 7);
        const gapX = Phaser.Math.Between(150, GAME_W - 150);
        for (let x = 55; x < GAME_W; x += 110) {
          if (Math.abs(x - gapX) < 150) continue;
          this.spawnMeteor(x, -70, 0, vyOf(s) * 0.75, false, 3, true);
        }
        this.spawnPickup('coin', gapX, -70, vyOf(s) * 0.75);
      }
    }
  }

  protected createBoss(): void {
    this.radioSay(S.radio.robot, true);
    this.hud.setGravity(0, false);
    this.gravityDir = 0;
    this.boss = new BossRobot(this);
    const b = this.boss;
    this.physics.add.overlap(this.bullets, b.cores, (o1, o2) => {
      const bullet = (o1 instanceof Bullet ? o1 : o2) as Bullet;
      const core = (bullet === o1 ? o2 : o1) as Phaser.Physics.Arcade.Image;
      if (!bullet.active || !b.exposed) return;
      bullet.kill();
      this.burst.setParticleTint(0xf876ff);
      this.burst.explode(4, bullet.x, bullet.y);
      this.sfx('bossHit');
      if (b.damageCore(core, bullet.damage)) {
        this.addScore(500, bullet.x, bullet.y);
        this.radioSay(b.alive ? 'نواة واحدة بقيت!' : S.radio.robotFixed, true);
      } else this.addScore(25);
    });
    this.physics.add.overlap(this.bullets, b.robot, (o1, o2) => {
      const bullet = (o1 instanceof Bullet ? o1 : o2) as Bullet;
      if (!bullet.active) return;
      bullet.kill();
      this.burst.setParticleTint(0xffec00);
      this.burst.explode(2, bullet.x, bullet.y);
    });
    this.physics.add.overlap(this.player, b.robot, () => {
      if (b.alive) this.hitPlayer();
    });
    this.physics.add.overlap(this.player, b['beam'], () => {
      if (b.alive) this.hitPlayer();
    });
  }

  protected bossUpdate(dt: number): void {
    this.boss?.update(dt);
  }

  protected bossAlive(): boolean {
    return this.boss ? this.boss.alive : true;
  }

  protected debugKillBoss(): void {
    if (!this.boss) return;
    this.boss.setExposed(true, 5);
    this.boss.cores.forEach((c) => this.boss?.damageCore(c, 999));
  }

  protected bossHpFrac(): number {
    return this.boss ? this.boss.hpFrac : 1;
  }
}

const vyOf = (s: { speed: readonly [number, number] | readonly number[] }) => (s.speed[0] + s.speed[1]) / 2;
