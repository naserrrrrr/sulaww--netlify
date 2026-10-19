// Stage 1: Mars meteor shower. Three waves then the Giant Rock. Also hosts the endless challenge mode.
import Phaser from 'phaser';
import { GAME_W, marsBossAt, marsWaves } from '../config/balance';
import { S } from '../config/strings.ar';
import { PlayScene } from './Play';
import { BossRock } from '../entities/BossRock';
import { Bullet } from '../entities/Bullet';

export class StageMars extends PlayScene {
  readonly stageId = 'mars' as const;
  readonly bossAt = marsBossAt;
  readonly bossName = 'الصخرة العملاقة';
  readonly planetKey = 'planet-mars';
  private boss: BossRock | null = null;
  private stormShown = false;
  private tutorial = [
    { at: 3, msg: S.radio.moveHint },
    { at: 8, msg: S.radio.autoFire },
  ];
  private tutIndex = 0;

  constructor() {
    super('StageMars');
  }

  protected stageCreate(): void {
    this.boss = null;
    this.stormShown = false;
    this.tutIndex = 0;
    if (this.challenge) this.hud.banner(S.challengeUI.title, true);
  }

  protected segmentKey(): string {
    if (this.bossStarted) return 'boss';
    const i = marsWaves.findIndex((w) => this.elapsed < w.to);
    return `w${i < 0 ? marsWaves.length : i}`;
  }

  private wave() {
    if (this.challenge) {
      const t = this.elapsed;
      return { rate: Math.min(3.6, 0.9 + t * 0.022), speed: [220 + t * 2.2, 320 + t * 2.6], big: Math.min(0.35, t / 300), diagonal: Math.min(0.5, Math.max(0, (t - 40) / 200)) };
    }
    return marsWaves.find((w) => this.elapsed < w.to) ?? marsWaves[marsWaves.length - 1];
  }

  protected currentSpeed(): number {
    const w = this.wave();
    return (w.speed[0] + w.speed[1]) / 2;
  }

  protected stageUpdate(dt: number): void {
    const w = this.wave();
    if (this.tutIndex < this.tutorial.length && this.elapsed >= this.tutorial[this.tutIndex].at && !this.challenge) {
      this.radioSay(this.tutorial[this.tutIndex++].msg);
    }
    if (!this.challenge && !this.stormShown && this.elapsed >= marsWaves[2].from) {
      this.stormShown = true;
      this.hud.banner('عاصفة!');
      this.hud.storm(true);
      this.radioSay(S.radio.storm, true);
      this.sfx('alarm');
    }
    this.spawnAcc += w.rate * this.assistFactor * dt;
    while (this.spawnAcc >= 1) {
      this.spawnAcc -= 1;
      const vy = Phaser.Math.Between(w.speed[0], w.speed[1]);
      const big = Math.random() < w.big;
      if (Math.random() < w.diagonal) {
        const fromLeft = Math.random() < 0.5;
        this.spawnMeteor(fromLeft ? -60 : GAME_W + 60, Phaser.Math.Between(-60, 300), (fromLeft ? 1 : -1) * Phaser.Math.Between(180, 260), vy * 0.8, big);
      } else {
        this.spawnMeteor(Phaser.Math.Between(50, GAME_W - 50), big ? -90 : -50, Phaser.Math.Between(-40, 40), vy, big);
      }
    }
  }

  protected createBoss(): void {
    this.radioSay(S.radio.bossRock, true);
    this.boss = new BossRock(this);
    const b = this.boss;
    this.physics.add.overlap(this.bullets, b.cracks, (o1, o2) => {
      const bullet = (o1 instanceof Bullet ? o1 : o2) as Bullet;
      const cr = (bullet === o1 ? o2 : o1) as Phaser.Physics.Arcade.Image;
      if (!bullet.active) return;
      bullet.kill();
      this.burst.setParticleTint(0xffec00);
      this.burst.explode(4, bullet.x, bullet.y);
      this.sfx('bossHit');
      if (b.damageCrack(cr, bullet.damage)) {
        this.addScore(400, bullet.x, bullet.y);
        this.radioSay(b.alive ? S.radio.rockCrack : S.radio.rockDone, true);
      } else this.addScore(25);
    });
    this.physics.add.overlap(this.player, b.rock, () => {
      if (!b.alive) return;
      this.hitPlayer();
    });
  }

  protected bossUpdate(dt: number): void {
    this.boss?.update(dt);
  }

  protected bossAlive(): boolean {
    return this.boss ? this.boss.alive : true;
  }

  protected debugKillBoss(): void {
    this.boss?.cracks.forEach((c) => this.boss?.damageCrack(c, 999));
  }

  protected bossHpFrac(): number {
    return this.boss ? this.boss.hpFrac : 1;
  }
}
