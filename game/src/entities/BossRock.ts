// Stage 1 boss: the Giant Rock. Three glowing cracks take damage; it fires telegraphed fragment rings.
import Phaser from 'phaser';
import { GAME_W, bossRock as B } from '../config/balance';
import type { PlayScene } from '../scenes/Play';

export class BossRock {
  rock: Phaser.Physics.Arcade.Image;
  cracks: Phaser.Physics.Arcade.Image[] = [];
  crackHp: number[] = [];
  alive = true;
  entered = false;
  private t = 0;
  private ringTimer = B.ringIntervalSeconds * 0.6;
  private warnRing: Phaser.GameObjects.Image;
  private warning = false;
  private dir = 1;
  private glow: Phaser.GameObjects.Image;

  constructor(private scene: PlayScene) {
    this.glow = scene.add.image(GAME_W / 2, -300, 'glow').setScale(7).setTint(0xfa500e).setAlpha(0.35).setDepth(9).setBlendMode(Phaser.BlendModes.ADD);
    this.rock = scene.physics.add.image(GAME_W / 2, -300, 'boss-rock').setDepth(10);
    const body = this.rock.body as Phaser.Physics.Arcade.Body;
    body.setCircle(148, 170 - 148, 170 - 148);
    body.setAllowGravity(false);
    body.setImmovable(true);
    const offsets = [
      { x: -70, y: -40 },
      { x: 60, y: -10 },
      { x: -10, y: 80 },
    ];
    offsets.forEach((o) => {
      const c = scene.physics.add.image(0, 0, 'crack').setDepth(11).setData('off', o);
      (c.body as Phaser.Physics.Arcade.Body).setCircle(30, 5, 5).setAllowGravity(false);
      this.cracks.push(c);
      this.crackHp.push(B.crackHp);
    });
    this.warnRing = scene.add.image(0, 0, 'ring').setDepth(12).setVisible(false).setBlendMode(Phaser.BlendModes.ADD);
    scene.tweens.add({ targets: this.rock, angle: 360, duration: 24000, repeat: -1 });
  }

  get hpFrac(): number {
    return this.crackHp.reduce((a, b) => a + b, 0) / (B.crackHp * B.cracks);
  }

  damageCrack(crack: Phaser.Physics.Arcade.Image, dmg: number): boolean {
    const i = this.cracks.indexOf(crack);
    if (i < 0 || this.crackHp[i] <= 0) return false;
    this.crackHp[i] -= dmg;
    crack.setScale(1.25);
    this.scene.tweens.add({ targets: crack, scale: 1, duration: 90 });
    if (this.crackHp[i] <= 0) {
      this.crackHp[i] = 0;
      crack.setVisible(false);
      (crack.body as Phaser.Physics.Arcade.Body).enable = false;
      this.scene.explode(crack.x, crack.y, 0xfa500e, 28);
      if (this.crackHp.every((h) => h <= 0)) this.defeat();
      return true;
    }
    return false;
  }

  private defeat(): void {
    this.alive = false;
    (this.rock.body as Phaser.Physics.Arcade.Body).enable = false;
    this.warnRing.setVisible(false);
    this.scene.tweens.add({ targets: this.rock, scale: 1.15, alpha: 0, duration: 700, ease: 'Quad.easeIn' });
    this.scene.tweens.add({ targets: this.glow, scale: 12, alpha: 0, duration: 800 });
    this.scene.time.addEvent({ delay: 90, repeat: 8, callback: () => this.scene.explode(this.rock.x + Phaser.Math.Between(-120, 120), this.rock.y + Phaser.Math.Between(-120, 120), 0x3a299b, 20) });
  }

  update(dt: number): void {
    if (!this.alive) return;
    this.t += dt;
    const targetY = 300;
    if (!this.entered) {
      this.rock.y = Phaser.Math.Linear(this.rock.y, targetY, 1 - Math.pow(0.97, dt * 60));
      if (Math.abs(this.rock.y - targetY) < 2) this.entered = true;
    } else {
      this.rock.x += this.dir * B.moveSpeed * dt;
      if (this.rock.x > GAME_W - 200) this.dir = -1;
      if (this.rock.x < 200) this.dir = 1;
      this.rock.y = targetY + Math.sin(this.t * 1.3) * 26;
      this.ringTimer -= dt;
      if (this.ringTimer <= B.warnSeconds && !this.warning) {
        this.warning = true;
        this.warnRing.setVisible(true).setScale(0.6).setAlpha(0.9);
        this.scene.tweens.add({ targets: this.warnRing, scale: 1.6, alpha: 0, duration: B.warnSeconds * 1000 });
        this.scene.sfx('alarm');
      }
      if (this.ringTimer <= 0) {
        this.fireRing();
        this.ringTimer = B.ringIntervalSeconds * (this.hpFrac < 0.4 ? 0.75 : 1);
        this.warning = false;
      }
    }
    (this.rock.body as Phaser.Physics.Arcade.Body).reset(this.rock.x, this.rock.y);
    this.glow.setPosition(this.rock.x, this.rock.y).setAlpha(0.25 + 0.15 * Math.sin(this.t * 4));
    this.warnRing.setPosition(this.rock.x, this.rock.y);
    const cos = Math.cos(this.rock.rotation);
    const sin = Math.sin(this.rock.rotation);
    this.cracks.forEach((c, i) => {
      const o = c.getData('off') as { x: number; y: number };
      const x = this.rock.x + o.x * cos - o.y * sin;
      const y = this.rock.y + o.x * sin + o.y * cos;
      c.setPosition(x, y).setRotation(this.rock.rotation);
      (c.body as Phaser.Physics.Arcade.Body).reset(x, y);
      if (this.crackHp[i] > 0) c.setAlpha(0.75 + 0.25 * Math.sin(this.t * 6 + i));
    });
  }

  private fireRing(): void {
    const n = B.ringFragments;
    const offset = Math.random() * Math.PI * 2;
    const gap = Phaser.Math.Between(0, n - 1); // one safe gap, always
    for (let i = 0; i < n; i++) {
      if (i === gap || i === (gap + 1) % n) continue;
      const a = offset + (i / n) * Math.PI * 2;
      const sp = B.fragmentSpeed * (this.hpFrac < 0.4 ? 1.15 : 1);
      this.scene.spawnHazard('fragment', this.rock.x + Math.cos(a) * 150, this.rock.y + Math.sin(a) * 150, Math.cos(a) * sp, Math.sin(a) * sp + 60);
    }
    this.scene.sfx('crack');
  }

  destroy(): void {
    this.rock.destroy();
    this.cracks.forEach((c) => c.destroy());
    this.warnRing.destroy();
    this.glow.destroy();
  }
}
