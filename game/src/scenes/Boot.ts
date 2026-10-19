import Phaser from 'phaser';
import { images, rocketSheets, ROCKET_FRAME } from '../assets';
import { buildTextures } from '../systems/textures';
import { GAME_H, GAME_W } from '../config/balance';

export class Boot extends Phaser.Scene {
  constructor() {
    super('Boot');
  }

  preload(): void {
    const g = this.add.graphics();
    g.fillStyle(0x5143de, 1);
    const bar = (p: number) => {
      g.clear();
      g.fillStyle(0x1e1b6b, 1);
      g.fillRoundedRect(GAME_W / 2 - 160, GAME_H / 2 - 10, 320, 20, 10);
      g.fillStyle(0x8cd9ff, 1);
      g.fillRoundedRect(GAME_W / 2 - 160, GAME_H / 2 - 10, 320 * p, 20, 10);
    };
    this.load.on('progress', bar);
    (Object.keys(images) as (keyof typeof images)[]).forEach((k) => {
      if ((rocketSheets as readonly string[]).includes(k)) this.load.spritesheet(k, images[k], { frameWidth: ROCKET_FRAME, frameHeight: ROCKET_FRAME });
      else this.load.image(k, images[k]);
    });
  }

  create(): void {
    buildTextures(this);
    this.buildStarfields();
    const jump = this.registry.get('devJump') as { key: string; jumpToBoss: boolean; challenge: boolean } | undefined;
    if (jump) this.scene.start(jump.key, { jumpToBoss: jump.jumpToBoss, challenge: jump.challenge });
    else this.scene.start('Menu');
  }

  private buildStarfields(): void {
    const mk = (key: string, count: number, maxR: number, tints: number[]) => {
      if (this.textures.exists(key)) return;
      const g = this.make.graphics({ x: 0, y: 0 }, false);
      for (let i = 0; i < count; i++) {
        g.fillStyle(tints[i % tints.length], Phaser.Math.FloatBetween(0.5, 1));
        g.fillCircle(Phaser.Math.Between(0, 512), Phaser.Math.Between(0, 512), Phaser.Math.FloatBetween(1, maxR));
      }
      g.generateTexture(key, 512, 512);
      g.destroy();
    };
    mk('stars-far', 90, 1.6, [0xffffff, 0x8cd9ff, 0xd5f7f7]);
    mk('stars-near', 36, 2.6, [0xffffff, 0xf876ff, 0x8cd9ff, 0xffec00]);
    if (!this.textures.exists('cloud')) {
      const g = this.make.graphics({ x: 0, y: 0 }, false);
      g.fillStyle(0x271466, 1);
      g.fillCircle(90, 110, 70);
      g.fillCircle(160, 90, 80);
      g.fillCircle(240, 120, 65);
      g.fillRoundedRect(60, 110, 220, 70, 35);
      g.generateTexture('cloud', 320, 200);
      g.destroy();
    }
  }
}
