// Always-running parallax space behind the DOM menus (not during play).
import Phaser from 'phaser';
import { GAME_H, GAME_W } from '../config/balance';

export class Background extends Phaser.Scene {
  private far!: Phaser.GameObjects.TileSprite;
  private near!: Phaser.GameObjects.TileSprite;
  private planets: Phaser.GameObjects.Image[] = [];

  constructor() {
    super('Background');
  }

  create(): void {
    this.cameras.main.setBackgroundColor('#101340');
    this.far = this.add.tileSprite(GAME_W / 2, GAME_H / 2, GAME_W, GAME_H, 'stars-far').setAlpha(0.7);
    this.near = this.add.tileSprite(GAME_W / 2, GAME_H / 2, GAME_W, GAME_H, 'stars-near');
    this.planets = [
      this.add.image(110, 260, 'planet-mars').setScale(0.9),
      this.add.image(GAME_W - 120, 900, 'planet-saturn').setScale(1.1),
      this.add.image(GAME_W - 90, 180, 'planet-moon').setScale(0.6).setAlpha(0.8),
    ];
    this.planets.forEach((p, i) => this.tweens.add({ targets: p, y: p.y - 24, duration: 3000 + i * 700, yoyo: true, repeat: -1, ease: 'Sine.easeInOut' }));
    for (let i = 0; i < 5; i++) {
      const s = this.add.image(Phaser.Math.Between(40, GAME_W - 40), Phaser.Math.Between(60, GAME_H - 60), 'star-6pt').setScale(Phaser.Math.FloatBetween(0.4, 0.9)).setAlpha(0.6).setTint(0xf876ff);
      this.tweens.add({ targets: s, alpha: 0.15, angle: 20, duration: 1800 + i * 300, yoyo: true, repeat: -1 });
    }
  }

  update(_t: number, d: number): void {
    this.far.tilePositionY -= 0.012 * d;
    this.near.tilePositionY -= 0.03 * d;
  }
}
