// Thin Phaser scenes that drive the DOM screens: Menu, Story, Map, Hangar.
import Phaser from 'phaser';
import { hideScreen, showScreen, on } from '../ui/dom';
import { menuScreen, storyScreen, mapScreen, bindHangar, openAchievements, openSettings, challengeScreen } from '../ui/screens';
import { save, StageId } from '../systems/save';
import { sfx } from '../systems/audio';

export function ensureBackground(scene: Phaser.Scene): void {
  if (!scene.scene.isActive('Background')) scene.scene.launch('Background');
  scene.scene.bringToTop(scene.scene.key);
}

export const stageScene = (id: StageId): string => (id === 'mars' ? 'StageMars' : 'StageSaturn');

export class Menu extends Phaser.Scene {
  constructor() {
    super('Menu');
  }

  create(): void {
    ensureBackground(this);
    document.getElementById('desk-logo')?.classList.add('show');
    this.render();
  }

  private render(): void {
    const root = showScreen(menuScreen(), { clear: true });
    on(root, '[data-act=play]', () => {
      sfx.unlock();
      hideScreen();
      this.scene.start(save.data.storySeen ? 'Map' : 'Story');
    });
    on(root, '[data-act=hangar]', () => {
      hideScreen();
      this.scene.start('Hangar', { back: 'Menu' });
    });
    on(root, '[data-act=ach]', () => openAchievements(() => this.render()));
    on(root, '[data-act=settings]', () => openSettings(() => this.render()));
    on(root, '[data-act=challenge]', () => {
      const r = showScreen(challengeScreen());
      on(r, '[data-act=start]', () => {
        hideScreen();
        this.scene.start('StageMars', { challenge: true });
      });
      on(r, '[data-act=back]', () => this.render());
    });
  }
}

export class Story extends Phaser.Scene {
  private index = 0;
  constructor() {
    super('Story');
  }

  create(): void {
    ensureBackground(this);
    this.index = 0;
    this.render();
  }

  private render(): void {
    const root = showScreen(storyScreen(this.index), { clear: true });
    const finish = () => {
      save.data.storySeen = true;
      save.write();
      hideScreen();
      this.scene.start('Map');
    };
    on(root, '[data-act=next]', () => {
      if (++this.index >= 3) finish();
      else this.render();
    });
    on(root, '[data-act=skip]', finish);
  }
}

export class MapScene extends Phaser.Scene {
  constructor() {
    super('Map');
  }

  create(): void {
    ensureBackground(this);
    const root = showScreen(mapScreen(), { clear: true });
    on(root, '[data-stage]', (el) => {
      const id = el.dataset.stage as StageId;
      sfx.unlock();
      sfx.play('click');
      hideScreen();
      this.scene.start(stageScene(id));
    });
    on(root, '[data-act=back]', () => {
      hideScreen();
      this.scene.start('Menu');
    });
    on(root, '[data-act=hangar]', () => {
      hideScreen();
      this.scene.start('Hangar', { back: 'Map' });
    });
  }
}

export class Hangar extends Phaser.Scene {
  private back = 'Menu';
  constructor() {
    super('Hangar');
  }

  create(data: { back?: string }): void {
    ensureBackground(this);
    this.back = data.back ?? 'Menu';
    this.render();
  }

  private render(): void {
    bindHangar(
      () => this.scene.start(this.back),
      () => this.render(),
    );
  }
}
