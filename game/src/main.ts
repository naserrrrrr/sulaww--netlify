import Phaser from 'phaser';
import './ui/ui.css';
import { GAME_H, GAME_W } from './config/balance';
import { save } from './systems/save';
import { sfx } from './systems/audio';
import { fitOverlay } from './ui/dom';
import { logos } from './assets';
import { Boot } from './scenes/Boot';
import { Background } from './scenes/Background';
import { Menu, Story, MapScene, Hangar } from './scenes/Menus';
import { StageMars } from './scenes/StageMars';
import { StageSaturn } from './scenes/StageSaturn';
import { Results } from './scenes/Results';

save.load();
document.documentElement.classList.toggle('reduce-motion', save.data.settings.reduceMotion);
(document.getElementById('desk-logo') as HTMLImageElement).src = logos.horizontalWhite;

const params = new URLSearchParams(location.search);
const debug = params.get('debug') === '1';

async function start(): Promise<void> {
  try {
    await Promise.all([document.fonts.load('900 32px MA'), document.fonts.load('700 32px MA')]);
  } catch {
    /* fallback font */
  }
  const game = new Phaser.Game({
    type: Phaser.AUTO,
    parent: 'game',
    width: GAME_W,
    height: GAME_H,
    backgroundColor: '#101340',
    scale: { mode: Phaser.Scale.FIT, autoCenter: Phaser.Scale.CENTER_BOTH, width: GAME_W, height: GAME_H },
    physics: { default: 'arcade', arcade: { debug, gravity: { x: 0, y: 0 } } },
    render: { antialias: true, pixelArt: false, powerPreference: 'high-performance', roundPixels: false },
    input: { activePointers: 2 },
    disableContextMenu: true,
    fps: { target: 60, forceSetTimeOut: false },
    scene: [Boot, Background, Menu, Story, MapScene, Hangar, StageMars, StageSaturn, Results],
  });
  if (debug) (window as unknown as { __game: Phaser.Game }).__game = game;
  const refit = () => fitOverlay(game.canvas);
  game.scale.on(Phaser.Scale.Events.RESIZE, refit);
  window.addEventListener('resize', () => setTimeout(refit, 50));
  game.events.once(Phaser.Core.Events.READY, () => setTimeout(refit, 0));
  setTimeout(refit, 300);

  // dev jump: ?stage=2&boss=1 (read by Boot)
  const stage = params.get('stage');
  if (stage) game.registry.set('devJump', { key: stage === '2' ? 'StageSaturn' : 'StageMars', jumpToBoss: params.get('boss') === '1', challenge: params.get('challenge') === '1' });
  document.addEventListener('pointerdown', () => sfx.unlock(), { once: true });
  document.addEventListener('keydown', () => sfx.unlock(), { once: true });
}

void start();
