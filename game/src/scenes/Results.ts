// Results: stars, save progress, achievements, then next/retry/hangar. Shows the victory scene after Saturn.
import Phaser from 'phaser';
import { ResultData } from './Play';
import { save, AchievementId } from '../systems/save';
import { stars as ST } from '../config/balance';
import { showScreen, hideScreen, on, mountLottie, $, toast } from '../ui/dom';
import { resultsScreen, victoryScreen, ResultView } from '../ui/screens';
import { lotties } from '../assets';
import { S } from '../config/strings.ar';
import { ensureBackground, stageScene } from './Menus';
import { sfx } from '../systems/audio';

export class Results extends Phaser.Scene {
  constructor() {
    super('Results');
  }

  create(r: ResultData): void {
    ensureBackground(this);
    const d = save.data;
    const unlocked: AchievementId[] = [];
    const ach = (id: AchievementId, cond: boolean) => {
      if (cond && save.unlockAchievement(id)) unlocked.push(id);
    };

    // bank coins
    d.coins += r.coins;
    d.totalCoins += r.coins;

    let view: ResultView;
    if (r.challenge) {
      d.challengeTop = [...d.challengeTop, r.score].sort((a, b) => b - a).slice(0, 5);
      view = { stageId: 'mars', won: false, score: r.score, coins: r.coins, maxCombo: r.maxCombo, nearMisses: r.nearMisses, stars: [false, false, false], newBest: d.challengeTop[0] === r.score, canCheckpoint: false, challenge: true, nextStage: null };
    } else {
      const hidden = d.hidden[r.stageId].map((h, i) => h || r.hiddenFound[i]);
      d.hidden[r.stageId] = hidden;
      const threshold = r.stageId === 'mars' ? ST.marsScore : ST.saturnScore;
      const stars: [boolean, boolean, boolean] = [r.won, r.won && r.score >= threshold, r.won && hidden.every(Boolean)];
      const count = stars.filter(Boolean).length;
      d.stars[r.stageId] = Math.max(d.stars[r.stageId], count);
      const newBest = r.score > d.best[r.stageId];
      if (newBest) d.best[r.stageId] = r.score;
      if (r.won) {
        if (r.stageId === 'mars') {
          d.unlocked.saturn = true;
          d.unlocked.challenge = true;
        }
        Object.keys(d.losses).filter((k) => k.startsWith(r.stageId)).forEach((k) => delete d.losses[k]);
      }
      view = { stageId: r.stageId, won: r.won, score: r.score, coins: r.coins, maxCombo: r.maxCombo, nearMisses: r.nearMisses, stars, newBest, canCheckpoint: r.canCheckpoint, challenge: false, nextStage: r.won && r.stageId === 'mars' ? 'saturn' : null };
    }
    ach('firstLaunch', true);
    ach('combo5', r.maxCombo >= 5);
    ach('nearMiss10', r.nearMisses >= 10);
    ach('noHit', r.won && r.hitsTaken === 0);
    ach('rockPerfect', r.won && r.stageId === 'mars' && r.bossPerfect);
    ach('fixRobot', r.won && r.stageId === 'saturn');
    ach('coins1000', d.totalCoins >= 1000);
    ach('allStars', d.hidden.mars.every(Boolean) && d.hidden.saturn.every(Boolean));
    save.write();

    if (r.won && r.stageId === 'saturn') this.showVictory(r, view, unlocked);
    else this.showResults(r, view, unlocked);
  }

  private showVictory(r: ResultData, view: ResultView, unlocked: AchievementId[]): void {
    const root = showScreen(victoryScreen(r.score), { clear: true });
    mountLottie($('#celebrateAnim', root), lotties.celebrate);
    sfx.play('victory');
    on(root, '[data-act=next]', () => this.showResults(r, view, unlocked));
  }

  private showResults(r: ResultData, view: ResultView, unlocked: AchievementId[]): void {
    const root = showScreen(resultsScreen(view), { clear: true });
    root.querySelectorAll<HTMLElement>('.big-stars .s').forEach((s, i) => {
      setTimeout(() => {
        s.classList.add('in');
        if (s.classList.contains('on')) sfx.play('star');
      }, 300 + i * 350);
    });
    unlocked.forEach((id, i) => setTimeout(() => toast(`🏅 ${S.achievementsUI.list[id].name}`), 1500 + i * 2000));
    const go = (key: string, data?: object) => {
      hideScreen();
      this.scene.start(key, data);
    };
    on(root, '[data-act=next]', () => go(stageScene(view.nextStage ?? 'mars')));
    on(root, '[data-act=retry]', () => go(stageScene(r.stageId), { challenge: r.challenge }));
    on(root, '[data-act=checkpoint]', () => go(stageScene(r.stageId), { fromCheckpoint: true, score: r.checkpoint?.score ?? 0, coins: r.checkpoint?.coins ?? 0 }));
    on(root, '[data-act=hangar]', () => go('Hangar', { back: 'Map' }));
    on(root, '[data-act=map]', () => go(r.challenge ? 'Menu' : 'Map'));
  }
}
