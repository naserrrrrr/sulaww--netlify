import { RocketSkin } from '../config/palette';

export type StageId = 'mars' | 'saturn';
export type UpgradeId = 'laser' | 'shield' | 'magnet' | 'turbo';
export type AchievementId =
  | 'firstLaunch'
  | 'combo5'
  | 'nearMiss10'
  | 'noHit'
  | 'allStars'
  | 'rockPerfect'
  | 'coins1000'
  | 'fixRobot';

export interface Settings {
  sound: boolean;
  volume: number; // 0..1
  vibrate: boolean;
  reduceMotion: boolean;
  quality: 'auto' | 'low' | 'high';
}

export interface SaveData {
  version: 1;
  coins: number;
  totalCoins: number;
  best: Record<StageId, number>;
  stars: Record<StageId, number>;
  hidden: Record<StageId, boolean[]>;
  upgrades: Record<UpgradeId, number>; // 0..3
  skins: RocketSkin[];
  skin: RocketSkin;
  achievements: Partial<Record<AchievementId, number>>; // unlock timestamp
  unlocked: { saturn: boolean; challenge: boolean };
  losses: Record<string, number>; // `${stage}:${segment}` -> consecutive losses
  challengeTop: number[];
  storySeen: boolean;
  settings: Settings;
}

const KEY = 'rb-galaxy-save-v1';

const defaults = (): SaveData => ({
  version: 1,
  coins: 0,
  totalCoins: 0,
  best: { mars: 0, saturn: 0 },
  stars: { mars: 0, saturn: 0 },
  hidden: { mars: [false, false, false], saturn: [false, false, false] },
  upgrades: { laser: 0, shield: 0, magnet: 0, turbo: 0 },
  skins: ['blue'],
  skin: 'blue',
  achievements: {},
  unlocked: { saturn: false, challenge: false },
  losses: {},
  challengeTop: [],
  storySeen: false,
  settings: { sound: true, volume: 0.8, vibrate: true, reduceMotion: false, quality: 'auto' },
});

class SaveStore {
  data: SaveData = defaults();

  load(): void {
    try {
      const raw = localStorage.getItem(KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw) as Partial<SaveData>;
      if (parsed.version !== 1) return;
      const d = defaults();
      this.data = {
        ...d,
        ...parsed,
        best: { ...d.best, ...parsed.best },
        stars: { ...d.stars, ...parsed.stars },
        hidden: { ...d.hidden, ...parsed.hidden },
        upgrades: { ...d.upgrades, ...parsed.upgrades },
        unlocked: { ...d.unlocked, ...parsed.unlocked },
        settings: { ...d.settings, ...parsed.settings },
      };
    } catch {
      this.data = defaults();
    }
  }

  write(): void {
    try {
      localStorage.setItem(KEY, JSON.stringify(this.data));
    } catch {
      /* storage unavailable: play without persistence */
    }
  }

  reset(): void {
    const settings = this.data.settings;
    this.data = { ...defaults(), settings };
    this.write();
  }

  unlockAchievement(id: AchievementId): boolean {
    if (this.data.achievements[id]) return false;
    this.data.achievements[id] = Date.now();
    this.write();
    return true;
  }
}

export const save = new SaveStore();
