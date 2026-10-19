// Every tuning number lives here. Units: px (720x1280 base), seconds, px/s.
export const GAME_W = 720;
export const GAME_H = 1280;

export const player = {
  maxSpeed: 900,
  smoothing: 0.22, // 0..1 per frame fraction toward target (light inertia)
  touchOffsetY: 110, // rocket sits above the finger
  hitRadius: 26,
  invulnSeconds: 1.0,
  blinkHz: 3,
  hearts: 3,
  marginX: 48,
  marginTop: 90,
  marginBottom: 120,
};

export const laser = {
  shotsPerSecond: 6,
  bulletSpeed: 1100,
  damageByLevel: [1, 1.35, 1.7, 2.1],
  tripleSpread: 14, // degrees
};

export const combo = {
  perStep: 8, // hits/pickups to raise the multiplier one step
  maxMult: 5,
  decaySeconds: 6, // no action for this long -> drop one step
};

export const nearMiss = {
  distance: 30, // px gap from meteor edge
  bonus: 50,
  comboCredit: 2,
  slowMs: 80,
  slowScale: 0.3,
};

export const powerups = {
  durationSeconds: 8,
  shieldByLevel: [8, 10, 12, 14],
  magnetByLevel: [160, 220, 290, 360],
  firstAppearAt: 42, // s into stage 1
  intervalMin: 18,
  intervalMax: 25,
};

export const turbo = {
  chargePerCoin: [0.04, 0.05, 0.06, 0.075],
  durationSeconds: 1.5,
  speedBoost: 1.6,
};

export const feel = {
  hitStopMs: 50,
  shakeHit: 0.004,
  shakeBoss: 0.008,
  shakeDuration: 180,
};

export const scoring = {
  meteorSmall: 30,
  meteorBig: 80,
  fragment: 15,
  coin: 10,
  diamond: 100, // 10 coins
  gate: 120,
  bossHit: 25,
  bossKill: 1500,
  secretStar: 300,
  heartLeftBonus: 400,
};

export const stars = {
  marsScore: 12000,
  saturnScore: 18000,
  hiddenPerStage: 3,
};

export const assist = {
  lossesBeforeHelp: 2,
  densityReduction: 0.15,
};

export const economy = {
  diamondValue: 10,
  upgradeCosts: [120, 300, 500],
  skinCost: 200,
};

export const bossRock = {
  crackHp: 25,
  cracks: 3,
  ringFragments: 12,
  ringIntervalSeconds: 2.6,
  warnSeconds: 0.8,
  fragmentSpeed: 300,
  moveSpeed: 160,
  coinRain: 30,
};

export const bossRobot = {
  coreHp: 30,
  cores: 2,
  beamWarnSeconds: 0.8,
  beamSweepSeconds: 1.4,
  beamWidth: 70,
  junkWaveInterval: 2.2,
  chaserSpeed: 150,
  chaserHp: 3,
  phase3At: 0.3,
  phase3SpeedMul: 1.35,
  coreExposeSeconds: 1.2,
  coinRain: 40,
};

export const saturn = {
  gravityAccel: 180,
  gravityFlipMin: 6,
  gravityFlipMax: 9,
  gravityWarnSeconds: 1.5,
  gateComboBonus: 1,
  gateDashSeconds: 0.4,
  iceSpeed: 220,
};

// Stage timelines: each wave = [startSec, endSec, spawnsPerSec, speedMin, speedMax]
export const marsWaves = [
  { from: 0, to: 40, rate: 0.9, speed: [220, 300], big: 0, diagonal: 0 },
  { from: 40, to: 90, rate: 1.4, speed: [300, 420], big: 0.2, diagonal: 0 },
  { from: 90, to: 130, rate: 2.0, speed: [380, 480], big: 0.25, diagonal: 0.4 },
] as const;
export const marsBossAt = 130;

export const saturnSections = [
  { from: 0, to: 60, rate: 1.0, speed: [300, 400], gates: true, ice: 0.3, narrow: false },
  { from: 60, to: 110, rate: 1.3, speed: [340, 460], gates: true, ice: 0.4, narrow: true },
] as const;
export const saturnBossAt = 110;
