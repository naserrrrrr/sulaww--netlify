// Base gameplay scene: vertical scroller loop, pooling, combo, near-miss, power-ups, turbo, HUD, radio.
import Phaser from 'phaser';
import { GAME_H, GAME_W, assist, combo as C, economy, feel, laser, nearMiss as NM, player as P, powerups as PW, scoring, turbo as T } from '../config/balance';
import { S } from '../config/strings.ar';
import { Player } from '../entities/Player';
import { Meteor } from '../entities/Meteor';
import { Bullet } from '../entities/Bullet';
import { Pickup, PickupKind } from '../entities/Pickup';
import { Hazard, HazardKind } from '../entities/Hazard';
import { Gate } from '../entities/Gate';
import { Hud } from '../ui/hud';
import { Radio } from '../ui/radio';
import { hideScreen, showScreen, on } from '../ui/dom';
import { pauseScreen, openSettings } from '../ui/screens';
import { Sfx, sfx, vibrate } from '../systems/audio';
import { save, StageId } from '../systems/save';

export interface PlayData {
  fromCheckpoint?: boolean;
  score?: number;
  coins?: number;
  challenge?: boolean;
  jumpToBoss?: boolean;
}

export interface ResultData {
  stageId: StageId;
  won: boolean;
  score: number;
  coins: number;
  maxCombo: number;
  nearMisses: number;
  hitsTaken: number;
  hiddenFound: boolean[];
  bossPerfect: boolean;
  canCheckpoint: boolean;
  checkpoint: { score: number; coins: number } | null;
  challenge: boolean;
  elapsed: number;
}

export abstract class PlayScene extends Phaser.Scene {
  abstract readonly stageId: StageId;
  abstract readonly bossAt: number;
  abstract readonly bossName: string;
  abstract readonly planetKey: string;

  strings = S;
  player!: Player;
  meteors!: Phaser.Physics.Arcade.Group;
  bullets!: Phaser.Physics.Arcade.Group;
  pickups!: Phaser.Physics.Arcade.Group;
  hazards!: Phaser.Physics.Arcade.Group;
  gates!: Phaser.Physics.Arcade.Group;
  hud!: Hud;
  radio!: Radio;
  burst!: Phaser.GameObjects.Particles.ParticleEmitter;
  lowQuality = false;
  challenge = false;

  elapsed = 0;
  score = 0;
  coins = 0;
  hearts = P.hearts;
  comboCount = 0;
  comboMult = 1;
  bonusMult = 0;
  comboIdle = 0;
  maxCombo = 1;
  nearMisses = 0;
  hitsTaken = 0;
  hiddenFound = [false, false, false];
  bossStarted = false;
  bossDone = false;
  bossHeartsAtStart = P.hearts;
  over = false;
  paused = false;
  protected spawnAcc = 0;
  protected coinTimer = 2.5;
  protected powerTimer = PW.firstAppearAt;
  protected starIndex = 0;
  protected starTimes: number[] = [];
  protected assistFactor = 1;
  protected slowUntil = 0;
  protected frozenUntil = 0;
  protected fireTimer = 0;
  protected gravityDir: -1 | 0 | 1 = 0;
  protected gravityDrift = 0;
  protected checkpoint: { score: number; coins: number } | null = null;
  private turboTold = false;
  protected cursors!: Phaser.Types.Input.Keyboard.CursorKeys;
  protected wasd!: Record<string, Phaser.Input.Keyboard.Key>;
  private farStars!: Phaser.GameObjects.TileSprite;
  private nearStars!: Phaser.GameObjects.TileSprite;
  private planet!: Phaser.GameObjects.Image;
  private clouds: Phaser.GameObjects.Image[] = [];
  private fpsText: Phaser.GameObjects.Text | null = null;
  private fpsSamples = 0;
  private fpsSum = 0;
  private debug = false;

  create(data: PlayData): void {
    this.debug = new URLSearchParams(location.search).get('debug') === '1';
    this.lowQuality = save.data.settings.quality === 'low' || (save.data.settings.quality === 'auto' && !!this.registry.get('lowQuality'));
    this.challenge = !!data.challenge;
    hideScreen();
    if (this.scene.isActive('Background')) this.scene.stop('Background');
    this.resetState();
    this.buildBackground();

    this.meteors = this.physics.add.group({ classType: Meteor, maxSize: 90, runChildUpdate: false });
    this.bullets = this.physics.add.group({ classType: Bullet, maxSize: 60 });
    this.pickups = this.physics.add.group({ classType: Pickup, maxSize: 80 });
    this.hazards = this.physics.add.group({ classType: Hazard, maxSize: 90 });
    this.gates = this.physics.add.group({ classType: Gate, maxSize: 8 });
    this.meteors.setDepth(16);
    this.pickups.setDepth(15);
    this.hazards.setDepth(17);
    this.bullets.setDepth(18);
    this.gates.setDepth(14);

    this.player = new Player(this, save.data.skin, this.lowQuality);
    this.burst = this.add.particles(0, 0, 'particle', {
      emitting: false,
      speed: { min: 90, max: 420 },
      scale: { start: 0.9, end: 0 },
      alpha: { start: 1, end: 0 },
      lifespan: { min: 250, max: 650 },
      blendMode: Phaser.BlendModes.ADD,
    }).setDepth(25);

    this.physics.add.overlap(this.bullets, this.meteors, (b, m) => this.bulletMeteor(b as Bullet, m as Meteor));
    this.physics.add.overlap(this.bullets, this.hazards, (b, h) => this.bulletHazard(b as Bullet, h as Hazard));
    this.physics.add.overlap(this.player, this.meteors, (_p, m) => this.playerMeteor(m as Meteor));
    this.physics.add.overlap(this.player, this.hazards, (_p, h) => this.playerHazard(h as Hazard));
    this.physics.add.overlap(this.player, this.pickups, (_p, k) => this.collect(k as Pickup));
    this.physics.add.overlap(this.player, this.gates, (_p, g) => this.passGate(g as Gate));

    this.hud = new Hud(() => this.pauseGame(), () => this.turbo());
    this.radio = new Radio();
    this.hud.setHearts(this.hearts);
    this.hud.setCoins(this.coins);
    this.hud.setScore(this.score);
    this.hud.setTurbo(0, false);

    this.setupInput();

    const key = `${this.stageId}:${this.segmentKey()}`;
    this.assistFactor = (save.data.losses[key] ?? 0) >= assist.lossesBeforeHelp ? 1 - assist.densityReduction : 1;

    if (data.fromCheckpoint || data.jumpToBoss) {
      this.score = data.score ?? 0;
      this.coins = data.coins ?? 0;
      this.elapsed = this.bossAt - 1.5;
      this.starIndex = 3;
      this.hud.setScore(this.score);
      this.hud.setCoins(this.coins);
    } else {
      this.radioSay(S.radio.start, true);
    }
    const pre = this.bossAt;
    this.starTimes = [0.22, 0.5, 0.78].map((f) => f * pre);
    if (this.challenge) this.starTimes = [];

    this.game.events.on(Phaser.Core.Events.HIDDEN, this.onHidden, this);
    this.events.once(Phaser.Scenes.Events.SHUTDOWN, () => {
      this.game.events.off(Phaser.Core.Events.HIDDEN, this.onHidden, this);
      this.hud.destroy();
      this.radio.destroy();
      hideScreen();
    });
    if (this.debug) this.fpsText = this.add.text(10, 180, '', { fontFamily: 'monospace', fontSize: '22px', color: '#00DF6A' }).setDepth(99);
    this.stageCreate();
  }

  protected resetState(): void {
    this.elapsed = 0;
    this.score = 0;
    this.coins = 0;
    this.hearts = P.hearts;
    this.comboCount = 0;
    this.comboMult = 1;
    this.bonusMult = 0;
    this.comboIdle = 0;
    this.maxCombo = 1;
    this.nearMisses = 0;
    this.hitsTaken = 0;
    this.hiddenFound = [false, false, false];
    this.bossStarted = false;
    this.bossDone = false;
    this.over = false;
    this.paused = false;
    this.spawnAcc = 0;
    this.coinTimer = 2.5;
    this.powerTimer = PW.firstAppearAt;
    this.starIndex = 0;
    this.slowUntil = 0;
    this.frozenUntil = 0;
    this.fireTimer = 0;
    this.gravityDir = 0;
    this.checkpoint = null;
    this.turboTold = false;
    this.clouds = [];
  }

  // ---------- hooks for stages ----------
  protected abstract stageCreate(): void;
  protected abstract stageUpdate(dt: number): void;
  protected abstract createBoss(): void;
  protected abstract bossUpdate(dt: number): void;
  protected abstract bossAlive(): boolean;
  protected abstract bossHpFrac(): number;
  protected abstract segmentKey(): string;
  protected abstract debugKillBoss(): void;

  // ---------- background ----------
  private buildBackground(): void {
    this.cameras.main.setBackgroundColor('#101340');
    this.farStars = this.add.tileSprite(GAME_W / 2, GAME_H / 2, GAME_W, GAME_H, 'stars-far').setDepth(1).setAlpha(0.7);
    this.nearStars = this.add.tileSprite(GAME_W / 2, GAME_H / 2, GAME_W, GAME_H, 'stars-near').setDepth(3);
    for (let i = 0; i < (this.lowQuality ? 2 : 4); i++) {
      const c = this.add.image(Phaser.Math.Between(0, GAME_W), Phaser.Math.Between(-200, GAME_H), 'cloud').setDepth(2).setAlpha(0.5).setScale(Phaser.Math.FloatBetween(1.2, 2.4));
      c.setData('vy', Phaser.Math.FloatBetween(18, 40));
      this.clouds.push(c);
    }
    this.planet = this.add.image(this.stageId === 'mars' ? 140 : GAME_W - 150, -220, this.planetKey).setDepth(2).setScale(this.stageId === 'mars' ? 1.1 : 1.25).setAlpha(0.6);
  }

  private updateBackground(dt: number, speedMul: number): void {
    this.farStars.tilePositionY -= 40 * dt * speedMul;
    this.nearStars.tilePositionY -= 120 * dt * speedMul;
    this.clouds.forEach((c) => {
      c.y += (c.getData('vy') as number) * dt * speedMul;
      if (c.y > GAME_H + 200) {
        c.y = -200;
        c.x = Phaser.Math.Between(0, GAME_W);
      }
    });
    const total = this.challenge ? 240 : this.bossAt + 70;
    this.planet.y = -220 + Math.min(1, this.elapsed / total) * (GAME_H + 440);
  }

  // ---------- input ----------
  private setupInput(): void {
    this.input.on('pointerdown', (p: Phaser.Input.Pointer) => {
      sfx.unlock();
      this.player.pointerDown(p);
    });
    this.input.on('pointermove', (p: Phaser.Input.Pointer) => this.player.pointerMove(p));
    this.input.on('pointerup', () => this.player.pointerUp());
    this.input.addPointer(1);
    const kb = this.input.keyboard;
    if (!kb) return;
    this.cursors = kb.createCursorKeys();
    this.wasd = kb.addKeys('W,A,S,D,P,ESC,SPACE') as Record<string, Phaser.Input.Keyboard.Key>;
    kb.on('keydown-SPACE', () => this.turbo());
    kb.on('keydown-P', () => this.pauseGame());
    kb.on('keydown-ESC', () => this.pauseGame());
    if (this.debug) {
      // dev hotkeys: K lose instantly, V skip to boss / defeat boss
      kb.on('keydown-K', () => {
        this.hearts = 1;
        this.player.invulnUntil = 0;
        this.hitPlayer();
      });
      kb.on('keydown-V', () => {
        if (!this.bossStarted) this.elapsed = this.bossAt;
        else this.debugKillBoss();
      });
    }
  }

  private onHidden(): void {
    if (!this.over && !this.paused) this.pauseGame();
  }

  pauseGame(): void {
    if (this.over || this.paused) return;
    this.paused = true;
    this.player.pointerUp();
    this.scene.pause();
    this.showPauseScreen();
  }

  private showPauseScreen(): void {
    const root = showScreen(pauseScreen());
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'p' || e.key === 'P' || e.key === 'Escape') {
        window.removeEventListener('keydown', onKey);
        if (this.paused && !this.over) this.resumeGame();
      }
    };
    window.addEventListener('keydown', onKey);
    on(root, '[data-act=resume]', () => this.resumeGame());
    on(root, '[data-act=restart]', () => {
      hideScreen();
      this.scene.restart({ challenge: this.challenge });
    });
    on(root, '[data-act=menu]', () => {
      hideScreen();
      this.scene.start('Menu');
    });
    on(root, '[data-act=settings]', () => openSettings(() => this.showPauseScreen()));
  }

  resumeGame(): void {
    hideScreen();
    this.paused = false;
    this.physics.resume();
    this.scene.resume();
  }

  turbo(): void {
    if (this.over || this.paused) return;
    if (this.player.tryTurbo()) {
      this.sfx('turbo');
      vibrate(30);
      this.hud.banner('توربو!', true);
    }
  }

  // ---------- main loop ----------
  update(time: number, delta: number): void {
    if (this.over) return;
    if (this.fpsText) {
      this.fpsText.setText(`${Math.round(this.game.loop.actualFps)} fps · m${this.meteors.countActive()} h${this.hazards.countActive()} p${this.pickups.countActive()}`);
    }
    if (save.data.settings.quality === 'auto' && this.fpsSamples < 240 && this.elapsed > 1.5) {
      this.fpsSum += this.game.loop.actualFps;
      if (++this.fpsSamples === 240 && this.fpsSum / 240 < 48) this.registry.set('lowQuality', true);
    }
    if (time < this.frozenUntil) return; // hit-stop
    const slow = time < this.slowUntil;
    this.physics.world.timeScale = slow ? 1 / NM.slowScale : 1;
    const dt = Math.min(delta, 100) / 1000 / (slow ? 1 / NM.slowScale : 1);
    this.elapsed += dt;

    this.player.update(dt, { keys: this.keyState() }, this.gravityDir, this.gravityDrift);
    this.updateBackground(dt, this.player.turboActive ? 3 : 1);
    this.autoFire(dt);
    this.updateCombo(dt);
    this.updateMagnet();
    this.checkNearMiss();
    this.hud.setTurbo(this.player.turboCharge, this.player.turboActive);

    if (!this.bossStarted) {
      if (!this.challenge && this.elapsed >= this.bossAt) this.startBoss();
      else {
        this.stageUpdate(dt);
        this.updateCommonSpawns(dt);
      }
    } else if (!this.bossDone) {
      this.bossUpdate(dt);
      this.hud.setBoss(this.bossHpFrac());
      if (!this.bossAlive()) this.bossDefeated();
    }
  }

  private keyState() {
    const c = this.cursors;
    const w = this.wasd;
    if (!c) return { left: false, right: false, up: false, down: false };
    return {
      left: c.left.isDown || w.A.isDown,
      right: c.right.isDown || w.D.isDown,
      up: c.up.isDown || w.W.isDown,
      down: c.down.isDown || w.S.isDown,
    };
  }

  private autoFire(dt: number): void {
    this.fireTimer -= dt;
    if (this.fireTimer > 0 || this.player.dead) return;
    this.fireTimer = 1 / laser.shotsPerSecond;
    const dmg = laser.damageByLevel[save.data.upgrades.laser] ?? 1;
    const triple = this.time.now < this.player.tripleUntil;
    const angles = triple ? [-laser.tripleSpread, 0, laser.tripleSpread] : [0];
    for (const a of angles) {
      const b = this.bullets.get(this.player.x, this.player.y - 70) as Bullet | null;
      b?.fire(this.player.x, this.player.y - 70, a, dmg);
    }
    this.sfx('laser');
  }

  private updateCombo(dt: number): void {
    this.comboIdle += dt;
    if (this.comboIdle > C.decaySeconds && this.comboCount > 0) {
      this.comboCount = Math.max(0, this.comboCount - C.perStep);
      this.comboIdle = 0;
      this.refreshCombo(false);
    }
  }

  private refreshCombo(bump: boolean): void {
    this.comboMult = Math.min(C.maxMult, 1 + Math.floor(this.comboCount / C.perStep) + this.bonusMult);
    this.maxCombo = Math.max(this.maxCombo, this.comboMult);
    this.hud.setCombo(this.comboMult, (this.comboCount % C.perStep) / C.perStep, bump);
  }

  private updateMagnet(): void {
    const range = this.time.now < this.player.magnetUntil ? PW.magnetByLevel[save.data.upgrades.magnet] ?? 160 : 70;
    const px = this.player.x;
    const py = this.player.y;
    this.pickups.getChildren().forEach((o) => {
      const k = o as Pickup;
      if (!k.active) return;
      const d = Phaser.Math.Distance.Between(px, py, k.x, k.y);
      if (d < range || k.pulled) {
        k.pulled = true;
        const a = Phaser.Math.Angle.Between(k.x, k.y, px, py);
        const sp = 700 + (range - d) * 2;
        (k.body as Phaser.Physics.Arcade.Body).setVelocity(Math.cos(a) * sp, Math.sin(a) * sp);
      }
    });
  }

  private checkNearMiss(): void {
    if (this.player.dead || this.player.turboActive) return;
    const px = this.player.x;
    const py = this.player.y;
    this.meteors.getChildren().forEach((o) => {
      const m = o as Meteor;
      if (!m.active || m.passedPlayer || m.wall) return;
      if (m.y >= py - 10) {
        m.passedPlayer = true;
        const gap = Phaser.Math.Distance.Between(px, py, m.x, m.y) - m.radius - P.hitRadius;
        if (gap > -4 && gap < NM.distance) this.nearMiss(m);
      }
    });
  }

  private nearMiss(m: Meteor): void {
    this.nearMisses++;
    this.comboCount += NM.comboCredit;
    this.comboIdle = 0;
    this.refreshCombo(true);
    this.addScore(NM.bonus, (this.player.x + m.x) / 2, this.player.y - 40, S.float.nearMiss, '#F876FF');
    if (!save.data.settings.reduceMotion) this.slowUntil = this.time.now + NM.slowMs;
    this.sfx('nearMiss');
    if (this.nearMisses === 3 || this.nearMisses === 10) this.radioSay(S.radio.nearMiss);
  }

  // ---------- spawning helpers ----------
  spawnMeteor(x: number, y: number, vx: number, vy: number, big: boolean, hp?: number, wall = false): Meteor | null {
    const m = this.meteors.get(x, y) as Meteor | null;
    if (!m) return null;
    return m.spawn(x, y, vx, vy, big, hp, wall);
  }

  spawnHazard(kind: HazardKind, x: number, y: number, vx: number, vy: number): Hazard | null {
    const h = this.hazards.get(x, y) as Hazard | null;
    if (!h) return null;
    return h.spawn(kind, x, y, vx, vy);
  }

  spawnPickup(kind: PickupKind, x: number, y: number, vy: number, vx = 0): Pickup | null {
    const k = this.pickups.get(x, y) as Pickup | null;
    if (!k) return null;
    return k.spawn(kind, x, y, vy, vx);
  }

  spawnGate(x: number, y: number, vy: number, sway = 0): Gate | null {
    const g = this.gates.get(x, y) as Gate | null;
    if (!g) return null;
    return g.spawn(x, y, vy, sway);
  }

  /** Coin trail patterns: line, arc, zigzag, diagonal. */
  spawnCoinPattern(vy: number, xMin = 70, xMax = GAME_W - 70, dense = false): void {
    const type = Phaser.Math.Between(0, 3);
    const n = dense ? 9 : Phaser.Math.Between(5, 7);
    const x0 = Phaser.Math.Between(xMin + 40, xMax - 40);
    const diamond = Math.random() < 0.07 ? Phaser.Math.Between(0, n - 1) : -1;
    for (let i = 0; i < n; i++) {
      let x = x0;
      if (type === 1) x = x0 + Math.sin((i / (n - 1)) * Math.PI) * 120 * (x0 < GAME_W / 2 ? 1 : -1);
      if (type === 2) x = x0 + (i % 2 === 0 ? -55 : 55);
      if (type === 3) x = x0 + (i - n / 2) * 50;
      x = Phaser.Math.Clamp(x, xMin, xMax);
      this.spawnPickup(i === diamond ? 'diamond' : 'coin', x, -40 - i * 68, vy);
    }
    if (this.hearts < P.hearts && Math.random() < 0.08) this.spawnPickup('heart', Phaser.Math.Between(xMin, xMax), -520, vy);
  }

  protected updateCommonSpawns(dt: number): void {
    this.coinTimer -= dt;
    if (this.coinTimer <= 0) {
      this.coinTimer = Phaser.Math.FloatBetween(3, 4.5);
      this.spawnCoinPattern(this.currentSpeed() * 0.8);
      if (this.elapsed < 20 && this.elapsed > 10) this.radioSay(S.radio.coins);
    }
    this.powerTimer -= dt;
    if (this.powerTimer <= 0) {
      this.powerTimer = Phaser.Math.FloatBetween(PW.intervalMin, PW.intervalMax);
      const kinds: PickupKind[] = ['shield', 'magnet', 'triple'];
      this.spawnPickup(kinds[Phaser.Math.Between(0, 2)], Phaser.Math.Between(90, GAME_W - 90), -60, this.currentSpeed() * 0.7);
    }
    if (this.starIndex < this.starTimes.length && this.elapsed >= this.starTimes[this.starIndex]) {
      const i = this.starIndex++;
      const x = i % 2 === 0 ? 60 : GAME_W - 60; // near the edge: worth the risk
      const star = this.spawnPickup('star', x, -60, this.currentSpeed() * 0.75);
      star?.setData('index', i);
    }
  }

  protected abstract currentSpeed(): number;

  // ---------- collisions ----------
  private bulletMeteor(b: Bullet, m: Meteor): void {
    if (!b.active || !m.active) return;
    b.kill();
    m.hp -= b.damage;
    this.burst.setParticleTint(0x8cd9ff);
    this.burst.explode(3, b.x, b.y);
    if (m.hp <= 0) this.destroyMeteor(m, true);
    else {
      m.setScale(1.08);
      this.tweens.add({ targets: m, scale: 1, duration: 80 });
    }
  }

  private destroyMeteor(m: Meteor, byBullet: boolean): void {
    m.kill();
    const pts = m.wall ? scoring.fragment : m.big ? scoring.meteorBig : scoring.meteorSmall;
    this.comboCount++;
    this.comboIdle = 0;
    this.refreshCombo(true);
    this.addScore(pts, m.x, m.y);
    this.explode(m.x, m.y, m.wall ? 0x5143de : 0x3a299b, m.big ? 22 : 10);
    this.sfx('crash');
    if (m.big && byBullet) {
      const n = Phaser.Math.Between(2, 3);
      for (let i = 0; i < n; i++) {
        const a = -Math.PI / 2 + (i - (n - 1) / 2) * 0.9 + Math.PI;
        this.spawnMeteor(m.x + Math.cos(a) * 30, m.y + Math.sin(a) * 30, Math.cos(a) * 160, Math.abs(Math.sin(a)) * 120 + this.currentSpeed() * 0.9, false);
      }
    }
    if (Math.random() < 0.18) this.spawnPickup('coin', m.x, m.y, this.currentSpeed() * 0.7);
  }

  private bulletHazard(b: Bullet, h: Hazard): void {
    if (!b.active || !h.active) return;
    b.kill();
    this.burst.setParticleTint(0x8cd9ff);
    this.burst.explode(3, b.x, b.y);
    if (h.kind === 'ice' || h.kind === 'edge') return;
    h.hp -= b.damage;
    if (h.hp <= 0) {
      h.kill();
      this.comboCount++;
      this.comboIdle = 0;
      this.refreshCombo(true);
      this.addScore(h.kind === 'chaser' ? scoring.meteorBig : scoring.fragment, h.x, h.y);
      this.explode(h.x, h.y, 0xfa500e, h.kind === 'chaser' ? 18 : 8);
      this.sfx('crash');
    }
  }

  private playerMeteor(m: Meteor): void {
    if (!m.active || this.player.dead) return;
    if (this.player.turboActive) {
      this.destroyMeteor(m, false);
      return;
    }
    if (this.hitPlayer()) m.passedPlayer = true;
  }

  private playerHazard(h: Hazard): void {
    if (!h.active || this.player.dead) return;
    if (this.player.turboActive && h.kind !== 'edge' && h.kind !== 'ice') {
      h.kill();
      this.addScore(scoring.fragment, h.x, h.y);
      this.explode(h.x, h.y, 0xfa500e, 8);
      return;
    }
    this.hitPlayer();
  }

  /** Returns true if damage was applied. */
  hitPlayer(): boolean {
    if (this.player.dead || this.over) return false;
    if (this.player.shielded) {
      this.sfx('crack');
      this.burst.setParticleTint(0x00df6a);
      this.burst.explode(10, this.player.x, this.player.y);
      return false;
    }
    if (this.player.invulnerable) return false;
    this.hearts--;
    this.hitsTaken++;
    this.comboCount = 0;
    this.bonusMult = 0;
    this.refreshCombo(false);
    this.player.invulnUntil = this.time.now + P.invulnSeconds * 1000;
    this.player.hitSquash();
    this.hud.setHearts(this.hearts);
    this.hud.hitFlash();
    this.sfx('hit');
    vibrate([40, 30, 40]);
    this.explode(this.player.x, this.player.y, 0xfa500e, 16);
    if (!save.data.settings.reduceMotion) {
      this.frozenUntil = this.time.now + feel.hitStopMs;
      this.physics.pause();
      this.time.delayedCall(feel.hitStopMs, () => {
        if (!this.over && !this.paused) this.physics.resume();
      });
      this.cameras.main.shake(feel.shakeDuration, feel.shakeHit);
    }
    if (this.hearts <= 0) this.gameOver();
    else this.radioSay(this.hearts === 1 ? S.radio.lastHeart : S.radio.hit);
    return true;
  }

  private collect(k: Pickup): void {
    if (!k.active) return;
    k.kill();
    const mult = this.comboMult;
    switch (k.kind) {
      case 'coin':
      case 'diamond': {
        const v = k.kind === 'coin' ? 1 : economy.diamondValue;
        this.coins += v;
        this.comboCount++;
        this.comboIdle = 0;
        this.refreshCombo(true);
        this.player.turboCharge = Math.min(1, this.player.turboCharge + (T.chargePerCoin[save.data.upgrades.turbo] ?? 0.04) * v);
        if (this.player.turboCharge >= 1 && !this.player.turboActive && !this.turboTold) {
          this.turboTold = true;
          this.radioSay(S.radio.turboReady);
        }
        this.hud.setCoins(this.coins, true);
        this.addScore(k.kind === 'coin' ? scoring.coin : scoring.diamond, k.x, k.y, k.kind === 'diamond' ? '+10' : undefined, k.kind === 'diamond' ? '#8CD9FF' : undefined);
        this.sfx(k.kind);
        break;
      }
      case 'heart':
        this.hearts = Math.min(P.hearts, this.hearts + 1);
        this.hud.setHearts(this.hearts);
        this.sfx('power');
        this.explode(k.x, k.y, 0x00df6a, 14);
        break;
      case 'star': {
        const i = k.getData('index') as number;
        this.hiddenFound[i] = true;
        this.addScore(scoring.secretStar, k.x, k.y, S.float.star, '#FFEC00');
        this.explode(k.x, k.y, 0xffec00, 24);
        this.sfx('star');
        break;
      }
      case 'shield':
        this.player.shieldUntil = this.time.now + (PW.shieldByLevel[save.data.upgrades.shield] ?? 8) * 1000;
        this.radioSay(S.radio.powerShield);
        this.powerFx(k, S.float.shield, '#00DF6A');
        break;
      case 'magnet':
        this.player.magnetUntil = this.time.now + PW.durationSeconds * 1000;
        this.radioSay(S.radio.powerMagnet);
        this.powerFx(k, S.float.magnet, '#EB459F');
        break;
      case 'triple':
        this.player.tripleUntil = this.time.now + PW.durationSeconds * 1000;
        this.radioSay(S.radio.powerTriple);
        this.powerFx(k, S.float.triple, '#F876FF');
        break;
    }
    if (mult < this.comboMult) {
      if (this.comboMult === 3) this.radioSay(S.radio.combo3);
      if (this.comboMult === 5) this.radioSay(S.radio.combo5);
    }
  }

  private powerFx(k: Pickup, label: string, color: string): void {
    this.sfx('power');
    this.explode(k.x, k.y, parseInt(color.slice(1), 16), 16);
    this.floatText(this.player.x, this.player.y - 90, label, color);
  }

  private passGate(g: Gate): void {
    if (g.passed || !g.active) return;
    g.passed = true;
    this.bonusMult = Math.min(C.maxMult - 1, this.bonusMult + 1);
    this.comboIdle = 0;
    this.refreshCombo(true);
    this.addScore(scoring.gate, g.x, g.y - 40, S.float.gate, '#8CD9FF');
    this.sfx('gate');
    this.explode(g.x, g.y, 0x8cd9ff, 14);
    this.tweens.add({ targets: g, alpha: 0.3, duration: 200 });
    this.tweens.add({ targets: this.player, scaleX: 0.85, scaleY: 1.18, duration: 110, yoyo: true });
  }

  // ---------- boss flow ----------
  private startBoss(): void {
    this.bossStarted = true;
    this.bossHeartsAtStart = this.hearts;
    this.checkpoint = { score: this.score, coins: this.coins };
    this.radioSay(S.radio.checkpoint, true);
    this.hud.banner(this.bossName);
    this.hud.storm(false);
    this.sfx('alarm');
    this.time.delayedCall(1800, () => {
      if (this.over) return;
      this.createBoss();
      this.hud.showBoss(this.bossName);
    });
  }

  private bossDefeated(): void {
    this.bossDone = true;
    this.hud.hideBoss();
    this.addScore(scoring.bossKill, GAME_W / 2, 400);
    this.sfx('victory');
    vibrate([60, 40, 60, 40, 120]);
    if (!save.data.settings.reduceMotion) this.cameras.main.shake(400, feel.shakeBoss);
    const rain = this.stageId === 'mars' ? 30 : 40;
    for (let i = 0; i < rain; i++) {
      this.time.delayedCall(i * 60, () => this.spawnPickup(i % 9 === 0 ? 'diamond' : 'coin', Phaser.Math.Between(60, GAME_W - 60), -40, Phaser.Math.Between(380, 520)));
    }
    this.player.magnetUntil = this.time.now + 6000;
    this.time.delayedCall(4200, () => this.stageClear());
  }

  private stageClear(): void {
    if (this.over) return;
    this.over = true;
    this.hud.banner(S.results.stageClear, true);
    const bonus = this.hearts * scoring.heartLeftBonus;
    this.score += bonus;
    this.tweens.add({ targets: this.player, y: -200, duration: 1100, ease: 'Quad.easeIn', delay: 600 });
    this.time.delayedCall(1900, () => this.finish(true));
  }

  private gameOver(): void {
    this.over = true;
    this.player.die();
    this.explode(this.player.x, this.player.y, 0xfa500e, 40);
    this.explode(this.player.x, this.player.y, 0xffec00, 30);
    this.tweens.add({ targets: this.player, alpha: 0, scale: 0.4, angle: 90, duration: 600 });
    const key = `${this.stageId}:${this.segmentKey()}`;
    save.data.losses[key] = (save.data.losses[key] ?? 0) + 1;
    save.write();
    this.time.delayedCall(1300, () => this.finish(false));
  }

  private finish(won: boolean): void {
    const data: ResultData = {
      stageId: this.stageId,
      won,
      score: this.score,
      coins: this.coins,
      maxCombo: this.maxCombo,
      nearMisses: this.nearMisses,
      hitsTaken: this.hitsTaken,
      hiddenFound: [...this.hiddenFound],
      bossPerfect: won && this.hearts === this.bossHeartsAtStart,
      canCheckpoint: !won && this.bossStarted && !this.challenge,
      checkpoint: this.checkpoint,
      challenge: this.challenge,
      elapsed: this.elapsed,
    };
    this.scene.start('Results', data);
  }

  // ---------- feedback helpers ----------
  addScore(base: number, x?: number, y?: number, label?: string, color?: string): void {
    const pts = Math.round(base * this.comboMult);
    this.score += pts;
    this.hud.setScore(this.score);
    if (x !== undefined && y !== undefined) this.floatText(x, y, label ? `${label} +${pts}` : `+${pts}`, color ?? (this.comboMult > 1 ? '#F876FF' : '#FFFFFF'));
  }

  floatText(x: number, y: number, text: string, color: string): void {
    const rtl = /[؀-ۿ]/.test(text);
    const t = this.add
      .text(x, y, text, { fontFamily: 'MA, sans-serif', fontSize: '30px', fontStyle: '900', color, stroke: '#101340', strokeThickness: 6, rtl })
      .setOrigin(0.5)
      .setDepth(40);
    this.tweens.add({ targets: t, y: y - 70, alpha: 0, duration: 800, ease: 'Quad.easeOut', onComplete: () => t.destroy() });
  }

  explode(x: number, y: number, tint: number, count: number, texture = 'particle'): void {
    if (texture !== 'particle') {
      const e = this.add.particles(x, y, texture, { speed: { min: 100, max: 400 }, lifespan: 500, scale: { start: 1.5, end: 0 }, tint, emitting: false }).setDepth(25);
      e.explode(count);
      this.time.delayedCall(600, () => e.destroy());
      return;
    }
    this.burst.setParticleTint(tint);
    this.burst.explode(this.lowQuality ? Math.ceil(count / 2) : count, x, y);
  }

  sfx(name: Sfx): void {
    sfx.play(name);
  }

  radioSay(msg: string, force = false): void {
    this.radio.say(msg, force);
  }
}
