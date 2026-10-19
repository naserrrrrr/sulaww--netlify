// HTML for every DOM screen. Scenes call showScreen(...) and bind buttons with on(...).
import { S } from '../config/strings.ar';
import { images, logos } from '../assets';
import { economy, stars as ST } from '../config/balance';
import { save, UpgradeId, AchievementId, StageId } from '../systems/save';
import { rocketSkins, RocketSkin } from '../config/palette';
import { fmt, showScreen, on, hideScreen, toast } from './dom';
import { sfx } from '../systems/audio';

const esc = (s: string) => s.replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c] as string);
const coinsPill = () => `<div class="coins-pill"><img src="${images['icon-coin']}" alt="">${fmt(save.data.coins)}</div>`;
const starsHtml = (n: number) => `<div class="stars">${[0, 1, 2].map((i) => `<i class="${i < n ? 'on' : ''}">★</i>`).join('')}</div>`;

export function menuScreen(): string {
  const d = save.data;
  const best = Math.max(d.best.mars, d.best.saturn);
  return `<div class="screen center">
    <img class="logo-hero" src="${logos.stackedWhite}" alt="ريان وبيان">
    <h1 class="title">${S.title}</h1>
    <p class="sub">${S.subtitle}</p>
    <div class="stack">
      <button class="btn big yellow" data-act="play">${d.unlocked.saturn ? S.continueGame : S.play}</button>
      <div class="row">
        <button class="btn" data-act="hangar"><img src="${images['icon-rocket']}" alt="">${S.hangar}</button>
        <button class="btn" data-act="ach">${S.achievements}</button>
      </div>
      <div class="row">
        <button class="btn ghost small" data-act="settings">${S.settings}</button>
        ${d.unlocked.challenge ? `<button class="btn ghost small" data-act="challenge">${S.challenge}</button>` : ''}
      </div>
    </div>
    ${best ? `<p class="hint">${S.bestScore}: <b>${fmt(best)}</b></p>` : ''}
    ${coinsPill()}
  </div>`;
}

export function storyScreen(i: number): string {
  const arts = [
    `<img class="robot" src="${images['robot-yellow']}" alt="" style="height:85%">`,
    `<img src="${images['rayan-point']}" alt="" style="height:90%"><img src="${images['bayan-open']}" alt="" style="height:85%">`,
    `<img src="${images['planet-mars']}" alt="" style="height:45%"><img src="${images['sticker-rocket']}" alt="" style="height:60%"><img src="${images['planet-saturn']}" alt="" style="height:50%">`,
  ];
  return `<div class="screen center">
    <div class="story-card">
      <div class="story-art">${arts[i]}</div>
      <p>${S.story[i]}</p>
      <div class="dots">${S.story.map((_, k) => `<i class="${k === i ? 'on' : ''}"></i>`).join('')}</div>
    </div>
    <div class="row">
      <button class="btn yellow" data-act="next">${i === S.story.length - 1 ? S.play : S.next}</button>
      <button class="btn ghost small" data-act="skip">${S.skip}</button>
    </div>
  </div>`;
}

export function mapScreen(): string {
  const d = save.data;
  const planet = (id: StageId, img: string, locked: boolean) => `
    <button class="planet ${locked ? 'locked' : 'active'} silent" data-stage="${id}" ${locked ? 'disabled' : ''}>
      <img src="${img}" alt="">
      <div class="info">
        <span class="name">${S.stages[id].name}</span>
        <span class="desc">${locked ? S.locked : S.stages[id].desc}</span>
        ${starsHtml(d.stars[id])}
        ${d.best[id] ? `<span class="best">${S.bestScore}: ${fmt(d.best[id])}</span>` : ''}
      </div>
    </button>`;
  return `<div class="screen scroll">
    <div class="topbar"><button class="btn ghost small" data-act="back">${S.back}</button><h2 class="title">${S.mapTitle}</h2>${coinsPill()}</div>
    <div class="map">
      ${planet('mars', images['planet-mars'], false)}
      ${planet('saturn', images['planet-saturn'], !d.unlocked.saturn)}
    </div>
    <button class="btn ghost" data-act="hangar"><img src="${images['icon-rocket']}" alt="">${S.hangar}</button>
  </div>`;
}

export function pauseScreen(): string {
  return `<div class="screen center">
    <h2 class="title">${S.paused}</h2>
    <div class="stack">
      <button class="btn big green" data-act="resume">${S.resume}</button>
      <button class="btn" data-act="restart">${S.restart}</button>
      <button class="btn ghost" data-act="settings">${S.settings}</button>
      <button class="btn ghost" data-act="menu">${S.toMenu}</button>
    </div>
  </div>`;
}

export interface ResultView {
  stageId: StageId;
  won: boolean;
  score: number;
  coins: number;
  maxCombo: number;
  nearMisses: number;
  stars: [boolean, boolean, boolean];
  newBest: boolean;
  canCheckpoint: boolean;
  challenge: boolean;
  nextStage: StageId | null;
}

export function resultsScreen(r: ResultView): string {
  const stage = S.stages[r.stageId].name;
  const threshold = r.stageId === 'mars' ? ST.marsScore : ST.saturnScore;
  return `<div class="screen center">
    <h2 class="title">${r.challenge ? S.challengeUI.title : r.won ? S.results.stageClear : S.results.gameOver}</h2>
    <p class="sub">${stage}</p>
    ${r.challenge ? '' : `<div class="big-stars">${r.stars.map((s, i) => `<div class="s ${s ? 'on' : ''}" style="animation-delay:${0.2 + i * 0.25}s">★</div>`).join('')}</div>
    <div class="star-legend">
      <span class="${r.stars[0] ? 'on' : ''}">★ ${S.results.starClear}</span>
      <span class="${r.stars[1] ? 'on' : ''}">★★ ${S.results.starScore} (${fmt(threshold)})</span>
      <span class="${r.stars[2] ? 'on' : ''}">★★★ ${S.results.starHidden}</span>
    </div>`}
    ${r.newBest ? `<div class="newbest">${S.results.newBest}</div>` : ''}
    <div class="result-box">
      <div class="line"><span>${S.results.coins}</span><b>+${fmt(r.coins)}</b></div>
      <div class="line"><span>${S.results.maxCombo}</span><b>×${r.maxCombo}</b></div>
      <div class="line"><span>${S.results.nearMisses}</span><b>${fmt(r.nearMisses)}</b></div>
      <div class="line total"><span>${S.results.score}</span><b>${fmt(r.score)}</b></div>
    </div>
    <div class="stack">
      ${r.won && r.nextStage ? `<button class="btn big yellow" data-act="next">${S.stages[r.nextStage].name} ←</button>` : ''}
      ${!r.won && r.canCheckpoint ? `<button class="btn big yellow" data-act="checkpoint">${S.results.retryFromCheckpoint}</button>` : ''}
      <div class="row">
        <button class="btn ${!r.won && !r.canCheckpoint ? 'yellow' : ''}" data-act="retry">${S.restart}</button>
        ${r.challenge ? '' : `<button class="btn" data-act="hangar"><img src="${images['icon-rocket']}" alt="">${S.hangar}</button>`}
      </div>
      <button class="btn ghost small" data-act="map">${r.challenge ? S.toMenu : S.toMap}</button>
    </div>
  </div>`;
}

export function victoryScreen(score: number): string {
  return `<div class="screen center">
    <h1 class="title">${S.victory.title}</h1>
    <div class="lottie-box" id="celebrateAnim"></div>
    <div class="story-card">
      <div class="story-art" style="height:calc(var(--px) * 240)"><img src="${images['robot-yellow']}" alt="" style="height:100%"><img src="${images['rayan-wave']}" alt="" style="height:90%"><img src="${images['bayan-walk']}" alt="" style="height:85%"></div>
      <p>${S.victory.text}</p>
      <p style="color:var(--c-yellow)">${S.victory.thanks}</p>
    </div>
    <p class="sub">${S.victory.finalScore}: <b>${fmt(score)}</b></p>
    <button class="btn big yellow" data-act="next">${S.next}</button>
  </div>`;
}

export function hangarScreen(): string {
  const d = save.data;
  const items: UpgradeId[] = ['laser', 'shield', 'magnet', 'turbo'];
  const icons: Record<UpgradeId, string> = { laser: '✦', shield: '◯', magnet: '∪', turbo: '➤' };
  const cards = items
    .map((id) => {
      const lvl = d.upgrades[id];
      const max = lvl >= economy.upgradeCosts.length;
      const cost = max ? 0 : economy.upgradeCosts[lvl];
      return `<div class="card">
        <div class="ic">${icons[id]}</div>
        <div class="txt"><span class="name">${S.hangarUI.items[id].name}</span><span class="desc">${S.hangarUI.items[id].desc}</span>
          <div class="pips">${[0, 1, 2].map((i) => `<i class="${i < lvl ? 'on' : ''}"></i>`).join('')}</div></div>
        <button class="btn small ${max ? 'ghost' : d.coins >= cost ? 'green' : 'ghost'}" data-up="${id}" ${max ? 'disabled' : ''}>${max ? S.hangarUI.max : `${S.hangarUI.buy} · ${cost}`}</button>
      </div>`;
    })
    .join('');
  const skins = rocketSkins
    .map((sk) => {
      const owned = d.skins.includes(sk);
      const onCls = d.skin === sk ? 'on' : '';
      return `<button class="skin ${onCls} silent" data-skin="${sk}">
        <div class="sw" style="background-image:url('${images[`rocket-${sk}` as keyof typeof images]}')"></div>
        <span class="lbl">${S.hangarUI.skinNames[sk]}</span>
        <span class="cost">${owned ? (d.skin === sk ? S.hangarUI.equipped : S.hangarUI.equip) : `${economy.skinCost} ◉`}</span>
      </button>`;
    })
    .join('');
  return `<div class="screen scroll">
    <div class="topbar"><button class="btn ghost small" data-act="back">${S.back}</button><h2 class="title">${S.hangarUI.title}</h2>${coinsPill()}</div>
    <h2 class="sub" style="align-self:flex-start">${S.hangarUI.upgrades}</h2>
    <div class="cards">${cards}</div>
    <h2 class="sub" style="align-self:flex-start">${S.hangarUI.skins}</h2>
    <div class="skins">${skins}</div>
  </div>`;
}

export function achievementsScreen(): string {
  const ids = Object.keys(S.achievementsUI.list) as AchievementId[];
  const d = save.data;
  return `<div class="screen scroll">
    <div class="topbar"><button class="btn ghost small" data-act="back">${S.back}</button><h2 class="title">${S.achievementsUI.title}</h2><span class="best">${ids.filter((i) => d.achievements[i]).length}/${ids.length}</span></div>
    ${ids
      .map((id) => {
        const onCls = d.achievements[id] ? 'on' : '';
        return `<div class="ach ${onCls}"><div class="medal">${onCls ? '✓' : '★'}</div><div class="txt"><div class="name">${S.achievementsUI.list[id].name}</div><div class="desc">${S.achievementsUI.list[id].desc}</div></div></div>`;
      })
      .join('')}
  </div>`;
}

export function settingsScreen(): string {
  const s = save.data.settings;
  const tog = (id: string, label: string, onv: boolean) => `<div class="setting"><span>${label}</span><button class="toggle ${onv ? 'on' : ''} silent" data-tog="${id}" aria-pressed="${onv}" aria-label="${label}"></button></div>`;
  return `<div class="screen scroll">
    <div class="topbar"><button class="btn ghost small" data-act="back">${S.back}</button><h2 class="title">${S.settingsUI.title}</h2><span></span></div>
    ${tog('sound', S.settingsUI.sound, s.sound)}
    <div class="setting"><span>${S.settingsUI.volume}</span><input id="vol" type="range" min="0" max="100" value="${Math.round(s.volume * 100)}" aria-label="${S.settingsUI.volume}"></div>
    ${tog('vibrate', S.settingsUI.vibrate, s.vibrate)}
    ${tog('reduceMotion', S.settingsUI.reduceMotion, s.reduceMotion)}
    <div class="setting"><span>${S.settingsUI.quality}</span><div class="seg">${(['auto', 'low', 'high'] as const).map((q) => `<button class="${s.quality === q ? 'on' : ''} silent" data-q="${q}">${S.settingsUI.qualities[q]}</button>`).join('')}</div></div>
    <div class="danger-box" id="resetBox"><button class="btn orange small" data-act="reset">${S.settingsUI.reset}</button></div>
  </div>`;
}

export function challengeScreen(): string {
  const top = save.data.challengeTop;
  return `<div class="screen center">
    <h2 class="title">${S.challengeUI.title}</h2>
    <p class="sub">${S.challengeUI.desc}</p>
    <div class="result-box">
      <div class="line"><span>${S.challengeUI.top5}</span></div>
      ${top.length ? top.map((v, i) => `<div class="line"><span>${i + 1}.</span><b>${fmt(v)}</b></div>`).join('') : `<div class="line"><span>—</span></div>`}
    </div>
    <button class="btn big yellow" data-act="start">${S.challengeUI.start}</button>
    <button class="btn ghost small" data-act="back">${S.back}</button>
  </div>`;
}

/** Settings & achievements are DOM-only overlays usable from menu and pause. */
export function openSettings(onBack: () => void): void {
  const root = showScreen(settingsScreen());
  const s = save.data.settings;
  on(root, '[data-tog]', (el) => {
    const id = el.dataset.tog as 'sound' | 'vibrate' | 'reduceMotion';
    s[id] = !s[id];
    el.classList.toggle('on', s[id]);
    el.setAttribute('aria-pressed', String(s[id]));
    if (id === 'reduceMotion') document.documentElement.classList.toggle('reduce-motion', s.reduceMotion);
    save.write();
    sfx.applySettings();
    sfx.play('click');
  });
  root.querySelector<HTMLInputElement>('#vol')?.addEventListener('input', (e) => {
    s.volume = Number((e.target as HTMLInputElement).value) / 100;
    save.write();
    sfx.applySettings();
    sfx.play('coin');
  });
  on(root, '[data-q]', (el) => {
    s.quality = el.dataset.q as 'auto' | 'low' | 'high';
    root.querySelectorAll('[data-q]').forEach((b) => b.classList.toggle('on', b === el));
    save.write();
  });
  on(root, '[data-act=reset]', () => {
    const box = root.querySelector('#resetBox')!;
    box.innerHTML = `<span>${S.settingsUI.resetConfirm}</span><div class="row"><button class="btn orange small" data-act="reset-yes">${S.settingsUI.yes}</button><button class="btn ghost small" data-act="reset-no">${S.settingsUI.no}</button></div>`;
    on(box, '[data-act=reset-yes]', () => {
      save.reset();
      hideScreen();
      onBack();
    });
    on(box, '[data-act=reset-no]', () => openSettings(onBack));
  });
  on(root, '[data-act=back]', () => {
    hideScreen();
    onBack();
  });
}

export function openAchievements(onBack: () => void): void {
  const root = showScreen(achievementsScreen());
  on(root, '[data-act=back]', () => {
    hideScreen();
    onBack();
  });
}

export function bindHangar(onBack: () => void, rerender: () => void): void {
  const root = showScreen(hangarScreen());
  const d = save.data;
  on(root, '[data-up]', (el) => {
    const id = el.dataset.up as UpgradeId;
    const lvl = d.upgrades[id];
    const cost = economy.upgradeCosts[lvl];
    if (cost === undefined) return;
    if (d.coins < cost) {
      toast(S.hangarUI.notEnough);
      return;
    }
    d.coins -= cost;
    d.upgrades[id] = lvl + 1;
    save.write();
    sfx.play('power');
    rerender();
  });
  on(root, '[data-skin]', (el) => {
    const sk = el.dataset.skin as RocketSkin;
    if (!d.skins.includes(sk)) {
      if (d.coins < economy.skinCost) {
        toast(S.hangarUI.notEnough);
        return;
      }
      d.coins -= economy.skinCost;
      d.skins.push(sk);
      sfx.play('power');
    } else sfx.play('click');
    d.skin = sk;
    save.write();
    rerender();
  });
  on(root, '[data-act=back]', () => {
    hideScreen();
    onBack();
  });
}

export const logoImg = (): string => `<img class="logo-corner" src="${logos.horizontalWhite}" alt="ريان وبيان">`;
export { esc };
