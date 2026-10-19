// Every brand asset the game uses, imported so Vite can inline them in the single-file build.
import rocketBlue from '@brand/rocket-blue.png';
import rocketOrange from '@brand/rocket-orange.png';
import rocketGreen from '@brand/rocket-green.png';
import rocketRose from '@brand/rocket-rose.png';
import iconCoin from '@brand/icon-coin.png';
import iconDiamond from '@brand/icon-diamond.png';
import iconGift from '@brand/icon-gift.png';
import iconArrow from '@brand/icon-arrow.png';
import iconRocket from '@brand/icon-rocket.png';
import iconNotification from '@brand/icon-notification.png';
import planetMars from '@brand/planet-mars.png';
import planetSaturn from '@brand/planet-saturn.png';
import planetMoon from '@brand/planet-moon.png';
import planetEarth from '@brand/planet-earth.png';
import robotYellow from '@brand/robot-yellow.png';
import star6 from '@brand/star-6pt.png';
import burstYellow from '@brand/burst-yellow.png';
import avatarBayan from '@brand/avatar-bayan.png';
import avatarRayan from '@brand/avatar-rayan.png';
import rayanPoint from '@brand/rayan-pose-point.png';
import rayanWave from '@brand/rayan-pose-wave.png';
import bayanOpen from '@brand/bayan-pose-open.png';
import bayanWalk from '@brand/bayan-pose-walk.png';
import stickerRocket from '@brand/sticker-rocket.png';
import logoHorizontalWhite from '@brand/logo-horizontal-white.svg';
import logoStackedWhite from '@brand/logo-stacked-white.svg';
import logoMarkWhite from '@brand/logo-mark-white.svg';
import lottieLogo from '@brand/lottie-logo-intro.json';
import lottieCelebrate from '@brand/lottie-celebrate.json';
import lottieRayanFly from '@brand/lottie-rayan-fly-2.json';
import lottieBayanFly from '@brand/lottie-bayan-fly-2.json';
import lottieStar from '@brand/lottie-star-1.json';

export const images = {
  'rocket-blue': rocketBlue,
  'rocket-orange': rocketOrange,
  'rocket-green': rocketGreen,
  'rocket-rose': rocketRose,
  'icon-coin': iconCoin,
  'icon-diamond': iconDiamond,
  'icon-gift': iconGift,
  'icon-arrow': iconArrow,
  'icon-rocket': iconRocket,
  'icon-notification': iconNotification,
  'planet-mars': planetMars,
  'planet-saturn': planetSaturn,
  'planet-moon': planetMoon,
  'planet-earth': planetEarth,
  'robot-yellow': robotYellow,
  'star-6pt': star6,
  'burst-yellow': burstYellow,
  'avatar-bayan': avatarBayan,
  'avatar-rayan': avatarRayan,
  'rayan-point': rayanPoint,
  'rayan-wave': rayanWave,
  'bayan-open': bayanOpen,
  'bayan-walk': bayanWalk,
  'sticker-rocket': stickerRocket,
} as const;

export const rocketSheets = ['rocket-blue', 'rocket-orange', 'rocket-green', 'rocket-rose'] as const;
export const ROCKET_FRAME = 180;
export const ROCKET_FRAMES = 12;

export const logos = { horizontalWhite: logoHorizontalWhite, stackedWhite: logoStackedWhite, markWhite: logoMarkWhite };
export const lotties = { logo: lottieLogo, celebrate: lottieCelebrate, rayanFly: lottieRayanFly, bayanFly: lottieBayanFly, star: lottieStar };
