// Brand colour tokens (Rayan w Bayan guideline, "Primary / Secondary colors").
export const palette = {
  primary: '#5143DE', // Majorelle Blue - buttons, main elements
  deep: '#101340', // Deadly Depths - space background, dark panels
  orange: '#FA500E', // Orioles Orange - danger, bosses, thrust flame
  yellow: '#FFEC00', // Middle Yellow - coins, stars
  green: '#00DF6A', // Tealish Green - success, shield, hearts
  rose: '#EB459F', // Rose Bonbon - power-ups
  sky: '#8CD9FF', // Kul Sharif Blue - UI highlight, ice
  flamingo: '#F876FF', // Electric Flamingo - glow, combo
  sugar: '#FFD7FF', // Sugarpills - soft backgrounds
  water: '#D5F7F7', // Water - light panels, text on dark
} as const;

export type PaletteKey = keyof typeof palette;

export const hex = (key: PaletteKey): number => parseInt(palette[key].slice(1), 16);

export const cssVars = (): string =>
  Object.entries(palette)
    .map(([k, v]) => `--c-${k}:${v};`)
    .join('');

// Rocket hull colours available in the hangar (pre-rendered spritesheets).
export const rocketSkins = ['blue', 'orange', 'green', 'rose'] as const;
export type RocketSkin = (typeof rocketSkins)[number];
