# مهمة المجرّة · Galaxy Mission

لعبة أركيد فضائية (تمرير عمودي) بهوية تطبيق «ريان وبيان». Vite + TypeScript + Phaser 3.90.

## التشغيل

```bash
cd game
npm install
npm run dev            # http://localhost:5173
npm run build          # dist/ (ملفات متعددة)
npm run build:single   # dist-single/index.html ثم تُنسخ إلى ../galaxy/index.html (ملف واحد ≈ 3.1MB)
npm run render:rocket  # يعيد توليد Spritesheet الصاروخ من ملف Lottie الرسمي (يحتاج Playwright/Chromium)
```

الأصول الرسمية في `../public/assets/brand/` تُولَّد بـ `python3 ../scripts/extract-assets.py` من مجلد `../brand/` (غير مرفوع في Git).

## وضع التطوير

| الرابط | الأثر |
|---|---|
| `?debug=1` | FPS وصناديق الاصطدام، ومفاتيح `V` (قفز إلى الزعيم / هزيمته فوراً) و`K` (خسارة فورية) |
| `?stage=1` أو `?stage=2` | بدء المرحلة مباشرة |
| `?stage=2&boss=1` | القفز إلى الزعيم (نقطة الحفظ) |
| `?stage=1&challenge=1` | وضع التحدي |

## الهيكل

- `src/config/`: `balance.ts` (كل أرقام التوازن)، `palette.ts` (ألوان الهوية)، `strings.ar.ts` (كل النصوص).
- `src/scenes/`: `Boot`، `Background`، `Menus` (Menu/Story/Map/Hangar)، `Play` (الحلقة الأساسية)، `StageMars`، `StageSaturn`، `Results`.
- `src/entities/`: `Player`، `Meteor`، `Bullet`، `Pickup`، `Hazard`، `Gate`، `BossRock`، `BossRobot`.
- `src/systems/`: `audio` (مؤثرات Web Audio مولّدة)، `save` (localStorage بمفتاح `rb-galaxy-save-v1`)، `textures` (أشكال مرسومة برمجياً).
- `src/ui/`: طبقة DOM (القوائم، HUD، صندوق رسائل بيان) و`ui.css`.
- `scripts/render-lottie.mjs`: صاروخ اللاعب من Lottie الرسمي بأربعة ألوان.
