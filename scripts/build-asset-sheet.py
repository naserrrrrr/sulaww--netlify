#!/usr/bin/env python3
"""Build docs/asset-sheet.html from public/assets/brand/manifest.json.

Usage:  python3 scripts/build-asset-sheet.py [--prefix PATH] [--out FILE]
Each asset is shown on a light and a dark tile with its name and source.
"""
import argparse
import html
import json
import os

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
CAT_LABELS = {
    "character": "الشخصيات",
    "avatar": "الأفاتار",
    "robot": "الروبوتات",
    "planet": "الكواكب",
    "icon": "الأيقونات",
    "sticker": "الملصقات",
    "spot": "رسوم صغيرة",
    "shape": "أشكال",
    "background": "خلفيات الفضاء",
    "logo": "الشعار",
    "lottie": "حركة Lottie",
    "font": "الخط",
}
PALETTE = [
    ("primary", "#5143DE", "Majorelle Blue", "الأزرار والعناصر الرئيسية"),
    ("deep", "#101340", "Deadly Depths", "خلفية الفضاء والألواح الداكنة"),
    ("orange", "#FA500E", "Orioles Orange", "الخطر والزعماء ولهب الدفع"),
    ("yellow", "#FFEC00", "Middle Yellow", "العملات والنجوم"),
    ("green", "#00DF6A", "Tealish Green", "النجاح والدرع والقلوب"),
    ("rose", "#EB459F", "Rose Bonbon", "القوى الخاصة"),
    ("sky", "#8CD9FF", "Kul Sharif Blue", "إبراز الواجهة والجليد"),
    ("flamingo", "#F876FF", "Electric Flamingo", "التوهج والكومبو"),
    ("sugar", "#FFD7FF", "Sugarpills", "خلفيات ناعمة"),
    ("water", "#D5F7F7", "Water", "الألواح الفاتحة والنص على الداكن"),
]


def tile(a, prefix):
    name = html.escape(a["name"])
    src = html.escape(a["source"])
    note = html.escape(a.get("note", ""))
    size = "×".join(str(v) for v in a.get("size", [])) if a.get("size") else ""
    files = " · ".join(html.escape(f) for f in a["files"])
    if a["category"] == "lottie":
        url = prefix + a["files"][0]
        media = "".join(
            f'<div class="pane {tone}"><div class="lottie" data-src="{html.escape(url)}"></div></div>' for tone in ("light", "dark")
        )
    elif a["category"] == "font":
        fam = "brand-" + a["name"]
        url = prefix + a["files"][0]
        media = (
            f'<style>@font-face{{font-family:"{fam}";src:url("{html.escape(url)}") format("woff2")}}</style>'
            + "".join(
                f'<div class="pane {tone} font" style="font-family:\'{fam}\'">مهمة المجرّة 0123456789</div>' for tone in ("light", "dark")
            )
        )
    else:
        img = next((f for f in a["files"] if f.endswith((".png", ".jpg"))), a["files"][0])
        url = html.escape(prefix + img)
        media = "".join(
            f'<div class="pane {tone}"><img decoding="async" src="{url}" alt="{name}"></div>' for tone in ("light", "dark")
        )
    return f"""<figure class="asset" data-cat="{a['category']}">
  <div class="panes">{media}</div>
  <figcaption><code dir="ltr">{name}</code><span class="meta">{note}</span>
  <span class="src" dir="ltr">{src}{(' · ' + size) if size else ''}</span><span class="files" dir="ltr">{files}</span></figcaption>
</figure>"""


def build(prefix, font_prefix, lottie_src, plan_html=""):
    m = json.load(open(os.path.join(ROOT, "public/assets/brand/manifest.json"), encoding="utf-8"))
    assets = m["assets"]
    cats = [c for c in CAT_LABELS if any(a["category"] == c for a in assets)]
    chips = '<button class="chip on" data-f="all">الكل <b>%d</b></button>' % len(assets) + "".join(
        f'<button class="chip" data-f="{c}">{CAT_LABELS[c]} <b>{sum(a["category"] == c for a in assets)}</b></button>' for c in cats
    )
    sections = ""
    for c in cats:
        items = "\n".join(tile(a, prefix) for a in assets if a["category"] == c)
        sections += f'<section data-cat="{c}"><h2>{CAT_LABELS[c]}</h2><div class="grid">{items}</div></section>'
    swatches = "".join(
        f'<div class="sw"><span style="background:{hx}"></span><code dir="ltr">{tok}</code><b dir="ltr">{hx}</b><small>{nm}</small><em>{use}</em></div>'
        for tok, hx, nm, use in PALETTE
    )
    excluded = "".join(f'<li><code dir="ltr">{html.escape(e["source"])}</code> — {html.escape(e["reason"])}</li>' for e in m["excluded"])
    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Galaxy Mission Assets</title>
<meta name="claude:viewport-fit-sides" content="cover">
<style>
@font-face{{font-family:"MA";src:url("{font_prefix}fonts/montserrat-arabic-regular.woff2") format("woff2");font-weight:400}}
@font-face{{font-family:"MA";src:url("{font_prefix}fonts/montserrat-arabic-bold.woff2") format("woff2");font-weight:700}}
@font-face{{font-family:"MA";src:url("{font_prefix}fonts/montserrat-arabic-black.woff2") format("woff2");font-weight:900}}
:root{{--bg:#F4F2FF;--fg:#101340;--card:#fff;--muted:#5b5f8a;--line:#e1ddff;--primary:#5143DE;--deep:#101340;--water:#D5F7F7}}
@media (prefers-color-scheme:dark){{:root:not([data-theme="light"]){{--bg:#0b0d2e;--fg:#D5F7F7;--card:#151a52;--muted:#9fa6d9;--line:#262c6e}}}}
:root[data-theme="dark"]{{--bg:#0b0d2e;--fg:#D5F7F7;--card:#151a52;--muted:#9fa6d9;--line:#262c6e}}
*{{box-sizing:border-box}}
body{{margin:0;background:var(--bg);color:var(--fg);font:15px/1.6 "MA",system-ui,sans-serif}}
header{{background:var(--deep);color:#D5F7F7;padding-block:28px 22px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px));position:relative;overflow:hidden}}
header .in{{max-width:1240px;margin:auto;display:flex;gap:20px;align-items:center;justify-content:space-between;flex-wrap:wrap}}
header h1{{margin:0;font-size:clamp(24px,4vw,38px);font-weight:900;letter-spacing:0}}
header p{{margin:4px 0 0;color:#8CD9FF}}
header img.logo{{height:56px;padding:14px}}
main{{max-width:1240px;margin:auto;padding-block:16px;padding-left:max(16px,env(safe-area-inset-left,0px));padding-right:max(16px,env(safe-area-inset-right,0px))}}
.plan{{background:var(--card);border:1px solid var(--line);border-radius:18px;padding:6px 20px 16px;margin-top:18px}}
.plan h1{{font-size:24px;font-weight:900}} .plan h2{{font-size:19px;margin:22px 0 8px}}
.plan .tw{{overflow-x:auto}} .plan table{{border-collapse:collapse;width:100%;font-size:14px;min-width:560px}}
.plan th,.plan td{{border-bottom:1px solid var(--line);padding:8px 10px;text-align:start;vertical-align:top}}
.plan th{{color:var(--primary);font-weight:700}} .plan code{{font-size:12px;background:var(--bg);padding:1px 5px;border-radius:6px}}
.chips{{top:env(safe-area-inset-top,0px)}}
.stats{{display:flex;gap:10px;flex-wrap:wrap;margin:14px 0}}
.stats div{{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:10px 16px}}
.stats b{{font-size:22px;font-weight:900;color:var(--primary);display:block}}
.chips{{display:flex;gap:8px;flex-wrap:wrap;position:sticky;top:0;background:var(--bg);padding:10px 0;z-index:3}}
.chip{{font:inherit;border:1px solid var(--line);background:var(--card);color:var(--fg);border-radius:999px;padding:8px 14px;min-height:40px;cursor:pointer}}
.chip.on{{background:var(--primary);border-color:var(--primary);color:#fff}}
.chip b{{opacity:.7;font-weight:400}}
h2{{font-weight:900;font-size:22px;margin:28px 0 12px}}
.grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:14px}}
.asset{{margin:0;background:var(--card);border:1px solid var(--line);border-radius:18px;overflow:hidden}}
.panes{{display:grid;grid-template-columns:1fr 1fr}}
.pane{{aspect-ratio:1;display:flex;align-items:center;justify-content:center;padding:12px}}
.pane.light{{background:#fff;background-image:linear-gradient(45deg,#f1f1f6 25%,transparent 25%,transparent 75%,#f1f1f6 75%),linear-gradient(45deg,#f1f1f6 25%,transparent 25%,transparent 75%,#f1f1f6 75%);background-size:16px 16px;background-position:0 0,8px 8px}}
.pane.dark{{background:#101340}}
.pane img{{max-width:100%;max-height:100%;object-fit:contain}}
.pane .lottie{{width:100%;height:100%}}
.pane.font{{font-size:17px;text-align:center;font-weight:700}}
.pane.light.font{{color:#101340}} .pane.dark.font{{color:#D5F7F7}}
section[data-cat="background"] .grid{{grid-template-columns:repeat(auto-fill,minmax(320px,1fr))}}
section[data-cat="background"] .panes{{grid-template-columns:1fr}} section[data-cat="background"] .pane.light{{display:none}}
section[data-cat="background"] .pane{{aspect-ratio:auto;padding:0}}
figcaption{{padding:10px 12px 12px;display:flex;flex-direction:column;gap:2px}}
figcaption code{{font-weight:700;font-size:14px;align-self:flex-start}}
.meta{{color:var(--fg)}} .src,.files{{color:var(--muted);font-size:12px;align-self:flex-start}}
.palette{{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px}}
.sw{{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:10px;display:grid;grid-template-columns:52px 1fr;gap:0 10px;align-items:center}}
.sw span{{grid-row:span 4;width:52px;height:52px;border-radius:14px;border:1px solid rgba(0,0,0,.08)}}
.sw small,.sw em{{color:var(--muted);font-style:normal;font-size:12px}}
.ex{{background:var(--card);border:1px solid var(--line);border-inline-start:5px solid #FA500E;border-radius:14px;padding:12px 18px}}
.ex li{{margin:4px 0}}
footer{{color:var(--muted);text-align:center;padding:30px 16px}}
</style>
</head>
<body>
<header><div class="in"><div><h1>ورقة مراجعة أصول «مهمة المجرّة»</h1><p>كل الأصول مستخرجة من الهوية الرسمية لريان وبيان — كل أصل على خلفية فاتحة وأخرى داكنة</p></div>
<img class="logo" src="{prefix}logo-horizontal-white.svg" alt="ريان وبيان"></div></header>
<main>
<div class="stats"><div><b>{len(assets)}</b>أصلاً جاهزاً</div><div><b>{sum(a['category'] == 'lottie' for a in assets)}</b>حركات Lottie</div>
<div><b>{len(m['excluded'])}</b>عناصر مستبعدة عمداً</div><div><b>0</b>بدائل مؤقتة</div></div>
{plan_html}
<h2>لوحة الألوان (Tokens)</h2><div class="palette">{swatches}</div>
<h2>مستبعد من اللعب</h2><ul class="ex">{excluded}</ul>
<div class="chips">{chips}</div>
{sections}
</main>
<footer>يُولَّد تلقائياً عبر <code dir="ltr">scripts/build-asset-sheet.py</code> من <code dir="ltr">manifest.json</code></footer>
<script src="{lottie_src}"></script>
<script>
document.querySelectorAll('.chip').forEach(function(c){{c.addEventListener('click',function(){{
  document.querySelectorAll('.chip').forEach(function(x){{x.classList.toggle('on',x===c)}});
  var f=c.dataset.f;document.querySelectorAll('main section').forEach(function(s){{s.hidden=!(f==='all'||s.dataset.cat===f)}});
}})}});
if(window.lottie){{document.querySelectorAll('.lottie').forEach(function(el){{
  lottie.loadAnimation({{container:el,renderer:'svg',loop:true,autoplay:true,path:el.dataset.src}});
}})}}
</script>
</body>
</html>
"""


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--prefix", default="../public/assets/brand/")
    ap.add_argument("--out", default=os.path.join(ROOT, "docs", "asset-sheet.html"))
    ap.add_argument("--lottie", default="https://cdnjs.cloudflare.com/ajax/libs/lottie-web/5.12.2/lottie.min.js")
    ap.add_argument("--artifact", action="store_true", help="omit document skeleton, embed docs/game-plan.md")
    args = ap.parse_args()
    plan = ""
    if args.artifact:
        import markdown
        md = open(os.path.join(ROOT, "docs", "game-plan.md"), encoding="utf-8").read()
        body = markdown.markdown(md, extensions=["tables"]).replace("<table>", '<div class="tw"><table>').replace("</table>", "</table></div>")
        plan = f'<section class="plan" id="plan">{body}</section>'
    out = build(args.prefix, args.prefix, args.lottie, plan)
    if args.artifact:
        import re
        out = re.sub(r"<!doctype html>\s*<html[^>]*>\s*<head>\s*<meta charset=\"utf-8\">\s*<meta name=\"viewport\"[^>]*>\s*", "", out)
        out = out.replace("</head>\n<body>\n", "").replace("</body>\n</html>\n", "")
    os.makedirs(os.path.dirname(args.out), exist_ok=True)
    open(args.out, "w", encoding="utf-8").write(out)
    print("wrote", os.path.relpath(args.out, ROOT))
