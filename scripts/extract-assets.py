#!/usr/bin/env python3
"""Extract game-ready brand assets from ./brand into public/assets/brand.

Usage:  python3 scripts/extract-assets.py
Needs:  pymupdf, pillow, fonttools, brotli   (pip install pymupdf pillow fonttools brotli)

Sources (official Rayan w Bayan delivery folder):
  - Guideline/PDF/Guideline.pdf  (full version, vector)  -> characters, avatars,
    planets, icons, stickers, spot illustrations, space backgrounds
  - Logo/*/SVG                    -> logos, copied unmodified (only the canvas is trimmed)
  - App Animation/*/data.json     -> Lottie animations, copied unmodified
  - Fonts/Montserrat-Arabic-*.ttf -> converted to woff2
Every illustration that contains Quranic text or surah names is skipped on purpose.
"""
import json
import os
import re
import shutil
import sys

import pymupdf as fitz
from PIL import Image

sys.path.insert(0, os.path.dirname(__file__))
from extract_lib import cluster, isolate, page_drawings, render_png, write_svg  # noqa: E402

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
BRAND = os.path.join(ROOT, "brand")
OUT = os.path.join(ROOT, "public", "assets", "brand")
GUIDE = os.path.join(BRAND, "Guideline", "PDF", "Guideline.pdf")

# (page, cluster index, name, category, png max side, drop first N drawings, note)
VECTOR_SPEC = [
    (26, 0, "rayan-pose-front", "character", 1024, 0, "Rayan (boy), front"),
    (26, 1, "rayan-pose-point", "character", 1024, 0, "Rayan pointing, headset"),
    (26, 2, "rayan-pose-side", "character", 1024, 0, "Rayan side view"),
    (26, 3, "rayan-pose-wave", "character", 1024, 0, "Rayan waving"),
    (27, 0, "bayan-pose-front", "character", 1024, 0, "Bayan (girl), front"),
    (27, 1, "bayan-pose-open", "character", 1024, 0, "Bayan arms open"),
    (27, 2, "bayan-pose-side", "character", 1024, 0, "Bayan side view"),
    (27, 3, "bayan-pose-walk", "character", 1024, 0, "Bayan walking"),
    (32, 0, "avatar-rayan", "avatar", 512, 0, "Rayan helmet avatar"),
    (32, 1, "avatar-bayan", "avatar", 512, 0, "Bayan helmet avatar (flower)"),
    (32, 2, "avatar-robot-pink", "avatar", 512, 0, "Robot head avatar"),
    (32, 3, "avatar-robot-orange", "avatar", 512, 0, "Robot head avatar, antennas"),
    (32, 4, "robot-sky", "robot", 512, 2, "Robot, background circle removed"),
    (32, 5, "robot-yellow", "robot", 512, 2, "Robot, background circle removed"),
    (32, 6, "robot-purple", "robot", 512, 2, "Robot, background circle removed"),
    (32, 7, "robot-orange", "robot", 512, 2, "Robot, background circle removed"),
    (33, 0, "planet-saturn", "planet", 512, 0, "Saturn"),
    (33, 1, "planet-earth", "planet", 512, 0, "Earth"),
    (33, 2, "planet-jupiter", "planet", 512, 0, "Jupiter"),
    (33, 3, "planet-uranus", "planet", 512, 0, "Uranus"),
    (33, 4, "planet-neptune", "planet", 512, 0, "Neptune"),
    (33, 5, "planet-venus", "planet", 512, 0, "Venus"),
    (33, 6, "planet-moon", "planet", 512, 0, "Moon"),
    (33, 7, "planet-mars", "planet", 512, 0, "Mars"),
    (34, 0, "icon-coin", "icon", 256, 0, "Coins"),
    (34, 1, "icon-notification", "icon", 256, 0, "Notification"),
    (34, 2, "icon-arrow", "icon", 256, 0, "Arrow"),
    (34, 3, "icon-diamond", "icon", 256, 0, "Diamond"),
    (34, 4, "icon-message", "icon", 256, 0, "Messages"),
    (34, 5, "icon-gift", "icon", 256, 0, "Gift"),
    (34, 6, "icon-rocket", "icon", 512, 0, "Rocket"),
    (34, 7, "icon-error", "icon", 256, 0, "Error"),
    (29, 1, "sticker-laptop", "sticker", 512, 0, "Character with laptop"),
    (29, 2, "sticker-bayan-gifts", "sticker", 512, 0, "Bayan celebrating with gifts"),
    (29, 3, "sticker-robot-couch", "sticker", 512, 0, "Robot on couch"),
    (29, 5, "sticker-robot-laptop", "sticker", 512, 0, "Robot on laptop"),
    (29, 10, "sticker-rocket", "sticker", 512, 0, "Rocket with flame"),
    (29, 4, "burst-green", "shape", 128, 0, "Sparkle burst"),
    (29, 8, "burst-purple", "shape", 128, 0, "Sparkle burst"),
    (29, 9, "burst-yellow", "shape", 128, 0, "Sparkle burst"),
    (29, 11, "burst-orange", "shape", 128, 0, "Sparkle burst"),
    (31, 1, "spot-plant", "spot", 512, 0, "Plant"),
    (31, 2, "spot-monitor", "spot", 512, 0, "Monitor"),
    (31, 4, "spot-moon", "spot", 512, 0, "Moon"),
    (31, 6, "spot-saturn", "spot", 512, 0, "Saturn with star"),
]

# Kept out of the game on purpose (Quranic text / surah names, or reading scenes).
EXCLUDED = [
    ("Guideline p31 cluster 5", "Open book with surah names (Al-Asr, Al-Fatiha, Al-Baqara, Ya-Sin, Al-Kawthar)"),
    ("Guideline p29 cluster 6", "Character reading a book (content unverifiable)"),
    ("App Animation/مشهد القمر", "Rayan reading a book on the moon (content unverifiable)"),
    ("App Animation/ترديد الآيات", "Quran recitation scene"),
    ("App Animation/تسميع القرآن", "Quran memorisation scene"),
]

LOTTIE_SPEC = [
    ("مشهد الصاروخ/data.json", "lottie-rocket.json", "Rocket with Rayan & Bayan, flame loop"),
    ("ريان/01/data.json", "lottie-rayan-fly-1.json", "Rayan flying, pointing"),
    ("ريان/02/data.json", "lottie-rayan-fly-2.json", "Rayan flying, arms open"),
    ("بيان/1/data.json", "lottie-bayan-fly-1.json", "Bayan flying"),
    ("بيان/2/data.json", "lottie-bayan-fly-2.json", "Bayan flying, waving"),
    ("النجوم/01/data.json", "lottie-star-1.json", "Smiling star"),
    ("النجوم/02/data.json", "lottie-star-2.json", "Smiling star"),
    ("ممتاز/data.json", "lottie-celebrate.json", "Bayan celebrating with gifts"),
    ("Logo splash/لوجو .json", "lottie-logo-intro.json", "Official logo intro scene"),
]

FONT_WEIGHTS = ["Regular", "Medium", "SemiBold", "Bold", "ExtraBold", "Black"]

manifest = {"generatedFrom": "brand/ (Rayan w Bayan official delivery)", "assets": [], "excluded": []}


def add(name, category, files, source, note="", **extra):
    entry = {"name": name, "category": category, "files": files, "source": source, "note": note}
    entry.update(extra)
    manifest["assets"].append(entry)


def nfc(path):
    """Drive delivers Arabic names in NFD on some systems; match regardless of form."""
    import unicodedata
    head, tail = os.path.split(path)
    if os.path.exists(path):
        return path
    for cand in os.listdir(head) if os.path.isdir(head) else []:
        if unicodedata.normalize("NFC", cand) == unicodedata.normalize("NFC", tail):
            return os.path.join(head, cand)
    return path


def resolve(rel):
    parts = rel.split("/")
    cur = os.path.join(BRAND, "App Animation")
    for p in parts:
        cur = nfc(os.path.join(cur, p))
    return cur


def extract_vectors(doc):
    cache = {}
    for page_no, idx, name, cat, px, drop, note in VECTOR_SPEC:
        if page_no not in cache:
            cache[page_no] = [c for c in cluster(page_drawings(doc[page_no - 1])) if c[0].width > 12 and c[0].height > 12]
        bb, group = cache[page_no][idx]
        group = group[drop:]
        if drop:
            bb = fitz.Rect(group[0]["rect"])
            for d in group[1:]:
                bb |= d["rect"]
        sub = isolate(group, bb)
        png = f"{name}.png"
        w, h = render_png(sub, os.path.join(OUT, png), px)
        write_svg(sub, os.path.join(OUT, f"{name}.svg"))
        add(name, cat, [png, f"{name}.svg"], f"Guideline.pdf p{page_no}", note, size=[w, h])


def extract_backgrounds(doc):
    page = doc[27]
    deep = (0.063, 0.075, 0.251)
    panels = [d["rect"] for d in page.get_drawings()
              if d.get("fill") and all(abs(a - b) < 0.01 for a, b in zip(d["fill"], deep)) and d["rect"].width > 400]
    panels = sorted({(round(r.x0), round(r.y0), round(r.x1), round(r.y1)) for r in panels}, key=lambda r: (r[1], r[0]))
    for i, r in enumerate(panels[:4], 1):
        clip = fitz.Rect(*r)
        zoom = 1600 / clip.width
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
        name = f"space-bg-{i}"
        pix.pil_save(os.path.join(OUT, f"{name}.jpg"), quality=88)
        add(name, "background", [f"{name}.jpg"], "Guideline.pdf p28", "Official space illustration panel", size=[pix.width, pix.height])
    # One 6-point star from the same page, as a reusable vector sprite.
    stars = [d for d in page.get_drawings() if len(d["items"]) == 7 and d.get("fill") and 6 < d["rect"].width < 30]
    if stars:
        s = max(stars, key=lambda d: d["rect"].width)
        s = dict(s, fill=(1, 1, 1))  # white so it can be tinted with palette colours in-game
        sub = isolate([s], s["rect"], margin=0.5)
        render_png(sub, os.path.join(OUT, "star-6pt.png"), 64)
        write_svg(sub, os.path.join(OUT, "star-6pt.svg"))
        add("star-6pt", "shape", ["star-6pt.png", "star-6pt.svg"], "Guideline.pdf p28", "Background star shape, white for tinting")


def svg_fill(path):
    m = re.search(r"fill:\s*(#[0-9a-fA-F]{3,6})\b", open(path, encoding="utf-8").read())
    return m.group(1).lower() if m else None


def trim_svg(src, dst, png_src):
    """Copy an official SVG, only tightening its viewBox to the drawn area (paths untouched)."""
    svg = open(src, encoding="utf-8").read()
    vb = [float(v) for v in re.search(r'viewBox="([^"]+)"', svg).group(1).split()]
    im = Image.open(png_src).convert("RGBA")
    l, t, r, b = im.getchannel("A").getbbox()
    sx, sy = vb[2] / im.width, vb[3] / im.height
    nvb = f"{vb[0] + l * sx:.1f} {vb[1] + t * sy:.1f} {(r - l) * sx:.1f} {(b - t) * sy:.1f}"
    svg = svg.replace(re.search(r'viewBox="[^"]+"', svg).group(0), f'viewBox="{nvb}"', 1)
    open(dst, "w", encoding="utf-8").write(svg)
    im.crop((l, t, r, b)).save(dst[:-4] + ".png", optimize=True)


def extract_logos():
    for folder, key in [("Horizontal Logo", "horizontal"), ("Stacked Logo", "stacked"), ("Mark", "mark")]:
        svgdir = os.path.join(BRAND, "Logo", folder, "SVG")
        for fn in sorted(os.listdir(svgdir)):
            src = os.path.join(svgdir, fn)
            png = os.path.join(BRAND, "Logo", folder, "PNG", fn[:-4] + ".png")
            if not os.path.exists(png):
                continue
            # The transparent artboards are the ones whose PNG has transparent corners.
            if Image.open(png).convert("RGBA").getpixel((2, 2))[3] != 0:
                continue
            tone = "white" if svg_fill(src) in ("#fff", "#fdfdfd", "#ffffff", "#fefefe") else "navy"
            name = f"logo-{key}-{tone}"
            trim_svg(src, os.path.join(OUT, f"{name}.svg"), png)
            add(name, "logo", [f"{name}.svg", f"{name}.png"], f"Logo/{folder}/SVG/{fn}",
                "Official logo, unmodified (canvas trimmed only)")


def copy_lottie():
    for rel, dst, note in LOTTIE_SPEC:
        src = resolve(rel)
        data = json.load(open(src, encoding="utf-8"))
        json.dump(data, open(os.path.join(OUT, dst), "w", encoding="utf-8"), separators=(",", ":"), ensure_ascii=False)
        add(dst[:-5], "lottie", [dst], f"App Animation/{rel}", note,
            size=[data["w"], data["h"]], seconds=round((data["op"] - data["ip"]) / data["fr"], 2))


def convert_fonts():
    from fontTools.ttLib import TTFont
    fdir = os.path.join(OUT, "fonts")
    os.makedirs(fdir, exist_ok=True)
    for w in FONT_WEIGHTS:
        src = os.path.join(BRAND, "Fonts", f"Montserrat-Arabic-{w}.ttf")
        if not os.path.exists(src):
            continue
        f = TTFont(src)
        f.flavor = "woff2"
        f.save(os.path.join(fdir, f"montserrat-arabic-{w.lower()}.woff2"))
        add(f"font-{w.lower()}", "font", [f"fonts/montserrat-arabic-{w.lower()}.woff2"], f"Fonts/Montserrat-Arabic-{w}.ttf",
            "Official brand typeface (Arabic + Latin)")


def main():
    if not os.path.isdir(BRAND):
        sys.exit("brand/ folder not found - copy the official delivery folder there first.")
    if os.path.isdir(OUT):
        shutil.rmtree(OUT)
    os.makedirs(OUT)
    doc = fitz.open(GUIDE)
    extract_vectors(doc)
    extract_backgrounds(doc)
    extract_logos()
    copy_lottie()
    convert_fonts()
    manifest["excluded"] = [{"source": s, "reason": r} for s, r in EXCLUDED]
    json.dump(manifest, open(os.path.join(OUT, "manifest.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print(f"{len(manifest['assets'])} assets -> {os.path.relpath(OUT, ROOT)}")


if __name__ == "__main__":
    main()
