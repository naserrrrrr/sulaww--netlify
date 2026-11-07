"""Shared helpers for the asset-portfolio-analysis skill.

Every script in this folder imports from here so that number parsing, date
parsing, Arabic name normalisation and status vocabulary behave identically
everywhere. Keep this module dependency-light (stdlib + optional hijridate).
"""
from __future__ import annotations

import datetime as _dt
import hashlib
import json
import os
import re
import unicodedata
from typing import Any, Optional, Tuple

CONTRACT_VERSION = "1.0"
HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)

# ---------------------------------------------------------------------------
# Status vocabulary (see references/data-contract.md §2)
# ---------------------------------------------------------------------------
S_VALUE = "value"                  # a usable value read from the file
S_DERIVED = "derived"              # computed from other cells of the same record (formula noted)
S_ZERO_RECORDED = "zero_recorded"  # a 0 written in the file (not verified as a true zero)
S_ZERO_CONFIRMED = "zero_confirmed"  # a 0 confirmed by explicit evidence (mapping override)
S_ZERO_AS_MISSING = "zero_as_missing"  # a 0 in a field where 0 is not a plausible value (e.g. valuation)
S_BLANK = "blank"                  # empty cell
S_FILLER = "filler"                # whitespace / NBSP-only cell
S_TEXT = "text"                    # non-numeric text in a numeric field (raw kept)
S_ERROR = "error"                  # Excel error (#REF!, #VALUE!, ...)
S_NOT_APPLICABLE = "not_applicable"
S_NOT_AVAILABLE = "not_available"  # explicit "غير متوفر" / "-" / N/A

KNOWN_STATUSES = frozenset({S_VALUE, S_DERIVED, S_ZERO_RECORDED, S_ZERO_CONFIRMED})
ALL_STATUSES = frozenset({
    S_VALUE, S_DERIVED, S_ZERO_RECORDED, S_ZERO_CONFIRMED, S_ZERO_AS_MISSING,
    S_BLANK, S_FILLER, S_TEXT, S_ERROR, S_NOT_APPLICABLE, S_NOT_AVAILABLE,
})

RECORD_TYPES = ("asset", "land_parcel", "unit", "contract", "operating_activity", "total", "other")
SCOPES = ("in_scope", "isolated", "excluded")
USES = ("investment", "service", "admin", "development", "unknown")
INCOME_NATURES = ("rent", "operating", "other", "none", "unknown")
OWNERSHIP_KINDS = ("owned", "partial", "usufruct", "leased_land_project", "managed_for_others", "unknown")
WAQF_VALUES = ("yes", "no", "unknown")

EXCEL_ERRORS = ("#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")

NA_TOKENS = {"غير متوفر", "غير متاح", "غير متوفره", "لا يوجد", "n/a", "na", "-", "—", "–", "؟", "?", "غير معروف"}
NOT_APPLICABLE_TOKENS = {"لا ينطبق", "غير منطبق", "لاينطبق"}

_AR_INDIC = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
_INVISIBLES = " ​‌‍‎‏⁦⁧⁨⁩‪‫‬‭‮﻿؜"
_INVIS_RE = re.compile("[" + _INVISIBLES + "]")
_DIACRITICS_RE = re.compile("[ؐ-ًؚ-ٰٟۖ-ۭ]")
_CURRENCY_RE = re.compile(r"(ريال(?:\s*سعودي)?|ر\.?\s*س\.?|رس|sar|sr|﷼)", re.IGNORECASE)


# ---------------------------------------------------------------------------
# JSON / file helpers
# ---------------------------------------------------------------------------
def load_json(path: str) -> Any:
    with open(path, "r", encoding="utf-8") as fh:
        return json.load(fh)


def save_json(path: str, data: Any) -> None:
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2, default=_json_default)
        fh.write("\n")


def _json_default(o: Any) -> Any:
    if isinstance(o, (_dt.date, _dt.datetime)):
        return o.isoformat()
    if isinstance(o, set):
        return sorted(o)
    raise TypeError(f"not JSON serialisable: {type(o)}")


def sha256_file(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def load_vocab() -> dict:
    return load_json(os.path.join(SKILL_DIR, "assets", "vocab.json"))


# ---------------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------------
def clean_text(v: Any) -> str:
    """Trim, drop invisible marks (NBSP, RLM, ...) and collapse whitespace. Keeps letters as-is."""
    if v is None:
        return ""
    s = _INVIS_RE.sub(" ", str(v))
    s = re.sub(r"\s+", " ", s).strip()
    return s


def is_filler(v: Any) -> bool:
    """True for a string that contains only whitespace / NBSP / bidi marks (not for None)."""
    return isinstance(v, str) and clean_text(v) == ""


def normalize_arabic(s: Any) -> str:
    """Normalise Arabic text for comparison only (never for display)."""
    s = clean_text(s)
    s = unicodedata.normalize("NFKC", s)
    s = _DIACRITICS_RE.sub("", s)
    s = s.replace("ـ", "")
    s = re.sub("[أإآٱ]", "ا", s)
    s = s.replace("ى", "ي").replace("ة", "ه").replace("ؤ", "و").replace("ئ", "ي")
    s = s.translate(_AR_INDIC)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def name_key(s: Any) -> str:
    """Looser key used ONLY to flag possibly-duplicate names for human review. Never merges records."""
    s = normalize_arabic(s)
    s = re.sub(r"[()\[\]{}.,،:;\-_/\\|\"'«»]", " ", s)
    tokens = [t for t in s.split() if t not in {"رقم", "حي", "بحي", "في", "ب"}]
    return " ".join(tokens)


def extract_years(s: Any) -> list:
    """Gregorian (19xx/20xx) and Hijri (13xx/14xx) four-digit years mentioned in a header."""
    t = clean_text(s).translate(_AR_INDIC)
    return [int(y) for y in re.findall(r"(?<!\d)(19\d{2}|20\d{2}|13\d{2}|14\d{2})(?!\d)", t)]


def word_count_ar(text: str) -> int:
    return len([w for w in re.split(r"\s+", clean_text(text)) if re.search(r"[\w؀-ۿ]", w)])


# ---------------------------------------------------------------------------
# Number parsing
# ---------------------------------------------------------------------------
def parse_number(v: Any) -> Tuple[Optional[float], str, Optional[str]]:
    """Return (number_or_None, status, note).

    Never turns unknown into zero. A numeric 0 returns status zero_recorded; the caller
    applies the column's zero policy (e.g. valuation 0 -> zero_as_missing).
    """
    if v is None:
        return None, S_BLANK, None
    if isinstance(v, bool):
        return None, S_TEXT, "قيمة منطقية في حقل رقمي"
    if isinstance(v, (int, float)):
        if v != v:  # NaN
            return None, S_ERROR, "NaN"
        return (float(v), S_ZERO_RECORDED if v == 0 else S_VALUE, None)
    if isinstance(v, (_dt.date, _dt.datetime)):
        return None, S_TEXT, "تاريخ في حقل رقمي"
    s = str(v)
    if is_filler(s):
        return None, S_FILLER, None
    c = clean_text(s)
    up = c.upper()
    if up in EXCEL_ERRORS or up.startswith("#REF"):
        return None, S_ERROR, c
    low = normalize_arabic(c)
    if low in {normalize_arabic(t) for t in NOT_APPLICABLE_TOKENS}:
        return None, S_NOT_APPLICABLE, c
    if low in {normalize_arabic(t) for t in NA_TOKENS}:
        return None, S_NOT_AVAILABLE, c
    t = c.translate(_AR_INDIC)
    t = _CURRENCY_RE.sub("", t)
    neg = False
    if re.fullmatch(r"\(\s*[\d.,٫٬\s]+\s*\)", t.strip()):
        neg = True
        t = t.strip()[1:-1]
    t = t.replace("٫", ".").replace("٬", ",").replace("،", ",")
    pct = t.strip().endswith("%")
    t = t.replace("%", "")
    t = re.sub(r"(?<=\d)[,\s](?=\d{3}(\D|$))", "", t).strip()
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", t):
        num = float(t)
        if neg:
            num = -num
        note = "نص رقمي حُوِّل إلى رقم" + (" (نسبة مئوية)" if pct else "")
        return num, (S_ZERO_RECORDED if num == 0 else S_VALUE), note
    return None, S_TEXT, c


# ---------------------------------------------------------------------------
# Date parsing (Gregorian / Hijri Umm al-Qura)
# ---------------------------------------------------------------------------
def _hijri_to_greg(y: int, m: int, d: int) -> Tuple[Optional[str], Optional[str]]:
    try:
        from hijridate import Hijri  # type: ignore
    except Exception:  # library missing
        return None, "مكتبة hijridate غير مثبتة؛ حُفظ التاريخ الهجري دون تحويل"
    try:
        return Hijri(y, m, d).to_gregorian().isoformat(), None
    except Exception as exc:  # out of range / invalid day (e.g. 30 in a 29-day month)
        try:
            # tolerate day overflow by clamping to the month length, but say so
            h = Hijri(y, m, 1)
            last = h.month_length()
            if d > last:
                return Hijri(y, m, last).to_gregorian().isoformat(), f"اليوم {d} يتجاوز طول الشهر ({last})؛ استُخدم آخر يوم"
        except Exception:
            pass
        return None, f"تاريخ هجري غير صالح: {exc}"


def parse_date(v: Any, calendar_hint: str = "auto") -> dict:
    """Return {"v": iso_gregorian|None, "s": status, "raw": str, "calendar": ..., "note": ...}.

    calendar_hint: auto | hijri | gregorian. The original text is always kept in raw.
    """
    out = {"v": None, "s": S_BLANK, "raw": None if v is None else str(v), "calendar": None, "note": None}
    if v is None:
        return out
    if isinstance(v, _dt.datetime):
        out.update(v=v.date().isoformat(), s=S_VALUE, calendar="gregorian")
        if calendar_hint == "hijri":
            out["note"] = "خلية تاريخ Excel (ميلادي) في عمود يُتوقع أن يكون هجريًا — تحقّق"
        return out
    if isinstance(v, _dt.date):
        out.update(v=v.isoformat(), s=S_VALUE, calendar="gregorian")
        return out
    if isinstance(v, (int, float)):
        out.update(s=S_TEXT, note="رقم في حقل تاريخ؛ لم يُحوَّل")
        return out
    s = str(v)
    if is_filler(s):
        out["s"] = S_FILLER
        return out
    c = clean_text(s).translate(_AR_INDIC)
    is_hijri_marked = bool(re.search(r"هـ|ه\b|هجري", c))
    nums = [int(x) for x in re.findall(r"\d+", c)]
    if len(nums) >= 3:
        a, b, d3 = nums[0], nums[1], nums[2]
        if a >= 1000:
            y, m, d = a, b, d3
        elif d3 >= 1000:
            y, m, d = d3, b, a
        else:
            out.update(s=S_TEXT, note="صيغة تاريخ غير محددة")
            return out
        hijri = is_hijri_marked or calendar_hint == "hijri" or (1300 <= y <= 1500 and calendar_hint != "gregorian")
        if hijri:
            out["calendar"] = "hijri"
            g, note = _hijri_to_greg(y, m, d)
            out.update(v=g, s=S_VALUE if g else S_TEXT, note=note)
            out["hijri"] = f"{y:04d}-{m:02d}-{d:02d}"
            return out
        try:
            out.update(v=_dt.date(y, m, d).isoformat(), s=S_VALUE, calendar="gregorian")
        except ValueError as exc:
            out.update(s=S_TEXT, note=f"تاريخ ميلادي غير صالح: {exc}")
        return out
    if len(nums) == 1 and 1300 <= nums[0] <= 1500:
        out.update(s=S_TEXT, calendar="hijri", note="سنة هجرية فقط")
        return out
    if len(nums) == 1 and 1900 <= nums[0] <= 2100:
        out.update(s=S_TEXT, calendar="gregorian", note="سنة ميلادية فقط")
        return out
    out.update(s=S_TEXT, note="نص غير قابل للتحويل إلى تاريخ")
    return out


# ---------------------------------------------------------------------------
# Personal-data detection (never rendered in the dashboard)
# ---------------------------------------------------------------------------
PII_HEADER_RE = re.compile(r"(جوال|هاتف|تواصل|اتصال|موبايل|ايميل|إيميل|بريد|mobile|phone|email|اسم\s*المشرف|المشرف|المسؤول\s*عن|هوية|سجل\s*مدني|iban|ايبان|آيبان)", re.IGNORECASE)
PHONE_RE = re.compile(r"(?<!\d)(?:\+?966|00966|0)?5\d{8}(?!\d)")
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")


def looks_like_phone(v: Any) -> bool:
    if v is None:
        return False
    s = clean_text(v).translate(_AR_INDIC).replace(" ", "").replace("-", "")
    if isinstance(v, float) and v.is_integer():
        s = str(int(v))
    return bool(PHONE_RE.fullmatch(s))


def ref(sheet: str, coord: str) -> str:
    """Canonical cell reference: exact sheet name (may contain trailing spaces) + '!' + A1 coordinate."""
    return f"{sheet}!{coord}"
