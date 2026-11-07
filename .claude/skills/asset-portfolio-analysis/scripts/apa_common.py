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
AMOUNT_KINDS = ("money", "area", "count")   # a '%' or a «12.500» here is never read as a plain decimal
AMBIGUOUS_THOUSANDS_NOTE = ("فاصل ملتبس: «{raw}» قد تكون نقطته فاصل آلاف ({alt:,}) أو فاصلة عشرية ({dec})؛ "
                            "لم تُحوَّل إلى رقم")
FOREIGN_CURRENCY_NOTE = "مبلغ بعملة أخرى ({cur}) في حقل بعملة {home}؛ لم يُحوَّل إلى رقم ولم يُجمع"
_FOREIGN_CURRENCY_RE = re.compile(
    r"(?<![A-Za-z])(USD|EUR|GBP|AED|KWD|QAR|BHD|OMR|EGP|JOD|US\$)(?![A-Za-z])|(\$|€|£)|(دولار|يورو|جنيه|درهم|دينار)",
    re.IGNORECASE)
_HOME_CURRENCY_TOKENS = {"SAR": ("SAR", "SR", "ريال", "﷼")}
_AREA_UNIT_RE = re.compile(r"(?:(?:m|M)\s*(?:2|²)|sq\.?\s*m|sqm|م\s*(?:2|²)|م\.?\s*م|(?:ال)?متر(?:\s*(?:ال)?مربع)?|متر\s*م|m\^2)\.?",
                           re.IGNORECASE)


def foreign_currency(v: Any, home: str = "SAR") -> Optional[str]:
    """The currency token of a text amount written in a currency other than the mapping's («USD 100,000», «$ 5,000»,
    «3000 دولار»), or None. Numbers are never in a foreign currency (the cell format is not read)."""
    if not isinstance(v, str):
        return None
    t = clean_text(v)
    m = _FOREIGN_CURRENCY_RE.search(t)
    if not m:
        return None
    tok = m.group(0)
    if tok.upper() == (home or "").upper():
        return None
    return tok


def parse_number(v: Any, kind: Optional[str] = None) -> Tuple[Optional[float], str, Optional[str]]:
    """Return (number_or_None, status, note).

    Never turns unknown into zero. A numeric 0 returns status zero_recorded; the caller
    applies the column's zero policy (e.g. valuation 0 -> zero_as_missing).
    kind (optional, the mapping column kind): for an amount kind (money/area/count) a text such as «12.500» — one dot
    followed by exactly three digits — is ambiguous (thousands separator or decimal) and stays text with
    AMBIGUOUS_THOUSANDS_NOTE; for kind area a trailing/leading square-metre unit («500 m2»، «٨٠٠ م²») is removed and noted.
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
    unit_note = None
    if kind == "area":
        t2 = _AREA_UNIT_RE.sub("", t)
        if t2 != t and re.search(r"\d", t2):
            t, unit_note = t2, "حُذفت وحدة المتر المربع من النص"
    neg = False
    if re.fullmatch(r"\(\s*[\d.,٫٬\s]+\s*\)", t.strip()):
        neg = True
        t = t.strip()[1:-1]
    t = t.replace("٫", ".").replace("٬", ",").replace("،", ",").replace("\u2212", "-")
    pct = t.strip().endswith("%")
    t = t.replace("%", "")
    if kind in AMOUNT_KINDS and not pct:
        m = re.fullmatch(r"\s*([-+]?)(\d{1,3})\.(\d{3})\s*", t)
        if m:
            alt = int(m.group(2) + m.group(3))
            dec = float(m.group(2) + "." + m.group(3))
            return None, S_TEXT, AMBIGUOUS_THOUSANDS_NOTE.format(raw=c, alt=alt, dec=f"{dec:g}")
    t = re.sub(r"(?<=\d)[,\s](?=\d{3}(\D|$))", "", t).strip()
    if re.fullmatch(r"[-+]?\d+(\.\d+)?", t):
        num = float(t)
        if neg:
            num = -num
        note = "نص رقمي حُوِّل إلى رقم" + (" (نسبة مئوية)" if pct else "") + (f"؛ {unit_note}" if unit_note else "")
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


AMBIGUOUS_DATE_NOTE = "تاريخ ملتبس: اليوم والشهر ≤ 12؛ قُرئ يوم/شهر"


def _join_notes(*notes) -> Optional[str]:
    parts = [n for n in notes if n]
    return "؛ ".join(parts) if parts else None


def parse_date(v: Any, calendar_hint: str = "auto") -> dict:
    """Return {"v": iso_gregorian|None, "s": status, "raw": str, "calendar": ..., "note": ...}.

    calendar_hint: auto | hijri | gregorian. The original text is always kept in raw.
    Text in day/month/year order is read as day/month; when both day and month are ≤ 12 (and differ) the note
    carries AMBIGUOUS_DATE_NOTE so the reading is checked against the source.
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
        ambiguous = None
        if a >= 1000:
            y, m, d = a, b, d3
        elif d3 >= 1000:
            y, m, d = d3, b, a
            if a <= 12 and b <= 12 and a != b:
                # «03/04/2021» may be 3 April (day/month) or 4 March (month/day): read as day/month, but say so
                ambiguous = AMBIGUOUS_DATE_NOTE
        else:
            out.update(s=S_TEXT, note="صيغة تاريخ غير محددة")
            return out
        hijri = is_hijri_marked or calendar_hint == "hijri" or (1300 <= y <= 1500 and calendar_hint != "gregorian")
        if hijri:
            out["calendar"] = "hijri"
            g, note = _hijri_to_greg(y, m, d)
            out.update(v=g, s=S_VALUE if g else S_TEXT, note=_join_notes(ambiguous, note))
            out["hijri"] = f"{y:04d}-{m:02d}-{d:02d}"
            return out
        try:
            out.update(v=_dt.date(y, m, d).isoformat(), s=S_VALUE, calendar="gregorian", note=ambiguous)
        except ValueError as exc:
            out.update(s=S_TEXT, note=_join_notes(ambiguous, f"تاريخ ميلادي غير صالح: {exc}"))
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
# Patterns live in assets/pii_patterns.json (shared with scripts/check_dashboard.mjs). Callers must translate
# Arabic-Indic digits to ASCII before matching values (pii_value_hits / mask_pii do it for you).
def _load_pii_patterns() -> dict:
    with open(os.path.join(SKILL_DIR, "assets", "pii_patterns.json"), encoding="utf-8") as fh:
        return json.load(fh)


def _flags(f: str) -> int:
    return re.IGNORECASE if "i" in (f or "") else 0


PII_PATTERNS = _load_pii_patterns()
REDACTED = PII_PATTERNS.get("redaction", "<محجوب>")
PII_HEADER_RE = re.compile("(" + PII_PATTERNS["header"]["pattern"] + ")", _flags(PII_PATTERNS["header"].get("flags")))
PII_KEY_RE = re.compile("(" + PII_PATTERNS["field_key"]["pattern"] + ")", _flags(PII_PATTERNS["field_key"].get("flags")))
PII_VALUE_RES = [(v["kind"], v.get("level", "block"), re.compile(v["pattern"], _flags(v.get("flags"))))
                 for v in PII_PATTERNS["values"]]
# Back-compat names used by older call sites.
PHONE_RE = next(rx for k, _l, rx in PII_VALUE_RES if k == "mobile")
EMAIL_RE = next(rx for k, _l, rx in PII_VALUE_RES if k == "email")


def _pii_text(v: Any) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return str(v)
    return clean_text(v).translate(_AR_INDIC)


def pii_value_hits(v: Any, levels: Tuple[str, ...] = ("block", "review")) -> list:
    """Kinds of personal data found in a value (e.g. ['mobile']). Never returns the matched text."""
    t = _pii_text(v)
    if not t:
        return []
    return [k for k, lvl, rx in PII_VALUE_RES if lvl in levels and rx.search(t)]


def mask_pii(v: Any, levels: Tuple[str, ...] = ("block",)) -> Tuple[str, list]:
    """Return (text with personal-data spans replaced by REDACTED, kinds found). Arabic-Indic digits are
    matched after translation; translation is 1:1 per character, so spans map back to the original text."""
    if v is None:
        return "", []
    orig = v if isinstance(v, str) else _pii_text(v)
    t = orig.translate(_AR_INDIC)
    spans, kinds = [], []
    for k, lvl, rx in PII_VALUE_RES:
        if lvl not in levels:
            continue
        for m in rx.finditer(t):
            spans.append((m.start(), m.end()))
            kinds.append(k)
    if not spans:
        return orig, []
    spans.sort()
    merged = [list(spans[0])]
    for a, b in spans[1:]:
        if a <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    out, last = [], 0
    for a, b in merged:
        out.append(orig[last:a])
        out.append(REDACTED)
        last = b
    out.append(orig[last:])
    return "".join(out), sorted(set(kinds))


def looks_like_phone(v: Any) -> bool:
    """True when the whole value is a phone number (mobile or landline), separators allowed."""
    t = _pii_text(v).strip()
    if not t:
        return False
    for k, _lvl, rx in PII_VALUE_RES:
        if k in ("mobile", "landline"):
            m = rx.search(t)
            if m and not re.sub(r"[\s\u00A0\-.()/+]", "", t[:m.start()] + t[m.end():]):
                return True
    return False


def _self_test_pii() -> None:
    tests = PII_PATTERNS.get("tests") or {}
    for x in tests.get("must_match", []):
        assert pii_value_hits(x), f"PII pattern miss: {x!r}"
    for x in tests.get("must_not_match", []):
        assert not pii_value_hits(x), f"PII false positive: {x!r} -> {pii_value_hits(x)}"
    for x in tests.get("block_must_match", []):
        assert pii_value_hits(x, ("block",)), f"PII block-level miss: {x!r}"
    for x in tests.get("block_must_not_match", []):
        assert not pii_value_hits(x, ("block",)), f"PII block-level false positive: {x!r} -> {pii_value_hits(x, ('block',))}"
    htests = PII_PATTERNS.get("header_tests") or {}
    for x in htests.get("must_match", []):
        assert is_pii_header(x), f"PII header miss: {x!r}"
    for x in htests.get("must_not_match", []):
        assert not is_pii_header(x), f"PII header false positive: {x!r}"


# Header test that tolerates spelling variants the raw pattern misses: tatweel («اســم المستــأجر»), hamza/taa-marbuta
# folding («المستاجر»، «هويه»). Same pattern source as PII_HEADER_RE, folded the way normalize_arabic folds text.
def _fold_ar_letters(s: str) -> str:
    s = re.sub("[أإآٱ]", "ا", s)
    return s.replace("ى", "ي").replace("ة", "ه").replace("ؤ", "و").replace("ئ", "ي")


PII_HEADER_FOLDED_RE = re.compile("(" + _fold_ar_letters(PII_PATTERNS["header"]["pattern"]) + ")",
                                  _flags(PII_PATTERNS["header"].get("flags")))


def is_pii_header(header: Any) -> bool:
    """True when a column header names personal data (contact, ID, IBAN, a person such as tenant/owner/beneficiary)."""
    if header is None:
        return False
    s = clean_text(header)
    if not s:
        return False
    return bool(PII_HEADER_RE.search(s) or PII_HEADER_FOLDED_RE.search(normalize_arabic(s)))


def ref(sheet: str, coord: str) -> str:
    """Canonical cell reference: exact sheet name (may contain trailing spaces) + '!' + A1 coordinate."""
    return f"{sheet}!{coord}"


# ---------------------------------------------------------------------------
# Total / subtotal row labels (shared by inspect_workbook.py and normalize.py)
# ---------------------------------------------------------------------------
# Applied to normalize_arabic(text).strip(): hamza/taa-marbuta are already folded (إجمالي → اجمالي, جملة → جمله).
TOTAL_LABEL_RE = re.compile(
    r"^(?:ال)?(?:مجموع|اجمالي|اجماليات|جمله|المجمل)(?:\s|$|ال)|^(?:grand\s+|sub-?)?totals?\b|^sum\b", re.IGNORECASE)


def is_total_label(v: Any) -> bool:
    """True when a cell text reads like a total/subtotal label («الإجمالي»، «جملة»، «Total»، «Grand total»)."""
    if v is None or not isinstance(v, str):
        return False
    t = normalize_arabic(clean_text(v)).strip()
    return bool(t) and bool(TOTAL_LABEL_RE.search(t))
