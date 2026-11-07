#!/usr/bin/env python3
"""normalize.py — build 03_records.json from the workbook and Claude's 02_mapping.json.

Usage:
    python3 scripts/normalize.py --xlsx <file.xlsx> --mapping RUN/02_mapping.json --out RUN/03_records.json

Implements data-contract.md §3–§4: typed values (TV) with source references, explicit
zero/blank/unknown statuses, derived fields, dimension rules, per-record overrides and
automatic data-quality issues (DQ-xxx). The workbook is opened read-only and never
modified. Personal-data columns are dropped (only counted).

Row handling (every source):
  * rows whose mapped cells are all blank/filler are skipped and counted (issue blank_rows_skipped);
  * a row skipped by skip_if_blank that still holds mapped content is reported (row_skipped_with_content)
    and listed in excluded_rows — never dropped silently;
  * total/subtotal rows (label in any mapped text cell, SUBTOTAL/SUM-range formula, or row kind total* in
    RUN/01_workbook_map.json) never become records unless 02_mapping.json/records overrides them; they are
    listed in excluded_rows with their numeric values (excluded_rows[].values) for reconciliation;
  * name/id cells inside a vertical merge read the merge anchor (source.merge_note).
Sources with "join" merge their fields into the records of another source through a confirmed link;
unmatched rows become isolated records. Controlled dimensions (use, income_nature, ownership, waqf,
managed_for_others) only hold values from assets/vocab.json.
Exit codes: 0 ok, 2 invalid mapping / unreadable input / a source that yields no rows.
"""
from __future__ import annotations

import argparse
import copy
import os
import re
import sys
import warnings
from collections import OrderedDict

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from apa_common import (  # noqa: E402
    ALL_STATUSES, CONTRACT_VERSION, EMAIL_RE, EXCEL_ERRORS, INCOME_NATURES, KNOWN_STATUSES, NA_TOKENS,
    NOT_APPLICABLE_TOKENS, OWNERSHIP_KINDS, S_BLANK, S_DERIVED, S_ERROR, S_FILLER,
    S_NOT_APPLICABLE, S_NOT_AVAILABLE, S_TEXT, S_VALUE, S_ZERO_AS_MISSING, S_ZERO_CONFIRMED, S_ZERO_RECORDED,
    SKILL_DIR, USES, WAQF_VALUES, clean_text, is_filler, is_pii_header, is_total_label, load_json, load_vocab,
    looks_like_phone, name_key, normalize_arabic, parse_date, parse_number, ref, save_json, sha256_file,
    REDACTED, mask_pii, pii_value_hits, foreign_currency, FOREIGN_CURRENCY_NOTE,
)
import formula_eval  # noqa: E402

NUMERIC_KINDS = ("money", "number", "area", "count", "percent")
AMOUNT_KINDS = ("money", "area", "count")       # a '%' here is never an amount
TEXT_KINDS = ("text", "url")
LIST_DIMS = ("ownership",)
DEFAULT_DIMS = OrderedDict([("use", "unknown"), ("income_nature", "unknown"),
                            ("ownership", ["unknown"]), ("waqf", "unknown")])
CONTROLLED_DIMS = ("use", "income_nature", "ownership", "waqf", "managed_for_others")
TOTAL_ROW_KINDS = ("total", "total_formula", "total_label")
ROW_KIND_AR = {"total": "إجمالي", "total_formula": "إجمالي بصيغة", "total_label": "إجمالي مدوّن"}
DECISION_SCOPE = {"primary": "in_scope", "supplementary": "in_scope", "isolated": "isolated", "excluded": "excluded"}
SEV_RANK = {"high": 0, "medium": 1, "low": 2, "info": 3}
TOTAL_WORD_RE = re.compile(r"(اجمالي|مجموع|جمله|المجمل|subtotal|total|(?<![a-z])sum(?![a-z]))", re.IGNORECASE)
VAL_RE = re.compile(r"^val\.([^.]+)\.(total|land|building)$")
MAX_REFS_IN_WHERE = 30
HOME_CURRENCY = ["SAR"]   # set from mapping.currency by normalize(); amounts in another currency stay text
_NA_NORM = frozenset(normalize_arabic(t) for t in NA_TOKENS)
_NOT_APPL_NORM = frozenset(normalize_arabic(t) for t in NOT_APPLICABLE_TOKENS)


def _vocab_keys() -> dict:
    try:
        v = load_vocab()
        out = {d: tuple(v[d]) for d in ("use", "income_nature", "ownership", "waqf")}
    except Exception:  # pragma: no cover - vocab.json ships with the skill
        out = {"use": USES, "income_nature": INCOME_NATURES, "ownership": OWNERSHIP_KINDS, "waqf": WAQF_VALUES}
    out["managed_for_others"] = (True, False)
    return out


VOCAB = _vocab_keys()


def eprint(*a):
    print(*a, file=sys.stderr)


# ---------------------------------------------------------------------------
# Mapping validation
# ---------------------------------------------------------------------------
def _arabic_error(err) -> str:
    v = err.validator
    if v == "required":
        m = re.search(r"'([^']+)' is a required property", err.message)
        return f"الحقل المطلوب «{m.group(1) if m else err.message}» مفقود"
    if v == "enum":
        allowed = "، ".join(str(x) for x in err.validator_value)
        return f"القيمة {err.instance!r} غير مسموحة؛ القيم المسموحة: {allowed}"
    if v == "const":
        return f"القيمة يجب أن تكون {err.validator_value!r}"
    if v == "type":
        return f"نوع القيمة غير صحيح ({type(err.instance).__name__})؛ المتوقع: {err.validator_value}"
    if v == "pattern":
        return f"القيمة {err.instance!r} لا تطابق الصيغة المطلوبة {err.validator_value}"
    if v == "additionalProperties":
        m = re.findall(r"'([^']+)'", err.message)
        return "مفتاح غير معروف: " + "، ".join(m) if m else "مفاتيح غير معروفة"
    if v in ("minItems", "minLength"):
        return f"القيمة قصيرة أو فارغة (الحد الأدنى {err.validator_value})"
    if v == "maxItems":
        return f"عدد العناصر أكبر من المسموح ({err.validator_value})"
    if v in ("minProperties", "maxProperties"):
        return "يجب أن يحتوي الكائن على عامل واحد فقط (مثل eq أو all)" if err.validator_value == 1 else \
            f"عدد المفاتيح غير مسموح ({v} = {err.validator_value})"
    if v == "uniqueItems":
        return "عناصر مكررة غير مسموحة"
    if v == "not":
        return f"القيمة {err.instance!r} غير مسموحة هنا"
    if v == "minimum":
        return f"القيمة {err.instance!r} أقل من الحد الأدنى {err.validator_value}"
    if v == "propertyNames":
        return f"اسم مفتاح غير صالح: {err.instance!r}"
    if v in ("anyOf", "oneOf"):
        return "القيمة لا تطابق أيًا من الصيغ المسموحة لهذا المفتاح"
    return f"لا يطابق المخطط ({v}): {err.message}"


def schema_errors(instance, schema_path: str) -> list:
    try:
        from jsonschema import Draft202012Validator
    except Exception:  # pragma: no cover - jsonschema is expected to be installed
        eprint("تحذير: مكتبة jsonschema غير مثبتة؛ تُجرى فحوص أساسية فقط.")
        try:
            required = load_json(schema_path).get("required") or []
        except Exception:
            required = []
        return [f"• المسار (الجذر): الحقل المطلوب «{k}» مفقود" for k in required if k not in instance]
    schema = load_json(schema_path)
    validator = Draft202012Validator(schema)
    out = []
    for e in sorted(validator.iter_errors(instance), key=lambda e: [str(p) for p in e.absolute_path]):
        path = "/".join(str(p) for p in e.absolute_path) or "(الجذر)"
        out.append(f"• المسار {path}: {_arabic_error(e)}")
    return out


def _dim_literal_ok(dim: str, v, allow_null: bool = False) -> bool:
    """A literal dimension value written in 02_mapping.json (dim_rules map/default, records dims)."""
    if v is None:
        return allow_null
    allowed = VOCAB[dim]
    if dim == "managed_for_others":
        return isinstance(v, bool)
    if dim in LIST_DIMS:
        items = v if isinstance(v, list) else [v]
        return bool(items) and all(isinstance(x, str) and x in allowed for x in items)
    return isinstance(v, str) and v in allowed


def _vocab_text(dim: str) -> str:
    return "true، false" if dim == "managed_for_others" else "، ".join(VOCAB[dim])


def semantic_errors(mapping: dict) -> list:
    """Checks the schema cannot express. Each entry is an Arabic message line."""
    errs = []
    for i, rule in enumerate(mapping.get("dim_rules") or []):
        dim = rule.get("dim")
        if dim not in CONTROLLED_DIMS:
            continue
        for k, v in (rule.get("map") or {}).items():
            if not _dim_literal_ok(dim, v):
                errs.append(f"• dim_rules/{i}/map/«{k}»: القيمة {v!r} خارج مفردات البعد «{dim}» "
                            f"(assets/vocab.json)؛ المسموح: {_vocab_text(dim)}")
        if "default" in rule and not _dim_literal_ok(dim, rule["default"], allow_null=(dim == "managed_for_others")):
            errs.append(f"• dim_rules/{i}/default: القيمة {rule['default']!r} خارج مفردات البعد «{dim}»؛ "
                        f"المسموح: {_vocab_text(dim)} (استخدم \"unknown\" لغير المحدد)")
    for rid, ov in (mapping.get("records") or {}).items():
        for dim, v in ((ov or {}).get("dims") or {}).items():
            if dim in CONTROLLED_DIMS and not _dim_literal_ok(dim, v, allow_null=(dim == "managed_for_others")):
                errs.append(f"• records/{rid}/dims/{dim}: القيمة {v!r} خارج مفردات البعد؛ المسموح: {_vocab_text(dim)}")
    sources = mapping.get("sources") or []
    by_id = {s.get("id"): s for s in sources}
    links = mapping.get("links") or []
    for s in sources:
        sid = s.get("id")
        dr = s.get("data_rows") or {}
        a, b = dr.get("from"), dr.get("to")
        if isinstance(a, int) and isinstance(b, int) and a > b:
            errs.append(f"• sources/{sid}/data_rows: بداية النطاق ({a}) بعد نهايته ({b})؛ صحّح from وto")
        j = s.get("join")
        if not j:
            continue
        into = by_id.get(j.get("into"))
        if into is None or j.get("into") == sid:
            errs.append(f"• sources/{sid}/join/into: المصدر الهدف «{j.get('into')}» غير موجود في sources (أو هو المصدر نفسه)")
        elif into.get("join"):
            errs.append(f"• sources/{sid}/join/into: المصدر الهدف «{j.get('into')}» مصدر ربط أيضًا؛ الربط يكون إلى مصدر سجلات فقط")
        else:
            on = j.get("on")
            if on == "id" and not (into.get("id_rule") or {}).get("column"):
                errs.append(f"• sources/{sid}/join/on: الربط بالمعرّف يتطلب id_rule.column في المصدر «{into['id']}»")
            if on == "name_key" and not into.get("name_column"):
                errs.append(f"• sources/{sid}/join/on: الربط بالاسم يتطلب name_column في المصدر «{into['id']}»")
            if on == "deed" and not any((c.get("field") or "") == "deed_no" or (c.get("field") or "").startswith("deed_no.")
                                        for c in (into.get("columns") or {}).values()):
                errs.append(f"• sources/{sid}/join/on: الربط بالصك يتطلب حقل deed_no في المصدر «{into['id']}»")
        if j.get("key_column") and j.get("key_column") in (s.get("columns") or {}) and \
                (s["columns"][j["key_column"]] or {}).get("kind") == "pii":
            errs.append(f"• sources/{sid}/join/key_column: عمود المفتاح «{j['key_column']}» معلَّم بيانات شخصية")
        link = [lk for lk in links if lk.get("id") == j.get("link_id")]
        if not link:
            errs.append(f"• sources/{sid}/join/link_id: لا يوجد في links[] رابط بالمعرّف «{j.get('link_id')}»؛ "
                        "لا تُنقل أرقام من ورقة أخرى إلا عبر رابط مؤكد")
        elif not any(lk.get("status") == "confirmed" or lk.get("match") == "confirmed" for lk in link):
            errs.append(f"• sources/{sid}/join/link_id: الرابط «{j.get('link_id')}» غير مؤكد؛ يلزم \"status\": \"confirmed\" "
                        "بدليل مطابقة (معرّف أو صك) قبل دمج بيانات المصدر")
    return errs


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def raw_str(val):
    if val is None:
        return None
    if isinstance(val, float) and val.is_integer():
        return str(int(val))
    if hasattr(val, "isoformat"):
        return val.isoformat()
    return str(val)


def tidy_num(x):
    """Integral floats become ints (cleaner JSON, identical arithmetic below 2**53)."""
    if isinstance(x, float) and x.is_integer() and abs(x) < 2 ** 53:
        return int(x)
    return x


def parse_rows(spec_rows) -> list:
    rows = []
    for item in spec_rows or []:
        if isinstance(item, int):
            rows.append(item)
        else:
            a, b = [int(x) for x in re.split(r"\s*-\s*", str(item).strip())]
            rows.extend(range(min(a, b), max(a, b) + 1))
    return rows


def col_index(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letter(n: int) -> str:
    return formula_eval.idx_to_col(n)


def cell_at(ws, r: int, c: int):
    """The cell object at (r, c) without creating it (openpyxl creates cells on ws[...] access)."""
    cells = getattr(ws, "_cells", None)
    if cells is not None:
        return cells.get((r, c))
    return ws.cell(row=r, column=c)


def cell_val(ws, r: int, c: int):
    cell = cell_at(ws, r, c)
    return None if cell is None else cell.value


def is_na_token(v) -> bool:
    return isinstance(v, str) and normalize_arabic(v) in _NA_NORM


def is_not_applicable_token(v) -> bool:
    return isinstance(v, str) and normalize_arabic(v) in _NOT_APPL_NORM


def is_blankish(v) -> bool:
    return v is None or is_filler(v)


def group_of(field: str, meta: dict) -> str:
    g = (meta.get(field) or {}).get("group")
    if g:
        return g
    if field.startswith("rev."):
        return "revenue"
    if field.startswith("val."):
        return "valuation"
    if field.startswith("exp."):
        return "expense"
    if field.startswith("area"):
        return "area"
    return "other"


def default_zero_policy(field: str, kind: str, meta: dict) -> str:
    if kind == "area" or group_of(field, meta) in ("valuation", "area") or field.startswith("val."):
        return "missing"
    return "recorded"


def impact_for(field: str, kind: str, meta: dict) -> str:
    g = group_of(field, meta)
    m = meta.get(field) or {}
    if g == "revenue":
        p = m.get("period") or "؟"
        return (f"يؤثر في مؤشرات الإيراد والعائد للفترة {p} (property_revenue، activity_revenue، gross_yield_*، "
                f"revenue_growth، revenue_concentration_top3، data_coverage) متى كان {field} هو حقل إيراد الفترة المعتمد")
    if g == "valuation":
        return (f"يؤثر في القيمة المسجلة والعائد (recorded_value، gross_yield_*، zero_rev_investment_value، "
                f"value_concentration_top3، data_coverage) متى كان {field} هو حقل التقييم المعتمد")
    if g == "expense":
        return "يؤثر في صافي الدخل التشغيلي (noi)"
    if kind == "count_mixed" or field == "units":
        return "لا يدخل في أي مؤشر مجمّع؛ لا يُجمع ولا يُستنتج منه الإشغال"
    if kind == "area" or g == "area":
        return "لا يدخل في المؤشرات الافتراضية؛ يؤثر في أي مؤشر للمتر المربع إن أضيف"
    if kind == "date":
        return "لا يدخل في المؤشرات الافتراضية؛ يؤثر في عرض تاريخ الأصل"
    return "لا يدخل في المؤشرات الافتراضية"


def join_refs(refs: list) -> str:
    refs = list(OrderedDict.fromkeys(r for r in refs if r))
    if len(refs) > MAX_REFS_IN_WHERE:
        return ", ".join(refs[:MAX_REFS_IN_WHERE]) + f", … (+{len(refs) - MAX_REFS_IN_WHERE})"
    return ", ".join(refs)


def rows_text(rows: list) -> str:
    """[3,4,5,9] -> '3–5، 9'"""
    rows = sorted(set(rows))
    out, i = [], 0
    while i < len(rows):
        j = i
        while j + 1 < len(rows) and rows[j + 1] == rows[j] + 1:
            j += 1
        out.append(str(rows[i]) if i == j else f"{rows[i]}–{rows[j]}")
        i = j + 1
    if len(out) > 12:
        return "، ".join(out[:12]) + f"، … (+{len(out) - 12})"
    return "، ".join(out)


_COUNT_FORMS = {  # one, two, 3–10, 11–99, hundreds (100, 101, 102 …)
    "row": ("صف واحد", "صفّان", "صفوف", "صفًا", "صف"),
    "record": ("سجل واحد", "سجلان", "سجلات", "سجلًا", "سجل"),
    "value": ("قيمة واحدة", "قيمتان", "قيم", "قيمة", "قيمة"),
}


def count_ar(n: int, noun: str) -> str:
    """Arabic number–noun agreement: 1 → «سجل واحد»، 2 → «سجلان»، 3–10 → «5 سجلات»، 11–99 → «12 سجلًا»، 100 → «100 سجل»."""
    one, two, few, many, hundred = _COUNT_FORMS[noun]
    if n == 1:
        return one
    if n == 2:
        return two
    r = n % 100
    if 3 <= r <= 10:
        return f"{n} {few}"
    if r >= 11:
        return f"{n} {many}"
    return f"{n} {hundred}"


def fmt(x) -> str:
    if x is None:
        return "—"
    if isinstance(x, float) and not x.is_integer():
        return f"{x:,.2f}"
    return f"{int(x):,}"


def unknown_of(dim: str):
    if dim in LIST_DIMS:
        return ["unknown"]
    if dim == "managed_for_others":
        return None
    return "unknown"


def is_unknown_dim(v) -> bool:
    return v is None or v == "unknown" or v == ["unknown"] or v == []


def coerce_dim(dim: str, v):
    """Value read from a cell for a controlled dimension -> (vocab value, ok)."""
    if dim == "managed_for_others":
        if isinstance(v, bool):
            return v, True
        t = clean_text(raw_str(v)).lower()
        if t in ("true", "false"):
            return t == "true", True
        return None, False
    allowed = VOCAB[dim]
    if dim in LIST_DIMS:
        parts = [clean_text(p).lower() for p in re.split(r"\s*[,،;/]\s*", clean_text(raw_str(v))) if clean_text(p)]
        if parts and all(p in allowed for p in parts):
            return list(OrderedDict.fromkeys(parts)), True
        return ["unknown"], False
    t = clean_text(raw_str(v)).lower()
    if t in allowed:
        return t, True
    return "unknown", False


def label_like_total(text) -> bool:
    """A short label such as «الإجمالي» / «Grand total (manual)» / «جملة القسم» in a non-name cell.
    Longer notes («إجمالي المساحة 500 م2») or texts with digits do not count."""
    if not isinstance(text, str) or not is_total_label(text):
        return False
    t = normalize_arabic(text)
    return not re.search(r"\d", t) and len(t.split()) <= 4


# ---------------------------------------------------------------------------
# Typed values
# ---------------------------------------------------------------------------
def make_tv(val, kind: str, cspec: dict, field: str, cell_ref: str, number_format: str, formula, meta: dict) -> dict:
    if kind in NUMERIC_KINDS or kind == "count_mixed":
        pct_amount = kind in AMOUNT_KINDS and (
            (isinstance(val, str) and ("%" in val or "٪" in val))
            or (isinstance(val, (int, float)) and not isinstance(val, bool) and number_format and "%" in number_format))
        if pct_amount:
            tv = {"v": None, "s": S_TEXT, "raw": raw_str(val), "ref": cell_ref,
                  "note": "نسبة مئوية في حقل مبلغ أو مساحة أو عدد؛ لم تُحوَّل إلى رقم", "percent_in_amount": True}
            if isinstance(val, (int, float)):
                tv["raw"] = f"{raw_str(val)} (تنسيق نسبة مئوية {number_format})"
        elif kind == "money" and foreign_currency(val, HOME_CURRENCY[0]):
            cur = foreign_currency(val, HOME_CURRENCY[0])
            tv = {"v": None, "s": S_TEXT, "raw": raw_str(val), "ref": cell_ref,
                  "note": FOREIGN_CURRENCY_NOTE.format(cur=cur, home=HOME_CURRENCY[0]), "foreign_currency": cur}
        else:
            v, s, note = parse_number(val, kind)
            ambiguous = s == S_TEXT and isinstance(note, str) and note.startswith("فاصل ملتبس")
            if kind == "percent" and isinstance(val, (int, float)) and not isinstance(val, bool) \
                    and number_format and "%" in number_format and v is not None:
                v = v * 100
                note = "خلية بتنسيق نسبة مئوية؛ خُزّنت القيمة بالنسبة المئوية (×100)"
            if s == S_ZERO_RECORDED:
                policy = cspec.get("zero_policy") or default_zero_policy(field, kind, meta)
                if policy == "missing":
                    s, v = S_ZERO_AS_MISSING, None
                    note = "صفر في حقل لا يُعد الصفر فيه قيمة مقبولة؛ عومل كغير متاح"
                else:
                    v = 0
            if s == S_TEXT and not ambiguous:
                note = ("وصف نصي مختلط؛ لا يُجمع" if kind == "count_mixed" else "نص في حقل رقمي؛ لم يُحوَّل")
            tv = {"v": tidy_num(v) if s in KNOWN_STATUSES else None, "s": s, "raw": raw_str(val),
                  "ref": cell_ref, "note": note}
            if ambiguous:
                tv["ambiguous_thousands"] = True
    elif kind == "date":
        if is_na_token(val) or is_not_applicable_token(val):
            tv = {"v": None, "s": S_NOT_APPLICABLE if is_not_applicable_token(val) else S_NOT_AVAILABLE,
                  "raw": raw_str(val), "ref": cell_ref, "calendar": None, "note": "رمز «غير متاح» وليس تاريخًا"}
        else:
            d = parse_date(val, cspec.get("calendar") or "auto")
            tv = {"v": d.get("v"), "s": d.get("s"), "raw": raw_str(val), "ref": cell_ref,
                  "calendar": d.get("calendar"), "note": d.get("note")}
            if d.get("hijri"):
                tv["hijri"] = d["hijri"]
    else:  # text, url
        if val is None:
            tv = {"v": None, "s": S_BLANK, "raw": None, "ref": cell_ref, "note": None}
        elif is_filler(val):
            tv = {"v": None, "s": S_FILLER, "raw": raw_str(val), "ref": cell_ref, "note": None}
        elif is_not_applicable_token(val):
            tv = {"v": None, "s": S_NOT_APPLICABLE, "raw": raw_str(val), "ref": cell_ref,
                  "note": "رمز «لا ينطبق» وليس قيمة أو فئة"}
        elif is_na_token(val):
            tv = {"v": None, "s": S_NOT_AVAILABLE, "raw": raw_str(val), "ref": cell_ref,
                  "note": "رمز «غير متاح» وليس قيمة أو فئة"}
        else:
            tv = {"v": clean_text(raw_str(val)), "s": S_VALUE, "raw": raw_str(val), "ref": cell_ref, "note": None}
    if isinstance(formula, str) and formula.startswith("="):
        tv["formula"] = formula
        if val is None:
            tv["note"] = "خلية صيغة بلا قيمة مخزنة (افتح الملف في Excel واحفظه لإعادة الحساب)"
            tv["no_cache"] = True
    return tv


def redact_if_contact(tv: dict, kind: str) -> bool:
    """Personal data in non-PII text columns is never stored (block-level patterns of assets/pii_patterns.json:
    phone, e-mail, IBAN, a labelled national ID). A value that is only a phone is dropped; otherwise the personal
    spans are replaced by the redaction mark and the rest of the text is kept."""
    if kind not in TEXT_KINDS + ("count_mixed",):
        return False
    raw = tv.get("raw")
    if raw is None:
        return False
    whole = looks_like_phone(raw) or any(looks_like_phone(tok) for tok in re.split(r"[\s,،/;|]+", clean_text(raw)) if tok)
    if not whole and not pii_value_hits(raw, ("block",)):
        return False
    masked, kinds = mask_pii(clean_text(raw), ("block",))
    rest = re.sub(r"[\s\u00A0\-.,:;()/+]", "", masked.replace(REDACTED, ""))
    if whole or not rest:
        tv.update(v=None, s=S_NOT_AVAILABLE, raw=None,
                  note="حُجبت قيمة تحمل بيانات شخصية (" + "، ".join(kinds or ["phone"]) + ")")
    else:
        tv.update(v=masked if tv.get("s") in KNOWN_STATUSES else tv.get("v"), raw=masked,
                  note="حُجبت بيانات شخصية من النص (" + "، ".join(kinds) + ")")
    return True


# ---------------------------------------------------------------------------
# Formula precedents: is a cached 0 just the result of blank inputs?
# ---------------------------------------------------------------------------
def _has_name_node(node) -> bool:
    if not isinstance(node, tuple) or not node:
        return False
    if node[0] == "name":
        return True
    return any(_has_name_node(x) for x in node[1:] if isinstance(x, tuple)) or \
        any(_has_name_node(a) for x in node[1:] if isinstance(x, list) for a in x)


class FormulaInputs:
    """Resolves the cells a formula reads (formula_eval parser; nothing is evaluated or written)."""
    MAX_CELLS = 5000
    MAX_DEPTH = 4

    def __init__(self, wb_v, wb_f):
        self.wb_v, self.wb_f = wb_v, wb_f
        self._tables = None
        self._memo = {}

    def _sheet(self, name):
        if name in self.wb_v.sheetnames:
            return name
        cands = [s for s in self.wb_v.sheetnames if clean_text(s).lower() == clean_text(name).lower()]
        return cands[0] if len(cands) == 1 else None

    def _table_list(self):
        if self._tables is None:
            from openpyxl.utils import range_boundaries
            self._tables = []
            for ws in self.wb_f.worksheets:
                for t in ws.tables.values():
                    try:
                        c1, r1, c2, r2 = range_boundaries(t.ref)
                    except Exception:
                        continue
                    hdr = 1 if t.headerRowCount is None else int(t.headerRowCount)
                    tot = int(t.totalsRowCount or 0)
                    self._tables.append({"name": t.name, "sheet": ws.title, "r1": r1, "c1": c1, "r2": r2, "c2": c2,
                                         "hdr": hdr, "tot": tot, "cols": [c.name for c in t.tableColumns]})
        return self._tables

    def _struct_cells(self, spec, sheet, row, col):
        tname = spec.get("table")
        tabs = self._table_list()
        if tname:
            t = next((x for x in tabs if (x["name"] or "").lower() == tname.lower()), None)
        else:
            t = next((x for x in tabs if x["sheet"] == sheet and x["r1"] <= row <= x["r2"] and x["c1"] <= col <= x["c2"]), None)
        if t is None:
            return None
        area, cols, is_span = formula_eval.parse_struct_spec(spec.get("spec") or "")
        if area == "this_row":
            if not t["r1"] <= row <= t["r2"]:
                return None
            rows = [row]
        elif area == "data":
            rows = list(range(t["r1"] + t["hdr"], t["r2"] - t["tot"] + 1))
        else:
            return None
        squeeze = lambda x: re.sub(r"\s+", " ", str(x or "")).strip().lower()  # noqa: E731
        names = [squeeze(c) for c in t["cols"]]
        if cols is None:
            cidx = list(range(t["c1"], t["c2"] + 1))
        else:
            try:
                idx = [t["c1"] + names.index(squeeze(c)) for c in cols]
            except ValueError:
                return None
            cidx = list(range(min(idx), max(idx) + 1)) if is_span else idx
        return [(t["sheet"], r, c) for r in rows for c in cidx]

    def cells(self, formula: str, sheet: str, row: int, col: int):
        """[(sheet, r, c)] read by the formula, or None when they cannot be resolved safely."""
        try:
            ast = formula_eval.parse(formula)
        except Exception:
            return None
        if _has_name_node(ast):
            return None
        out = []
        for kind, spec in formula_eval.iter_ref_nodes(ast):
            if kind == "ref":
                sh = self._sheet(spec["sheet"]) if spec.get("sheet") else sheet
                if sh is None:
                    return None
                ws = self.wb_v[sh]
                r1, r2, c1, c2 = spec["r1"], spec["r2"], spec["c1"], spec["c2"]
                if r1 is None:
                    r1, r2 = 1, max(1, ws.max_row)
                if c1 is None:
                    c1, c2 = 1, max(1, ws.max_column)
                if len(out) + (r2 - r1 + 1) * (c2 - c1 + 1) > self.MAX_CELLS:
                    return None
                out.extend((sh, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1))
            else:
                try:
                    res = self._struct_cells(spec, sheet, row, col)
                except Exception:
                    res = None
                if res is None or len(out) + len(res) > self.MAX_CELLS:
                    return None
                out.extend(res)
        return out

    def all_blank(self, formula: str, sheet: str, row: int, col: int, depth: int = 0) -> bool:
        """True when every cell the formula reads is blank, filler, a «غير متاح» token, an error or a formula
        without a cached value (recursively: a precedent formula whose own inputs are all blank)."""
        key = (sheet, row, col)
        if key in self._memo:
            return self._memo[key]
        self._memo[key] = False  # guards against circular references
        cells = self.cells(formula, sheet, row, col)
        ok = bool(cells)
        for sh, r, c in cells or []:
            v = cell_val(self.wb_v[sh], r, c)
            if is_blankish(v) or is_na_token(v) or is_not_applicable_token(v):
                continue
            if isinstance(v, str) and (clean_text(v).upper() in EXCEL_ERRORS or clean_text(v).upper().startswith("#REF")):
                continue
            f = cell_val(self.wb_f[sh], r, c)
            if isinstance(v, (int, float)) and not isinstance(v, bool) and v == 0 and isinstance(f, str) \
                    and f.startswith("=") and depth < self.MAX_DEPTH and self.all_blank(f, sh, r, c, depth + 1):
                continue
            ok = False
            break
        self._memo[key] = ok
        return ok

    def zero_from_blank(self, ws_name: str, row: int, col: int) -> bool:
        """The cell holds a formula whose cached result is 0 and whose inputs are all blank/unknown."""
        v = cell_val(self.wb_v[ws_name], row, col)
        if not (isinstance(v, (int, float)) and not isinstance(v, bool) and v == 0):
            return False
        f = cell_val(self.wb_f[ws_name], row, col)
        return isinstance(f, str) and f.startswith("=") and self.all_blank(f, ws_name, row, col)


# ---------------------------------------------------------------------------
# Derived-field formulas: + - * / ( ) numbers and field names only
# ---------------------------------------------------------------------------
_TOKEN_RE = re.compile(r"\s*(?:(\d+(?:\.\d+)?)|([A-Za-z_][A-Za-z0-9_.]*)|(\S))")


def _tokenize(formula: str) -> list:
    toks, pos = [], 0
    formula = formula.strip()
    while pos < len(formula):
        m = _TOKEN_RE.match(formula, pos)
        if not m or m.end() == pos:
            break
        pos = m.end()
        if m.group(1):
            toks.append(("num", float(m.group(1))))
        elif m.group(2):
            toks.append(("field", m.group(2)))
        elif m.group(3):
            if m.group(3) not in "+-*/()":
                raise ValueError(f"رمز غير مسموح في الصيغة: {m.group(3)}")
            toks.append(("op", m.group(3)))
    return toks


def formula_fields(formula: str) -> list:
    return [t[1] for t in _tokenize(formula) if t[0] == "field"]


def eval_formula(formula: str, fields: dict):
    """Return (value, missing_fields, inputs). Missing => value None."""
    toks = _tokenize(formula)
    missing, inputs = [], []
    for kind, name in toks:
        if kind == "field":
            tv = fields.get(name)
            if not tv or tv.get("s") not in KNOWN_STATUSES:
                missing.append(name)
            else:
                inputs.append({"field": name, "v": tv.get("v"), "ref": tv.get("ref")})
    if missing:
        return None, sorted(set(missing)), inputs
    pos = [0]

    def peek():
        return toks[pos[0]] if pos[0] < len(toks) else (None, None)

    def take():
        t = peek()
        pos[0] += 1
        return t

    def expr():
        v = term()
        while peek() in (("op", "+"), ("op", "-")):
            op = take()[1]
            r = term()
            v = v + r if op == "+" else v - r
        return v

    def term():
        v = factor()
        while peek() in (("op", "*"), ("op", "/")):
            op = take()[1]
            r = factor()
            if op == "/":
                if r == 0:
                    raise ZeroDivisionError("قسمة على صفر")
                v = v / r
            else:
                v = v * r
        return v

    def factor():
        k, x = take()
        if k == "num":
            return x
        if k == "field":
            return fields[x]["v"]
        if (k, x) == ("op", "-"):
            return -factor()
        if (k, x) == ("op", "("):
            v = expr()
            if take() != ("op", ")"):
                raise ValueError("قوس غير مغلق")
            return v
        raise ValueError("صيغة غير صالحة")

    value = expr()
    if pos[0] != len(toks):
        raise ValueError("صيغة غير صالحة (رموز زائدة)")
    return tidy_num(value), [], inputs


# ---------------------------------------------------------------------------
# Issues
# ---------------------------------------------------------------------------
class Issues:
    def __init__(self):
        self.items = []

    def add(self, type_, severity, where, record_ids, detail, impact, treatment, **extra):
        it = OrderedDict([("id", None), ("type", type_), ("severity", severity), ("where", where),
                          ("record_ids", sorted(set(record_ids or []))), ("detail", detail),
                          ("impact", impact), ("treatment", treatment)])
        it.update(extra)
        self.items.append(it)
        return it

    def finalize(self, by_id: dict):
        """Issues that touch only isolated/excluded records say so in `impact` and sort after the issues that touch
        in-scope records (or no record); severity orders each group. Ids are assigned in that order (DQ-001…)."""
        groups = []
        for it in self.items:
            rids = [r for r in it["record_ids"] if r in by_id]
            scopes = {by_id[r]["scope"] for r in rids}
            if not rids or "in_scope" in scopes:
                groups.append(0)
                continue
            groups.append(1)
            n = len(rids)
            if scopes == {"isolated"}:
                what = "سجل معزول" if n == 1 else "سجلات معزولة"
            elif scopes == {"excluded"}:
                what = "سجل مستبعد" if n == 1 else "سجلات مستبعدة"
            else:
                what = "سجلات معزولة أو مستبعدة"
            it["impact"] = f"لا أثر على المؤشرات ({what})"
        order = sorted(range(len(self.items)),
                       key=lambda i: (groups[i], SEV_RANK.get(self.items[i]["severity"], 9), i))
        self.items = [self.items[i] for i in order]
        for i, it in enumerate(self.items, 1):
            it["id"] = f"DQ-{i:03d}"
        return self.items


# ---------------------------------------------------------------------------
# Workbook helpers
# ---------------------------------------------------------------------------
def open_workbook(path: str):
    import openpyxl
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb_v = openpyxl.load_workbook(path, data_only=True)
        wb_f = openpyxl.load_workbook(path, data_only=False)
    return wb_v, wb_f


def resolve_sheet(wb, name: str, issues: Issues):
    if name in wb.sheetnames:
        return name
    cands = [s for s in wb.sheetnames if clean_text(s) == clean_text(name)]
    if len(cands) == 1:
        issues.add("mapping_sheet_name", "low", cands[0] + "!", [],
                   f"اسم الورقة في mapping «{name}» لا يطابق حرفيًا؛ استُخدمت الورقة «{cands[0]}» (فرق مسافات).",
                   "لا أثر على الأرقام", "صحّح اسم الورقة في 02_mapping.json ليطابق الاسم الحرفي")
        return cands[0]
    raise SystemExit(f"خطأ: الورقة «{name}» غير موجودة. الأوراق المتاحة: " +
                     "، ".join(f"«{s}»" for s in wb.sheetnames))


def is_total_formula(formula) -> bool:
    if not isinstance(formula, str) or not formula.startswith("="):
        return False
    f = formula.upper()
    if "SUBTOTAL(" in f or "AGGREGATE(" in f:
        return True
    for m in re.finditer(r"SUM\(\s*\$?[A-Z]{1,3}\$?(\d+)\s*:\s*\$?[A-Z]{1,3}\$?(\d+)", f):
        if m.group(1) != m.group(2):
            return True
    if re.search(r"SUM\([^)]*\[", f) and "#THIS ROW" not in f and "[@" not in f:
        return True
    return False


def vertical_merges(ws) -> dict:
    """(row, col) of every non-anchor cell of a merge spanning several rows -> (anchor_row, anchor_col, 'B2:B3')."""
    out = {}
    for mr in ws.merged_cells.ranges:
        if mr.max_row <= mr.min_row:
            continue
        for r in range(mr.min_row, mr.max_row + 1):
            for c in range(mr.min_col, mr.max_col + 1):
                if (r, c) != (mr.min_row, mr.min_col):
                    out[(r, c)] = (mr.min_row, mr.min_col, mr.coord)
    return out


def row_kinds_by_sheet(wb_map) -> dict:
    """{sheet name: {row: kind}} from 01_workbook_map.json row_kinds ({"3-17": "data", "18": "total_formula"})."""
    out = {}
    for sh in (wb_map or {}).get("sheets") or []:
        kinds = {}
        for span, kind in (sh.get("row_kinds") or {}).items():
            m = re.fullmatch(r"\s*(\d+)\s*(?:-\s*(\d+))?\s*", str(span))
            if not m:
                continue
            a = int(m.group(1))
            b = int(m.group(2) or a)
            if kind in TOTAL_ROW_KINDS:
                for r in range(a, b + 1):
                    kinds[r] = kind
        out[sh.get("name")] = kinds
    return out


def section_rows_by_sheet(wb_map) -> dict:
    """{sheet name: {rows}} that 01_workbook_map.json classified section_label."""
    out = {}
    for sh in (wb_map or {}).get("sheets") or []:
        rows = set()
        for span, kind in (sh.get("row_kinds") or {}).items():
            m = re.fullmatch(r"\s*(\d+)\s*(?:-\s*(\d+))?\s*", str(span))
            if m and kind == "section_label":
                rows.update(range(int(m.group(1)), int(m.group(2) or m.group(1)) + 1))
        out[sh.get("name")] = rows
    return out


def find_sheet_decl(decls: list, sheet: str):
    exact = [d for d in decls if d.get("sheet") == sheet]
    if exact:
        return exact[0]
    loose = [d for d in decls if clean_text(d.get("sheet")) == clean_text(sheet)]
    return loose[0] if loose else None


# ---------------------------------------------------------------------------
# Main normalisation
# ---------------------------------------------------------------------------
class Ctx:
    def __init__(self, xlsx: str, mapping: dict, wb_map):
        self.mapping = mapping
        self.meta = mapping.get("field_meta") or {}
        self.issues = Issues()
        self.wb_v, self.wb_f = open_workbook(xlsx)
        self.finputs = FormulaInputs(self.wb_v, self.wb_f)
        self.row_kinds = row_kinds_by_sheet(wb_map)
        self.section_rows = section_rows_by_sheet(wb_map)
        self.regions = {sh.get("name"): sh.get("regions") or [] for sh in (wb_map or {}).get("sheets") or []}
        self.overrides = mapping.get("records") or {}
        self.decls = mapping.get("sheets") or []
        self.records = []
        self.excluded_rows = []
        self.seen_ids = {}
        self.id_texts = {}
        self.field_kinds = {}
        self.row_owner = OrderedDict()   # (sheet, row) -> [record ids] of non-join sources
        self.scope_problems = []         # (source_id, sheet, [Arabic problems])
        self.managed_sources = []        # (source_id, sheet) on managed_properties sheets
        self.stats = OrderedDict([("pii_columns", 0), ("pii_values", 0), ("rows_skipped", 0), ("rows_excluded", 0),
                                  ("rows_skipped_with_content", 0), ("rows_total_excluded", 0), ("rows_joined", 0),
                                  ("rows_total_after_range", 0)])


def _candidate_id(prefix: str, id_text, row: int) -> str:
    return prefix + id_text if id_text else f"{prefix}r{row}"


def _overridden_as_record(ctx: Ctx, rid: str) -> bool:
    ov = ctx.overrides.get(rid) or {}
    return "record_type" in ov or "scope" in ov


def source_scope(ctx: Ctx, src: dict, sheet: str) -> tuple:
    """Effective default scope of a source + the Arabic problems found (A9)."""
    ws = ctx.wb_f[sheet]
    state = getattr(ws, "sheet_state", "visible") or "visible"
    hidden = state in ("hidden", "veryHidden")
    decl = find_sheet_decl(ctx.decls, sheet)
    stated = src.get("default_scope")
    problems = []
    expected = DECISION_SCOPE.get((decl or {}).get("decision")) if decl else None
    if decl is None:
        problems.append(f"الورقة «{sheet}» غير مدرجة في sheets[] في 02_mapping.json؛ لا يوجد قرار نطاق موثق لها")
    if stated is None:
        eff = expected or ("isolated" if hidden else "in_scope")
    else:
        eff = stated
    if hidden and (stated is None or not clean_text(src.get("scope_reason"))):
        problems.append(f"المصدر يقرأ ورقة مخفية ({state}) دون default_scope وscope_reason صريحين"
                        + (f"؛ طُبّق «{eff}» احتياطًا" if stated is None else ""))
    if expected and eff != expected:
        problems.append(f"default_scope = «{eff}» لا يوافق قرار الورقة «{decl['decision']}» في sheets[] "
                        f"(المتوقع «{expected}»)")
    if src.get("join") and decl and decl.get("decision") not in ("primary", "supplementary"):
        problems.append(f"مصدر ربط يدمج بياناته في سجلات المصدر «{src['join'].get('into')}» بينما قرار ورقته "
                        f"«{decl.get('decision')}»؛ الدمج يتطلب قرار primary أو supplementary")
    if problems:
        ctx.scope_problems.append((src["id"], sheet, problems))
    if decl and decl.get("role") == "managed_properties":
        ctx.managed_sources.append((src["id"], sheet))
    reason = src.get("scope_reason")
    if reason is None and stated is None and eff != "in_scope" and decl:
        reason = decl.get("reason")
    return eff, reason


def read_source(ctx: Ctx, src: dict, sheet: str) -> list:
    """Read the data rows of one source (row dicts); exclusions, blank and total rows are handled here."""
    sid = src["id"]
    issues, meta = ctx.issues, ctx.meta
    ws, wsf = ctx.wb_v[sheet], ctx.wb_f[sheet]
    last_row = ws.max_row
    header_row = src.get("header_row")
    cols = sorted(src["columns"].items(), key=lambda kv: col_index(kv[0]))
    pii_auto = set()
    for col, cspec in cols:
        if cspec["kind"] == "pii":
            ctx.stats["pii_columns"] += 1
            continue
        if header_row and not cspec.get("pii_ok"):
            header = cell_val(ws, header_row, col_index(col))
            if header is not None and is_pii_header(header) and cspec["kind"] != "ignore":
                pii_auto.add(col)
                ctx.stats["pii_columns"] += 1
                issues.add("pii_header_not_marked", "high", ref(sheet, f"{col}{header_row}"), [],
                           f"عنوان العمود {col} يدل على بيانات شخصية لكنه لم يُعلَّم pii في mapping.",
                           "لا أثر على المؤشرات",
                           "عومل العمود كبيانات شخصية ولم تُخزَّن قيمه؛ إن لم يكن شخصيًا فأضف \"pii_ok\": true")
        if cspec["kind"] != "ignore":
            ctx.field_kinds.setdefault(cspec.get("field"), cspec["kind"])
    data_cols = [(col, cspec, col_index(col)) for col, cspec in cols
                 if cspec["kind"] not in ("ignore", "pii") and col not in pii_auto]
    numeric_cols = [(col, cspec, ci) for col, cspec, ci in data_cols if cspec["kind"] in NUMERIC_KINDS]
    dr = src["data_rows"]
    start, end = dr["from"], dr["to"]
    excl = {}
    for ex in src.get("exclude_rows") or []:
        for r in parse_rows(ex.get("rows")):
            excl[r] = ex.get("reason")
    id_rule = src.get("id_rule") or {}
    prefix = id_rule.get("prefix")
    if prefix is None:
        prefix = sid + "-"
    id_col = id_rule.get("column")
    name_col = src.get("name_column")
    skip_col = src.get("skip_if_blank")
    join = src.get("join") or {}
    key_col = join.get("key_column")
    pii_label_cols = []   # name/id columns dropped as personal data: still scanned for total labels, never stored
    # Safety net: a name/id column whose header (or column spec) marks personal data never becomes a record name/id.
    for role, col in (("name_column", name_col), ("id_rule.column", id_col)):
        if not col:
            continue
        cs = src["columns"].get(col) or {}
        header = cell_val(ws, header_row, col_index(col)) if header_row else None
        if cs.get("kind") == "pii" or col in pii_auto or (not cs.get("pii_ok") and header is not None
                                                          and is_pii_header(header)):
            issues.add("pii_header_not_marked", "high", ref(sheet, f"{col}{header_row or start}"), [],
                       f"العمود {col} المستخدم في {role} للمصدر {sid} يحمل بيانات شخصية"
                       + (f" (العنوان «{clean_text(header)}»)" if header is not None else "") + ".",
                       "لا أثر على المؤشرات",
                       "لم تُخزَّن قيمه اسمًا أو معرّفًا للسجلات (السجلات بلا اسم أو بمعرّف الصف)؛ اختر عمود اسم الأصل أو "
                       "معرّفه، أو أضف \"pii_ok\": true في columns إن لم يكن شخصيًا")
            pii_label_cols.append(col)
            if col == name_col:
                name_col = None
            if col == id_col:
                id_col = None
    merges = vertical_merges(ws)
    kinds01 = ctx.row_kinds.get(sheet) or {}
    # a row is blank when its data cells and name cell are empty: an id/key alone (pre-numbered template rows) is not content
    content_cols = [(col, ci) for col, _cs, ci in data_cols if col not in (id_col, key_col)]
    if name_col and name_col not in [c for c, _ in content_cols] and name_col not in pii_auto and \
            (src["columns"].get(name_col) or {}).get("kind") != "pii":
        content_cols.append((name_col, col_index(name_col)))

    def resolved(col, row):
        """Value of an id/name/key/skip cell; a cell inside a vertical merge reads the merge anchor."""
        ci = col_index(col)
        v = cell_val(ws, row, ci)
        if is_blankish(v) and (row, ci) in merges:
            ar, ac, coord = merges[(row, ci)]
            av = cell_val(ws, ar, ac)
            if not is_blankish(av):
                return av, f"«{col}{row}» داخل خلية مدمجة رأسيًا {coord}؛ قُرئت قيمتها من {col_letter(ac)}{ar}"
        return v, None

    def has_content(ci, row):
        v = cell_val(ws, row, ci)
        if is_blankish(v):
            return False
        if isinstance(v, (int, float)) and not isinstance(v, bool) and v == 0 and ctx.finputs.zero_from_blank(sheet, row, ci):
            return False
        return True

    def total_reasons(row, id_text, name):
        """(strong, weak) reasons for a total row. Strong: the name/id cell itself reads as a total label. Weak: a total
        label in another mapped cell, a vertical SUM/SUBTOTAL, or 01's row kind — signs a named asset row can also carry
        (a «إجمالي/صافي» rent-basis column, a parent building summing its unit rows)."""
        strong, weak = [], []
        for lab, txt, col in (("الاسم", name, name_col), ("المعرّف", id_text, id_col)):
            if txt and is_total_label(txt):
                strong.append(f"{lab} «{txt}» ({col}{row}) يدل على إجمالي")
        for col in pii_label_cols:
            if is_total_label(cell_val(ws, row, col_index(col))):
                strong.append(f"الخلية {col}{row} تدل على إجمالي")
        seen_formula = False
        for col, cspec, ci in data_cols:
            if col in (name_col, id_col):
                continue
            v = cell_val(ws, row, ci)
            if label_like_total(v):
                weak.append(f"الخلية {col}{row} «{clean_text(v)}» تدل على إجمالي")
            if cspec["kind"] in NUMERIC_KINDS and not seen_formula and is_total_formula(cell_val(wsf, row, ci)):
                weak.append(f"صيغة تجميع في {col}{row} (SUBTOTAL/SUM لنطاق)")
                seen_formula = True   # one formula reason is enough
        k = kinds01.get(row)
        if k:
            weak.append(f"خريطة المصنف (01) صنّفت الصف «{ROW_KIND_AR.get(k, k)}»")
        return strong, weak

    def row_values(row):
        vals = OrderedDict()
        for col, cspec, ci in numeric_cols:
            cell = cell_at(ws, row, ci)
            v = None if cell is None else cell.value
            f = cell_val(wsf, row, ci)
            if is_blankish(v) and not (isinstance(f, str) and f.startswith("=")):
                continue
            vals[cspec["field"]] = make_tv(v, cspec["kind"], cspec, cspec["field"], ref(sheet, f"{col}{row}"),
                                           cell.number_format if cell is not None else "General", f, meta)
        return vals

    sections01 = ctx.section_rows.get(sheet) or set()

    def section_label_of(row):
        """The text of a section-label row («Jeddah», «عقارات الرياض») whose only content is in the id/key column (or
        that 01 classified section_label): reported as a section label, not as a blank row."""
        for col in [c for c in (id_col, key_col) if c]:
            v = cell_val(ws, row, col_index(col))
            if isinstance(v, str) and not is_blankish(v):
                t = clean_text(v)
                if row in sections01 or (re.search(r"[A-Za-z\u0600-\u06FF]{3,}", t) and not re.search(r"\d", t)):
                    return t
        return None

    rows_out = []
    blank_rows, beyond = [], 0
    section_rows = []      # (row, label)
    skipped_content = []   # (row, [refs])
    total_rows = []        # (row, reasons, where)
    for row in range(start, end + 1):
        if row in excl:
            ctx.stats["rows_excluded"] += 1
            entry = OrderedDict([("source_id", sid), ("sheet", sheet), ("row", row), ("reason", excl[row])])
            if row <= last_row:
                iv = resolved(id_col, row)[0] if id_col else None
                nv = resolved(name_col, row)[0] if name_col else None
                is_tot = any(total_reasons(row, clean_text(raw_str(iv)) or None, clean_text(raw_str(nv)) or None)) \
                    or bool(TOTAL_WORD_RE.search(normalize_arabic(excl[row] or "")))
                if is_tot:
                    vals = row_values(row)
                    if vals:
                        entry["values"] = vals
            ctx.excluded_rows.append(entry)
            continue
        if row > last_row:
            beyond += 1
            continue
        content = [col for col, ci in content_cols if has_content(ci, row)]
        if not content:
            lab = section_label_of(row)
            if lab:
                section_rows.append((row, lab))
            else:
                blank_rows.append(row)
            continue
        if skip_col:
            sv, _note = resolved(skip_col, row)
            if is_blankish(sv):
                others = [c for c in content if c != skip_col]
                if not others:
                    blank_rows.append(row)
                    continue
                refs = [ref(sheet, f"{c}{row}") for c in others]
                skipped_content.append((row, refs))
                ctx.stats["rows_skipped_with_content"] += 1
                ctx.excluded_rows.append(OrderedDict([
                    ("source_id", sid), ("sheet", sheet), ("row", row),
                    ("reason", f"تُخطّي الصف لأن خانة {skip_col}{row} (skip_if_blank) فارغة رغم وجود بيانات في: "
                               + "، ".join(f"{c}{row}" for c in others))]))
                continue
        merge_notes = []
        id_text = None
        if id_col:
            iv, note = resolved(id_col, row)
            if not is_blankish(iv):
                id_text = clean_text(raw_str(iv))
                if note:
                    merge_notes.append(note)
        name = None
        if name_col:
            nv, note = resolved(name_col, row)
            name = clean_text(raw_str(nv)) or None
            if name and note:
                merge_notes.append(note)
        strong, weak = total_reasons(row, id_text, name)
        cand = _candidate_id(prefix, id_text, row)
        named = bool(id_text and name)   # a real id and a non-total name: weak signs alone do not make it a total
        reasons = strong + weak if (strong or not named) else []
        if reasons and not (not join and _overridden_as_record(ctx, cand)):
            where = ref(sheet, f"{name_col or id_col or (data_cols[0][0] if data_cols else 'A')}{row}")
            total_rows.append((row, reasons, where, cand))
            ctx.stats["rows_total_excluded"] += 1
            entry = OrderedDict([("source_id", sid), ("sheet", sheet), ("row", row),
                                 ("reason", "صف إجمالي مكتشف آليًا: " + "؛ ".join(reasons)), ("auto", True)])
            vals = row_values(row)
            if vals:
                entry["values"] = vals
            ctx.excluded_rows.append(entry)
            continue
        rr = {"row": row, "id_text": id_text, "name": name, "merge_notes": merge_notes, "total_reasons": reasons,
              "possible_total": weak if (weak and not strong and named) else None,
              "prefix": prefix, "id_col": id_col, "name_col": name_col, "sheet": sheet, "source_id": sid}
        if key_col:
            kv, note = resolved(key_col, row)
            rr["key_raw"] = None if is_blankish(kv) else kv
            rr["key_ref"] = ref(sheet, f"{key_col}{row}")
        dims = OrderedDict((k, copy.deepcopy(v)) for k, v in DEFAULT_DIMS.items())
        dims_ev = OrderedDict()
        fields = OrderedDict()
        oov = OrderedDict()
        for col, cspec in cols:
            kind = cspec["kind"]
            if kind == "ignore":
                continue
            ci = col_index(col)
            cell = cell_at(ws, row, ci)
            val = None if cell is None else cell.value
            if kind == "pii" or col in pii_auto:
                if not is_blankish(val):
                    ctx.stats["pii_values"] += 1
                continue
            cell_ref = ref(sheet, f"{col}{row}")
            formula = cell_val(wsf, row, ci)
            field = cspec["field"]
            tv = make_tv(val, kind, cspec, field, cell_ref, cell.number_format if cell is not None else "General",
                         formula, meta)
            if kind in NUMERIC_KINDS and tv.get("s") in (S_ZERO_RECORDED, S_ZERO_AS_MISSING) and \
                    ctx.finputs.zero_from_blank(sheet, row, ci):
                tv.update(v=None, s=S_NOT_AVAILABLE, zero_from_blank_inputs=True,
                          note="صيغة ناتجها المخزن 0 لأن كل الخلايا التي تعتمد عليها فارغة أو غير متاحة؛ "
                               "عوملت كغير متاحة لا كصفر")
            if redact_if_contact(tv, kind):
                ctx.stats["pii_values"] += 1
                tv["redacted"] = True
            fields[field] = tv
            dim = cspec.get("dim")
            if dim:
                dv = tv.get("v") if tv.get("s") in KNOWN_STATUSES else None
                if dim in CONTROLLED_DIMS:
                    if dv is None:
                        dims[dim] = unknown_of(dim)
                    else:
                        val2, ok = coerce_dim(dim, dv)
                        dims[dim] = val2
                        if not ok:
                            oov[dim] = (cell_ref, raw_str(tv.get("raw") if tv.get("raw") is not None else dv), field)
                elif dim in LIST_DIMS:
                    if dv is not None:
                        dims[dim] = [dv]
                else:
                    dims[dim] = dv
                dims_ev[dim] = cell_ref
        rr.update(dims=dims, dims_ev=dims_ev, fields=fields, oov=oov)
        rows_out.append(rr)

    if blank_rows or beyond:
        n = len(blank_rows) + beyond
        ctx.stats["rows_skipped"] += n
        wcol = skip_col or name_col or (content_cols[0][0] if content_cols else "A")
        lo = min(blank_rows) if blank_rows else last_row + 1
        hi = max(blank_rows) if blank_rows else end
        if beyond:
            hi = end
        empty = "فارغ" if n == 1 else ("فارغان" if n == 2 else "فارغة")
        detail = f"{count_ar(n, 'row')} {empty} ضمن نطاق البيانات ({start}–{end}) في المصدر {sid}"
        if blank_rows:
            detail += f": الصفوف {rows_text(blank_rows)}"
        if beyond:
            detail += f"؛ منها {count_ar(beyond, 'row')} بعد آخر صف مستخدم في الورقة ({last_row})"
        issues.add("blank_rows_skipped", "low", ref(sheet, f"{wcol}{lo}:{wcol}{hi}"), [],
                   detail + " — كل خاناتها المربوطة فارغة أو مسافات.",
                   "لا أثر على المؤشرات؛ لم تُنشأ منها سجلات",
                   "تُجووزت الصفوف وعُدّت فقط؛ ضيّق data_rows إن كانت خارج السجل", source_id=sid, count=n)
    if section_rows:
        lab_col = id_col or key_col or "A"
        issues.add("section_label_skipped", "low", join_refs([ref(sheet, f"{lab_col}{r}") for r, _ in section_rows]), [],
                   f"{count_ar(len(section_rows), 'row')} في المصدر {sid} "
                   + ("عنوان قسم لا سجل" if len(section_rows) == 1 else "عناوين أقسام لا سجلات") + ": "
                   + "، ".join(f"الصف {r} «{lab}»" for r, lab in section_rows[:6]) + ".",
                   "لا أثر على المؤشرات؛ لم تُنشأ منها سجلات",
                   "تُجووزت كعناوين أقسام؛ إن كانت تصنّف السطور التي تحتها (مدينة، حي) فأضف التصنيف بـ dim_rules",
                   source_id=sid, rows=[r for r, _ in section_rows])
    after_rows = totals_after_range(ctx, src, sheet, ws, wsf, end, last_row, id_col, name_col, data_cols, row_values)
    for row, why in after_rows:
        entry = OrderedDict([("source_id", sid), ("sheet", sheet), ("row", row),
                             ("reason", "صف إجمالي بعد نطاق البيانات — للمطابقة فقط: " + why), ("auto", True),
                             ("after_data_rows", True)])
        vals = row_values(row)
        if any(tv.get("s") in KNOWN_STATUSES for tv in vals.values()):
            entry["values"] = vals
        else:
            entry["reason"] += " (لا قيم معروفة في خاناته المربوطة)"
        ctx.excluded_rows.append(entry)
        ctx.stats["rows_total_after_range"] += 1
    if skipped_content:
        refs = [r for _, rs in skipped_content for r in rs]
        issues.add("row_skipped_with_content", "medium", join_refs(refs), [],
                   f"تُخطّي في المصدر {sid} {count_ar(len(skipped_content), 'row')} لأن خانة {skip_col} (skip_if_blank) "
                   f"فارغة رغم وجود بيانات مربوطة: الصفوف {rows_text([r for r, _ in skipped_content])}.",
                   "بيانات هذه الصفوف لا تدخل أي مؤشر؛ قد يكون الإجمالي ناقصًا إن كانت سطورًا تابعة لأصل",
                   "لم تُنشأ منها سجلات وأُدرجت في excluded_rows؛ صنّفها في mapping (عمود اسم آخر، أو مصدر ربط، أو "
                   "استبعاد صريح بسبب)", source_id=sid, rows=[r for r, _ in skipped_content])
    for row, reasons, where, cand in total_rows:
        issues.add("total_like_row", "high", where, [],
                   f"الصف {row} في المصدر {sid} يبدو صف إجمالي: " + "؛ ".join(reasons) + ".",
                   "لا أثر على المؤشرات بعد استبعاده؛ لو دخل النطاق لضاعف العدد والقيمة والإيراد",
                   "لم يُنشأ منه سجل؛ أُدرج في excluded_rows مع قيمه للمطابقة مع مجموع السجلات. إن لم يكن إجماليًا "
                   f"فأضف records/{cand} بنوعه ونطاقه الصحيحين مع الدليل", source_id=sid, row=row)
    return rows_out


def totals_after_range(ctx: Ctx, src: dict, sheet: str, ws, wsf, end: int, last_row: int, id_col, name_col,
                       data_cols, row_values) -> list:
    """Recorded total rows just after data_rows.to (inspect's proposal usually ends before them): [(row, why)].
    From 01: the region holding the source's rows lists them in total_rows. Without 01: the next rows (up to 6, stopping
    at the first row that is neither blank nor a total) whose name/id cell or another mapped text cell is a total label.
    Rows read by another source of the same sheet are skipped. They are kept only for reconciliation (A14)."""
    others = set()
    for o in ctx.mapping.get("sources") or []:
        if o is src or o.get("sheet") != src.get("sheet"):
            continue
        dr = o.get("data_rows") or {}
        if isinstance(dr.get("from"), int) and isinstance(dr.get("to"), int):
            others.update(range(dr["from"], dr["to"] + 1))
    excl = set()
    for ex in src.get("exclude_rows") or []:
        excl.update(parse_rows(ex.get("rows")))
    start = src["data_rows"]["from"]
    out = []
    regions = ctx.regions.get(sheet)
    if regions:
        for reg in regions:
            rows = reg.get("rows") or []
            if len(rows) != 2 or not (rows[0] <= start <= rows[1] or rows[0] <= end <= rows[1]):
                continue
            for r in reg.get("total_rows") or []:
                if isinstance(r, int) and end < r <= min(rows[1], last_row) and r not in others and r not in excl:
                    out.append((r, f"خريطة المصنف (01) تعده صف إجمالي في منطقة الجدول ({rows[0]}–{rows[1]})"))
        return sorted(set(out))
    for r in range(end + 1, min(end + 6, last_row) + 1):
        if r in others or r in excl:
            break
        texts = [cell_val(ws, r, col_index(c)) for c in (name_col, id_col) if c]
        texts += [cell_val(ws, r, ci) for _c, cs, ci in data_cols if cs["kind"] in TEXT_KINDS]
        vals = [cell_val(ws, r, ci) for _c, cs, ci in data_cols]
        if all(is_blankish(v) for v in vals + texts):
            continue
        lab = next((clean_text(t) for t in texts if isinstance(t, str) and is_total_label(t)), None)
        if not lab:
            break
        out.append((r, f"الصف يبدأ بعبارة إجمالي «{lab}»"))
    return out


def new_record(ctx: Ctx, src: dict, rr: dict, scope: str, scope_reason, record_type=None) -> OrderedDict:
    sid, sheet, row = rr["source_id"], rr["sheet"], rr["row"]
    prefix, id_col, name_col, id_text = rr["prefix"], rr["id_col"], rr["name_col"], rr["id_text"]
    base_id = _candidate_id(prefix, id_text, row)
    rid = base_id
    if rid in ctx.seen_ids:
        rid = f"{base_id}-r{row}"
        n = 2
        while rid in ctx.seen_ids:
            rid = f"{base_id}-r{row}-{n}"
            n += 1
        first = ctx.seen_ids[base_id]
        ctx.issues.add("duplicate_id", "high",
                       join_refs([first["source"].get("ref_id"), ref(sheet, f"{id_col}{row}") if id_col else None]),
                       [first["id"], rid],
                       f"المعرّف «{base_id}» مكرر؛ أُعطي السجل في الصف {row} المعرّف «{rid}».",
                       "قد يؤدي إلى ازدواج الأصل في العدد والقيمة والإيراد إن كانا السجل نفسه",
                       "لم يُدمج السجلان؛ يلزم التحقق من المعرف والصك والموقع وتصحيح الترقيم أو استبعاد المكرر في mapping")
    rec = OrderedDict()
    rec["id"] = rid
    rec["name"] = rr["name"]
    rec["source"] = OrderedDict([("source_id", sid), ("sheet", sheet), ("row", row),
                                 ("ref_name", ref(sheet, f"{name_col}{row}") if name_col else None)])
    if id_col:
        rec["source"]["ref_id"] = ref(sheet, f"{id_col}{row}")
    if rr.get("merge_notes"):
        rec["source"]["merge_note"] = "؛ ".join(rr["merge_notes"])
    rec["record_type"] = record_type or src.get("default_record_type") or "asset"
    rec["scope"] = scope
    rec["scope_reason"] = scope_reason
    rec["dims"] = rr["dims"]
    rec["dims_evidence"] = rr["dims_ev"]
    rec["fields"] = rr["fields"]
    rec["issue_ids"] = []
    if rr.get("oov"):
        rec["_oov"] = OrderedDict(rr["oov"])
    if rr.get("total_reasons"):
        rec["_total"] = rr["total_reasons"]
    if rr.get("possible_total"):
        rec["_possible_total"] = rr["possible_total"]
    ctx.seen_ids[rid] = rec
    ctx.id_texts[rid] = id_text
    ctx.records.append(rec)
    return rec


def _join_norm(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    t = normalize_arabic(raw_str(v))
    if re.fullmatch(r"\d+\.0+", t):
        t = t.split(".")[0]
    return t


def _deed_key(v) -> str:
    k = re.sub(r"[\s()\[\]{}]", "", normalize_arabic(raw_str(v)))
    return k if k and re.search(r"\d", k) and not re.fullmatch(r"0+", k) else ""


def _join_keys(ctx: Ctx, rec: dict, on: str) -> list:
    if on == "id":
        k = _join_norm(ctx.id_texts.get(rec["id"]))
        return [k] if k else []
    if on == "name_key":
        k = name_key(rec.get("name")) if rec.get("name") else ""
        return [k] if k else []
    keys = []
    for f, tv in rec["fields"].items():
        if (f == "deed_no" or f.startswith("deed_no.")) and _known(tv):
            k = _deed_key(tv.get("v"))
            if k and k not in keys:
                keys.append(k)
    return keys


def _row_key(rr: dict, on: str) -> str:
    raw = rr.get("key_raw")
    if raw is None:
        return ""
    if on == "id":
        return _join_norm(clean_text(raw_str(raw)))
    if on == "name_key":
        return name_key(raw)
    return _deed_key(raw)


JOIN_TARGET_TYPES = ("asset", "land_parcel", "operating_activity")


def _join_eligible(ctx: Ctx, rec: dict) -> bool:
    """A join target that is not excluded and is an asset/land parcel/operating activity, counting records/ overrides
    (applied later in normalize) so that units set to excluded as validate instructs never compete with their parent."""
    ov = ctx.overrides.get(rec["id"]) or {}
    scope = ov.get("scope") or rec.get("scope")
    rtype = ov.get("record_type") or rec.get("record_type")
    return scope != "excluded" and rtype in JOIN_TARGET_TYPES


def merge_join_source(ctx: Ctx, src: dict, sheet: str, rows: list) -> None:
    """Merge a join source's mapped fields into the records of `join.into` (A7). Never creates in-scope records."""
    sid, j = src["id"], src["join"]
    into, on, link_id = j["into"], j["on"], j["link_id"]
    key_field = ((src["columns"].get(j["key_column"]) or {}).get("field"))
    targets = [r for r in ctx.records if r["source"]["source_id"] == into]
    index = {}
    for t in targets:
        for k in _join_keys(ctx, t, on):
            index.setdefault(k, []).append(t)
    unmatched, ambiguous, conflicts = [], [], []
    on_ar = {"id": "المعرّف", "deed": "رقم الصك", "name_key": "الاسم"}[on]
    for rr in rows:
        k = _row_key(rr, on)
        cands = index.get(k, []) if k else []
        # units/contracts and excluded rows usually share their parent's deed: match the parent (an in-scope asset,
        # land parcel or activity, after records/ overrides); fall back to every candidate only when none qualifies
        cands = [t for t in cands if _join_eligible(ctx, t)] or cands
        if len(cands) == 1:
            tgt = cands[0]
            ctx.stats["rows_joined"] += 1
            tgt["source"].setdefault("joined", []).append(
                OrderedDict([("source_id", sid), ("sheet", sheet), ("row", rr["row"]), ("link_id", link_id)]))
            for f, tv in rr["fields"].items():
                if f == key_field:
                    continue
                tv = dict(tv)
                tv["joined_from"] = sid
                cur = tgt["fields"].get(f)
                if cur is not None and cur.get("s") not in (S_BLANK, S_FILLER):
                    same = _known(cur) and _known(tv) and cur.get("v") == tv.get("v")
                    if not same and tv.get("s") not in (S_BLANK, S_FILLER):
                        conflicts.append((tgt["id"], f, cur.get("ref"), tv.get("ref")))
                    continue
                if tv.get("s") in (S_BLANK, S_FILLER) and cur is not None:
                    continue
                if cur is not None:
                    tv["note"] = ((tv.get("note") or "") + f" (حلّت محل خانة فارغة في {cur.get('ref')})").strip()
                tgt["fields"][f] = tv
            for dim, ev in rr["dims_ev"].items():
                val = rr["dims"].get(dim)
                if dim in rr.get("oov", {}):
                    if is_unknown_dim(tgt["dims"].get(dim)):
                        tgt.setdefault("_oov", OrderedDict())[dim] = rr["oov"][dim]
                    continue
                if is_unknown_dim(val):
                    continue
                cur = tgt["dims"].get(dim)
                if is_unknown_dim(cur):
                    tgt["dims"][dim] = copy.deepcopy(val)
                    tgt["dims_evidence"][dim] = f"{ev} (مدمج من المصدر {sid})"
                elif cur != val:
                    conflicts.append((tgt["id"], f"dims.{dim}", tgt["dims_evidence"].get(dim), ev))
            continue
        if len(cands) > 1:
            reason = (f"صف من مصدر الربط {sid}: {on_ar} «{clean_text(raw_str(rr.get('key_raw')))}» يطابق أكثر من سجل في "
                      f"{into} ({'، '.join(c['id'] for c in cands)})؛ لم يُدمج")
            rec = new_record(ctx, src, rr, "isolated", reason, record_type=src.get("default_record_type") or "other")
            ambiguous.append((rec, cands, rr))
        else:
            reason = (f"صف من مصدر الربط {sid} لم يطابق أي سجل في {into} ({on_ar} "
                      f"«{clean_text(raw_str(rr.get('key_raw'))) or 'فارغ'}»)؛ عُزل ولم يُدمج")
            rec = new_record(ctx, src, rr, "isolated", reason, record_type=src.get("default_record_type") or "other")
            unmatched.append((rec, rr))
    if unmatched:
        ctx.issues.add("join_unmatched", "medium", join_refs([rr.get("key_ref") for _, rr in unmatched]),
                       [r["id"] for r, _ in unmatched],
                       f"في مصدر الربط {sid} {count_ar(len(unmatched), 'row')} بلا سجل مطابق ب{on_ar} في المصدر {into}.",
                       "لا تدخل أرقامها أي مؤشر؛ قد تكون بيانات أصل غير مسجل أو مفتاحًا مكتوبًا بصيغة مختلفة",
                       "حُفظت الصفوف سجلاتٍ معزولة (isolated) ولم تُدمج؛ تحقق من المفتاح قبل أي استخدام", source_id=sid)
    for rec, cands, rr in ambiguous:
        ctx.issues.add("join_ambiguous", "high", join_refs([rr.get("key_ref")] + [c["source"].get("ref_name") for c in cands]),
                       [rec["id"]] + [c["id"] for c in cands],
                       f"مفتاح الربط في {rr.get('key_ref')} ({on_ar} «{clean_text(raw_str(rr.get('key_raw')))}») يطابق "
                       f"{count_ar(len(cands), 'record')} في {into}: {'، '.join(c['id'] for c in cands)}.",
                       "لم تُدمج البيانات في أي منها؛ لو دُمجت بالتخمين لنُسبت لأصل خاطئ",
                       "عُزل الصف سجلًا مستقلًا؛ صحّح المفتاح أو استخدم مفتاحًا أدق (المعرّف أو الصك)", source_id=sid)
    if conflicts:
        ctx.issues.add("join_conflict", "medium", join_refs([c[3] for c in conflicts] + [c[2] for c in conflicts]),
                       sorted({c[0] for c in conflicts}),
                       f"في مصدر الربط {sid} {count_ar(len(conflicts), 'value')} لحقول موجودة أصلًا بقيم أخرى في سجلات {into}: "
                       + "؛ ".join(f"{c[0]}.{c[1]}" for c in conflicts[:8]) + ("…" if len(conflicts) > 8 else "") + ".",
                       "بقيت قيم السجل الهدف كما هي؛ إن كانت قيمة المصدر المربوط هي الصحيحة فالمؤشرات تستخدم حاليًا القيمة الأخرى",
                       "أُبقيت قيمة السجل الهدف ولم تُستبدل؛ وثّق أي تصحيح عبر field_overrides بدليل", source_id=sid)


def normalize(xlsx: str, mapping: dict, wb_map=None) -> dict:
    cur = mapping.get("currency")
    HOME_CURRENCY[0] = (cur.get("value") if isinstance(cur, dict) else cur) or "SAR"
    ctx = Ctx(xlsx, mapping, wb_map)
    issues, meta = ctx.issues, ctx.meta
    sources = mapping["sources"]
    plain = [s for s in sources if not s.get("join")]
    joins = [s for s in sources if s.get("join")]
    for src in plain + joins:
        sid = src["id"]
        sheet = resolve_sheet(ctx.wb_v, src["sheet"], issues)
        scope, scope_reason = source_scope(ctx, src, sheet)
        rows = read_source(ctx, src, sheet)
        dr = src["data_rows"]
        if not rows:
            raise SystemExit(f"خطأ: المصدر «{sid}» لم يُنتج أي سجل من الصفوف {dr['from']}–{dr['to']} في الورقة «{sheet}» "
                             "(كل الصفوف فارغة أو مستبعدة أو إجماليات). صحّح data_rows أو exclude_rows أو skip_if_blank "
                             "في 02_mapping.json.")
        if src.get("join"):
            merge_join_source(ctx, src, sheet, rows)
            continue
        for rr in rows:
            rec = new_record(ctx, src, rr, scope, scope_reason)
            ctx.row_owner.setdefault((sheet, rr["row"]), []).append(rec["id"])

    records = ctx.records
    by_id = {r["id"]: r for r in records}
    apply_derived(records, mapping.get("derived_fields") or [], issues, meta)
    apply_dim_rules(records, mapping.get("dim_rules") or [])
    apply_overrides(records, by_id, ctx.overrides, issues, meta)
    apply_rollups(records, by_id, ctx.overrides, issues, meta)
    for r in records:
        d = r["dims"]
        if "managed_for_others" not in d or d.get("managed_for_others") is None:
            own = d.get("ownership") or []
            if "managed_for_others" in own:
                d["managed_for_others"] = True
            elif not own or own == ["unknown"]:
                d["managed_for_others"] = None
            else:
                d["managed_for_others"] = False

    # ---- automatic checks -------------------------------------------------
    for rec in records:
        weak = rec.pop("_possible_total", None)
        if weak:
            src = rec["source"]
            issues.add("possible_total_row", "medium", src.get("ref_name") or ref(src["sheet"], f"A{src['row']}"),
                       [rec["id"]], f"الصف {src['row']} في المصدر {src['source_id']} له معرّف واسم أصل («{rec.get('name')}») "
                       "لكن فيه علامات إجمالي: " + "؛ ".join(weak) + ". أُبقي سجلًا ولم يُستبعد آليًا.",
                       "إن كان يجمع سطورًا تابعة (وحدات أو عقود) مدرجة ضمن النطاق فقد تتضاعف المبالغ",
                       "تحقق من الصف: إن كان إجماليًا فأضفه إلى exclude_rows؛ وإن كان أصلًا تجمع صيغته وحداته فأبقِ أحد "
                       "المستويين فقط ضمن النطاق لتجنب الازدواج")
        why = rec.pop("_total", None)
        if why:
            issues.add("total_like_row", "high", rec["source"]["ref_name"] or ref(rec["source"]["sheet"], f"A{rec['source']['row']}"),
                       [rec["id"]], f"صف داخل نطاق البيانات يبدو صف إجمالي: {'؛ '.join(why)}. أبقاه تجاوز صريح في "
                       f"02_mapping.json/records/{rec['id']} (النوع «{rec['record_type']}»، النطاق «{rec['scope']}»).",
                       "إن كان إجماليًا فسيُضاعف العدد والقيمة والإيراد",
                       "طُبّق تجاوز mapping كما ورد؛ يلزم دليل صريح على أنه ليس إجماليًا", mapping_override=True)
    check_scope_decisions(ctx)
    check_overlaps(ctx)
    check_out_of_vocabulary(records, issues)
    check_valuation_totals(records, issues, meta)
    check_vat(records, issues, meta)
    check_field_statuses(records, issues, meta, ctx.field_kinds)
    check_names(records, issues)
    check_deeds(records, issues)
    check_classification(records, issues)

    items = issues.finalize(by_id)
    for it in items:
        for rid in it["record_ids"]:
            if rid in by_id:
                by_id[rid]["issue_ids"].append(it["id"])
    for r in records:
        r.pop("_oov", None)

    st = ctx.stats
    stats = OrderedDict()
    stats["records"] = len(records)
    stats["by_type"] = dict(sorted(_count(r["record_type"] for r in records).items()))
    stats["by_scope"] = dict(sorted(_count(r["scope"] for r in records).items()))
    stats["by_source"] = dict(sorted(_count(r["source"]["source_id"] for r in records).items()))
    stats["pii_columns_dropped"] = st["pii_columns"]
    stats["pii_values_dropped"] = st["pii_values"]
    stats["rows_skipped_blank"] = st["rows_skipped"]
    stats["rows_skipped_with_content"] = st["rows_skipped_with_content"]
    stats["rows_excluded"] = st["rows_excluded"]
    stats["rows_total_excluded"] = st["rows_total_excluded"]
    stats["rows_total_after_range"] = st["rows_total_after_range"]
    stats["rows_joined"] = st["rows_joined"]
    stats["issues"] = len(items)
    stats["issues_by_severity"] = dict(sorted(_count(i["severity"] for i in items).items()))

    out = OrderedDict()
    out["contract_version"] = CONTRACT_VERSION
    out["source"] = OrderedDict([("file", xlsx), ("name", os.path.basename(xlsx)), ("sha256", sha256_file(xlsx)),
                                 ("mapping_source_file", mapping.get("source_file"))])
    out["context"] = OrderedDict((k, copy.deepcopy(mapping.get(k))) for k in
                                 ("entity", "data_period", "study_date", "analysis_date", "currency", "valuation_dates",
                                  "sheets"))
    out["currency"] = mapping.get("currency")
    out["field_meta"] = copy.deepcopy(meta)
    out["records"] = records
    out["excluded_rows"] = ctx.excluded_rows
    out["issues"] = items
    out["stats"] = stats
    return out


def _count(it):
    d = {}
    for x in it:
        d[x] = d.get(x, 0) + 1
    return d


# ---------------------------------------------------------------------------
# Derived fields, dim rules, overrides
# ---------------------------------------------------------------------------
def apply_derived(records: list, rules: list, issues: Issues, meta: dict) -> None:
    for rule in rules:
        field, formula, when = rule["field"], rule["formula"], rule["when"]
        try:
            formula_fields(formula)
        except ValueError as exc:
            raise SystemExit(f"خطأ في derived_fields ({field}): {exc}")
        only = set(rule.get("records") or [])
        only_src = set(rule.get("sources") or [])
        changed, failed, refs = [], [], []
        for r in records:
            if only and r["id"] not in only:
                continue
            if only_src and r["source"]["source_id"] not in only_src:
                continue
            cur = r["fields"].get(field)
            if when == "if_unknown" and cur is not None and cur.get("s") in KNOWN_STATUSES:
                continue
            try:
                value, missing, inputs = eval_formula(formula, r["fields"])
                err = None
            except (ValueError, ZeroDivisionError) as exc:
                value, missing, inputs, err = None, [], [], str(exc)
            if value is None:
                if when == "if_unknown":
                    failed.append(r["id"])
                    continue
                new = {"v": None, "s": S_NOT_AVAILABLE, "raw": None, "ref": None,
                       "note": f"تعذر الاشتقاق ({formula}): " + (err or ("حقول غير معروفة: " + "، ".join(missing)))}
                failed.append(r["id"])
            else:
                new = {"v": value, "s": S_DERIVED, "raw": None,
                       "ref": " + ".join(i["ref"] for i in inputs if i.get("ref")) or None,
                       "note": f"مشتق: {formula} — {rule['reason']}", "formula": formula, "inputs": inputs}
                changed.append(r["id"])
            if cur is not None and when == "always":
                r["fields"][field + "__original"] = copy.deepcopy(cur)
            if cur is not None and cur.get("ref"):
                refs.append(cur["ref"])
            new["derivation"] = {"rule": formula, "when": when}
            r["fields"][field] = new
        if changed or failed:
            if when == "always":
                treat = (f"استُبدلت قيمة {field} بالقيمة المشتقة ({formula}) وحُفظت القيمة الأصلية في الحقل "
                         f"{field}__original للمراجعة")
            else:
                treat = f"مُلئت قيمة {field} غير المعروفة بالقيمة المشتقة ({formula}) دون تغيير القيم المعروفة"
            detail = f"اشتقاق صريح من mapping: {rule['reason']}. سجلات مشتقة: {len(changed)}"
            if failed:
                detail += f"؛ تعذر الاشتقاق لـ {len(failed)} سجل (بقيت غير معروفة)"
            issues.add("derived_field", "medium" if when == "always" else "low", join_refs(refs) or field,
                       changed + failed, detail, impact_for(field, "money", meta), treat, field=field)


def _norm_key(x) -> str:
    return normalize_arabic(x) if x is not None else ""


def apply_dim_rules(records: list, rules: list) -> None:
    for rule in rules:
        dim, src_field = rule["dim"], rule["from_field"]
        mp = [(_norm_key(k), k, v) for k, v in rule["map"].items()]
        contains = rule.get("match") == "contains"
        only_src = set(rule.get("sources") or [])
        for r in records:
            if only_src and r["source"]["source_id"] not in only_src:
                continue
            tv = r["fields"].get(src_field)
            text = None
            if tv and tv.get("s") in KNOWN_STATUSES and tv.get("v") is not None:
                text = raw_str(tv["v"])
            key = _norm_key(text) if text else ""
            matched = None
            if key:
                for nk, k, v in mp:
                    if nk == key:
                        matched = v
                        break
                if matched is None and contains:
                    for nk, k, v in sorted(mp, key=lambda t: (-len(t[0]), t[0])):
                        if nk and nk in key:
                            matched = v
                            break
            dims = r["dims"]
            if matched is not None:
                if dim in LIST_DIMS:
                    vals = matched if isinstance(matched, list) else [matched]
                    cur = [x for x in (dims.get(dim) or []) if x != "unknown"]
                    for x in vals:
                        if x not in cur:
                            cur.append(x)
                    dims[dim] = cur or ["unknown"]
                else:
                    dims[dim] = matched
                r["dims_evidence"][dim] = f"{tv.get('ref')} («{text}»)"
                (r.get("_oov") or {}).pop(dim, None)
            elif "default" in rule:
                cur = dims.get(dim)
                unset = cur is None or cur == "unknown" or cur == ["unknown"] or cur == []
                if unset:
                    d = rule["default"]
                    if dim in LIST_DIMS:
                        d = d if isinstance(d, list) else ([d] if d is not None else ["unknown"])
                    dims[dim] = d
                    if text:
                        r["dims_evidence"].setdefault(dim, f"{tv.get('ref')} («{text}» — غير مطابق؛ القيمة الافتراضية)")


def apply_overrides(records: list, by_id: dict, overrides: dict, issues: Issues, meta: dict) -> None:
    for rid in sorted(overrides):
        ov = overrides[rid]
        r = by_id.get(rid)
        if r is None:
            issues.add("mapping_unknown_record", "medium", "02_mapping.json/records/" + rid, [],
                       f"يوجد تجاوز في mapping للمعرّف «{rid}» لكن لا يوجد سجل بهذا المعرّف.",
                       "لم يُطبَّق التصنيف المقصود؛ قد يبقى سجل بلا تصنيف",
                       "صحّح المعرّف في 02_mapping.json/records")
            eprint(f"تحذير: تجاوز لمعرّف غير موجود: {rid}")
            continue
        for k in ("record_type", "scope", "scope_reason", "name"):
            if k in ov:
                r[k] = ov[k]
        if "notes" in ov:
            r["notes"] = ov["notes"]
        for k, v in (ov.get("dims") or {}).items():
            r["dims"][k] = copy.deepcopy(v)
            (r.get("_oov") or {}).pop(k, None)
        for k, v in (ov.get("dims_evidence") or {}).items():
            r["dims_evidence"][k] = v
        for field, fo in sorted((ov.get("field_overrides") or {}).items()):
            s = fo["s"]
            if s not in ALL_STATUSES:
                raise SystemExit(f"خطأ: حالة غير معروفة «{s}» في records/{rid}/field_overrides/{field}")
            cur = r["fields"].get(field)
            new = copy.deepcopy(cur) if cur else {"v": None, "s": None, "raw": None, "ref": None, "note": None}
            prev_s = cur.get("s") if cur else None
            prev_v = cur.get("v") if cur else None
            if s in (S_ZERO_RECORDED, S_ZERO_CONFIRMED):
                v = 0
            elif s in (S_VALUE, S_DERIVED):
                v = fo.get("v")
                if not isinstance(v, (int, float)) or isinstance(v, bool):
                    raise SystemExit(f"خطأ: records/{rid}/field_overrides/{field}: الحالة {s} تتطلب قيمة رقمية v")
                v = tidy_num(float(v)) if isinstance(v, float) else v
            else:
                v = None
            new["v"] = v
            new["s"] = s
            new["override"] = {"from_s": prev_s, "from_v": prev_v, "evidence": fo["evidence"]}
            new["note"] = fo.get("note") or f"تجاوز موثق: {fo['evidence']}"
            r["fields"][field] = new
            issues.add("field_override", "low", (cur or {}).get("ref") or f"{rid}:{field}", [rid],
                       f"تعديل يدوي موثق في mapping: {field} من الحالة {prev_s} (القيمة {fmt(prev_v)}) إلى {s}"
                       f" (القيمة {fmt(v)}). الدليل: {fo['evidence']}",
                       impact_for(field, "money", meta), "طُبّق التعديل كما ورد في 02_mapping.json مع حفظ الحالة السابقة",
                       field=field)


def apply_rollups(records: list, by_id: dict, overrides: dict, issues: Issues, meta: dict) -> None:
    """records/<parent>.rollup: sum the children's known values into the parent's unknown fields (revenue recorded on
    unit rows while the building row holds only the valuation). Children go out of scope so nothing is counted twice."""
    for rid in sorted(overrides):
        ru = (overrides[rid] or {}).get("rollup")
        if not ru:
            continue
        parent = by_id.get(rid)
        missing = [c for c in [rid] + list(ru["from"]) if c not in by_id]
        if parent is None or missing:
            issues.add("rollup_unknown_record", "high", f"02_mapping.json/records/{rid}/rollup", [],
                       "تجميع الوحدات في الأصل يذكر معرّفات لا سجل لها: " + "، ".join(missing) + ".",
                       "لم يُطبَّق التجميع؛ تبقى مبالغ الوحدات خارج الأصل", "صحّح المعرّفات في records/<id>/rollup")
            continue
        kids = [by_id[c] for c in ru["from"]]
        for k in kids:
            if "scope" not in (overrides.get(k["id"]) or {}):
                k["scope"] = "excluded"
                k["scope_reason"] = f"جُمعت مبالغه في الأصل {rid} (records/{rid}/rollup)"
        done, incomplete, conflicts = [], [], []
        for f in ru["fields"]:
            tvs = [(k, (k.get("fields") or {}).get(f)) for k in kids]
            unknown = [k["id"] for k, tv in tvs if not _known(tv)]
            total = None if unknown else tidy_num(sum(tv["v"] for _k, tv in tvs))
            cur = parent["fields"].get(f)
            if _known(cur):
                if total is None or abs(float(cur["v"]) - float(total)) > 0.5:
                    conflicts.append((f, cur.get("v"), total, unknown))
                continue
            if unknown:
                incomplete.append((f, unknown))
                continue
            refs = [tv.get("ref") for _k, tv in tvs if tv.get("ref")]
            new = OrderedDict([("v", total), ("s", S_DERIVED), ("raw", None), ("ref", join_refs(refs)),
                               ("note", "مجموع " + "، ".join(f"{k['id']} ({fmt(tv['v'])})" for k, tv in tvs)
                                + f" — تجميع موثق: {ru['evidence']}"),
                               ("rollup", OrderedDict([("from", list(ru["from"])), ("evidence", ru["evidence"])]))])
            if cur:
                new["from_parent_cell"] = OrderedDict([("s", cur.get("s")), ("ref", cur.get("ref"))])
            parent["fields"][f] = new
            done.append((f, total))
        ids = [rid] + list(ru["from"])
        if done:
            issues.add("rollup_applied", "low", f"02_mapping.json/records/{rid}/rollup", ids,
                       f"جُمعت مبالغ {len(kids)} سجلات تابعة في الأصل {rid}: "
                       + "؛ ".join(f"{f} = {fmt(t)}" for f, t in done) + f". الدليل: {ru['evidence']}",
                       "تدخل المبالغ مرة واحدة على مستوى الأصل؛ السجلات التابعة خارج النطاق",
                       "طُبّق التجميع كما ورد في 02_mapping.json")
        for f, unknown in incomplete:
            issues.add("rollup_incomplete", "medium", f"02_mapping.json/records/{rid}/rollup", ids,
                       f"تعذر تجميع {f} في الأصل {rid}: القيمة غير معروفة في " + "، ".join(unknown) + ".",
                       impact_for(f, "money", meta),
                       "بقي الحقل غير معروف في الأصل؛ لا يُجمع مجموع جزئي. وثّق القيم الناقصة أو أخرجها من التجميع")
        for f, cur_v, total, unknown in conflicts:
            issues.add("rollup_conflict", "medium", (parent["fields"].get(f) or {}).get("ref") or rid, ids,
                       f"الأصل {rid} يحمل قيمة معروفة في {f} ({fmt(cur_v)}) "
                       + (f"تخالف مجموع السجلات التابعة ({fmt(total)})" if total is not None
                          else "وبعض السجلات التابعة غير معروفة القيمة") + ".",
                       impact_for(f, "money", meta),
                       "أُبقيت قيمة الأصل ولم يُستبدل بها المجموع؛ والسجلات التابعة خارج النطاق")


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------
def _known(tv):
    return tv is not None and tv.get("s") in KNOWN_STATUSES


def check_scope_decisions(ctx: Ctx) -> None:
    """A9: source scope vs sheets[] decision, hidden sheets, managed-properties sheets."""
    issues = ctx.issues
    for sid, sheet, problems in ctx.scope_problems:
        rids = [r["id"] for r in ctx.records if r["source"]["source_id"] == sid]
        issues.add("sheet_scope_mismatch", "high", f"02_mapping.json/sources/{sid} ({sheet})", rids,
                   f"قرار نطاق المصدر {sid} (الورقة «{sheet}») غير متسق: " + "؛ ".join(problems) + ".",
                   "قد تدخل المؤشرات بيانات ورقة لم يتحقق انتماؤها للمحفظة، أو تُستبعد بيانات معتمدة",
                   "لم يُغيَّر نطاق السجلات المذكور صراحة؛ اجعل default_scope وscope_reason للمصدر موافقين لقرار الورقة "
                   "في sheets[]", source_id=sid)
    for sid, sheet in ctx.managed_sources:
        bad = [r for r in ctx.records if r["source"]["source_id"] == sid and r["scope"] == "in_scope"
               and r["dims"].get("managed_for_others") is not True]
        if bad:
            issues.add("managed_unflagged", "high", join_refs([r["source"].get("ref_name") for r in bad]),
                       [r["id"] for r in bad],
                       f"{count_ar(len(bad), 'record')} ضمن النطاق من ورقة «{sheet}» (دورها managed_properties) بلا "
                       "dims.managed_for_others = true.",
                       "تدخل قيمة عقارات الغير وكامل إيراداتها في قيمة الجمعية وإيراداتها وعوائدها",
                       "صنّف كل سجل من الورقة بـ managed_for_others: true (وownership يتضمن managed_for_others)، أو اعزل "
                       "المصدر", source_id=sid)


def check_overlaps(ctx: Ctx) -> None:
    """A10: the same sheet row read into more than one record (join sources excluded)."""
    groups = OrderedDict()
    for (sheet, row), rids in ctx.row_owner.items():
        if len(rids) < 2:
            continue
        srcs = tuple(sorted({ctx.seen_ids[r]["source"]["source_id"] for r in rids}))
        g = groups.setdefault((sheet, srcs), {"rows": [], "ids": []})
        g["rows"].append(row)
        g["ids"].extend(rids)
    for (sheet, srcs), g in groups.items():
        ctx.issues.add("overlapping_source_rows", "high", f"{sheet}!{rows_text(g['rows'])}", g["ids"],
                       f"المصادر {'، '.join(srcs)} تقرأ الصفوف نفسها من الورقة «{sheet}» ({rows_text(g['rows'])})؛ "
                       f"كل صف صار أكثر من سجل: {'، '.join(g['ids'][:12])}" + ("…" if len(g["ids"]) > 12 else "") + ".",
                       "يتضاعف عدد الأصول وقيمتها وإيرادها لهذه الصفوف",
                       "لم تُدمج السجلات؛ اجعل نطاقات data_rows للمصادر غير متداخلة أو استبعد الصفوف المكررة من أحدها",
                       rows=g["rows"])


def check_out_of_vocabulary(records: list, issues: Issues) -> None:
    """A6: cell values of controlled dimensions outside assets/vocab.json are stored as unknown."""
    groups = OrderedDict()
    for r in records:
        for dim, (cref, raw, field) in (r.get("_oov") or {}).items():
            g = groups.setdefault((dim, field), {"ids": [], "refs": [], "examples": []})
            g["ids"].append(r["id"])
            g["refs"].append(cref)
            if raw is not None and raw not in g["examples"] and len(g["examples"]) < 4:
                g["examples"].append(raw)
    for (dim, field), g in groups.items():
        issues.add("dim_out_of_vocabulary", "high", join_refs(g["refs"]), g["ids"],
                   f"قيم البعد «{dim}» من الحقل {field} خارج المفردات المعتمدة في {count_ar(len(g['ids']), 'record')}: "
                   + "؛ ".join(f"«{x}»" for x in g["examples"]) + f". المسموح: {_vocab_text(dim)}.",
                   "خُزّن البعد «غير محدد» فتُستبعد هذه السجلات من المؤشرات المقيدة بهذا البعد",
                   f"لم يُخمَّن التصنيف؛ حوّل النصوص إلى الرموز عبر dim_rules (dim: {dim}) أو records بدليل",
                   field=field, dim=dim)


def check_valuation_totals(records: list, issues: Issues, meta: dict) -> None:
    for r in records:
        years = {}
        for f in r["fields"]:
            m = VAL_RE.match(f)
            if m:
                years.setdefault(m.group(1), {})[m.group(2)] = f
        for y in sorted(years):
            parts = years[y]
            for comp in ("land", "building"):
                ctv = r["fields"].get(parts.get(comp)) if parts.get(comp) else None
                if ctv and _known(ctv) and ctv.get("formula"):
                    issues.add("valuation_component_formula", "low", ctv.get("ref"), [r["id"]],
                               f"قيمة {'الأرض' if comp == 'land' else 'المبنى'} لتقييم {y} محسوبة بصيغة ({ctv['formula'][:120]}) "
                               "وليست مدخلة مباشرة؛ قد تكون مشتقة من الإجمالي وليست تقييمًا مستقلًا.",
                               impact_for(parts[comp], "money", meta),
                               "استُخدمت القيمة المخزنة كما هي؛ لا تُجمع مع الإجمالي، ويلزم التحقق من تقرير التقييم")
            if not all(k in parts for k in ("total", "land", "building")):
                continue
            t, la, b = (r["fields"][parts[k]] for k in ("total", "land", "building"))
            if not (_known(t) and _known(la) and _known(b)):
                continue
            s = la["v"] + b["v"]
            diff = t["v"] - s
            if abs(diff) > max(1.0, abs(t["v"]) * 1e-9):
                extra = ""
                if t.get("formula"):
                    extra = f" الصيغة المدونة: {t['formula'][:160]}"
                issues.add("valuation_total_mismatch", "high", join_refs([t.get("ref"), la.get("ref"), b.get("ref")]),
                           [r["id"]],
                           f"إجمالي تقييم {y} المدون = {fmt(t['v'])} بينما الأرض + المبنى = {fmt(la['v'])} + {fmt(b['v'])}"
                           f" = {fmt(s)} (فرق {fmt(diff)}).{extra}",
                           impact_for(parts["total"], "money", meta),
                           "لم يُصحَّح تلقائيًا؛ إن ثبت خطأ الصيغة فأضف derived_fields صريحًا أو استخدم سنة تقييم أخرى")


def _vat_pairs(meta: dict, fields_present: set) -> dict:
    """period -> (excl_field, incl_field)"""
    per = {}
    for f in sorted(set(meta) | fields_present):
        m = meta.get(f) or {}
        if group_of(f, meta) != "revenue":
            continue
        vat = m.get("vat")
        if vat is None:
            if "excl" in f:
                vat = "excl"
            elif "incl" in f:
                vat = "incl"
        if vat not in ("excl", "incl"):
            continue
        period = m.get("period")
        if not period:
            mm = re.match(r"^rev\.([^.]+)\.", f)
            period = mm.group(1) if mm else None
        if not period:
            continue
        per.setdefault(period, {}).setdefault(vat, f)
    return {p: (d["excl"], d["incl"]) for p, d in per.items() if "excl" in d and "incl" in d}


def check_vat(records: list, issues: Issues, meta: dict) -> None:
    present = set()
    for r in records:
        present.update(r["fields"].keys())
    pairs = _vat_pairs(meta, present)
    for p in sorted(pairs):
        ex_f, in_f = pairs[p]
        equal_ids, equal_refs = [], []
        for r in records:
            ex, inc = r["fields"].get(ex_f), r["fields"].get(in_f)
            if not (_known(ex) and _known(inc)):
                continue
            if inc["v"] < ex["v"]:
                issues.add("vat_incl_lt_excl", "high", join_refs([inc.get("ref"), ex.get("ref")]), [r["id"]],
                           f"الإيراد مع الضريبة ({fmt(inc['v'])}) أقل من الإيراد بدون ضريبة ({fmt(ex['v'])}) للفترة {p}.",
                           impact_for(ex_f, "money", meta),
                           "لم تُعدَّل القيم؛ يلزم التحقق من المصدر قبل اعتماد أي منهما")
            elif inc["v"] == ex["v"] and ex["v"] != 0:
                equal_ids.append(r["id"])
                equal_refs.extend([ex.get("ref"), inc.get("ref")])
        if equal_ids:
            issues.add("vat_equal", "low", join_refs(equal_refs), equal_ids,
                       f"الإيراد مع الضريبة يساوي الإيراد بدون ضريبة للفترة {p} في {len(equal_ids)} سجل؛ "
                       "قد يعني عدم خضوع الإيراد للضريبة أو عدم فصلها في الملف — لا يثبت أيهما من الملف وحده.",
                       impact_for(ex_f, "money", meta),
                       "استُخدم الحقل بدون ضريبة كما هو؛ لم تُطبَّق نسبة ضريبة افتراضية")


def check_field_statuses(records: list, issues: Issues, meta: dict, kinds: dict) -> None:
    groups = OrderedDict()

    def add(key, rid, tv, extra=None):
        g = groups.setdefault(key, {"ids": [], "refs": [], "examples": []})
        g["ids"].append(rid)
        g["refs"].append(tv.get("ref"))
        if extra and len(g["examples"]) < 3 and extra not in g["examples"]:
            g["examples"].append(extra)

    for r in records:
        for f, tv in r["fields"].items():
            if f.endswith("__original"):
                continue
            kind = kinds.get(f)
            s = tv.get("s")
            if tv.get("redacted"):
                add(("pii_redacted", f), r["id"], tv)
                continue
            if tv.get("no_cache"):
                add(("formula_no_cache", f), r["id"], tv)
            if tv.get("zero_from_blank_inputs"):
                add(("zero_from_blank_inputs", f), r["id"], tv, f"{(tv.get('ref') or '').split('!')[-1]}: {tv.get('formula')}")
                continue
            if kind in NUMERIC_KINDS:
                if tv.get("percent_in_amount"):
                    add(("percent_in_amount", f), r["id"], tv, tv.get("raw"))
                elif tv.get("foreign_currency"):
                    add(("foreign_currency", f), r["id"], tv, tv.get("raw"))
                elif tv.get("ambiguous_thousands"):
                    add(("ambiguous_thousands_separator", f), r["id"], tv, tv.get("raw"))
                elif s == S_TEXT:
                    add(("text_in_numeric", f), r["id"], tv, tv.get("raw"))
                elif s == S_ERROR:
                    add(("excel_error", f), r["id"], tv, tv.get("raw"))
                elif s == S_ZERO_AS_MISSING:
                    add(("zero_as_missing", f), r["id"], tv)
                elif tv.get("note") and "نص رقمي" in (tv.get("note") or ""):
                    add(("numeric_text_converted", f), r["id"], tv, tv.get("raw"))
            elif kind == "count_mixed":
                if s == S_TEXT:
                    add(("count_mixed_text", f), r["id"], tv, tv.get("raw"))
                elif s == S_ERROR:
                    add(("excel_error", f), r["id"], tv, tv.get("raw"))
            elif kind == "date":
                if tv.get("calendar") == "hijri" and tv.get("v"):
                    add(("hijri_conversion", f), r["id"], tv, f"{tv.get('raw')} → {tv.get('v')}")
                if s == S_TEXT:
                    add(("date_unparsed", f), r["id"], tv, f"{tv.get('raw')}: {tv.get('note')}")
                elif tv.get("note") and s not in (S_NOT_AVAILABLE, S_NOT_APPLICABLE):
                    add(("date_note", f), r["id"], tv, f"{tv.get('raw')}: {tv.get('note')}")

    # one percent field holding both fractions (0.5) and percentages (50): the scale is mixed and shares cannot be summed
    for f, kind in kinds.items():
        if kind != "percent":
            continue
        small, big = [], []
        for r in records:
            tv = (r.get("fields") or {}).get(f)
            if not isinstance(tv, dict) or tv.get("s") not in KNOWN_STATUSES or not isinstance(tv.get("v"), (int, float)):
                continue
            (small if 0 < tv["v"] <= 1 else big if tv["v"] > 1 else []).append((r["id"], tv))
        if small and big:
            ex = [f"{rid}: {tv.get('raw')}" for rid, tv in (small[:2] + big[:2])]
            issues.add("percent_scale_mixed", "medium", join_refs([tv.get("ref") for _, tv in small + big]),
                       [rid for rid, _ in small + big],
                       f"الحقل {f} يخلط نسبًا مدونة كسورًا (≤ 1: {len(small)} سجل) ونسبًا مئوية (> 1: {len(big)} سجل)"
                       " — أمثلة: " + "؛ ".join(f"«{x}»" for x in ex) + ".",
                       impact_for(f, kind, meta),
                       "خُزّنت القيم كما هي؛ وحّد المقياس (0.5 تعني 50%؟) في records أو derived_fields مع الدليل قبل "
                       "استعمالها في حصة ملكية أو نسبة", field=f)

    texts = {
        "text_in_numeric": ("medium", "نص في حقل رقمي ({f}) في {n} سجل",
                            "لم يُحوَّل النص إلى رقم ولم يُعامل كصفر؛ القيمة غير معروفة مع حفظ النص الأصلي"),
        "percent_in_amount": ("medium", "نسبة مئوية مدونة في حقل مبلغ أو مساحة أو عدد ({f}) في {n} سجل",
                              "لم تُحوَّل إلى رقم ولم تُجمع؛ القيمة غير معروفة مع حفظ النص الأصلي، ويلزم التحقق من "
                              "المبلغ الفعلي في المصدر"),
        "foreign_currency": ("high", "مبالغ بعملة غير عملة الدراسة في الحقل {f} ({n} سجل)",
                             "لم تُحوَّل إلى أرقام ولم تُجمع (status text)؛ لا يُجمع مبلغان بعملتين. حوّل المبلغ بسعر "
                             "صرف موثق في derived_fields أو records مع الدليل، أو اتركه غير معروف"),
        "ambiguous_thousands_separator": ("medium", "أرقام نصية بفاصل ملتبس في الحقل {f} ({n} سجل)",
                                          "نقطة يتبعها ثلاث خانات بالضبط (مثل «12.500») قد تكون فاصل آلاف (12,500) "
                                          "أو فاصلة عشرية (12.5)؛ لم تُحوَّل إلى رقم. تحقق من المصدر وثبّت القيمة في "
                                          "records مع الدليل"),
        "zero_from_blank_inputs": ("medium", "صيغ ناتجها المخزن 0 لأن مدخلاتها فارغة في الحقل {f} ({n} سجل)",
                                   "عوملت القيمة كغير متاحة (not_available) لا كصفر مسجل؛ الصفر هنا نتيجة خانات فارغة "
                                   "وليس رقمًا مدونًا"),
        "excel_error": ("high", "خطأ صيغة Excel في الحقل {f} في {n} سجل",
                        "عوملت القيمة كغير معروفة (error)؛ يلزم تصحيح الصيغة في المصدر"),
        "zero_as_missing": ("medium", "قيمة صفرية في الحقل {f} ({n} سجل) حيث لا يُعد الصفر قيمة مقبولة",
                            "عومل الصفر كغير متاح (zero_as_missing) ولم يدخل في المجاميع"),
        "numeric_text_converted": ("low", "أرقام مدونة كنص في الحقل {f} ({n} سجل)",
                                   "حُوِّلت إلى أرقام مع حفظ النص الأصلي في raw"),
        "count_mixed_text": ("medium", "الحقل {f} يخلط أرقامًا وأوصافًا نصية ({n} سجل وصفه نصي)",
                             "حُفظ الوصف النصي كما هو (status text) ولا يُجمع الحقل ولا يُعرض إجماليًا مكتملًا"),
        "hijri_conversion": ("low", "تواريخ هجرية في الحقل {f} حُوِّلت إلى الميلادي ({n} سجل)",
                             "حُوِّلت بتقويم أم القرى مع حفظ النص الهجري الأصلي (raw، hijri)"),
        "date_note": ("low", "ملاحظات تحويل تاريخ في الحقل {f} ({n} سجل)",
                      "حُوِّل التاريخ مع الملاحظة المذكورة؛ يلزم التحقق إن كان مؤثرًا"),
        "date_unparsed": ("medium", "تعذر تحويل التاريخ في الحقل {f} ({n} سجل)",
                          "حُفظ النص الأصلي دون تحويل (status text)"),
        "formula_no_cache": ("medium", "خلايا صيغ بلا قيم مخزنة في الحقل {f} ({n} سجل)",
                             "عوملت القيم كغير معروفة؛ افتح الملف في Excel واحفظه لإعادة الحساب"),
        "pii_redacted": ("high", "قيم تحمل بيانات شخصية (جوال أو بريد أو آيبان أو رقم هوية) في الحقل {f} ({n} سجل)",
                         "حُجب الجزء الشخصي ولم يُخزَّن؛ علّم العمود pii في mapping إن كان شخصيًا كله"),
    }
    for (typ, f), g in groups.items():
        sev, det, treat = texts[typ]
        if typ == "zero_as_missing" and group_of(f, meta) not in ("valuation",):
            sev = "low"
        detail = det.format(f=f, n=len(g["ids"]))
        if g["examples"] and typ not in ("pii_redacted",):
            detail += " — أمثلة: " + "؛ ".join(f"«{x}»" for x in g["examples"])
        issues.add(typ, sev, join_refs(g["refs"]), g["ids"], detail + ".",
                   impact_for(f, kinds.get(f) or "", meta), treat, field=f)


def check_names(records: list, issues: Issues) -> None:
    groups = OrderedDict()
    for r in records:
        if not r.get("name"):
            continue
        k = name_key(r["name"])
        if k:
            groups.setdefault(k, []).append(r)
    for k, rs in groups.items():
        if len(rs) < 2:
            continue
        same_src = len({x["source"]["source_id"] for x in rs}) == 1
        issues.add("possible_duplicate_name", "medium" if same_src else "low",
                   join_refs([x["source"]["ref_name"] for x in rs]), [x["id"] for x in rs],
                   f"أسماء متطابقة أو متشابهة جدًا ({len(rs)} سجلات): «{rs[0]['name']}»"
                   + ("" if same_src else " — في مصادر مختلفة"),
                   "إن كانت السجلات أصلًا واحدًا فقد يتضاعف العدد والقيمة والإيراد؛ وقد تكون قطعًا مستقلة",
                   "لم تُدمج السجلات؛ يلزم مطابقة المعرّف والصك والموقع والمساحة قبل أي دمج")


def check_deeds(records: list, issues: Issues) -> None:
    groups = OrderedDict()
    for r in records:
        for f, tv in r["fields"].items():
            if not (f == "deed_no" or f.startswith("deed_no.")):
                continue
            if not _known(tv) or tv.get("v") in (None, ""):
                continue
            k = re.sub(r"[\s()\[\]{}]", "", normalize_arabic(raw_str(tv["v"])))
            if not k or not re.search(r"\d", k) or re.fullmatch(r"0+", k):
                continue  # deed "numbers" without digits are notes (e.g. «لم تفرغ باسم الجمعية»), not identifiers
            groups.setdefault(k, []).append((r, tv))
    for k, items in groups.items():
        ids = sorted({r["id"] for r, _ in items})
        if len(ids) < 2:
            continue
        issues.add("shared_deed", "medium", join_refs([tv.get("ref") for _, tv in items]), ids,
                   f"رقم الصك «{items[0][1]['v']}» مشترك بين {len(ids)} سجلات.",
                   "قد يمثل أصلًا واحدًا مسجلًا أكثر من مرة (ازدواج) أو أصلًا يحتاج فصلًا",
                   "لم تُدمج السجلات؛ يلزم التحقق من الصك وحدود كل سجل")


def check_classification(records: list, issues: Issues) -> None:
    labels = {"use": ("الاستخدام", "gross_yield_investment، gross_yield_income_producing، zero_rev_investment_value"),
              "income_nature": ("طبيعة الدخل", "property_revenue، activity_revenue، gross_yield_*، noi، revenue_growth")}
    for dim, (lab, metrics) in labels.items():
        rs = [r for r in records if r["scope"] == "in_scope" and r["record_type"] != "total"
              and (r["dims"].get(dim) in (None, "unknown"))]
        if not rs:
            continue
        issues.add("missing_classification", "medium", join_refs([r["source"]["ref_name"] for r in rs]),
                   [r["id"] for r in rs],
                   f"{len(rs)} سجل ضمن النطاق بلا تصنيف «{lab}» (unknown).",
                   f"تُستبعد هذه السجلات من المؤشرات المقيدة بهذا البعد: {metrics}",
                   "لم يُفترض تصنيف؛ صنّفها في 02_mapping.json (dim_rules أو records) بدليل من الملف أو من الجهة")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def load_workbook_map(out_path: str, xlsx: str):
    """RUN/01_workbook_map.json next to --out, when it belongs to the same workbook (or its converted copy)."""
    p = os.path.join(os.path.dirname(os.path.abspath(out_path)), "01_workbook_map.json")
    if not os.path.isfile(p):
        return None
    try:
        wm = load_json(p)
    except Exception as exc:
        eprint(f"تحذير: تعذر قراءة {p} ({exc})؛ لن تُستخدم أنواع الصفوف منه.")
        return None
    f = wm.get("file") or {}
    sha = sha256_file(xlsx)
    if sha not in (f.get("sha256"), (f.get("converted") or {}).get("sha256")):
        eprint("تحذير: 01_workbook_map.json في مجلد التشغيل لا يخص الملف المعطى (بصمة مختلفة)؛ لن تُستخدم أنواع الصفوف منه.")
        return None
    return wm


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="بناء 03_records.json من المصنف و02_mapping.json")
    ap.add_argument("--xlsx", required=True, help="مسار ملف Excel المصدر (لا يُعدَّل)")
    ap.add_argument("--mapping", required=True, help="مسار 02_mapping.json")
    ap.add_argument("--out", required=True, help="مسار 03_records.json (يُقرأ 01_workbook_map.json من المجلد نفسه إن وُجد)")
    ap.add_argument("--schema", default=os.path.join(SKILL_DIR, "schemas", "mapping.schema.json"))
    args = ap.parse_args(argv)

    if not os.path.isfile(args.xlsx):
        eprint(f"خطأ: الملف غير موجود: {args.xlsx}")
        return 2
    try:
        mapping = load_json(args.mapping)
    except Exception as exc:
        eprint(f"خطأ: تعذر قراءة mapping ({args.mapping}): {exc}")
        return 2
    errs = schema_errors(mapping, args.schema)
    if errs:
        eprint(f"ملف mapping غير مطابق للمخطط ({len(errs)} خطأ):")
        for e in errs:
            eprint(e)
        return 2
    errs = semantic_errors(mapping)
    if errs:
        eprint(f"ملف mapping غير متسق ({len(errs)} خطأ):")
        for e in errs:
            eprint(e)
        return 2
    meta = mapping.get("field_meta") or {}
    used = set()
    for src in mapping["sources"]:
        for col, c in src["columns"].items():
            if c.get("kind") not in ("pii", "ignore") and c.get("field"):
                used.add(c["field"])
    missing_meta = sorted(f for f in used if f not in meta and (f.startswith(("rev.", "val.", "exp."))))
    if missing_meta:
        eprint("تحذير: حقول مالية بلا وصف في field_meta (يحتاجها --init-config): " + "، ".join(missing_meta))
    src_file = mapping.get("source_file")
    if src_file and os.path.basename(src_file) != os.path.basename(args.xlsx):
        eprint(f"تحذير: source_file في mapping ({os.path.basename(src_file)}) لا يطابق الملف المعطى "
               f"({os.path.basename(args.xlsx)}).")
    wb_map = load_workbook_map(args.out, args.xlsx)
    try:
        out = normalize(args.xlsx, mapping, wb_map)
    except SystemExit as exc:
        eprint(str(exc))
        return 2
    save_json(args.out, out)
    st = out["stats"]
    sev = st["issues_by_severity"]
    print(f"تم: {st['records']} سجل → {args.out}")
    print(f"  حسب النوع: {st['by_type']} | حسب النطاق: {st['by_scope']}")
    print(f"  مشكلات البيانات: {st['issues']} (مرتفع {sev.get('high', 0)}، متوسط {sev.get('medium', 0)}، "
          f"منخفض {sev.get('low', 0)}) | أعمدة شخصية محذوفة: {st['pii_columns_dropped']}")
    print(f"  صفوف فارغة متجاوزة: {st['rows_skipped_blank']} | مستبعدة: {st['rows_excluded']} | إجماليات مستبعدة آليًا: "
          f"{st['rows_total_excluded']} | متخطاة رغم وجود بيانات: {st['rows_skipped_with_content']} | مدمجة بالربط: "
          f"{st['rows_joined']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
