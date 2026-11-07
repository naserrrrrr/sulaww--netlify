#!/usr/bin/env python3
"""normalize.py — build 03_records.json from the workbook and Claude's 02_mapping.json.

Usage:
    python3 scripts/normalize.py --xlsx <file.xlsx> --mapping RUN/02_mapping.json --out RUN/03_records.json

Implements data-contract.md §3–§4: typed values (TV) with source references, explicit
zero/blank/unknown statuses, derived fields, dimension rules, per-record overrides and
automatic data-quality issues (DQ-xxx). The workbook is opened read-only and never
modified. Personal-data columns are dropped (only counted).
Exit codes: 0 ok, 2 invalid mapping / unreadable input.
"""
from __future__ import annotations

import argparse
import copy
import os
import re
import sys
import warnings
from collections import OrderedDict

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from apa_common import (  # noqa: E402
    ALL_STATUSES, CONTRACT_VERSION, EMAIL_RE, KNOWN_STATUSES, PII_HEADER_RE, S_BLANK, S_DERIVED,
    S_ERROR, S_FILLER, S_NOT_AVAILABLE, S_TEXT, S_VALUE, S_ZERO_AS_MISSING, S_ZERO_CONFIRMED,
    S_ZERO_RECORDED, SKILL_DIR, clean_text, is_filler, load_json, looks_like_phone, name_key,
    normalize_arabic, parse_date, parse_number, ref, save_json, sha256_file,
)

NUMERIC_KINDS = ("money", "number", "area", "count", "percent")
TEXT_KINDS = ("text", "url")
LIST_DIMS = ("ownership",)
DEFAULT_DIMS = OrderedDict([("use", "unknown"), ("income_nature", "unknown"),
                            ("ownership", ["unknown"]), ("waqf", "unknown")])
TOTAL_RE = re.compile(r"(?<!\w)(المجموع|مجموع|الاجمالي|اجمالي|الاجماليات|total|subtotal|grand total)(?!\w)",
                      re.IGNORECASE)
VAL_RE = re.compile(r"^val\.([^.]+)\.(total|land|building)$")
MAX_REFS_IN_WHERE = 30


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
    return f"لا يطابق المخطط ({v}): {err.message}"


def schema_errors(instance, schema_path: str) -> list:
    try:
        from jsonschema import Draft202012Validator
    except Exception:  # pragma: no cover - jsonschema is expected to be installed
        eprint("تحذير: مكتبة jsonschema غير مثبتة؛ تُجرى فحوص أساسية فقط.")
        errs = []
        for k in ("contract_version", "sources", "field_meta"):
            if k not in instance:
                errs.append(f"• المسار (الجذر): الحقل المطلوب «{k}» مفقود")
        return errs
    schema = load_json(schema_path)
    validator = Draft202012Validator(schema)
    out = []
    for e in sorted(validator.iter_errors(instance), key=lambda e: [str(p) for p in e.absolute_path]):
        path = "/".join(str(p) for p in e.absolute_path) or "(الجذر)"
        out.append(f"• المسار {path}: {_arabic_error(e)}")
    return out


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


def fmt(x) -> str:
    if x is None:
        return "—"
    if isinstance(x, float) and not x.is_integer():
        return f"{x:,.2f}"
    return f"{int(x):,}"


# ---------------------------------------------------------------------------
# Typed values
# ---------------------------------------------------------------------------
def make_tv(val, kind: str, cspec: dict, field: str, cell_ref: str, number_format: str, formula, meta: dict) -> dict:
    if kind in NUMERIC_KINDS or kind == "count_mixed":
        v, s, note = parse_number(val)
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
        if s == S_TEXT:
            note = ("وصف نصي مختلط؛ لا يُجمع" if kind == "count_mixed" else "نص في حقل رقمي؛ لم يُحوَّل")
        tv = {"v": tidy_num(v) if s in KNOWN_STATUSES else None, "s": s, "raw": raw_str(val),
              "ref": cell_ref, "note": note}
    elif kind == "date":
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
        else:
            tv = {"v": clean_text(raw_str(val)), "s": S_VALUE, "raw": raw_str(val), "ref": cell_ref, "note": None}
    if isinstance(formula, str) and formula.startswith("="):
        tv["formula"] = formula
        if val is None:
            tv["note"] = "خلية صيغة بلا قيمة مخزنة (افتح الملف في Excel واحفظه لإعادة الحساب)"
            tv["no_cache"] = True
    return tv


def redact_if_contact(tv: dict, kind: str) -> bool:
    """Contact-like values in non-PII text columns are never stored."""
    if kind not in TEXT_KINDS + ("count_mixed",):
        return False
    raw = tv.get("raw")
    if raw is None:
        return False
    if looks_like_phone(raw) or EMAIL_RE.search(str(raw)) or \
            any(looks_like_phone(tok) for tok in re.split(r"[\s,،/;|]+", clean_text(raw)) if tok):
        tv.update(v=None, s=S_NOT_AVAILABLE, raw=None, note="حُجبت قيمة تشبه بيانات اتصال شخصية")
        return True
    return False


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

    def finalize(self):
        for i, it in enumerate(self.items, 1):
            it["id"] = f"DQ-{i:03d}"
        return self.items


# ---------------------------------------------------------------------------
# Main normalisation
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


def normalize(xlsx: str, mapping: dict) -> dict:
    issues = Issues()
    meta = mapping.get("field_meta") or {}
    wb_v, wb_f = open_workbook(xlsx)
    records = []
    excluded_rows = []
    pii_columns = 0
    pii_values = 0
    rows_skipped = 0
    rows_excluded = 0
    seen_ids = {}
    field_kinds = {}
    overrides = mapping.get("records") or {}

    # per-record raw info kept for checks
    total_like = []  # (rec, reason)

    for src in mapping["sources"]:
        sid = src["id"]
        sheet = resolve_sheet(wb_v, src["sheet"], issues)
        ws, wsf = wb_v[sheet], wb_f[sheet]
        header_row = src.get("header_row")
        cols = sorted(src["columns"].items(), key=lambda kv: col_index(kv[0]))
        # PII by header (safety net)
        pii_auto = set()
        for col, cspec in cols:
            if cspec["kind"] == "pii":
                pii_columns += 1
                continue
            if header_row and not cspec.get("pii_ok"):
                header = ws[f"{col}{header_row}"].value
                if header is not None and PII_HEADER_RE.search(clean_text(header)) and cspec["kind"] != "ignore":
                    pii_auto.add(col)
                    pii_columns += 1
                    issues.add("pii_header_not_marked", "high", ref(sheet, f"{col}{header_row}"), [],
                               f"عنوان العمود {col} يدل على بيانات شخصية لكنه لم يُعلَّم pii في mapping.",
                               "لا أثر على المؤشرات",
                               "عومل العمود كبيانات شخصية ولم تُخزَّن قيمه؛ إن لم يكن شخصيًا فأضف \"pii_ok\": true")
            if cspec["kind"] != "ignore":
                field_kinds.setdefault(cspec.get("field"), cspec["kind"])
        dr = src["data_rows"]
        start = dr["from"]
        end = dr.get("to") or ws.max_row
        excl = {}
        for ex in src.get("exclude_rows") or []:
            for r in parse_rows(ex.get("rows")):
                excl[r] = ex.get("reason")
        for r, reason in sorted(excl.items()):
            if start <= r <= end:
                excluded_rows.append({"source_id": sid, "sheet": sheet, "row": r, "reason": reason})
        id_rule = src.get("id_rule") or {}
        prefix = id_rule.get("prefix")
        if prefix is None:
            prefix = sid + "-"
        id_col = id_rule.get("column")
        name_col = src.get("name_column")
        skip_col = src.get("skip_if_blank")
        for row in range(start, end + 1):
            if row in excl:
                rows_excluded += 1
                continue
            if skip_col:
                sv = ws[f"{skip_col}{row}"].value
                if sv is None or is_filler(sv):
                    rows_skipped += 1
                    continue
            # id
            base_id = None
            id_text = None
            if id_col:
                iv = ws[f"{id_col}{row}"].value
                if iv is not None and not is_filler(iv):
                    id_text = clean_text(raw_str(iv))
                    base_id = prefix + id_text
            if not base_id:
                base_id = f"{prefix}r{row}"
            rid = base_id
            if rid in seen_ids:
                rid = f"{base_id}-r{row}"
                n = 2
                while rid in seen_ids:
                    rid = f"{base_id}-r{row}-{n}"
                    n += 1
                first = seen_ids[base_id]
                issues.add("duplicate_id", "high",
                           join_refs([first["source"].get("ref_id"), ref(sheet, f"{id_col}{row}") if id_col else None]),
                           [first["id"], rid],
                           f"المعرّف «{base_id}» مكرر؛ أُعطي السجل في الصف {row} المعرّف «{rid}».",
                           "قد يؤدي إلى ازدواج الأصل في العدد والقيمة والإيراد إن كانا السجل نفسه",
                           "لم يُدمج السجلان؛ يلزم التحقق من المعرف والصك والموقع وتصحيح الترقيم أو استبعاد المكرر في mapping")
            name = None
            if name_col:
                nv = ws[f"{name_col}{row}"].value
                name = clean_text(raw_str(nv)) or None
            rec = OrderedDict()
            rec["id"] = rid
            rec["name"] = name
            rec["source"] = OrderedDict([("source_id", sid), ("sheet", sheet), ("row", row),
                                         ("ref_name", ref(sheet, f"{name_col}{row}") if name_col else None)])
            if id_col:
                rec["source"]["ref_id"] = ref(sheet, f"{id_col}{row}")
            rec["record_type"] = src.get("default_record_type") or "asset"
            rec["scope"] = src.get("default_scope") or "in_scope"
            rec["scope_reason"] = src.get("scope_reason")
            dims = OrderedDict((k, copy.deepcopy(v)) for k, v in DEFAULT_DIMS.items())
            dims_ev = OrderedDict()
            fields = OrderedDict()
            row_total_formula = False
            for col, cspec in cols:
                kind = cspec["kind"]
                cell_ref = ref(sheet, f"{col}{row}")
                if kind == "ignore":
                    continue
                cell = ws[f"{col}{row}"]
                val = cell.value
                if kind == "pii" or col in pii_auto:
                    if val is not None and not is_filler(val):
                        pii_values += 1
                    continue
                formula = wsf[f"{col}{row}"].value
                if kind in NUMERIC_KINDS and is_total_formula(formula):
                    row_total_formula = True
                field = cspec["field"]
                tv = make_tv(val, kind, cspec, field, cell_ref, cell.number_format, formula, meta)
                if redact_if_contact(tv, kind):
                    pii_values += 1
                    tv["redacted"] = True
                fields[field] = tv
                dim = cspec.get("dim")
                if dim:
                    dv = tv.get("v") if tv.get("s") in KNOWN_STATUSES else None
                    if dim in LIST_DIMS:
                        if dv is not None:
                            dims[dim] = [dv]
                    else:
                        dims[dim] = dv
                    dims_ev[dim] = cell_ref
            rec["dims"] = dims
            rec["dims_evidence"] = dims_ev
            rec["fields"] = fields
            rec["issue_ids"] = []
            seen_ids[rid] = rec
            records.append(rec)
            reasons = []
            for txt in (name, id_text):
                if txt and TOTAL_RE.search(normalize_arabic(txt)):
                    reasons.append(f"النص «{txt}» يدل على إجمالي")
            if row_total_formula:
                reasons.append("الصف يحتوي صيغة تجميع (SUBTOTAL/SUM لنطاق)")
            if reasons:
                total_like.append((rec, "؛ ".join(reasons)))

    by_id = {r["id"]: r for r in records}
    apply_derived(records, mapping.get("derived_fields") or [], issues, meta)
    apply_dim_rules(records, mapping.get("dim_rules") or [])
    apply_overrides(records, by_id, overrides, issues, meta)
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
    for rec, why in total_like:
        issues.add("total_like_row", "high", rec["source"]["ref_name"] or ref(rec["source"]["sheet"], f"A{rec['source']['row']}"),
                   [rec["id"]], f"صف داخل نطاق البيانات يبدو صف إجمالي: {why}.",
                   "إن كان إجماليًا فسيُضاعف العدد والقيمة والإيراد",
                   "لم يُستبعد تلقائيًا؛ إن ثبت أنه إجمالي فأضفه إلى exclude_rows أو اجعل record_type = total في mapping")
    check_valuation_totals(records, issues, meta)
    check_vat(records, issues, meta)
    check_field_statuses(records, issues, meta, field_kinds)
    check_names(records, issues)
    check_deeds(records, issues)
    check_classification(records, issues)

    items = issues.finalize()
    for it in items:
        for rid in it["record_ids"]:
            if rid in by_id:
                by_id[rid]["issue_ids"].append(it["id"])

    stats = OrderedDict()
    stats["records"] = len(records)
    stats["by_type"] = dict(sorted(_count(r["record_type"] for r in records).items()))
    stats["by_scope"] = dict(sorted(_count(r["scope"] for r in records).items()))
    stats["by_source"] = dict(sorted(_count(r["source"]["source_id"] for r in records).items()))
    stats["pii_columns_dropped"] = pii_columns
    stats["pii_values_dropped"] = pii_values
    stats["rows_skipped_blank"] = rows_skipped
    stats["rows_excluded"] = rows_excluded
    stats["issues"] = len(items)
    stats["issues_by_severity"] = dict(sorted(_count(i["severity"] for i in items).items()))

    out = OrderedDict()
    out["contract_version"] = CONTRACT_VERSION
    out["source"] = OrderedDict([("file", xlsx), ("name", os.path.basename(xlsx)), ("sha256", sha256_file(xlsx)),
                                 ("mapping_source_file", mapping.get("source_file"))])
    out["currency"] = mapping.get("currency")
    out["field_meta"] = copy.deepcopy(meta)
    out["records"] = records
    out["excluded_rows"] = excluded_rows
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


# ---------------------------------------------------------------------------
# Checks
# ---------------------------------------------------------------------------
def _known(tv):
    return tv is not None and tv.get("s") in KNOWN_STATUSES


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
            if kind in NUMERIC_KINDS:
                if s == S_TEXT:
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
                elif tv.get("note"):
                    add(("date_note", f), r["id"], tv, f"{tv.get('raw')}: {tv.get('note')}")

    texts = {
        "text_in_numeric": ("medium", "نص في حقل رقمي ({f}) في {n} سجل",
                            "لم يُحوَّل النص إلى رقم ولم يُعامل كصفر؛ القيمة غير معروفة مع حفظ النص الأصلي"),
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
        "pii_redacted": ("high", "قيم تشبه بيانات اتصال شخصية في الحقل {f} ({n} سجل)",
                         "حُجبت القيم ولم تُخزَّن؛ علّم العمود pii في mapping"),
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
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="بناء 03_records.json من المصنف و02_mapping.json")
    ap.add_argument("--xlsx", required=True, help="مسار ملف Excel المصدر (لا يُعدَّل)")
    ap.add_argument("--mapping", required=True, help="مسار 02_mapping.json")
    ap.add_argument("--out", required=True, help="مسار 03_records.json")
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
    try:
        out = normalize(args.xlsx, mapping)
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
    return 0


if __name__ == "__main__":
    sys.exit(main())
