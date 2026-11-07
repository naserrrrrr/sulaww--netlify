#!/usr/bin/env python3
"""compute_metrics.py — metric configuration and authoritative metric results.

Usage:
    # 1) draft the metric configuration from 03_records.json (Claude reviews/edits it afterwards)
    python3 scripts/compute_metrics.py --init-config --records RUN/03_records.json --out RUN/04_metrics_config.json
    # 2) compute 05_metrics.json (default filters + parity cases for the dashboard)
    python3 scripts/compute_metrics.py --records RUN/03_records.json --config RUN/04_metrics_config.json --out RUN/05_metrics.json
    # 3) ad-hoc check of a filtered view (prints an Arabic table; --out optional)
    python3 scripts/compute_metrics.py --records RUN/03_records.json --config RUN/04_metrics_config.json \
        --filters '{"use":"investment"}' --print

The arithmetic lives in metric_engine.py (Python twin of assets/metric_engine.js).
When --filters is combined with --out, 05_metrics.json is still computed with the default
filters (it is the authoritative file); --filters only changes the printed table.
Exit codes: 0 ok, 2 invalid input/config.
"""
from __future__ import annotations

import argparse
import copy
import json
import os
import re
import sys
from collections import OrderedDict

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from apa_common import CONTRACT_VERSION, KNOWN_STATUSES, SKILL_DIR, load_json, load_vocab, save_json  # noqa: E402
import metric_engine as me  # noqa: E402
from normalize import schema_errors as _schema_errors  # noqa: E402

ASSET_TYPES = ["asset", "land_parcel"]
# Revenue is read only from whole assets, land parcels and operating activities: unit/contract rows of an
# asset already counted would double its revenue (brief §3.5–3.6, §13.7).
REVENUE_RECORD_TYPES = ["asset", "land_parcel", "operating_activity"]
W_ASSET = {"in": ["record_type", ASSET_TYPES]}
W_REV_REC = {"in": ["record_type", REVENUE_RECORD_TYPES]}
W_NOT_MANAGED = {"not": {"eq": ["managed_for_others", True]}}
W_RENT = {"eq": ["income_nature", "rent"]}
W_NOT_OPERATING = {"not": {"eq": ["income_nature", "operating"]}}
W_INVEST = {"eq": ["use", "investment"]}
BREAKDOWN_DIMS = ["asset_type", "location", "use", "ownership", "income_nature", "rent_allocation", "restriction"]
FILTER_DIMS = ["asset_type", "location", "use"]
# ratios of revenue to the single valuation field: valid only for the valuation year and the year after (C2)
VAL_RATIO_IDS = ("gross_yield_investment", "gross_yield_owned", "gross_yield_income_producing")

VAT_AR = {"excl": "بدون ضريبة", "incl": "شامل الضريبة", "none_evidenced": "لا ضريبة مثبتة في المصدر",
          "unknown": "أساس الضريبة غير موضح"}
BASIS_AR = {"actual": "فعلي", "expected": "متوقع", "recorded": "مسجل", "contract": "تعاقدي", "unknown": "أساس الإيراد غير موضح"}
CASH_AR = {"collected": "محصّل", "accrued": "مستحق", "unknown": "الاستحقاق والتحصيل غير موضح"}
COMPLETENESS_AR = {"full_year": "سنة كاملة", "partial": "فترة جزئية", "unknown": "اكتمال الفترة غير موضح"}
# Comparability classes (B1): two known values in different classes make two periods not comparable.
VAT_CLASS = {"excl": "without_vat", "none_evidenced": "without_vat", "incl": "with_vat"}
BASIS_CLASS = {"actual": "realised", "recorded": "realised", "expected": "expected", "contract": "expected"}
CASH_CLASS = {"collected": "collected", "accrued": "accrued"}
COMPLETENESS_CLASS = {"full_year": "full_year", "partial": "partial"}
COMPARE_KEYS = (("vat", VAT_AR, "الضريبة", VAT_CLASS),
                ("basis", BASIS_AR, "أساس الإيراد", BASIS_CLASS),
                ("cash_basis", CASH_AR, "الاستحقاق والتحصيل", CASH_CLASS),
                ("completeness", COMPLETENESS_AR, "اكتمال الفترة", COMPLETENESS_CLASS))
CONFIRM_KEYS = ("vat", "basis", "completeness")
VAL_BASIS_AR = {"market": "تقييم سوقي", "book": "قيمة دفترية", "cost": "تكلفة اقتناء", "unknown": "أساس التقييم غير موضح في الملف"}
STATUS_AR_FALLBACK = {"complete": "مكتمل", "partial": "جزئي", "not_computable": "غير قابل للاحتساب",
                      "not_comparable": "غير قابل للمقارنة", "empty": "لا سجلات"}


def eprint(*a):
    print(*a, file=sys.stderr)


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------
def schema_errors(instance, schema_path: str) -> list:
    """Arabic, path-qualified JSON-schema errors (shared translator from normalize.py)."""
    return _schema_errors(instance, schema_path)


def list_dims_of(records: list) -> set:
    """Dims that hold a list in at least one record (ownership, …)."""
    out = set()
    for r in records or []:
        for k, v in (r.get("dims") or {}).items():
            if isinstance(v, list):
                out.add(k)
    return out


def _where_list_dim_errors(where, list_dims: set, path: str) -> list:
    """eq/in on a dim that is a list never matches (strict equality): reject it and suggest has (B12)."""
    errs = []
    if not isinstance(where, dict):
        return errs
    for op, arg in where.items():
        if op in ("all", "any") and isinstance(arg, list):
            for i, w in enumerate(arg):
                errs += _where_list_dim_errors(w, list_dims, f"{path}/{op}/{i}")
        elif op == "not":
            errs += _where_list_dim_errors(arg, list_dims, f"{path}/not")
        elif op in ("eq", "in") and isinstance(arg, list) and arg and arg[0] in list_dims:
            dim = arg[0]
            if op == "eq":
                hint = json.dumps({"has": [dim, arg[1] if len(arg) > 1 else None]}, ensure_ascii=False)
            else:
                vals = arg[1] if len(arg) > 1 and isinstance(arg[1], list) else []
                hint = json.dumps({"any": [{"has": [dim, v]} for v in vals]}, ensure_ascii=False)
            errs.append(f"• {path}: «{op}» على البُعد «{dim}» وهو قائمة في السجلات فلا يطابق أي سجل؛ استخدم {hint}")
    return errs


def semantic_errors(config: dict, records: list = None) -> list:
    errs = []
    ids = [p.get("id") for p in config.get("periods") or []]
    if len(ids) != len(set(ids)):
        errs.append("• periods: معرّفات فترات مكررة")
    for p in config.get("periods") or []:
        if p.get("prev_id") and p["prev_id"] not in ids:
            errs.append(f"• periods/{p['id']}: prev_id «{p['prev_id']}» غير موجود؛ عند حذف الفترة السابقة من periods "
                        "اجعل prev_id = null (يبقى prev_rev_field صالحًا لحساب النمو) أو احذف prev_rev_field")
    mids = [m.get("id") for m in config.get("metrics") or []]
    dup = sorted({x for x in mids if mids.count(x) > 1})
    if dup:
        errs.append("• metrics: معرّفات مؤشرات مكررة: " + "، ".join(dup))
    bids = [b.get("id") for b in config.get("breakdowns") or []]
    dupb = sorted({x for x in bids if bids.count(x) > 1})
    if dupb:
        errs.append("• breakdowns: معرّفات مكررة: " + "، ".join(dupb))
    for m in config.get("metrics") or []:
        mid = m.get("id")
        vp = m.get("valid_periods")
        if isinstance(vp, list):
            unknown = [x for x in vp if x not in ids]
            if unknown:
                errs.append(f"• metrics/{mid}: valid_periods تذكر فترات غير معرّفة في periods: "
                            + "، ".join(str(x) for x in unknown) + "؛ احذفها من valid_periods (أو أعد الفترة إلى periods)")
        if m.get("empty_as_zero") and m.get("kind") not in ("sum", "count"):
            errs.append(f"• metrics/{mid}: empty_as_zero يصح لمؤشرات sum وcount فقط (النوع هنا {m.get('kind')})")
    if records is not None:
        ld = list_dims_of(records)
        errs += _where_list_dim_errors(config.get("base_where"), ld, "base_where")
        for key in ("metrics", "breakdowns"):
            for x in config.get(key) or []:
                errs += _where_list_dim_errors(x.get("where"), ld, f"{key}/{x.get('id')}/where")
        errs += _where_list_dim_errors((config.get("trend") or {}).get("where"), ld, "trend/where")
    return errs


def comparability_errors(config: dict, meta: dict) -> list:
    """Periods/trend marked comparable although field_meta shows known, different bases (B1).

    The engine computes a growth % for every status but not_comparable, so a known VAT / revenue-basis /
    accrual / completeness or period mismatch must be marked not_comparable before 05 is written.
    """
    errs = []
    periods = config.get("periods") or []
    for p in periods:
        rf, pf = p.get("rev_field"), p.get("prev_rev_field")
        if not rf or not pf:
            continue
        st = (p.get("comparable_with_prev") or {}).get("status")
        if st == "not_comparable":
            continue
        d = comparability(meta.get(rf) or {}, meta.get(pf) or {}, p.get("id"), p.get("prev_id"),
                          period_of(rf, meta), period_of(pf, meta))
        if d["status"] == "not_comparable":
            errs.append(f"• periods/{p.get('id')}: comparable_with_prev «{st or 'غير محددة'}» يخالف وصف الحقول "
                        f"({rf} مقابل {pf}): {d['reason']}؛ اجعلها not_comparable أو صحّح field_meta بدليل من المصدر")
    t = config.get("trend") or {}
    if t.get("enabled") and (t.get("comparable") or {}).get("status") != "not_comparable":
        d = trend_comparability([(x.get("period"), x.get("field")) for x in t.get("fields") or []], meta)
        if d["status"] == "not_comparable":
            errs.append(f"• trend: comparable «{(t.get('comparable') or {}).get('status') or 'غير محددة'}» يخالف وصف "
                        f"الحقول: {d['reason']}؛ اجعلها not_comparable")
    return errs


def numeric_fields_used(config: dict) -> set:
    out = set()
    for p in config.get("periods") or []:
        for k in ("rev_field", "prev_rev_field", "exp_field"):
            if p.get(k):
                out.add(p[k])
    if (config.get("valuation") or {}).get("field"):
        out.add(config["valuation"]["field"])
    for t in (config.get("trend") or {}).get("fields") or []:
        out.add(t.get("field"))

    def lit(x):
        if isinstance(x, str) and not (x.startswith("{") and x.endswith("}")):
            out.add(x)

    def walk_where(w):
        if not isinstance(w, dict):
            return
        for op, arg in w.items():
            if op in ("all", "any"):
                for x in arg:
                    walk_where(x)
            elif op == "not":
                walk_where(arg)
            elif op == "known":
                lit(arg)
            elif op in ("status", "gt"):
                lit(arg[0])

    for m in list(config.get("metrics") or []) + list(config.get("breakdowns") or []):
        for k in ("field", "num", "den", "a", "b", "prev", "curr"):
            lit(m.get(k))
        for x in m.get("fields") or []:
            lit(x)
        walk_where(m.get("where"))
    walk_where(config.get("base_where"))
    out.discard(None)
    return out


# ---------------------------------------------------------------------------
# --init-config
# ---------------------------------------------------------------------------
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
    return "other"


def period_of(field: str, meta: dict):
    m = meta.get(field) or {}
    p = m.get("period") or m.get("year")
    if p:
        return str(p)
    mm = re.match(r"^(?:rev|exp|val)\.([^.]+)\.", field)
    return mm.group(1) if mm else None


def _period_sort_key(p: str):
    m = re.match(r"^(\d+)", str(p))
    return (int(m.group(1)) if m else -1, str(p))


def _known_count(records: list, field: str) -> int:
    n = 0
    for r in records:
        tv = (r.get("fields") or {}).get(field)
        if isinstance(tv, dict) and tv.get("s") in KNOWN_STATUSES:
            n += 1
    return n


def _all_fields(records: list, meta: dict) -> list:
    """Field names in field_meta order, then any extra field present in records (sorted)."""
    seen = OrderedDict()
    for f in meta:
        seen[f] = True
    extra = set()
    for r in records:
        extra.update((r.get("fields") or {}).keys())
    for f in sorted(extra):
        if f not in seen and not f.endswith("__original"):
            seen[f] = True
    return list(seen)


def _rev_desc(m: dict) -> str:
    parts = [BASIS_AR.get(m.get("basis") or "unknown", m.get("basis")), VAT_AR.get(m.get("vat") or "unknown", m.get("vat"))]
    if m.get("completeness") == "partial":
        parts.append(COMPLETENESS_AR["partial"])
    return "، ".join(p for p in parts if p)


def comparability(cm: dict, pm: dict, cur_id, prev_id, cur_field_period=None, prev_field_period=None) -> dict:
    """Comparability of two revenue fields from their field_meta (brief §6: same period length, VAT and accounting basis).

    not_comparable — a partial period on either side; a field whose own period differs from the period it stands for;
                     or a key known on both sides in different classes: VAT (incl vs excl/none_evidenced),
                     revenue basis (actual/recorded vs expected/contract), accrual (collected vs accrued),
                     completeness (full_year vs partial).
    confirmed      — vat, basis and completeness known and equal on both sides (and no difference above).
    unverified     — anything else (a key unknown on one side, or different within one class, e.g. actual vs recorded).
    """
    cur_lab = str(cur_id) if cur_id is not None else "الحالية"
    prev_lab = str(prev_id) if prev_id is not None else (str(prev_field_period) if prev_field_period else "السابقة")
    hard = []
    if cur_field_period is not None and cur_id is not None and str(cur_field_period) != str(cur_id):
        hard.append(f"حقل إيراد الفترة {cur_lab} يخص الفترة {cur_field_period}")
    if prev_field_period is not None and prev_id is not None and str(prev_field_period) != str(prev_id):
        hard.append(f"حقل إيراد الفترة السابقة {prev_lab} يخص الفترة {prev_field_period}")
    partial = [x for x, m in ((cur_lab, cm), (prev_lab, pm)) if m.get("completeness") == "partial"]
    if partial:
        hard.append(f"الفترة {'، '.join(partial)} جزئية؛ لا تُقارن بفترة أخرى ما لم يثبت تساوي طوليهما، "
                    "ولا تُحوَّل إلى أداء سنوي")
    for k, ar, lab, cls in COMPARE_KEYS:
        if k == "completeness" and partial:
            continue
        a, b = cm.get(k) or "unknown", pm.get(k) or "unknown"
        if cls.get(a) and cls.get(b) and cls[a] != cls[b]:
            hard.append(f"{lab}: {ar.get(a, a)} ({cur_lab}) مقابل {ar.get(b, b)} ({prev_lab})")
    if hard:
        return {"status": "not_comparable", "reason": "الفترتان غير قابلتين للمقارنة — " + "؛ ".join(hard)}
    if all(cm.get(k) == pm.get(k) and cm.get(k) not in (None, "unknown") for k in CONFIRM_KEYS):
        return {"status": "confirmed", "reason": "أساس الضريبة والإيراد واكتمال الفترة متطابقة وفق وصف الحقول"}
    diffs = []
    for k, ar, lab, _ in COMPARE_KEYS:
        if k not in CONFIRM_KEYS:
            continue
        a, b = cm.get(k) or "unknown", pm.get(k) or "unknown"
        if a == b:
            if a == "unknown":
                diffs.append(f"{lab}: غير موضح في الفترتين")
        else:
            diffs.append(f"{lab}: {ar.get(a, a)} ({cur_lab}) مقابل {ar.get(b, b)} ({prev_lab})")
    return {"status": "unverified", "reason": "لم يثبت تماثل الفترتين — " + "؛ ".join(diffs)}


def trend_comparability(points: list, meta: dict) -> dict:
    """Comparability of all trend periods: every pair (not only neighbours) must pass `comparability`.

    points = [(period_id, field), ...]. not_comparable if any pair is; confirmed if every pair is; else unverified.
    """
    hard, all_confirmed = [], True
    for i in range(len(points)):
        for j in range(i + 1, len(points)):
            (pa, fa), (pb, fb) = points[j], points[i]
            d = comparability(meta.get(fa) or {}, meta.get(fb) or {}, pa, pb, period_of(fa, meta) if fa else None,
                              period_of(fb, meta) if fb else None)
            if d["status"] == "not_comparable":
                hard.append(d["reason"].replace("الفترتان غير قابلتين للمقارنة — ", f"{pb}/{pa}: "))
            if d["status"] != "confirmed":
                all_confirmed = False
    if hard:
        return {"status": "not_comparable", "reason": "فترات التطور الزمني غير قابلة للمقارنة؛ القيم تُعرض للاطلاع فقط "
                                                      "ولا تُقرأ نموًا — " + " | ".join(hard)}
    if all_confirmed and len(points) >= 2:
        return {"status": "confirmed", "reason": "أساس الضريبة والإيراد واكتمال الفترات متطابقة وفق وصف الحقول"}
    return {"status": "unverified",
            "reason": "لم يثبت تماثل أساس الضريبة والإيراد واكتمال الفترات بين السنوات؛ تُعرض القيم دون حكم بالنمو"}


def _pool(records: list) -> list:
    """Records used to decide which fields are live: in-scope ones only.

    A revenue/valuation field known only on isolated or excluded records (e.g. a hidden allocation sheet)
    never drives a period or the valuation (C22); with no in-scope record there are no periods at all.
    """
    return [r for r in records if r.get("scope") == "in_scope"]


def _pick_expense(records: list, rev_field: str, cands: list):
    """The first expense field of the period that co-occurs (both known) with the revenue field on a record.

    An expense column from another sheet that never sits next to the chosen revenue field would
    only produce a misleading NOI; then exp_field stays null (noi -> not_computable).
    """
    for f in cands:
        for r in records:
            fl = r.get("fields") or {}
            a, b = fl.get(rev_field), fl.get(f)
            if isinstance(a, dict) and isinstance(b, dict) and a.get("s") in KNOWN_STATUSES and b.get("s") in KNOWN_STATUSES:
                return f
    return None


def build_periods(records: list, meta: dict, warnings: list = None) -> list:
    """One period per revenue year known on in-scope records. Appends init warnings to `warnings` (B3)."""
    warnings = warnings if warnings is not None else []
    by_period = OrderedDict()
    for f in _all_fields(records, meta):
        if group_of(f, meta) != "revenue":
            continue
        p = period_of(f, meta)
        if not p:
            continue
        by_period.setdefault(p, []).append(f)
    exp_cands = OrderedDict()
    for f in _all_fields(records, meta):
        if group_of(f, meta) == "expense":
            p = period_of(f, meta)
            if p:
                exp_cands.setdefault(p, []).append(f)
    # Only fields with known values in the analysed pool count (in-scope records; all records if none).
    pool = _pool(records)
    for pid in list(by_period):
        live = [f for f in by_period[pid] if _known_count(pool, f) > 0]
        if live:
            by_period[pid] = live
        else:
            del by_period[pid]
    pids = sorted(by_period, key=_period_sort_key, reverse=True)
    periods = []
    for pid in pids:
        cands = by_period[pid]
        chosen = None
        for f in cands:
            if (meta.get(f) or {}).get("vat") == "excl":
                chosen = f
                break
        if chosen is None:
            for f in cands:
                if (meta.get(f) or {}).get("vat") != "incl":
                    chosen = f
                    break
        if chosen is None:
            chosen = cands[0]
        m = meta.get(chosen) or {}
        if m.get("vat") == "incl":
            warnings.append(OrderedDict([
                ("code", "revenue_incl_vat"), ("period", pid), ("field", chosen),
                ("text", f"الفترة {pid}: لا يوجد حقل إيراد بدون ضريبة بقيم معروفة ضمن النطاق، فاستُخدم «{m.get('label') or chosen}» "
                         "الشامل للضريبة. الإيراد والعائد لهذه الفترة شاملان للضريبة المحصلة للغير ولا يطابقان تعريف عائد "
                         "الدخل الإجمالي (بدون الضرائب)؛ وُسمت تسمياتها «شامل الضريبة». افصل الضريبة بمستند قبل المقارنة.")]))
        periods.append(OrderedDict([
            ("id", pid),
            ("label", f"{pid} — {_rev_desc(m)}"),
            ("rev_field", chosen),
            ("exp_field", _pick_expense(pool, chosen, exp_cands.get(pid) or [])),
            ("prev_id", None),
            ("prev_rev_field", None),
            ("comparable_with_prev", None),
            ("meta", {"vat": m.get("vat", "unknown"), "basis": m.get("basis", "unknown"),
                      "cash_basis": m.get("cash_basis", "unknown"), "completeness": m.get("completeness", "unknown"),
                      "rev_label": m.get("label") or chosen, "alternatives": [f for f in cands if f != chosen]}),
        ]))
    for i, p in enumerate(periods):
        if i + 1 < len(periods):
            prev = periods[i + 1]
            p["prev_id"] = prev["id"]
            p["prev_rev_field"] = prev["rev_field"]
            p["comparable_with_prev"] = comparability(meta.get(p["rev_field"]) or {}, meta.get(prev["rev_field"]) or {},
                                                      p["id"], prev["id"], period_of(p["rev_field"], meta),
                                                      period_of(prev["rev_field"], meta))
        else:
            # "unverified" (not "not_comparable") so growth is not_computable whatever the check order
            p["comparable_with_prev"] = {"status": "unverified", "reason": "لا توجد فترة سابقة بحقل إيراد"}
    return periods


def build_valuation(records: list, meta: dict) -> dict:
    cands = []
    for f in _all_fields(records, meta):
        if group_of(f, meta) != "valuation":
            continue
        m = meta.get(f) or {}
        is_total = m.get("component") == "total" or bool(re.match(r"^val\.[^.]+\.total$", f))
        if not is_total:
            continue
        y = str(m.get("year") or period_of(f, meta) or "")
        if not y:
            continue
        cands.append((y, f))
    cands.sort(key=lambda t: _period_sort_key(t[0]), reverse=True)
    pool = _pool(records)
    for y, f in cands:
        if _known_count(pool, f) > 0:
            m = meta.get(f) or {}
            code = m.get("valuation_basis") or "unknown"
            others = [c for c in cands if c[1] != f]
            note = "أحدث تقييم مسجل في الملف ولا يعني قيمة سوقية حالية؛ لا تُجمع تقييمات سنوات مختلفة للأصل نفسه"
            if others:
                note += "؛ تقييمات سنوات أخرى متاحة للمراجعة: " + "، ".join(o[0] for o in others)
            return OrderedDict([("field", f), ("year", y), ("basis", VAL_BASIS_AR.get(code, code)), ("basis_code", code),
                                ("date_text", m.get("valuation_date") or f"{y} (اليوم غير محدد)"),
                                ("label", f"تقييم {y}"), ("note", note), ("reason", None)])
    return OrderedDict([("field", None), ("year", None), ("basis", None), ("basis_code", None), ("date_text", None),
                        ("label", "لا يوجد تقييم"), ("note", None),
                        ("reason", "لا يوجد حقل تقييم إجمالي (val.<سنة>.total) بقيم معروفة في السجلات")])


def _sources_text(records: list, file_name: str) -> str:
    """Sheets of the records a metric actually reads (C1): never sheets whose records are isolated or excluded."""
    sheets = []
    for r in records:
        s = (r.get("source") or {}).get("sheet")
        if s and s not in sheets:
            sheets.append(s)
    if not sheets:
        return f"الملف {file_name or '؟'} — لا توجد سجلات ضمن نطاق هذا المؤشر"
    word = "الورقة" if len(sheets) == 1 else "الأوراق"
    return f"الملف {file_name or '؟'} — {word}: " + "، ".join(f"«{s.strip()}»" for s in sheets)


def _records_for(records: list, base_where, where, ctx: dict) -> list:
    """Records satisfying base_where ∧ where for a period context (where with an unresolved placeholder → base only)."""
    try:
        w = me.and_where(base_where, me.resolve_where(where, ctx))
    except me.Unresolved:
        w = base_where
    return [r for r in records if me.eval_where(w, r)]


def metric_texts(mid: str, p: dict, val: dict, meta: dict, src: str) -> dict:
    """Arabic texts for a metric in period p (p may be None when there are no periods)."""
    P = p["id"] if p else "؟"
    PREV = (p or {}).get("prev_id") or "؟"
    Y = val.get("year") or "؟"
    rev_f = (p or {}).get("rev_field")
    rm = meta.get(rev_f) or {}
    rev_desc = _rev_desc(rm) if rev_f else "لا يوجد حقل إيراد"
    rev_lab = rm.get("label") or rev_f or "—"
    val_lab = (meta.get(val.get("field")) or {}).get("label") or val.get("field") or "—"
    vdate = val.get("date_text") or "لا يوجد تقييم"
    vbasis = val.get("basis") or "—"
    period_text = (p or {}).get("label") or "—"
    rev_src = f"{src}؛ حقل الإيراد: {rev_lab}"
    val_src = f"{src}؛ حقل التقييم: {val_lab}"
    both_src = f"{src}؛ الإيراد: {rev_lab}؛ التقييم: {val_lab}"
    no_val = val.get("reason") or "لا يوجد تقييم إجمالي مسجل"
    no_rev = f"لا يوجد حقل إيراد للفترة {P}" if p else "لا توجد حقول إيراد (group = revenue) في الملف"
    no_both = no_val if not val.get("field") else no_rev
    # B3: a VAT-inclusive revenue field (fallback when no excl field is known) is named in every label
    incl = rm.get("vat") == "incl"
    incl_tag = " (شامل الضريبة)" if incl else ""
    before_exp = "قبل المصروفات، شامل الضريبة" if incl else "قبل المصروفات"
    # worded without «بدون/دون ضريبة»: validate_analysis.label_vat_basis reads that phrase as a claim about this field
    incl_def = (" الإيراد هنا شامل الضريبة المحصلة للغير، إذ لا يتضمن الملف حقل إيراد معروف القيم يستثني الضريبة؛ فالنسبة "
                "أعلى من عائد الدخل الإجمالي المحسوب على إيراد مستثنى منه الضريبة ولا تُقارن به.") if incl else ""
    no_units = "؛ صفوف الوحدات والعقود التابعة لأصل لا تُجمع معه تجنبًا للازدواج"
    T = {}
    T["asset_count"] = dict(
        label="عدد الأصول ضمن النطاق", short_label="الأصول",
        definition="عدد سجلات الأصول وقطع الأراضي ضمن نطاق الدراسة، كل سجل مرة واحدة، دون العقارات المدارة للغير (تُعرض منفصلة)؛ لا يُستنتج العدد من آخر رقم مسلسل ولا تُعد الوحدات أصولًا مستقلة.",
        scope_text="السجلات من نوع «أصل» أو «قطعة أرض» المصنفة ضمن النطاق، عدا المدار للغير", basis="عدّ سجلات",
        period_text="كما في الملف", source_text=src, empty_text="لا توجد أصول أو قطع أراضٍ ضمن النطاق")
    T["managed_records_count"] = dict(
        label="عدد العقارات المدارة للغير", short_label="مدارة للغير",
        definition="عدد العقارات التي تديرها الجهة لصالح الغير ضمن النطاق؛ تُعرض منفصلة، ولا تُضاف إلى عدد أصول الجهة، ولا تدخل إيراداتها أو قيمها في مؤشرات الجهة.",
        scope_text="السجلات المصنفة «إدارة أملاك للغير» ضمن النطاق", basis="عدّ سجلات",
        period_text="كما في الملف", source_text=src, empty_text="لا توجد عقارات مدارة للغير ضمن النطاق")
    T["operating_records_count"] = dict(
        label="عدد السجلات التشغيلية", short_label="السجلات التشغيلية",
        definition="عدد سجلات الأنشطة التشغيلية (مثل المدارس) ضمن النطاق؛ تُعرض منفصلة ولا تُضاف إلى عدد الأصول.",
        scope_text="السجلات من نوع «نشاط تشغيلي» ضمن النطاق", basis="عدّ سجلات", period_text="كما في الملف",
        source_text=src, empty_text="لا توجد أنشطة تشغيلية ضمن النطاق")
    T["recorded_value"] = dict(
        label="القيمة المسجلة للأصول المقيمة", short_label="القيمة المسجلة",
        definition=f"مجموع إجمالي {val.get('label') or 'التقييم'} المسجل في الملف للأصول المقيمة ضمن النطاق، مرة واحدة لكل أصل ودون جمع الأرض والمبنى مع إجماليهما. أحدث تقييم مسجل وليس قيمة سوقية حالية، ولا يمثل القيمة الكاملة للمحفظة إذا وُجدت أصول غير مقيمة (تظهر في المستبعدات).",
        scope_text="الأصول وقطع الأراضي ضمن النطاق، عدا العقارات المدارة للغير", basis=vbasis,
        valuation_date=vdate, period_text=f"تاريخ التقييم: {vdate}", source_text=val_src, unavailable_reason=no_val)
    T["property_revenue"] = dict(
        label=f"الإيرادات العقارية المسجلة {P}{incl_tag}", short_label="الإيرادات العقارية",
        definition=f"مجموع «{rev_lab}» لسجلات الأصول وقطع الأراضي ذات طبيعة الدخل «إيجارات عقارية» ضمن النطاق، دون العقارات المدارة للغير ودون إيرادات الأنشطة التشغيلية{no_units}. الخانات الفارغة لا تُعامل كصفر.",
        scope_text="سجلات الإيجارات العقارية (أصول وقطع أراضٍ) ضمن النطاق، عدا المدار للغير والوحدات والعقود", basis=f"{rev_desc}؛ قبل المصروفات",
        period_text=period_text, source_text=rev_src, unavailable_reason=no_rev)
    T["activity_revenue"] = dict(
        label=f"إيرادات الأنشطة التشغيلية {P}{incl_tag}", short_label="إيرادات الأنشطة",
        definition=f"مجموع «{rev_lab}» للسجلات ذات طبيعة الدخل «نشاط تشغيلي» (مثل المدارس){no_units}. إيراد وليس ربحًا، ولا يُعامل كإيجار عقاري ولا يدخل في عائد العقارات.",
        scope_text="سجلات الأنشطة التشغيلية ضمن النطاق، دون الوحدات والعقود", basis=f"{rev_desc}؛ إيراد إجمالي قبل تكاليف النشاط",
        period_text=period_text, source_text=rev_src, unavailable_reason=no_rev,
        empty_text="لا توجد أنشطة تشغيلية ضمن النطاق")
    yield_def = (f"مجموع إيراد {P} ÷ مجموع {val.get('label') or 'التقييم'} للأصول نفسها التي يتوافر لها الرقمان × 100، "
                 "من المجاميع المتطابقة لا من متوسط عوائد الأصول، ويشمل الأصول ذات الإيراد الصفري المسجل. "
                 "إذا كانت الحالة «جزئي» فهو «عائد الجزء القابل للاحتساب». ليس عائدًا صافيًا ولا عائدًا استثماريًا كليًا."
                 + incl_def)
    T["gross_yield_investment"] = dict(
        label=f"الإيراد المسجل لعام {P} إلى تقييم {Y} للأصول الاستثمارية نفسها، {before_exp}",
        short_label="عائد الدخل الإجمالي — الأصول الاستثمارية",
        definition=yield_def,
        scope_text="الأصول الاستثمارية ضمن النطاق، عدا الأنشطة التشغيلية والعقارات المدارة للغير",
        basis=f"قبل المصروفات؛ {rev_desc}؛ {vbasis}", valuation_date=vdate, period_text=period_text,
        source_text=both_src, unavailable_reason=no_both)
    T["gross_yield_owned"] = dict(
        label=f"الإيراد المسجل لعام {P} إلى تقييم {Y} لجميع الأصول المملوكة، {before_exp}",
        short_label="عائد الدخل الإجمالي — جميع الأصول المملوكة",
        definition=yield_def + " يشمل الأصول الخدمية والإدارية؛ إدراجها في نسبة عامة لا يكفي للحكم على أدائها لأن غرضها خدمة المستفيدين.",
        scope_text="جميع الأصول المملوكة ضمن النطاق، عدا الأنشطة التشغيلية", basis=f"قبل المصروفات؛ {rev_desc}؛ {vbasis}",
        valuation_date=vdate, period_text=period_text, source_text=both_src, unavailable_reason=no_both)
    T["gross_yield_income_producing"] = dict(
        label=f"الإيراد المسجل لعام {P} إلى تقييم {Y} للأصول الاستثمارية المدرة فقط، {before_exp}",
        short_label="عائد الدخل الإجمالي — الأصول المدرة فقط",
        definition=yield_def + f" يقتصر على الأصول الاستثمارية ذات إيراد موجب في {P}؛ للمقارنة فقط ولا يحل محل النطاق الاستثماري الكامل.",
        scope_text="الأصول الاستثمارية ذات الإيراد الموجب ضمن النطاق", basis=f"قبل المصروفات؛ {rev_desc}؛ {vbasis}",
        valuation_date=vdate, period_text=period_text, source_text=both_src, unavailable_reason=no_both)
    T["noi"] = dict(
        label=f"صافي الدخل التشغيلي {P}{incl_tag}", short_label="صافي الدخل التشغيلي",
        definition="الإيراد العقاري ناقص مصروفات التشغيل المرتبطة به لكل أصل يتوافر له الرقمان، دون الإهلاك والتمويل والنفقات الرأسمالية. لا يساوي المبلغ المتاح لخدمة الأيتام.",
        scope_text="سجلات الإيجارات العقارية (أصول وقطع أراضٍ) ضمن النطاق، عدا المدار للغير والوحدات والعقود", basis=f"{rev_desc}",
        period_text=period_text, source_text=rev_src,
        unavailable_reason="لا توجد مصروفات تشغيل مربوطة بسجلات النطاق لهذه الفترة (إن كانت في ورقة مستقلة فاربطها "
                           "بالسجلات عبر join برابط مؤكد في 02)؛ غياب المصروفات لا يعني أنها صفر")
    T["zero_rev_investment_value"] = dict(
        label="قيمة الأصول الاستثمارية المسجلة بإيراد صفري", short_label="قيمة الاستثماري بإيراد صفري",
        definition=f"مجموع {val.get('label') or 'التقييم'} للأصول الاستثمارية التي إيرادها في {P} صفر مدوّن في الملف (مؤكد بدليل أو غير متحقق من سببه). الصفر لا يثبت الشغور أو تعثر التحصيل، والخانات الفارغة لا تُحسب هنا.",
        scope_text="الأصول الاستثمارية ضمن النطاق ذات إيراد صفري مدوّن، عدا المدار للغير", basis=vbasis,
        valuation_date=vdate, period_text=period_text, source_text=both_src, unavailable_reason=no_both,
        empty_text="لا توجد أصول استثمارية بإيراد صفري مسجل")
    T["data_coverage"] = dict(
        label="تغطية بيانات الإيراد والتقييم", short_label="تغطية البيانات",
        definition=f"أدنى نسبة من الأصول ضمن النطاق يتوافر لها إيراد {P} معروف وتقييم {Y} معروف (تُحسب لكل حقل ويؤخذ الأدنى). الخانة «لا ينطبق» ليست نقصًا في البيانات فتُستبعد من مقام ذلك الحقل وتُعرض منفصلة. تغطية بعدد السجلات لا بالقيمة.",
        scope_text="الأصول وقطع الأراضي ضمن النطاق، عدا المدار للغير", basis="عدّ سجلات ذات قيم معروفة", valuation_date=vdate,
        period_text=period_text, source_text=both_src, unavailable_reason=no_both)
    comp = (p or {}).get("comparable_with_prev") or {}
    T["revenue_growth"] = dict(
        label=f"تغير الإيراد العقاري المسجل {PREV} ← {P}{incl_tag}", short_label="تغير الإيراد العقاري",
        definition=("التغير في مجموع الإيراد للأصول نفسها المعروف إيرادها في الفترتين؛ الأصول المعروفة في فترة واحدة تُعرض منفصلة "
                    "(إضافة/تخارج أو نقص بيانات). الأصل الذي كان إيراده في الفترة السابقة صفرًا مسجلًا وأصبح موجبًا يُعرض «بدء إيراد» "
                    "بتغيره بالريال خارج نسبة النمو. "
                    f"قابلية المقارنة: {comp.get('reason') or 'غير متحققة'}."),
        scope_text="سجلات الإيجارات العقارية (أصول وقطع أراضٍ) ضمن النطاق، عدا المدار للغير والوحدات والعقود", basis=f"{rev_desc}",
        period_text=f"{PREV} مقابل {P}", source_text=rev_src,
        unavailable_reason="لا توجد فترة سابقة بحقل إيراد قابل للمقارنة")
    T["value_concentration_top3"] = dict(
        label="تركّز القيمة المسجلة في أكبر 3 أصول", short_label="تركّز القيمة (أكبر 3)",
        definition="حصة أكبر 3 أصول من مجموع القيمة المسجلة الموجبة المعروفة × 100، مع مؤشر هيرفندال (HHI).",
        scope_text="الأصول وقطع الأراضي ضمن النطاق ذات تقييم معروف موجب، عدا المدار للغير", basis=vbasis,
        valuation_date=vdate, period_text=f"تاريخ التقييم: {vdate}", source_text=val_src, unavailable_reason=no_val)
    T["revenue_concentration_top3"] = dict(
        label=f"تركّز الإيراد المسجل {P} في أكبر 3 سجلات", short_label="تركّز الإيراد (أكبر 3)",
        definition="حصة أكبر 3 سجلات من مجموع الإيراد الموجب المعروف × 100 (الإيجارات والأنشطة معًا) لقياس الاعتماد على مصدر دخل واحد.",
        scope_text="سجلات الأصول وقطع الأراضي والأنشطة التشغيلية ضمن النطاق ذات إيراد موجب معروف، عدا المدار للغير والوحدات والعقود", basis=f"{rev_desc}",
        period_text=period_text, source_text=rev_src, unavailable_reason=no_rev)
    return T.get(mid, {})


def standard_metrics() -> list:
    """Structure of the standard metrics (texts are added by init_config).

    Every revenue-based metric also requires record_type ∈ REVENUE_RECORD_TYPES (A2): unit/contract rows of an
    asset are never added to the asset's revenue.
    """
    W_ASSET_OWN = {"all": [W_ASSET, W_NOT_MANAGED]}
    return [
        dict(id="asset_count", kind="count", where=W_ASSET_OWN, unit="asset", card=True, card_group="counts",
             card_order=1, kind_tag="fact"),
        dict(id="operating_records_count", kind="count", where={"eq": ["record_type", "operating_activity"]},
             unit="count", card=True, card_group="counts", card_order=2, kind_tag="fact"),
        # shown only when the scope holds managed-for-others records (see init_config)
        dict(id="managed_records_count", kind="count", where={"eq": ["managed_for_others", True]},
             unit="count", card=True, card_group="counts", card_order=2, kind_tag="fact"),
        dict(id="recorded_value", kind="sum", field="{VAL}", where=W_ASSET_OWN, unit="SAR", card=True,
             card_group="recorded_value", card_order=3, kind_tag="derived"),
        dict(id="property_revenue", kind="sum", field="{REV}", where={"all": [W_RENT, W_REV_REC, W_NOT_MANAGED]},
             unit="SAR", card=True, card_group="property_revenue", card_order=4, kind_tag="derived"),
        dict(id="activity_revenue", kind="sum", field="{REV}",
             where={"all": [{"eq": ["income_nature", "operating"]}, W_REV_REC]},
             unit="SAR", card=True, card_group="activity_revenue", card_order=5, kind_tag="derived"),
        dict(id="gross_yield_investment", kind="ratio", num="{REV}", den="{VAL}", scale=100,
             where={"all": [W_INVEST, W_ASSET, W_NOT_OPERATING, W_NOT_MANAGED]}, unit="%", card=True,
             card_group="gross_yield_investment", card_order=6, kind_tag="derived"),
        dict(id="gross_yield_owned", kind="ratio", num="{REV}", den="{VAL}", scale=100,
             where={"all": [W_ASSET, {"has": ["ownership", "owned"]}, W_NOT_OPERATING, W_NOT_MANAGED]}, unit="%",
             card=False, card_group=None, card_order=None, kind_tag="derived"),
        dict(id="gross_yield_income_producing", kind="ratio", num="{REV}", den="{VAL}", scale=100,
             where={"all": [W_INVEST, W_ASSET, W_NOT_OPERATING, W_NOT_MANAGED, {"gt": ["{REV}", 0]}]}, unit="%",
             card=False, card_group=None, card_order=None, kind_tag="derived"),
        dict(id="noi", kind="diff", a="{REV}", b="{EXP}", where={"all": [W_RENT, W_REV_REC, W_NOT_MANAGED]},
             unit="SAR", card=True, card_group="noi", card_order=7, kind_tag="derived"),
        # no qualifying asset in the whole scope is a real zero (nothing of value earns zero), not "no data" (C3)
        dict(id="zero_rev_investment_value", kind="sum", field="{VAL}",
             where={"all": [W_INVEST, W_ASSET, W_NOT_MANAGED,
                            {"status": ["{REV}", ["zero_recorded", "zero_confirmed"]]}]},
             unit="SAR", card=True, card_group="zero_rev_investment_value", card_order=8, kind_tag="derived",
             empty_as_zero=True),
        dict(id="data_coverage", kind="coverage", fields=["{REV}", "{VAL}"], where=W_ASSET_OWN, unit="%", card=True,
             card_group="data_coverage", card_order=9, kind_tag="derived"),
        dict(id="revenue_growth", kind="growth", prev="{REV_PREV}", curr="{REV}",
             where={"all": [W_RENT, W_REV_REC, W_NOT_MANAGED]}, unit="%", card=False, card_group=None,
             card_order=None, kind_tag="derived"),
        dict(id="value_concentration_top3", kind="concentration", field="{VAL}", top=3, where=W_ASSET_OWN, unit="%",
             card=False, card_group=None, card_order=None, kind_tag="derived"),
        dict(id="revenue_concentration_top3", kind="concentration", field="{REV}", top=3,
             where={"all": [W_REV_REC, W_NOT_MANAGED]},
             unit="%", card=False, card_group=None, card_order=None, kind_tag="derived"),
    ]


PERIOD_DEPENDENT = {"property_revenue", "activity_revenue", "gross_yield_investment", "gross_yield_owned",
                    "gross_yield_income_producing", "noi", "zero_rev_investment_value", "data_coverage",
                    "revenue_growth", "revenue_concentration_top3"}
BASE_WHERE = {"all": [{"eq": ["scope", "in_scope"]}]}
SPEC_TEXT_KEYS = ("definition", "scope_text", "basis", "valuation_date", "period_text", "source_text",
                  "unavailable_reason", "empty_text")


def _dim_values(records: list, dim: str) -> list:
    vals = set()
    for r in records:
        v = (r.get("dims") or {}).get(dim)
        items = v if isinstance(v, list) else [v]
        for x in items:
            if x is None or isinstance(x, (list, dict)):
                continue
            vals.add(x)
    return sorted(vals, key=lambda x: str(x))


def _distinct_known(records: list, dim: str) -> int:
    """Number of distinct classified values of a dim (unknown/missing not counted; list dims by element)."""
    return len([v for v in _dim_values(records, dim) if v != "unknown"])


def _valid_yield_periods(periods: list, val: dict):
    """Periods whose year is the valuation year or the year after (C2); None when the valuation year is unknown."""
    y = _period_sort_key(val.get("year") or "")[0] if val.get("year") else -1
    if y < 0:
        return None
    return [p["id"] for p in periods if _period_sort_key(p["id"])[0] in (y, y + 1)]


def init_config(rec_doc: dict) -> dict:
    records = rec_doc.get("records") or []
    meta = rec_doc.get("field_meta") or {}
    file_name = (rec_doc.get("source") or {}).get("name")
    vocab = load_vocab()
    dim_labels = vocab.get("dims") or {}
    warnings = []
    in_scope = [r for r in records if r.get("scope") == "in_scope"]
    if not in_scope:
        warnings.append(OrderedDict([
            ("code", "no_in_scope_records"), ("period", None), ("field", None),
            ("text", "لا توجد سجلات ضمن النطاق في 03؛ لم تُحدد فترات إيراد ولا حقل تقييم، ولا تُستخدم حقول السجلات "
                     "المعزولة أو المستبعدة في المؤشرات. راجع قرارات النطاق في 02 ثم أعد التشغيل.")]))
    periods = build_periods(records, meta, warnings)
    val = build_valuation(records, meta)
    p0 = periods[0] if periods else None
    valid_yield = _valid_yield_periods(periods, val) if periods else None
    if valid_yield is not None and len(valid_yield) < len(periods):
        outside = [p["id"] for p in periods if p["id"] not in valid_yield]
        warnings.append(OrderedDict([
            ("code", "yield_periods_limited"), ("period", None), ("field", val.get("field")),
            ("text", f"تقييم {val.get('year')} يقابل إيراد الفترات {'، '.join(valid_yield) or 'لا شيء'} فقط؛ مؤشرات العائد "
                     f"للفترات {'، '.join(outside)} «غير قابلة للمقارنة» (valid_periods). لحذف فترة من 04: احذف كائنها من "
                     "periods ومعرّفها من valid_periods، واجعل prev_id للفترة التالية لها null مع إبقاء prev_rev_field "
                     "إن أردت حساب النمو.")]))

    def ctx_of(p):
        return me.placeholder_context({"periods": [p] if p else [], "valuation": val}, p["id"] if p else None)

    src_cache = {}

    def src_of(where, p):
        key = (json.dumps(where, sort_keys=True, ensure_ascii=False), p["id"] if p else None)
        if key not in src_cache:
            src_cache[key] = _sources_text(_records_for(records, BASE_WHERE, where, ctx_of(p)), file_name)
        return src_cache[key]

    metrics = []
    for base in standard_metrics():
        mid = base["id"]
        spec = OrderedDict()
        spec["id"] = mid
        texts = metric_texts(mid, p0, val, meta, src_of(base.get("where"), p0))
        spec["label"] = texts.get("label")
        spec["short_label"] = texts.get("short_label")
        for k in ("card", "card_order", "card_group", "kind"):
            spec[k] = base.get(k)
        for k in ("field", "num", "den", "a", "b", "prev", "curr", "fields", "top", "scale"):
            if k in base:
                spec[k] = copy.deepcopy(base[k])
        spec["where"] = copy.deepcopy(base.get("where"))
        spec["unit"] = base["unit"]
        for k in SPEC_TEXT_KEYS:
            if k in texts or k != "empty_text":
                spec[k] = texts.get(k)
        if "empty_as_zero" in base:
            spec["empty_as_zero"] = base["empty_as_zero"]
        if mid in VAL_RATIO_IDS and valid_yield is not None:
            spec["valid_periods"] = list(valid_yield)
            spec["invalid_period_reason"] = (
                f"تقييم {val.get('year')} لا يقابل إيراد هذه الفترة؛ يُحسب العائد لفترات الإيراد "
                f"{'، '.join(valid_yield) or '—'} فقط (سنة التقييم والسنة التالية لها)، ولا يُقسم إيراد سنة على تقييم "
                "سنة بعيدة عنها")
        spec["kind_tag"] = base["kind_tag"]
        if mid == "managed_records_count" and not any(
                (r.get("dims") or {}).get("managed_for_others") is True for r in in_scope):
            spec["card"] = False
            spec["card_group"] = None
            spec["card_order"] = None
        if mid in PERIOD_DEPENDENT and len(periods) > 1:
            spec["text_by_period"] = OrderedDict(
                (p["id"], {k: v for k, v in metric_texts(mid, p, val, meta, src_of(base.get("where"), p)).items()})
                for p in periods)
        metrics.append(spec)

    # C19: a breakdown is drawn only when its own records hold at least 2 distinct classified values
    val_where = {"all": [W_ASSET, W_NOT_MANAGED]}
    rev_where = {"all": [W_REV_REC, W_NOT_MANAGED]}
    val_recs = _records_for(records, BASE_WHERE, val_where, {})
    rev_recs = _records_for(records, BASE_WHERE, rev_where, {})
    breakdowns = []
    for dim in BREAKDOWN_DIMS:
        if _distinct_known(val_recs, dim) < 2:
            continue
        dl = dim_labels.get(dim, dim)
        overlap = " (قد تتداخل الفئات)" if dim == "ownership" else ""
        breakdowns.append(OrderedDict([
            ("id", f"value_by_{dim}"), ("label", f"توزيع القيمة المسجلة حسب {dl}{overlap}"), ("field", "{VAL}"),
            ("by", dim), ("where", copy.deepcopy(val_where)), ("unit", "SAR"),
            ("valuation_date", val.get("date_text")),
            ("unavailable_reason", val.get("reason") or "لا يوجد تقييم إجمالي مسجل")]))

    def rev_tag(p):
        return " (شامل الضريبة)" if p and (meta.get(p["rev_field"]) or {}).get("vat") == "incl" else ""

    for dim in BREAKDOWN_DIMS:
        if _distinct_known(rev_recs, dim) < 2:
            continue
        dl = dim_labels.get(dim, dim)
        overlap = " (قد تتداخل الفئات)" if dim == "ownership" else ""
        b = OrderedDict([
            ("id", f"revenue_by_{dim}"),
            ("label", f"توزيع الإيراد المسجل {p0['id'] if p0 else ''}{rev_tag(p0)} حسب {dl}{overlap}".replace("  ", " ")),
            ("field", "{REV}"), ("by", dim), ("where", copy.deepcopy(rev_where)), ("unit", "SAR"),
            ("period_text", p0["label"] if p0 else None),
            ("unavailable_reason", "لا يوجد حقل إيراد لهذه الفترة")])
        if len(periods) > 1:
            b["text_by_period"] = OrderedDict(
                (p["id"], {"label": f"توزيع الإيراد المسجل {p['id']}{rev_tag(p)} حسب {dl}{overlap}", "period_text": p["label"]})
                for p in periods)
        breakdowns.append(b)

    asc = sorted(periods, key=lambda p: _period_sort_key(p["id"]))
    tcomp = trend_comparability([(p["id"], p["rev_field"]) for p in asc], meta)
    trend = OrderedDict([
        ("enabled", len(asc) >= 2),
        ("label", "تطور الإيراد العقاري المسجل عبر الفترات"),
        ("unit", "SAR"),
        ("fields", [OrderedDict([("period", p["id"]), ("field", p["rev_field"]), ("label", p["label"])]) for p in asc]),
        ("where", {"all": [W_RENT, W_REV_REC, W_NOT_MANAGED]}),
        ("comparable", tcomp),
        ("note", "مجموع السجلات المعروف إيرادها في كل فترة؛ عدد السجلات المعروفة قد يختلف بين الفترات (n_known)، لذا "
                 "تُحسب أيضًا سلسلة للسجلات نفسها المعروفة في كل الفترات (same_records)"),
    ])

    filters = []
    for dim in FILTER_DIMS:
        if len(_dim_values(in_scope, dim)) >= 2:
            filters.append(dim)

    cfg = OrderedDict()
    cfg["contract_version"] = CONTRACT_VERSION
    cfg["currency"] = rec_doc.get("currency") or "SAR"
    cfg["periods"] = periods
    cfg["valuation"] = val
    cfg["base_where"] = copy.deepcopy(BASE_WHERE)
    cfg["metrics"] = metrics
    cfg["breakdowns"] = breakdowns
    cfg["trend"] = trend
    cfg["filters"] = filters
    cfg["parity_max_values_per_dim"] = 3
    cfg["init_warnings"] = warnings
    return cfg


# ---------------------------------------------------------------------------
# Compute
# ---------------------------------------------------------------------------
def parity_cases(eng: me.Engine, config: dict) -> list:
    cases = []
    seen = set()

    def add(f):
        key = json.dumps(f, sort_keys=True, ensure_ascii=False)
        if key in seen:
            return
        seen.add(key)
        res = eng.compute_all(f)
        cases.append(OrderedDict([("filters", f),
                                  ("metrics", OrderedDict((mid, me.compact(r)) for mid, r in res["metrics"].items()))]))

    default = {"period": eng.default_period} if eng.default_period is not None else {}
    add(dict(default))
    for pid in eng.period_ids:
        add({"period": pid})
    nmax = int(config.get("parity_max_values_per_dim", 3) or 0)
    base_recs = eng.filter_records(default)
    for dim in config.get("filters") or []:
        for v in _dim_values(base_recs, dim)[:nmax]:
            f = dict(default)
            f[dim] = v
            add(f)
    return cases


def compute(rec_doc: dict, config: dict, records_path: str = None, config_path: str = None) -> dict:
    eng = me.Engine(config, rec_doc.get("records") or [])
    default = {"period": eng.default_period} if eng.default_period is not None else {}
    res = eng.compute_all(default)
    out = OrderedDict()
    out["contract_version"] = CONTRACT_VERSION
    out["inputs"] = OrderedDict([("records", records_path), ("config", config_path),
                                 ("source_sha256", (rec_doc.get("source") or {}).get("sha256"))])
    out["default_filters"] = default
    out["metrics"] = res["metrics"]
    out["breakdowns"] = res["breakdowns"]
    out["trend"] = res["trend"]
    out["records_in_scope"] = res["records"]
    out["parity_cases"] = parity_cases(eng, config)
    out["engine_notes"] = me.IMPLEMENTATION_NOTES
    return out


# ---------------------------------------------------------------------------
# Printing
# ---------------------------------------------------------------------------
def fmt_value(v, unit):
    if v is None:
        return "—"
    if unit == "%":
        return f"{v:,.2f}%"
    if unit == "SAR":
        return f"{v:,.2f} ريال" if isinstance(v, float) and not float(v).is_integer() else f"{int(v):,} ريال"
    if isinstance(v, float) and not v.is_integer():
        return f"{v:,.2f}"
    return f"{int(v):,}"


def print_table(res: dict, filters: dict) -> None:
    try:
        vocab = load_vocab()
        st_ar = {k: v.get("ar") for k, v in (vocab.get("metric_status") or {}).items()}
    except Exception:
        st_ar = {}
    desc = "، ".join(f"{k} = {v}" for k, v in filters.items() if v is not None) or "دون تصفية (لا توجد فترات)"
    print(f"التصفية: {desc}")
    print(f"السجلات ضمن النطاق بعد التصفية: {len(res['records'])}")
    rows = [("المؤشر", "القيمة", "الحالة", "المشمول/النطاق")]
    for mid, r in res["metrics"].items():
        st = r.get("status")
        st_txt = st_ar.get(st) or STATUS_AR_FALLBACK.get(st, str(st))
        if st == "empty" and r.get("empty_cause") == "scope":
            st_txt = "لا توجد سجلات مؤهلة في الملف"  # the vocab text speaks of a filter (C3)
        rows.append((f"{r.get('short_label') or r.get('label')} [{mid}]", fmt_value(r.get("value"), r.get("unit")),
                     st_txt, f"{r.get('n_included')}/{r.get('n_scope')}"))
    widths = [max(len(str(row[i])) for row in rows) for i in range(4)]
    for i, row in enumerate(rows):
        print(" | ".join(str(c).ljust(widths[j]) for j, c in enumerate(row)))
        if i == 0:
            print("-+-".join("-" * w for w in widths))
    for mid, r in res["metrics"].items():
        if r.get("reason") and r.get("status") in ("not_computable", "not_comparable", "empty"):
            print(f"  • {mid}: {r['reason']}")
        elif r.get("note"):
            print(f"  • {mid}: {r['note']}")
        if r.get("revenue_start_ids"):
            print(f"  • {mid}: بدء إيراد (خارج النسبة): {'، '.join(r['revenue_start_ids'])} "
                  f"بتغير {fmt_value(r.get('start_change'), 'SAR')}")
    t = res.get("trend") or {}
    if t.get("points"):
        pts = "، ".join(f"{p['period']}: {fmt_value(p['value'], 'SAR')} ({p['n_known']}/{p['n']})" for p in t["points"])
        print(f"التطور الزمني ({t.get('comparable', {}).get('status')}): {pts}")
        sr = t.get("same_records") or {}
        if t.get("n_known_varies") and sr.get("points"):
            spts = "، ".join(f"{p['period']}: {fmt_value(p['value'], 'SAR')}" for p in sr["points"])
            print(f"  تنبيه: السجلات المعروفة تختلف بين الفترات؛ السجلات نفسها ({sr.get('n')}) في كل الفترات: {spts}")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="إعداد تعريفات المؤشرات وحسابها (05_metrics.json)")
    ap.add_argument("--records", required=True, help="مسار 03_records.json")
    ap.add_argument("--config", help="مسار 04_metrics_config.json (للحساب)")
    ap.add_argument("--out", help="مسار الإخراج: 04_metrics_config.json مع --init-config أو 05_metrics.json")
    ap.add_argument("--init-config", action="store_true", help="إنشاء مسودة 04_metrics_config.json من السجلات")
    ap.add_argument("--filters", help="تصفية JSON للعرض، مثل '{\"use\":\"investment\",\"period\":\"2024\"}'")
    ap.add_argument("--print", dest="do_print", action="store_true", help="طباعة جدول عربي مختصر")
    ap.add_argument("--schema", default=os.path.join(SKILL_DIR, "schemas", "metrics_config.schema.json"))
    args = ap.parse_args(argv)

    try:
        rec_doc = load_json(args.records)
    except Exception as exc:
        eprint(f"خطأ: تعذر قراءة السجلات ({args.records}): {exc}")
        return 2
    errs = me.check_records(rec_doc.get("records") or [])
    if errs:
        eprint("أخطاء في السجلات:\n" + "\n".join(errs))
        return 2

    if args.init_config:
        if not args.out:
            eprint("خطأ: --out مطلوب مع --init-config")
            return 2
        cfg = init_config(rec_doc)
        verrs = (schema_errors(cfg, args.schema) + semantic_errors(cfg, rec_doc.get("records") or [])
                 + comparability_errors(cfg, rec_doc.get("field_meta") or {}))
        if verrs:
            eprint("الإعداد المولَّد غير مطابق للمخطط (خلل داخلي):\n" + "\n".join(verrs))
            return 2
        save_json(args.out, cfg)
        print(f"تم إنشاء مسودة الإعداد: {args.out}")
        print(f"  الفترات: {', '.join(p['id'] for p in cfg['periods']) or 'لا توجد'} | "
              f"التقييم: {cfg['valuation']['field'] or 'لا يوجد'} | المرشحات: {', '.join(cfg['filters']) or 'لا توجد'}")
        print(f"  المؤشرات: {len(cfg['metrics'])} | التوزيعات: {len(cfg['breakdowns'])} | "
              f"التطور الزمني: {'مفعّل' if cfg['trend']['enabled'] else 'غير مفعّل'}")
        if not cfg["periods"]:
            print("  تنبيه: لا توجد حقول إيراد (group = revenue) في field_meta؛ مؤشرات الإيراد غير قابلة للاحتساب.")
        if not cfg["valuation"]["field"]:
            print(f"  تنبيه: {cfg['valuation']['reason']}")
        for w in cfg.get("init_warnings") or []:
            print(f"  تنبيه [{w['code']}]: {w['text']}")
        return 0

    if not args.config:
        eprint("خطأ: --config مطلوب للحساب (أو استخدم --init-config)")
        return 2
    try:
        config = load_json(args.config)
    except Exception as exc:
        eprint(f"خطأ: تعذر قراءة الإعداد ({args.config}): {exc}")
        return 2
    verrs = (schema_errors(config, args.schema) + semantic_errors(config, rec_doc.get("records") or [])
             + comparability_errors(config, rec_doc.get("field_meta") or {}))
    if verrs:
        eprint(f"ملف الإعداد غير مطابق ({len(verrs)} خطأ):")
        for e in verrs:
            eprint(e)
        return 2
    used = numeric_fields_used(config)
    ferrs = me.check_records(rec_doc.get("records") or [], used)
    if ferrs:
        eprint("أخطاء في قيم الحقول المستخدمة في المؤشرات:\n" + "\n".join(ferrs))
        return 2
    if not args.out and not args.do_print:
        eprint("خطأ: حدّد --out أو --print")
        return 2

    if args.out:
        out = compute(rec_doc, config, args.records, args.config)
        save_json(args.out, out)
        n_part = sum(1 for r in out["metrics"].values() if r["status"] == "partial")
        n_nc = sum(1 for r in out["metrics"].values() if r["status"] in ("not_computable", "not_comparable"))
        print(f"تم: {len(out['metrics'])} مؤشر → {args.out} (جزئي {n_part}، غير قابل للاحتساب/المقارنة {n_nc}، "
              f"حالات تطابق {len(out['parity_cases'])})")

    if args.do_print:
        try:
            filters = json.loads(args.filters) if args.filters else {}
        except json.JSONDecodeError as exc:
            eprint(f"خطأ: --filters ليس JSON صالحًا: {exc}")
            return 2
        eng = me.Engine(config, rec_doc.get("records") or [])
        try:
            f = eng.normalize_filters(filters)
        except ValueError as exc:
            eprint(f"خطأ: {exc}")
            return 2
        print_table(eng.compute_all(f), f)
    return 0


if __name__ == "__main__":
    sys.exit(main())
