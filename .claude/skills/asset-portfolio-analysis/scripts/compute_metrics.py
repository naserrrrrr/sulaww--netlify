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

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from apa_common import CONTRACT_VERSION, KNOWN_STATUSES, SKILL_DIR, load_json, load_vocab, save_json  # noqa: E402
import metric_engine as me  # noqa: E402
from normalize import schema_errors as _schema_errors  # noqa: E402

ASSET_TYPES = ["asset", "land_parcel"]
W_ASSET = {"in": ["record_type", ASSET_TYPES]}
W_NOT_MANAGED = {"not": {"eq": ["managed_for_others", True]}}
W_RENT = {"eq": ["income_nature", "rent"]}
W_NOT_OPERATING = {"not": {"eq": ["income_nature", "operating"]}}
W_INVEST = {"eq": ["use", "investment"]}
BREAKDOWN_DIMS = ["asset_type", "location", "use", "ownership", "income_nature", "rent_allocation", "restriction"]
FILTER_DIMS = ["asset_type", "location", "use"]

VAT_AR = {"excl": "بدون ضريبة", "incl": "شامل الضريبة", "none_evidenced": "لا ضريبة مثبتة في المصدر",
          "unknown": "أساس الضريبة غير موضح"}
BASIS_AR = {"actual": "فعلي", "expected": "متوقع", "recorded": "مسجل", "contract": "تعاقدي", "unknown": "أساس الإيراد غير موضح"}
COMPLETENESS_AR = {"full_year": "سنة كاملة", "partial": "فترة جزئية", "unknown": "اكتمال الفترة غير موضح"}
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


def semantic_errors(config: dict) -> list:
    errs = []
    ids = [p.get("id") for p in config.get("periods") or []]
    if len(ids) != len(set(ids)):
        errs.append("• periods: معرّفات فترات مكررة")
    for p in config.get("periods") or []:
        if p.get("prev_id") and p["prev_id"] not in ids:
            errs.append(f"• periods/{p['id']}: prev_id «{p['prev_id']}» غير موجود")
    mids = [m.get("id") for m in config.get("metrics") or []]
    dup = sorted({x for x in mids if mids.count(x) > 1})
    if dup:
        errs.append("• metrics: معرّفات مؤشرات مكررة: " + "، ".join(dup))
    bids = [b.get("id") for b in config.get("breakdowns") or []]
    dupb = sorted({x for x in bids if bids.count(x) > 1})
    if dupb:
        errs.append("• breakdowns: معرّفات مكررة: " + "، ".join(dupb))
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


def comparability(cm: dict, pm: dict, cur_id: str, prev_id: str) -> dict:
    keys = (("vat", VAT_AR, "الضريبة"), ("basis", BASIS_AR, "أساس الإيراد"), ("completeness", COMPLETENESS_AR, "اكتمال الفترة"))
    if cm.get("completeness") == "partial" or pm.get("completeness") == "partial":
        which = [x for x, m in ((cur_id, cm), (prev_id, pm)) if m.get("completeness") == "partial"]
        return {"status": "not_comparable",
                "reason": f"الفترة {'، '.join(which)} جزئية؛ لا تُقارن بفترة كاملة ولا تُحوَّل إلى أداء سنوي"}
    same = all(cm.get(k) == pm.get(k) and cm.get(k) not in (None, "unknown") for k, _, _ in keys)
    if same:
        return {"status": "confirmed", "reason": "أساس الضريبة والإيراد واكتمال الفترة متطابقة وفق وصف الحقول"}
    diffs = []
    for k, ar, lab in keys:
        a, b = cm.get(k) or "unknown", pm.get(k) or "unknown"
        if a == b:
            if a == "unknown":
                diffs.append(f"{lab}: غير موضح في الفترتين")
        else:
            diffs.append(f"{lab}: {ar.get(a, a)} ({cur_id}) مقابل {ar.get(b, b)} ({prev_id})")
    return {"status": "unverified", "reason": "لم يثبت تماثل الفترتين — " + "؛ ".join(diffs)}


def _pool(records: list) -> list:
    """Records used to decide which fields are live: in-scope ones, or all when none is in scope."""
    ins = [r for r in records if r.get("scope") == "in_scope"]
    return ins or list(records)


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


def build_periods(records: list, meta: dict) -> list:
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
                                                      p["id"], prev["id"])
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
    sheets = []
    for r in records:
        s = (r.get("source") or {}).get("sheet")
        if s and s not in sheets:
            sheets.append(s)
    return f"الملف {file_name or '؟'} — الورقة: " + "، ".join(f"«{s.strip()}»" for s in sheets)


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
    T = {}
    T["asset_count"] = dict(
        label="عدد الأصول ضمن النطاق", short_label="الأصول",
        definition="عدد سجلات الأصول وقطع الأراضي ضمن نطاق الدراسة، كل سجل مرة واحدة، دون العقارات المدارة للغير (تُعرض منفصلة)؛ لا يُستنتج العدد من آخر رقم مسلسل ولا تُعد الوحدات أصولًا مستقلة.",
        scope_text="السجلات من نوع «أصل» أو «قطعة أرض» المصنفة ضمن النطاق، عدا المدار للغير", basis="عدّ سجلات",
        period_text="كما في الملف", source_text=src)
    T["managed_records_count"] = dict(
        label="عدد العقارات المدارة للغير", short_label="مدارة للغير",
        definition="عدد العقارات التي تديرها الجهة لصالح الغير ضمن النطاق؛ تُعرض منفصلة، ولا تُضاف إلى عدد أصول الجهة، ولا تدخل إيراداتها أو قيمها في مؤشرات الجهة.",
        scope_text="السجلات المصنفة «إدارة أملاك للغير» ضمن النطاق", basis="عدّ سجلات",
        period_text="كما في الملف", source_text=src)
    T["operating_records_count"] = dict(
        label="عدد السجلات التشغيلية", short_label="السجلات التشغيلية",
        definition="عدد سجلات الأنشطة التشغيلية (مثل المدارس) ضمن النطاق؛ تُعرض منفصلة ولا تُضاف إلى عدد الأصول.",
        scope_text="السجلات من نوع «نشاط تشغيلي» ضمن النطاق", basis="عدّ سجلات", period_text="كما في الملف",
        source_text=src)
    T["recorded_value"] = dict(
        label="القيمة المسجلة للأصول المقيمة", short_label="القيمة المسجلة",
        definition=f"مجموع إجمالي {val.get('label') or 'التقييم'} المسجل في الملف للأصول المقيمة ضمن النطاق، مرة واحدة لكل أصل ودون جمع الأرض والمبنى مع إجماليهما. أحدث تقييم مسجل وليس قيمة سوقية حالية، ولا يمثل القيمة الكاملة للمحفظة إذا وُجدت أصول غير مقيمة (تظهر في المستبعدات).",
        scope_text="الأصول وقطع الأراضي ضمن النطاق، عدا العقارات المدارة للغير", basis=vbasis,
        valuation_date=vdate, period_text=f"تاريخ التقييم: {vdate}", source_text=val_src, unavailable_reason=no_val)
    T["property_revenue"] = dict(
        label=f"الإيرادات العقارية المسجلة {P}", short_label="الإيرادات العقارية",
        definition=f"مجموع «{rev_lab}» للسجلات ذات طبيعة الدخل «إيجارات عقارية» ضمن النطاق، دون العقارات المدارة للغير ودون إيرادات الأنشطة التشغيلية. الخانات الفارغة لا تُعامل كصفر.",
        scope_text="سجلات الإيجارات العقارية ضمن النطاق، عدا المدار للغير", basis=f"{rev_desc}؛ قبل المصروفات",
        period_text=period_text, source_text=rev_src, unavailable_reason=no_rev)
    T["activity_revenue"] = dict(
        label=f"إيرادات الأنشطة التشغيلية {P}", short_label="إيرادات الأنشطة",
        definition=f"مجموع «{rev_lab}» للسجلات ذات طبيعة الدخل «نشاط تشغيلي» (مثل المدارس). إيراد وليس ربحًا، ولا يُعامل كإيجار عقاري ولا يدخل في عائد العقارات.",
        scope_text="سجلات الأنشطة التشغيلية ضمن النطاق", basis=f"{rev_desc}؛ إيراد إجمالي قبل تكاليف النشاط",
        period_text=period_text, source_text=rev_src, unavailable_reason=no_rev)
    yield_def = (f"مجموع إيراد {P} ÷ مجموع {val.get('label') or 'التقييم'} للأصول نفسها التي يتوافر لها الرقمان × 100، "
                 "من المجاميع المتطابقة لا من متوسط عوائد الأصول، ويشمل الأصول ذات الإيراد الصفري المسجل. "
                 "إذا كانت الحالة «جزئي» فهو «عائد الجزء القابل للاحتساب». ليس عائدًا صافيًا ولا عائدًا استثماريًا كليًا.")
    T["gross_yield_investment"] = dict(
        label=f"الإيراد المسجل لعام {P} إلى تقييم {Y} للأصول الاستثمارية نفسها، قبل المصروفات",
        short_label="عائد الدخل الإجمالي — الأصول الاستثمارية",
        definition=yield_def,
        scope_text="الأصول الاستثمارية ضمن النطاق، عدا الأنشطة التشغيلية والعقارات المدارة للغير",
        basis=f"قبل المصروفات؛ {rev_desc}؛ {vbasis}", valuation_date=vdate, period_text=period_text,
        source_text=both_src, unavailable_reason=no_both)
    T["gross_yield_owned"] = dict(
        label=f"الإيراد المسجل لعام {P} إلى تقييم {Y} لجميع الأصول المملوكة، قبل المصروفات",
        short_label="عائد الدخل الإجمالي — جميع الأصول المملوكة",
        definition=yield_def + " يشمل الأصول الخدمية والإدارية؛ إدراجها في نسبة عامة لا يكفي للحكم على أدائها لأن غرضها خدمة المستفيدين.",
        scope_text="جميع الأصول المملوكة ضمن النطاق، عدا الأنشطة التشغيلية", basis=f"قبل المصروفات؛ {rev_desc}؛ {vbasis}",
        valuation_date=vdate, period_text=period_text, source_text=both_src, unavailable_reason=no_both)
    T["gross_yield_income_producing"] = dict(
        label=f"الإيراد المسجل لعام {P} إلى تقييم {Y} للأصول الاستثمارية المدرة فقط، قبل المصروفات",
        short_label="عائد الدخل الإجمالي — الأصول المدرة فقط",
        definition=yield_def + f" يقتصر على الأصول الاستثمارية ذات إيراد موجب في {P}؛ للمقارنة فقط ولا يحل محل النطاق الاستثماري الكامل.",
        scope_text="الأصول الاستثمارية ذات الإيراد الموجب ضمن النطاق", basis=f"قبل المصروفات؛ {rev_desc}؛ {vbasis}",
        valuation_date=vdate, period_text=period_text, source_text=both_src, unavailable_reason=no_both)
    T["noi"] = dict(
        label=f"صافي الدخل التشغيلي {P}", short_label="صافي الدخل التشغيلي",
        definition="الإيراد العقاري ناقص مصروفات التشغيل المرتبطة به لكل أصل يتوافر له الرقمان، دون الإهلاك والتمويل والنفقات الرأسمالية. لا يساوي المبلغ المتاح لخدمة الأيتام.",
        scope_text="سجلات الإيجارات العقارية ضمن النطاق، عدا المدار للغير", basis=f"{rev_desc}",
        period_text=period_text, source_text=rev_src,
        unavailable_reason="لا توجد بيانات مصروفات تشغيل في الملف لهذه الفترة؛ غياب المصروفات لا يعني أنها صفر")
    T["zero_rev_investment_value"] = dict(
        label="قيمة الأصول الاستثمارية المسجلة بإيراد صفري", short_label="قيمة الاستثماري بإيراد صفري",
        definition=f"مجموع {val.get('label') or 'التقييم'} للأصول الاستثمارية التي إيرادها في {P} صفر مدوّن في الملف (مؤكد بدليل أو غير متحقق من سببه). الصفر لا يثبت الشغور أو تعثر التحصيل، والخانات الفارغة لا تُحسب هنا.",
        scope_text="الأصول الاستثمارية ضمن النطاق ذات إيراد صفري مدوّن، عدا المدار للغير", basis=vbasis,
        valuation_date=vdate, period_text=period_text, source_text=both_src, unavailable_reason=no_both)
    T["data_coverage"] = dict(
        label="تغطية بيانات الإيراد والتقييم", short_label="تغطية البيانات",
        definition=f"أدنى نسبة من الأصول ضمن النطاق يتوافر لها إيراد {P} معروف وتقييم {Y} معروف (تُحسب لكل حقل ويؤخذ الأدنى). تغطية بعدد السجلات لا بالقيمة.",
        scope_text="الأصول وقطع الأراضي ضمن النطاق، عدا المدار للغير", basis="عدّ سجلات ذات قيم معروفة", valuation_date=vdate,
        period_text=period_text, source_text=both_src, unavailable_reason=no_both)
    comp = (p or {}).get("comparable_with_prev") or {}
    T["revenue_growth"] = dict(
        label=f"تغير الإيراد العقاري المسجل {PREV} ← {P}", short_label="تغير الإيراد العقاري",
        definition=("التغير في مجموع الإيراد للأصول نفسها المعروف إيرادها في الفترتين؛ الأصول المعروفة في فترة واحدة تُعرض منفصلة "
                    "(إضافة/تخارج أو نقص بيانات). إذا كان إيراد الفترة السابقة صفرًا يُعرض «بدء إيراد» والتغير بالريال دون نسبة. "
                    f"قابلية المقارنة: {comp.get('reason') or 'غير متحققة'}."),
        scope_text="سجلات الإيجارات العقارية ضمن النطاق، عدا المدار للغير", basis=f"{rev_desc}",
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
        scope_text="جميع السجلات ضمن النطاق ذات إيراد موجب معروف، عدا المدار للغير", basis=f"{rev_desc}",
        period_text=period_text, source_text=rev_src, unavailable_reason=no_rev)
    return T.get(mid, {})


def standard_metrics() -> list:
    """Structure of the standard metrics (texts are added by init_config)."""
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
        dict(id="property_revenue", kind="sum", field="{REV}", where={"all": [W_RENT, W_NOT_MANAGED]}, unit="SAR",
             card=True, card_group="property_revenue", card_order=4, kind_tag="derived"),
        dict(id="activity_revenue", kind="sum", field="{REV}", where={"eq": ["income_nature", "operating"]},
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
        dict(id="noi", kind="diff", a="{REV}", b="{EXP}", where={"all": [W_RENT, W_NOT_MANAGED]}, unit="SAR",
             card=True, card_group="noi", card_order=7, kind_tag="derived"),
        dict(id="zero_rev_investment_value", kind="sum", field="{VAL}",
             where={"all": [W_INVEST, W_ASSET, W_NOT_MANAGED,
                            {"status": ["{REV}", ["zero_recorded", "zero_confirmed"]]}]},
             unit="SAR", card=True, card_group="zero_rev_investment_value", card_order=8, kind_tag="derived"),
        dict(id="data_coverage", kind="coverage", fields=["{REV}", "{VAL}"], where=W_ASSET_OWN, unit="%", card=True,
             card_group="data_coverage", card_order=9, kind_tag="derived"),
        dict(id="revenue_growth", kind="growth", prev="{REV_PREV}", curr="{REV}",
             where={"all": [W_RENT, W_NOT_MANAGED]}, unit="%", card=False, card_group=None, card_order=None,
             kind_tag="derived"),
        dict(id="value_concentration_top3", kind="concentration", field="{VAL}", top=3, where=W_ASSET_OWN, unit="%",
             card=False, card_group=None, card_order=None, kind_tag="derived"),
        dict(id="revenue_concentration_top3", kind="concentration", field="{REV}", top=3, where=W_NOT_MANAGED,
             unit="%", card=False, card_group=None, card_order=None, kind_tag="derived"),
    ]


PERIOD_DEPENDENT = {"property_revenue", "activity_revenue", "gross_yield_investment", "gross_yield_owned",
                    "gross_yield_income_producing", "noi", "zero_rev_investment_value", "data_coverage",
                    "revenue_growth", "revenue_concentration_top3"}


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


def _dim_informative(records: list, dim: str) -> bool:
    return any(v not in ("unknown",) for v in _dim_values(records, dim))


def init_config(rec_doc: dict) -> dict:
    records = rec_doc.get("records") or []
    meta = rec_doc.get("field_meta") or {}
    vocab = load_vocab()
    dim_labels = vocab.get("dims") or {}
    src = _sources_text(records, (rec_doc.get("source") or {}).get("name"))
    periods = build_periods(records, meta)
    val = build_valuation(records, meta)
    p0 = periods[0] if periods else None
    in_scope = [r for r in records if r.get("scope") == "in_scope"]

    metrics = []
    for base in standard_metrics():
        mid = base["id"]
        spec = OrderedDict()
        spec["id"] = mid
        texts = metric_texts(mid, p0, val, meta, src)
        spec["label"] = texts.get("label")
        spec["short_label"] = texts.get("short_label")
        for k in ("card", "card_order", "card_group", "kind"):
            spec[k] = base.get(k)
        for k in ("field", "num", "den", "a", "b", "prev", "curr", "fields", "top", "scale"):
            if k in base:
                spec[k] = copy.deepcopy(base[k])
        spec["where"] = copy.deepcopy(base.get("where"))
        spec["unit"] = base["unit"]
        for k in ("definition", "scope_text", "basis", "valuation_date", "period_text", "source_text", "unavailable_reason"):
            spec[k] = texts.get(k)
        spec["kind_tag"] = base["kind_tag"]
        if mid == "managed_records_count" and not any(
                (r.get("dims") or {}).get("managed_for_others") is True for r in in_scope):
            spec["card"] = False
            spec["card_group"] = None
            spec["card_order"] = None
        if mid in PERIOD_DEPENDENT and len(periods) > 1:
            spec["text_by_period"] = OrderedDict(
                (p["id"], {k: v for k, v in metric_texts(mid, p, val, meta, src).items()}) for p in periods)
        metrics.append(spec)

    breakdowns = []
    for dim in BREAKDOWN_DIMS:
        if not _dim_informative(in_scope, dim):
            continue
        dl = dim_labels.get(dim, dim)
        overlap = " (قد تتداخل الفئات)" if dim == "ownership" else ""
        breakdowns.append(OrderedDict([
            ("id", f"value_by_{dim}"), ("label", f"توزيع القيمة المسجلة حسب {dl}{overlap}"), ("field", "{VAL}"),
            ("by", dim), ("where", {"all": [W_ASSET, W_NOT_MANAGED]}), ("unit", "SAR"),
            ("valuation_date", val.get("date_text")),
            ("unavailable_reason", val.get("reason") or "لا يوجد تقييم إجمالي مسجل")]))
    for dim in BREAKDOWN_DIMS:
        if not _dim_informative(in_scope, dim):
            continue
        dl = dim_labels.get(dim, dim)
        overlap = " (قد تتداخل الفئات)" if dim == "ownership" else ""
        b = OrderedDict([
            ("id", f"revenue_by_{dim}"),
            ("label", f"توزيع الإيراد المسجل {p0['id'] if p0 else ''} حسب {dl}{overlap}".replace("  ", " ")),
            ("field", "{REV}"), ("by", dim), ("where", W_NOT_MANAGED), ("unit", "SAR"),
            ("period_text", p0["label"] if p0 else None),
            ("unavailable_reason", "لا يوجد حقل إيراد لهذه الفترة")])
        if len(periods) > 1:
            b["text_by_period"] = OrderedDict(
                (p["id"], {"label": f"توزيع الإيراد المسجل {p['id']} حسب {dl}{overlap}", "period_text": p["label"]})
                for p in periods)
        breakdowns.append(b)

    asc = sorted(periods, key=lambda p: _period_sort_key(p["id"]))
    tmetas = [meta.get(p["rev_field"]) or {} for p in asc]
    if any(m.get("completeness") == "partial" for m in tmetas):
        tcomp = {"status": "not_comparable", "reason": "فترة واحدة على الأقل جزئية؛ القيم تُعرض للاطلاع فقط ولا تُقرأ كنمو"}
    elif tmetas and all(all(m.get(k) == tmetas[0].get(k) and m.get(k) not in (None, "unknown")
                            for k in ("vat", "basis", "completeness")) for m in tmetas):
        tcomp = {"status": "confirmed", "reason": "أساس الضريبة والإيراد واكتمال الفترات متطابقة وفق وصف الحقول"}
    else:
        tcomp = {"status": "unverified",
                 "reason": "لم يثبت تماثل أساس الضريبة والإيراد واكتمال الفترات بين السنوات؛ تُعرض القيم دون حكم بالنمو"}
    trend = OrderedDict([
        ("enabled", len(asc) >= 2),
        ("label", "تطور الإيراد العقاري المسجل عبر الفترات"),
        ("unit", "SAR"),
        ("fields", [OrderedDict([("period", p["id"]), ("field", p["rev_field"]), ("label", p["label"])]) for p in asc]),
        ("where", {"all": [W_RENT, W_NOT_MANAGED]}),
        ("comparable", tcomp),
        ("note", "مجموع السجلات المعروف إيرادها في كل فترة؛ عدد السجلات المعروفة قد يختلف بين الفترات (n_known)"),
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
    cfg["base_where"] = {"all": [{"eq": ["scope", "in_scope"]}]}
    cfg["metrics"] = metrics
    cfg["breakdowns"] = breakdowns
    cfg["trend"] = trend
    cfg["filters"] = filters
    cfg["parity_max_values_per_dim"] = 3
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
        rows.append((f"{r.get('short_label') or r.get('label')} [{mid}]", fmt_value(r.get("value"), r.get("unit")),
                     st_ar.get(r.get("status")) or STATUS_AR_FALLBACK.get(r.get("status"), str(r.get("status"))),
                     f"{r.get('n_included')}/{r.get('n_scope')}"))
    widths = [max(len(str(row[i])) for row in rows) for i in range(4)]
    for i, row in enumerate(rows):
        print(" | ".join(str(c).ljust(widths[j]) for j, c in enumerate(row)))
        if i == 0:
            print("-+-".join("-" * w for w in widths))
    for mid, r in res["metrics"].items():
        if r.get("reason") and r.get("status") in ("not_computable", "not_comparable"):
            print(f"  • {mid}: {r['reason']}")
    t = res.get("trend") or {}
    if t.get("points"):
        pts = "، ".join(f"{p['period']}: {fmt_value(p['value'], 'SAR')} ({p['n_known']}/{p['n']})" for p in t["points"])
        print(f"التطور الزمني ({t.get('comparable', {}).get('status')}): {pts}")


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
        verrs = schema_errors(cfg, args.schema) + semantic_errors(cfg)
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
        return 0

    if not args.config:
        eprint("خطأ: --config مطلوب للحساب (أو استخدم --init-config)")
        return 2
    try:
        config = load_json(args.config)
    except Exception as exc:
        eprint(f"خطأ: تعذر قراءة الإعداد ({args.config}): {exc}")
        return 2
    verrs = schema_errors(config, args.schema) + semantic_errors(config)
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
