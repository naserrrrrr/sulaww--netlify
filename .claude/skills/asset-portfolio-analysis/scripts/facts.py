r"""Facts that live outside 05_metrics.json and may be cited in the narrative via {{f:<key>}} tokens.

Why: rule 12 of SKILL.md says numbers reach the narrative only through tokens. Some numbers the board needs are not
metrics (e.g. how many records were isolated, how many sheets are hidden, the recorded vs recomputed value of a
workbook anomaly). Counting them as metrics would break the in-scope rule, so they are exposed here instead.

API (used by build_dashboard.py to render and by validate_analysis.py to resolve tokens):
    compute_facts(run_dir) -> OrderedDict[key] = {"value": number, "unit": "count"|"SAR"|"number",
                                                  "label": Arabic label, "source": "file:path"}
Keys are stable, ASCII, dot-separated:
    records.total | records.in_scope | records.isolated | records.excluded | records.excluded_rows
    records.isolated_assetlike            isolated records typed asset/land_parcel/unit, or read from a sheet whose
                                          role is asset_register/managed_properties/allocation (as the dashboard)
    records.in_scope.<record_type>        e.g. records.in_scope.asset
    sources.<source_id>.records | sources.<source_id>.isolated
    sheets.total | sheets.hidden
    issues.total | issues.high | issues.medium | issues.low
    anomalies.total | anomalies.high
    wb.<WB-id>.<evidence_key>             numeric evidence of a 01 anomaly (e.g. wb.WB-004.recomputed)
    excluded.<source_id>.r<row>.<field>   numeric values kept on excluded total rows (when normalize stores them)
    sources.<source_id>.<field>.sum | .n_known | .n_zero | .max | .max_share
                                          a revenue/expense field over EVERY record of a source (isolated included):
                                          to describe an isolated statement apart from the metrics (max_share is the
                                          largest line as a % of the sum, unit "pct")
    sections.<source_id>.r<row>.records   records between the previous total row of the source and total row <row>
Token grammar: {{f:<key>}} with key matching [A-Za-z0-9_.\-]+ (see FACT_TOKEN_RE).
"""
from __future__ import annotations

import json
import os
import re
import sys
from collections import OrderedDict

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

PROPERTY_ROLES = ("asset_register", "managed_properties", "allocation")
FACT_TOKEN_RE = re.compile(r"\{\{\s*f:\s*([A-Za-z0-9_.\-]+)\s*\}\}")
_MONEY_KEYS = ("value", "recorded", "recomputed", "recomputed_from_details", "cached", "difference", "diff",
               "sum", "total", "expected", "sections_sum")


def _load(path):
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def _num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def compute_facts(run_dir: str) -> "OrderedDict[str, dict]":
    out: "OrderedDict[str, dict]" = OrderedDict()

    def put(key, value, unit, label, source):
        if _num(value):
            out[key] = {"value": value, "unit": unit, "label": label, "source": source}

    rec = _load(os.path.join(run_dir, "03_records.json")) or {}
    records = [r for r in (rec.get("records") or []) if isinstance(r, dict)]
    if records:
        by_scope = {}
        for r in records:
            by_scope[r.get("scope")] = by_scope.get(r.get("scope"), 0) + 1
        put("records.total", len(records), "count", "عدد السجلات المقروءة من الملف", "03_records.json:records")
        put("records.in_scope", by_scope.get("in_scope", 0), "count", "السجلات ضمن النطاق", "03_records.json:records[scope=in_scope]")
        put("records.isolated", by_scope.get("isolated", 0), "count", "السجلات المعزولة", "03_records.json:records[scope=isolated]")
        put("records.excluded", by_scope.get("excluded", 0), "count", "السجلات المستبعدة", "03_records.json:records[scope=excluded]")
        # same rule as the dashboard's coverage caveat: an isolated record typed asset/land/unit, or read from a sheet
        # whose role in 02 sheets[] names properties (asset_register, managed_properties, allocation)
        mp = _load(os.path.join(run_dir, "02_mapping.json")) or {}
        decls = (rec.get("context") or {}).get("sheets") or mp.get("sheets") or []
        role_of = {str(d.get("sheet") or "").strip(): d.get("role") for d in decls if isinstance(d, dict)}
        iso_like = sum(1 for r in records if r.get("scope") == "isolated" and (
            r.get("record_type") in ("asset", "land_parcel", "unit")
            or role_of.get(str((r.get("source") or {}).get("sheet") or "").strip()) in PROPERTY_ROLES))
        put("records.isolated_assetlike", iso_like, "count", "سجلات معزولة تمثل أصولًا أو أراضي أو وحدات",
            "03_records.json:records[scope=isolated, record_type∈asset|land_parcel|unit أو ورقتها دورها "
            "asset_register|managed_properties|allocation]")
        types = {}
        for r in records:
            if r.get("scope") == "in_scope":
                types[r.get("record_type")] = types.get(r.get("record_type"), 0) + 1
        for t, n in sorted(types.items(), key=lambda kv: str(kv[0])):
            if t:
                put(f"records.in_scope.{t}", n, "count", f"سجلات ضمن النطاق من نوع {t}", f"03_records.json:records[scope=in_scope,record_type={t}]")
        by_src = {}
        for r in records:
            sid = (r.get("source") or {}).get("source_id") or r.get("source_id")
            if not sid:
                continue
            d = by_src.setdefault(sid, {"records": 0, "isolated": 0})
            d["records"] += 1
            if r.get("scope") == "isolated":
                d["isolated"] += 1
        for sid, d in sorted(by_src.items()):
            put(f"sources.{sid}.records", d["records"], "count", f"سجلات المصدر {sid}", f"03_records.json:records[source={sid}]")
            put(f"sources.{sid}.isolated", d["isolated"], "count", f"السجلات المعزولة من المصدر {sid}", f"03_records.json:records[source={sid},scope=isolated]")
    # descriptive facts of each source's revenue/expense fields over all its records (an isolated statement is
    # described apart from the metrics, never added to them)
    meta = rec.get("field_meta") or {}

    def money_field(f):
        g = (meta.get(f) or {}).get("group")
        return g in ("revenue", "expense") or (g is None and str(f).startswith(("rev.", "exp.")))
    src_vals = OrderedDict()
    for r in records:
        sid = (r.get("source") or {}).get("source_id") or r.get("source_id")
        for f, tv in (r.get("fields") or {}).items():
            if not sid or f.endswith("__original") or not money_field(f) or not isinstance(tv, dict):
                continue
            if tv.get("joined_from"):
                continue
            if tv.get("s") in ("value", "derived", "zero_recorded", "zero_confirmed") and _num(tv.get("v")):
                src_vals.setdefault((sid, f), []).append(tv["v"])
    for (sid, f), vals in src_vals.items():
        tot = sum(vals)
        base = f"03_records.json:records[source={sid}].fields.{f}"
        put(f"sources.{sid}.{f}.sum", round(tot, 2), "SAR", f"مجموع {f} لكل سجلات المصدر {sid}", base)
        put(f"sources.{sid}.{f}.n_known", len(vals), "count", f"سجلات المصدر {sid} المعروف فيها {f}", base)
        put(f"sources.{sid}.{f}.n_zero", sum(1 for v in vals if v == 0), "count", f"سجلات المصدر {sid} بقيمة صفر مسجلة في {f}", base)
        put(f"sources.{sid}.{f}.max", max(vals), "SAR", f"أكبر قيمة {f} في المصدر {sid}", base)
        if tot > 0:
            put(f"sources.{sid}.{f}.max_share", round(max(vals) / tot * 100, 4), "pct",
                f"حصة أكبر سطر من مجموع {f} في المصدر {sid} (%)", base)
    exr = rec.get("excluded_rows") or []
    if isinstance(exr, list) and records:
        tot_rows = {}
        for e in exr:
            if isinstance(e, dict) and isinstance(e.get("row"), int) and isinstance(e.get("values"), dict) and e["values"]:
                tot_rows.setdefault(e.get("source_id"), []).append(e["row"])
        for sid, rows in tot_rows.items():
            rows = sorted(set(rows))
            src_rows = [(r.get("source") or {}).get("row") for r in records
                        if ((r.get("source") or {}).get("source_id") or r.get("source_id")) == sid]
            for i, row in enumerate(rows):
                lo = rows[i - 1] if i else None
                n = sum(1 for x in src_rows if isinstance(x, int) and x < row and (lo is None or x > lo))
                if n:
                    put(f"sections.{sid}.r{row}.records", n, "count", f"سجلات القسم فوق صف الإجمالي {row} في المصدر {sid}",
                        f"03_records.json:records[source={sid}, row<{row}]")
    if isinstance(exr, list):
        put("records.excluded_rows", len(exr), "count", "الصفوف المستبعدة (إجماليات وعناوين)", "03_records.json:excluded_rows")
        for e in exr:
            if not isinstance(e, dict):
                continue
            vals = e.get("values") or {}
            if isinstance(vals, dict):
                for f, v in vals.items():
                    vv = v.get("v") if isinstance(v, dict) else v
                    put(f"excluded.{e.get('source_id')}.r{e.get('row')}.{f}", vv, "number",
                        f"القيمة المسجلة في صف الإجمالي المستبعد {e.get('sheet')}!{e.get('row')} للحقل {f}",
                        f"03_records.json:excluded_rows[{e.get('source_id')}:{e.get('row')}].values.{f}")
    issues = [i for i in (rec.get("issues") or []) if isinstance(i, dict)]
    if issues:
        put("issues.total", len(issues), "count", "مشكلات البيانات المسجلة", "03_records.json:issues")
        for sev in ("high", "medium", "low"):
            put(f"issues.{sev}", sum(1 for i in issues if i.get("severity") == sev), "count",
                f"مشكلات البيانات بخطورة {sev}", f"03_records.json:issues[severity={sev}]")

    wb = _load(os.path.join(run_dir, "01_workbook_map.json")) or {}
    sheets = [s for s in (wb.get("sheets") or []) if isinstance(s, dict)]
    if sheets:
        put("sheets.total", len(sheets), "count", "عدد أوراق المصنف", "01_workbook_map.json:sheets")
        hidden = sum(1 for s in sheets if s.get("state") in ("hidden", "veryHidden") or s.get("hidden") is True)
        put("sheets.hidden", hidden, "count", "الأوراق المخفية", "01_workbook_map.json:sheets[state≠visible]")
    anomalies = [a for a in (wb.get("anomalies") or []) if isinstance(a, dict)]
    if anomalies:
        put("anomalies.total", len(anomalies), "count", "ملاحظات فحص المصنف", "01_workbook_map.json:anomalies")
        put("anomalies.high", sum(1 for a in anomalies if a.get("severity") == "high"), "count",
            "ملاحظات فحص المصنف عالية الخطورة", "01_workbook_map.json:anomalies[severity=high]")
        for a in anomalies:
            aid = a.get("id")
            ev = a.get("evidence") or {}
            if not aid or not isinstance(ev, dict):
                continue
            for k, v in ev.items():
                if _num(v) and re.fullmatch(r"[A-Za-z0-9_\-]+", str(k)):
                    unit = "number" if k in _MONEY_KEYS else "count"
                    put(f"wb.{aid}.{k}", v, unit, f"{aid} ({a.get('type')}): {k}", f"01_workbook_map.json:anomalies[{aid}].evidence.{k}")
    return out


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="List the {{f:key}} facts available to the narrative for a run folder.")
    ap.add_argument("--run-dir", required=True)
    a = ap.parse_args(argv)
    for k, f in compute_facts(a.run_dir).items():
        print(f"{{{{f:{k}}}}}\t{f['value']}\t{f['unit']}\t{f['label']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
