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
    rec.<record_id>.<field_id>            the known numeric value of one field of any record of 03 (in scope, isolated
                                          or excluded), e.g. rec.STM-r7.rev.2025.recorded — to state a figure that
                                          another source disputes beside the one the metric uses. Money fields
                                          (revenue/expense/valuation, unit SAR) render as money. An unknown record, an
                                          unknown field or a value that is not a known number is not a fact (the token
                                          is unresolvable: fact_tokens blocks). compute_facts lists the rec.* keys that
                                          06_narrative.json cites ({{f:…}} tokens and «f:…» evidence refs);
                                          record_fact() resolves any other one.
Token grammar: {{f:<key>}} with key matching [A-Za-z0-9_.\-]+ (see FACT_TOKEN_RE); a record id with other characters
cannot be cited as a fact.
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
FACT_REF_RE = re.compile(r"^\s*f:\s*([A-Za-z0-9_.\-]+)\s*$")      # an evidence ref «f:<key>»
REC_PREFIX = "rec."
KNOWN = ("value", "derived", "zero_recorded", "zero_confirmed")
_MONEY_GROUPS = ("revenue", "expense", "valuation")
_MONEY_KEYS = ("value", "recorded", "recomputed", "recomputed_from_details", "cached", "difference", "diff",
               "sum", "total", "expected", "sections_sum")


def _load(path):
    if not os.path.isfile(path):
        return None
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)


def _num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def _unit_of(field: str, meta: dict) -> str:
    """SAR for revenue/expense/valuation fields (formatted as money like other money facts), pct for percent fields."""
    m = meta.get(field) or meta.get(field.replace("__original", "")) or {}
    unit = str(m.get("unit") or "").strip()
    if m.get("group") in _MONEY_GROUPS or unit.upper() in ("SAR", "SR") or (
            not m.get("group") and field.startswith(("rev.", "exp.", "val.", "noi."))):
        return "SAR"
    if unit in ("%", "pct", "percent"):
        return "pct"
    return "number"


def record_fact(rec_doc: dict, key: str):
    """Resolve «rec.<record_id>.<field_id>» against 03_records.json: (fact dict, None) or (None, reason code).

    Reason codes: not_rec (the key does not start with rec.), unknown_record, unknown_field, not_numeric (the value is
    blank, text, an error, «غير متاح» … — unknown values are never facts). A record id may hold dots (REG-1.2): every
    split of the key is tried and the one whose record holds the field wins."""
    if not isinstance(key, str) or not key.startswith(REC_PREFIX):
        return None, "not_rec"
    rest = key[len(REC_PREFIX):]
    records = {}
    for r in (rec_doc or {}).get("records") or []:
        if isinstance(r, dict) and isinstance(r.get("id"), str):
            records.setdefault(r["id"], r)
    meta = (rec_doc or {}).get("field_meta") or {}
    seen_record = None
    for i, ch in enumerate(rest):
        if ch != ".":
            continue
        rid, field = rest[:i], rest[i + 1:]
        r = records.get(rid)
        if r is None or not field:
            continue
        seen_record = seen_record or r
        tv = (r.get("fields") or {}).get(field)
        if not isinstance(tv, dict):
            continue
        if tv.get("s") not in KNOWN or not _num(tv.get("v")):
            return None, "not_numeric"
        lab = (meta.get(field) or {}).get("label") or field
        name = r.get("name") or rid
        return {"value": tv["v"], "unit": _unit_of(field, meta), "label": f"{lab} — {name} ({rid})",
                "source": f"03_records.json:records[{rid}].fields.{field}", "ref": tv.get("ref")}, None
    return None, ("unknown_field" if seen_record is not None else "unknown_record")


REC_REASON_AR = {
    "unknown_record": "لا يوجد سجل بهذا المعرّف في 03_records.json",
    "unknown_field": "السجل موجود لكن الحقل غير موجود فيه",
    "not_numeric": "قيمة الحقل ليست رقمًا معروفًا (فارغة أو نص أو خطأ أو غير متاحة)",
    "not_rec": "المفتاح لا يبدأ بـ rec.",
}


def _cited_keys(narrative) -> list:
    """Fact keys cited in 06: {{f:…}} tokens in any string and «f:…» evidence refs (strings or {ref} objects)."""
    out = []

    def walk(x):
        if isinstance(x, str):
            out.extend(m.group(1) for m in FACT_TOKEN_RE.finditer(x))
            m = FACT_REF_RE.match(x)
            if m:
                out.append(m.group(1))
        elif isinstance(x, list):
            for y in x:
                walk(y)
        elif isinstance(x, dict):
            for y in x.values():
                walk(y)
    walk(narrative)
    return list(dict.fromkeys(out))


def record_fact_example(rec_doc: dict):
    """(key, fact) of one resolvable rec.* key for the listing: an isolated money field first, else any number."""
    best = None
    meta = (rec_doc or {}).get("field_meta") or {}
    for r in (rec_doc or {}).get("records") or []:
        if not isinstance(r, dict) or not isinstance(r.get("id"), str) or not re.fullmatch(r"[A-Za-z0-9_.\-]+", r["id"]):
            continue
        for f, tv in (r.get("fields") or {}).items():
            if not isinstance(tv, dict) or tv.get("s") not in KNOWN or not _num(tv.get("v")):
                continue
            rank = (0 if r.get("scope") != "in_scope" else 1) + (0 if _unit_of(f, meta) == "SAR" else 2)
            if best is None or rank < best[0]:
                best = (rank, f"{REC_PREFIX}{r['id']}.{f}")
            if rank == 0:
                break
        if best and best[0] == 0:
            break
    if best is None:
        return None, None
    return best[1], record_fact(rec_doc, best[1])[0]


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
    # rec.<record_id>.<field_id>: the record values 06 cites (tokens and evidence refs); any other one resolves
    # through record_fact() — never one entry per record and field
    if records:
        narrative = _load(os.path.join(run_dir, "06_narrative.json"))
        for key in _cited_keys(narrative) if narrative else []:
            if key.startswith(REC_PREFIX) and key not in out:
                fact, _why = record_fact(rec, key)
                if fact is not None:
                    out[key] = fact
    return out


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser(description="List the {{f:key}} facts available to the narrative for a run folder.")
    ap.add_argument("--run-dir", required=True)
    a = ap.parse_args(argv)
    for k, f in compute_facts(a.run_dir).items():
        if k.startswith(REC_PREFIX):
            continue   # record values: the pattern below, not one line per record and field
        print(f"{{{{f:{k}}}}}\t{f['value']}\t{f['unit']}\t{f['label']}")
    rec = _load(os.path.join(a.run_dir, "03_records.json")) or {}
    key, fact = record_fact_example(rec)
    print("{{f:rec.<record_id>.<field_id>}}\t—\t—\tالقيمة الرقمية المعروفة لحقل أي سجل في 03_records.json (ضمن النطاق "
          "أو معزول أو مستبعد)؛ حقول الإيراد والمصروف والتقييم تُعرض مبلغًا"
          + (f"؛ مثال: {{{{f:{key}}}}} = {fact['value']}" if key and fact else ""))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
