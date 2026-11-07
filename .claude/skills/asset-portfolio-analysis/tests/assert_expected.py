#!/usr/bin/env python3
"""Compare a run folder built from the synthetic fixture with tests/synthetic/expected.json.

Usage:
    python3 tests/assert_expected.py --run-dir RUN --expected tests/synthetic/expected.json [--xlsx FIXTURE.xlsx]

Checks (each printed as PASS/FAIL; exit 0 when all pass, 1 otherwise):
  01  sheet states, Excel tables, anomaly types and locations, PII columns, formulas without cache
  03  record counts, every record's type/scope/dims and key typed values, data-issue types
  04  periods, revenue/expense/valuation fields, comparability, filters
  05  every standard metric (value within 1e-6, status, n_scope, n_included, included/excluded ids and extras),
      trend points, and filtered views recomputed with scripts/metric_engine.py (+ parity cases of 05)
  07  no blocking error; expected warnings present
  08  browser check passed
  version_comparison.json  (run compared with itself) has zero changes
  privacy: no fictional supervisor name or phone number appears in any output file (json/md/html/xlsx)
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL_DIR = os.path.dirname(HERE)
sys.dont_write_bytecode = True  # C23: importing metric_engine must not leave scripts/__pycache__
sys.path.insert(0, os.path.join(SKILL_DIR, "scripts"))

TOL = 1e-6


class Report:
    def __init__(self):
        self.passed, self.failed = 0, []

    def check(self, name: str, ok: bool, detail: str = "") -> bool:
        if ok:
            self.passed += 1
        else:
            self.failed.append(f"{name}: {detail}")
            print(f"  FAIL  {name} — {detail}")
        return ok

    def section(self, title: str):
        print(f"[{title}]")


def num_eq(a, b) -> bool:
    if a is None or b is None:
        return a is None and b is None
    if isinstance(a, bool) or isinstance(b, bool):
        return a is b
    try:
        return abs(float(a) - float(b)) <= TOL or abs(float(a) - float(b)) <= 1e-12 * max(abs(float(b)), 1.0)
    except (TypeError, ValueError):
        return a == b


def load(run, name, rep):
    p = os.path.join(run, name)
    if not os.path.isfile(p):
        rep.check(f"file {name}", False, "الملف غير موجود (الخطوة التي تنتجه فشلت أو لم تُنفذ)")
        return None
    try:
        with open(p, encoding="utf-8") as fh:
            return json.load(fh)
    except ValueError as exc:
        rep.check(f"file {name}", False, f"JSON غير صالح: {exc}")
        return None


# ---------------------------------------------------------------------------
def check_workbook_map(wm, exp, rep):
    rep.section("01_workbook_map.json")
    ew = exp["workbook"]
    sheets = {s["name"]: s for s in wm.get("sheets", [])}
    for name, state in ew["sheets"].items():
        rep.check(f"sheet «{name}» state", name in sheets and sheets[name].get("state") == state,
                  f"expected {state}, got {sheets.get(name, {}).get('state') if name in sheets else 'missing'}")
    for t in ew["tables"]:
        tabs = sheets.get(t["sheet"], {}).get("tables") or []
        ok = any(x.get("header_row") == t["header_row"] and list(x.get("data_rows") or []) == t["data_rows"]
                 and x.get("totals_row") == t["totals_row"] for x in tabs)
        rep.check(f"table on «{t['sheet']}»", ok, f"expected {t}, got {tabs}")
    types = {a.get("type") for a in wm.get("anomalies", [])}
    missing = sorted(set(ew["anomaly_types"]) - types)
    rep.check("anomaly types present", not missing, f"missing: {missing}")
    for a in ew["anomalies_at"]:
        cands = [x for x in wm.get("anomalies", []) if x.get("type") == a["type"]
                 and str(x.get("where", "")).startswith(a["sheet"] + "!")]
        if a.get("cell"):
            cands = [x for x in cands if str(x["where"])[len(a["sheet"]) + 1:].split(":")[0] == a["cell"]]
        rep.check(f"anomaly {a['type']} @ {a['sheet']}!{a.get('cell') or '*'}", bool(cands), "not found")
    pii = {(p.get("sheet"), p.get("col")) for p in wm.get("pii_columns", [])}
    for p in ew["pii_columns"]:
        rep.check(f"pii column {p['sheet']}!{p['col']}", (p["sheet"], p["col"]) in pii, "not detected")
    reg = next(iter(ew["sheets"]))
    cm = ((sheets.get(reg) or {}).get("formulas") or {}).get("cached_missing", 0)
    rep.check("formulas without cached value counted", cm >= ew["formulas_cached_missing_min"],
              f"cached_missing = {cm}, expected >= {ew['formulas_cached_missing_min']}")


def check_records(rd, exp, rep):
    rep.section("03_records.json")
    er = exp["records"]
    st = rd.get("stats", {})
    rep.check("record count", len(rd.get("records", [])) == er["count"], f"{len(rd.get('records', []))} != {er['count']}")
    for k in ("by_type", "by_scope", "by_source", "pii_columns_dropped"):
        rep.check(f"stats.{k}", st.get(k) == er[k], f"{st.get(k)} != {er[k]}")
    by_id = {r["id"]: r for r in rd.get("records", [])}
    for rid, e in er["items"].items():
        r = by_id.get(rid)
        if not rep.check(f"record {rid} exists", r is not None, "missing"):
            continue
        src = r.get("source") or {}
        bad = []
        if src.get("source_id") != e["source_id"] or src.get("row") != e["row"]:
            bad.append(f"source {src.get('source_id')}:{src.get('row')} != {e['source_id']}:{e['row']}")
        for k in ("record_type", "scope"):
            if r.get(k) != e[k]:
                bad.append(f"{k} {r.get(k)} != {e[k]}")
        for k, v in e["dims"].items():
            if (r.get("dims") or {}).get(k) != v:
                bad.append(f"dims.{k} {(r.get('dims') or {}).get(k)!r} != {v!r}")
        for f, tv in e["fields"].items():
            got = (r.get("fields") or {}).get(f)
            if got is None:
                bad.append(f"field {f} absent")
                continue
            if got.get("s") != tv["s"] or not num_eq(got.get("v"), tv["v"]):
                bad.append(f"{f} = ({got.get('s')}, {got.get('v')}) != ({tv['s']}, {tv['v']})")
        rep.check(f"record {rid}", not bad, "; ".join(bad))
    itypes = {i.get("type") for i in rd.get("issues", [])}
    missing = sorted(set(exp["issue_types"]) - itypes)
    rep.check("data-issue types present", not missing, f"missing: {missing}")
    unexpected = sorted(set(exp["issues_absent"]) & itypes)
    rep.check("no unexpected data-issue types", not unexpected, f"unexpected: {unexpected}")
    pii_fields = [f for r in rd.get("records", []) for f in (r.get("fields") or {}) if f in ("supervisor", "phone")]
    rep.check("pii columns dropped from records", not pii_fields, f"{pii_fields[:5]}")


def check_config(cfg, exp, rep):
    rep.section("04_metrics_config.json")
    ec = exp["config"]
    periods = cfg.get("periods") or []
    rep.check("periods (default first)", [p.get("id") for p in periods] == ec["periods"],
              f"{[p.get('id') for p in periods]} != {ec['periods']}")
    for p in periods:
        pid = p.get("id")
        for key, src in (("rev_field", "rev_field"), ("exp_field", "exp_field"), ("prev_rev_field", "prev_rev_field")):
            if pid in ec[src]:
                rep.check(f"period {pid}.{key}", p.get(key) == ec[src][pid], f"{p.get(key)} != {ec[src][pid]}")
        if pid in ec["comparable_with_prev"]:
            got = (p.get("comparable_with_prev") or {}).get("status")
            rep.check(f"period {pid}.comparable_with_prev", got == ec["comparable_with_prev"][pid],
                      f"{got} != {ec['comparable_with_prev'][pid]}")
    rep.check("valuation field", (cfg.get("valuation") or {}).get("field") == ec["valuation_field"],
              f"{(cfg.get('valuation') or {}).get('field')} != {ec['valuation_field']}")
    rep.check("filters", cfg.get("filters") == ec["filters"], f"{cfg.get('filters')} != {ec['filters']}")
    rep.check("trend comparability", ((cfg.get("trend") or {}).get("comparable") or {}).get("status") == ec["trend_comparable"],
              f"{((cfg.get('trend') or {}).get('comparable') or {}).get('status')}")
    groups = []
    for m in cfg.get("metrics") or []:
        if m.get("card") is True:
            g = m.get("card_group") or m.get("id")
            if g not in groups:
                groups.append(g)
    rep.check("visual cards between 6 and 8", 6 <= len(groups) <= 8, f"{len(groups)} cards: {groups}")


def _cmp_metric(mid, got, e, rep, label):
    bad = []
    for k in ("value", "status", "n_included", "n_scope"):
        if k not in e:
            continue
        if k == "value":
            if not num_eq(got.get("value"), e["value"]):
                bad.append(f"value {got.get('value')!r} != {e['value']!r}")
        elif got.get(k) != e[k]:
            bad.append(f"{k} {got.get(k)!r} != {e[k]!r}")
    if "included_ids" in e and sorted(got.get("included_ids") or []) != sorted(e["included_ids"]):
        bad.append(f"included_ids {sorted(got.get('included_ids') or [])} != {e['included_ids']}")
    if "excluded" in e:
        g = sorted(({"id": x.get("id"), "reason": x.get("reason")} for x in got.get("excluded") or []),
                   key=lambda x: x["id"])
        if g != e["excluded"]:
            bad.append(f"excluded {g} != {e['excluded']}")
    for k in ("num", "den", "num_known_total", "coverage_num_share", "a_sum", "b_sum", "prev_sum", "curr_sum",
              "abs_change", "total", "hhi"):
        if k in e and not num_eq(got.get(k), e[k]):
            bad.append(f"{k} {got.get(k)!r} != {e[k]!r}")
    for k in ("revenue_start", "only_curr", "only_prev", "top_ids"):
        if k in e:
            g = got.get(k)
            if isinstance(g, list) and k != "top_ids":
                g = sorted(g)
            if g != e[k]:
                bad.append(f"{k} {g!r} != {e[k]!r}")
    if "per_field" in e:
        g = [{"field": p.get("field"), "n_known": p.get("n_known"), "n": p.get("n")} for p in got.get("per_field") or []]
        if g != e["per_field"]:
            bad.append(f"per_field {g} != {e['per_field']}")
    if "comparability_status" in e:
        g = (got.get("comparability") or {}).get("status")
        if g != e["comparability_status"]:
            bad.append(f"comparability {g} != {e['comparability_status']}")
    rep.check(f"{label} {mid}", not bad, "; ".join(bad))


def check_metrics(md, exp, rep):
    rep.section("05_metrics.json (default: period 2025, valuation 2024)")
    rep.check("default filters", md.get("default_filters") == {"period": exp["config"]["default_period"]},
              f"{md.get('default_filters')}")
    mets = md.get("metrics") or {}
    for mid, e in exp["metrics"].items():
        if rep.check(f"metric {mid} present", mid in mets, "missing"):
            _cmp_metric(mid, mets[mid], e, rep, "metric")
    pts = [{"period": p.get("period"), "value": p.get("value"), "n_known": p.get("n_known"), "n": p.get("n")}
           for p in (md.get("trend") or {}).get("points") or []]
    rep.check("trend points", len(pts) == len(exp["trend"]) and all(
        a["period"] == b["period"] and num_eq(a["value"], b["value"]) and a["n_known"] == b["n_known"] and a["n"] == b["n"]
        for a, b in zip(pts, exp["trend"])), f"{pts} != {exp['trend']}")


def check_filtered(rd, cfg, md, exp, rep):
    rep.section("filtered views (scripts/metric_engine.py + parity cases of 05)")
    try:
        import metric_engine as me
    except Exception as exc:  # pragma: no cover
        rep.check("import metric_engine", False, str(exc))
        return
    eng = me.Engine(cfg, rd.get("records") or [])
    parity = {}
    for pc in (md or {}).get("parity_cases") or []:
        parity[json.dumps(pc.get("filters"), sort_keys=True, ensure_ascii=False)] = pc.get("metrics") or {}
    for case in exp["filtered_cases"]:
        flt = case["filters"]
        res = eng.compute_all(flt)["metrics"]
        tag = json.dumps(flt, ensure_ascii=False)
        for mid, e in case["metrics"].items():
            _cmp_metric(mid, res.get(mid) or {}, e, rep, f"engine {tag}")
        norm = dict(flt)
        norm.setdefault("period", exp["config"]["default_period"])
        key = json.dumps(norm, sort_keys=True, ensure_ascii=False)
        if key in parity:
            for mid, e in case["metrics"].items():
                _cmp_metric(mid, parity[key].get(mid) or {}, e, rep, f"parity {tag}")


def check_validation(v, exp, rep, label="07_validation.json"):
    rep.section(label)
    bc = v.get("blocking_count")
    blocking = [c.get("id") for c in v.get("checks") or [] if not c.get("passed") and c.get("level") == "blocking"]
    rep.check("no blocking error", bc == exp["validation"]["blocking_count"] and not blocking,
              f"blocking_count = {bc}: {blocking}")
    warns = {c.get("id") for c in v.get("checks") or [] if not c.get("passed") and c.get("level") == "warning"}
    for w in exp["validation"]["expected_warnings"]:
        rep.check(f"expected warning {w}", w in warns, f"warnings raised: {sorted(warns)}")


def check_dashboard(d, rep):
    rep.section("08_dashboard_check.json")
    failed = [c.get("name") for c in d.get("checks") or [] if not c.get("passed")]
    rep.check("browser checks passed", d.get("passed") is True and not failed, f"failed: {failed}")


def check_version(vc, exp, rep):
    rep.section("version_comparison.json (run vs itself)")
    counts = vc.get("counts") or {}
    total = sum(v for v in counts.values() if isinstance(v, int))
    m = (vc.get("matching") or {}).get("counts") or {}
    e = exp["version_comparison_self"]
    rep.check("zero changes", total == e["total_changes"], f"counts = {counts}")
    rep.check("all records matched", m.get("only_prev") == e["only_prev"] and m.get("only_curr") == e["only_curr"],
              f"matching = {m}")


def _texts_of(path):
    if path.endswith(".xlsx"):
        out = []
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if n.endswith(".xml"):
                    out.append(z.read(n).decode("utf-8", "replace"))
        return "\n".join(out)
    with open(path, "rb") as fh:
        return fh.read().decode("utf-8", "replace")


def check_privacy(run, exp, xlsx, rep):
    rep.section("privacy (fictional supervisor names and phones never leave the workbook)")
    values = exp["pii_values"]
    hits = []
    for root, _, files in os.walk(run):
        for fn in files:
            p = os.path.join(root, fn)
            if xlsx and os.path.abspath(p) == os.path.abspath(xlsx):
                continue
            if not fn.endswith((".json", ".md", ".html", ".xlsx", ".txt", ".csv")):
                continue
            txt = _texts_of(p)
            for v in values:
                if v in txt or json.dumps(v)[1:-1] in txt:
                    hits.append(f"{os.path.relpath(p, run)} ⊃ <{'phone' if v[:1].isdigit() else 'name'}>")
    rep.check("no personal value in outputs", not hits, "; ".join(sorted(set(hits))[:10]))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="مقارنة مخرجات التشغيل على المصنف التجريبي بالنتائج المتوقعة")
    ap.add_argument("--run-dir", required=True)
    ap.add_argument("--expected", default=os.path.join(HERE, "synthetic", "expected.json"))
    ap.add_argument("--xlsx", help="المصنف التجريبي (يُستثنى من فحص الخصوصية إن كان داخل مجلد التشغيل)")
    ap.add_argument("--skip-dashboard", action="store_true", help="لا تشترط 08_dashboard_check.json")
    args = ap.parse_args(argv)
    with open(args.expected, encoding="utf-8") as fh:
        exp = json.load(fh)
    rep = Report()
    run = args.run_dir
    wm = load(run, "01_workbook_map.json", rep)
    if wm:
        check_workbook_map(wm, exp, rep)
    rd = load(run, "03_records.json", rep)
    if rd:
        check_records(rd, exp, rep)
    cfg = load(run, "04_metrics_config.json", rep)
    if cfg:
        check_config(cfg, exp, rep)
    md = load(run, "05_metrics.json", rep)
    if md:
        check_metrics(md, exp, rep)
    if rd and cfg:
        check_filtered(rd, cfg, md, exp, rep)
    v = load(run, "07_validation.json", rep)
    if v:
        check_validation(v, exp, rep)
    if not args.skip_dashboard:
        d = load(run, "08_dashboard_check.json", rep)
        if d:
            check_dashboard(d, rep)
    vc = load(run, "version_comparison.json", rep)
    if vc:
        check_version(vc, exp, rep)
    rep.check("analysis.json exists", os.path.isfile(os.path.join(run, "analysis.json")), "missing")
    rep.check("supporting_workbook.xlsx exists", os.path.isfile(os.path.join(run, "supporting_workbook.xlsx")), "missing")
    check_privacy(run, exp, args.xlsx, rep)
    total = rep.passed + len(rep.failed)
    if rep.failed:
        print(f"النتيجة: فشل {len(rep.failed)} من {total} فحصًا")
        return 1
    print(f"النتيجة: نجحت جميع الفحوص ({total})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
