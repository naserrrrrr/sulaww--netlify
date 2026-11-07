#!/usr/bin/env python3
"""Unit tests for scripts/metric_engine.py (data-contract.md §5 semantics).

Pure Python, synthetic in-memory records only (no sample data).
Run:  python3 tests/test_metric_engine.py      (or python3 -m pytest tests/test_metric_engine.py)
"""
import copy
import math
import os
import random
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "scripts"))

import metric_engine as me  # noqa: E402


# ---------------------------------------------------------------------------
# Fixture helpers
# ---------------------------------------------------------------------------
def tv(v=None, s=None):
    if s is None:
        s = "blank" if v is None else ("zero_recorded" if v == 0 else "value")
    return {"v": v if s in ("value", "derived", "zero_recorded", "zero_confirmed") else None, "s": s,
            "raw": None if v is None else str(v), "ref": "T!A1", "note": None}


def rec(rid, rtype="asset", scope="in_scope", use="investment", inc="rent", own=("owned",), atype="عمارة",
        loc="جنوب", mfo=False, **fields):
    return {"id": rid, "name": rid, "record_type": rtype, "scope": scope,
            "dims": {"use": use, "income_nature": inc, "ownership": list(own), "waqf": "unknown",
                     "asset_type": atype, "location": loc, "managed_for_others": mfo},
            "fields": {k.replace("_", ".", 1): v for k, v in fields.items()}}


def records():
    return [
        # A1: investment rent, rev 100 / val 1000 in 2025; rev 80 in 2024; exp 10
        rec("A1", rev_2025=tv(100), rev_2024=tv(80), val_2024=tv(1000), exp_2025=tv(10)),
        # A2: investment rent, recorded zero revenue (included in ratio), val 500
        rec("A2", rev_2025=tv(0), rev_2024=tv(0), val_2024=tv(500), exp_2025=tv(5)),
        # A3: investment rent, revenue blank (unknown) -> num_unknown; val 2000
        rec("A3", rev_2025=tv(None), rev_2024=tv(50), val_2024=tv(2000), loc="شمال"),
        # A4: investment rent, revenue 300, valuation zero_as_missing -> den_unknown
        rec("A4", rev_2025=tv(300), rev_2024=tv(None), val_2024=tv(None, "zero_as_missing"), atype="أرض"),
        # A5: service asset, no revenue fields at all, val 700
        rec("A5", use="service", inc="none", val_2024=tv(700), loc="شمال"),
        # A6: managed for others (excluded from property revenue)
        rec("A6", own=("managed_for_others",), mfo=True, rev_2025=tv(999), val_2024=tv(999)),
        # A7: operating activity (school)
        rec("A7", rtype="operating_activity", use="service", inc="operating", own=("usufruct",), rev_2025=tv(400)),
        # A8: isolated (outside base_where)
        rec("A8", scope="isolated", rev_2025=tv(5000), val_2024=tv(5000)),
        # A9: investment, revenue confirmed zero, known negative den -> den_nonpositive
        rec("A9", rev_2025=tv(0, "zero_confirmed"), val_2024={"v": -1, "s": "value", "raw": "-1", "ref": "T!A9"}),
    ]


ASSET = {"in": ["record_type", ["asset", "land_parcel"]]}
NOT_MFO = {"not": {"eq": ["managed_for_others", True]}}


def config(**over):
    cfg = {
        "contract_version": "1.0",
        "periods": [
            {"id": "2025", "label": "2025", "rev_field": "rev.2025", "exp_field": "exp.2025", "prev_id": "2024",
             "prev_rev_field": "rev.2024", "comparable_with_prev": {"status": "unverified", "reason": "x"}},
            {"id": "2024", "label": "2024", "rev_field": "rev.2024", "exp_field": None, "prev_id": None,
             "prev_rev_field": None, "comparable_with_prev": {"status": "unverified", "reason": "none"}},
        ],
        "valuation": {"field": "val.2024", "year": "2024"},
        "base_where": {"all": [{"eq": ["scope", "in_scope"]}]},
        "metrics": [],
        "breakdowns": [],
        "trend": {"enabled": True, "label": "t", "fields": [{"period": "2024", "field": "rev.2024"},
                                                            {"period": "2025", "field": "rev.2025"}],
                  "where": {"all": [{"eq": ["income_nature", "rent"]}, NOT_MFO]},
                  "comparable": {"status": "unverified", "reason": "u"}},
        "filters": ["asset_type", "location", "use"],
        "parity_max_values_per_dim": 3,
    }
    cfg.update(over)
    return cfg


def engine(recs=None, **over):
    return me.Engine(config(**over), records() if recs is None else recs)


def close(a, b, tol=1e-12):
    return a is not None and b is not None and math.isclose(a, b, rel_tol=tol, abs_tol=tol)


# ---------------------------------------------------------------------------
# Where
# ---------------------------------------------------------------------------
def test_where_ops():
    r = records()[0]
    ev = me.eval_where
    assert ev({"eq": ["use", "investment"]}, r)
    assert ev({"eq": ["id", "A1"]}, r)
    assert not ev({"eq": ["missing_dim", "x"]}, r)
    assert ev({"eq": ["missing_dim", None]}, r)          # missing attr => null
    assert ev({"in": ["record_type", ["asset", "land_parcel"]]}, r)
    assert ev({"has": ["ownership", "owned"]}, r)
    assert not ev({"has": ["ownership", "usufruct"]}, r)
    assert not ev({"eq": ["ownership", "owned"]}, r)      # eq never matches a list
    assert ev({"not": {"eq": ["managed_for_others", True]}}, r)
    assert not ev({"eq": ["managed_for_others", 0]}, r)   # strict: False != 0
    assert ev({"known": "rev.2025"}, r)
    assert not ev({"known": "rev.missing"}, r)
    assert ev({"status": ["rev.2025", ["value"]]}, r)
    assert not ev({"status": ["rev.missing", ["blank"]]}, r)  # absent field has status null
    assert ev({"gt": ["rev.2025", 0]}, r)
    assert not ev({"gt": ["rev.2025", 100]}, r)
    assert ev({"all": []}, r) and not ev({"any": []}, r)
    assert ev({"any": [{"eq": ["use", "x"]}, {"eq": ["use", "investment"]}]}, r)
    r2 = records()[1]
    assert not ev({"gt": ["rev.2025", 0]}, r2)
    assert ev({"status": ["rev.2025", ["zero_recorded", "zero_confirmed"]]}, r2)


def test_placeholders_in_where():
    eng = engine()
    spec = {"id": "zr", "kind": "sum", "field": "{VAL}", "unit": "SAR",
            "where": {"all": [{"eq": ["use", "investment"]}, ASSET, NOT_MFO,
                              {"status": ["{REV}", ["zero_recorded", "zero_confirmed"]]}]}}
    res = eng.compute(spec, {"period": "2025"})
    # A2 (zero_recorded, val 500) and A9 (zero_confirmed, val -1)
    assert res["included_ids"] == ["A2", "A9"], res
    assert res["value"] == 499 and res["status"] == "complete"
    res24 = eng.compute(spec, {"period": "2024"})
    assert res24["included_ids"] == ["A2"] and res24["value"] == 500


# ---------------------------------------------------------------------------
# Kinds
# ---------------------------------------------------------------------------
def test_count():
    eng = engine()
    spec = {"id": "c", "kind": "count", "where": ASSET, "unit": "asset"}
    res = eng.compute(spec)
    assert res["value"] == 7 and res["status"] == "complete" and res["n_included"] == 7  # A1-A6, A9
    res = eng.compute({"id": "c", "kind": "count", "where": {"eq": ["record_type", "unit"]}, "unit": "count"})
    assert res["value"] == 0 and res["status"] == "empty" and res["n_included"] == 0


def test_sum_statuses():
    eng = engine()
    spec = {"id": "pr", "kind": "sum", "field": "{REV}", "where": {"all": [{"eq": ["income_nature", "rent"]}, NOT_MFO]},
            "unit": "SAR"}
    res = eng.compute(spec)
    # R = A1,A2,A3,A4,A9 ; K = A1(100), A2(0), A4(300), A9(0) ; U = A3(blank)
    assert res["value"] == 400, res
    assert res["status"] == "partial"
    assert res["n_scope"] == 5 and res["n_included"] == 4
    assert res["excluded"] == [{"id": "A3", "reason": "blank"}]
    # absent field reason
    res = eng.compute({"id": "x", "kind": "sum", "field": "rev.2025", "where": {"eq": ["id", "A5"]}, "unit": "SAR"})
    assert res["status"] == "not_computable" and res["value"] is None
    assert res["excluded"] == [{"id": "A5", "reason": "absent"}]
    # empty
    res = eng.compute({"id": "x", "kind": "sum", "field": "rev.2025", "where": {"eq": ["id", "ZZ"]}, "unit": "SAR"})
    assert res["status"] == "empty" and res["value"] is None and res["n_scope"] == 0
    # complete + zero_as_missing never summed
    res = eng.compute({"id": "v", "kind": "sum", "field": "{VAL}", "where": {"all": [ASSET, NOT_MFO]}, "unit": "SAR"})
    assert res["value"] == 1000 + 500 + 2000 + 700 - 1 and res["status"] == "partial"
    assert {"id": "A4", "reason": "zero_as_missing"} in res["excluded"]


def test_ratio_zero_included_and_exclusion_reasons():
    eng = engine()
    spec = {"id": "gy", "kind": "ratio", "num": "{REV}", "den": "{VAL}", "scale": 100, "unit": "%",
            "where": {"all": [{"eq": ["use", "investment"]}, ASSET, {"not": {"eq": ["income_nature", "operating"]}}, NOT_MFO]}}
    res = eng.compute(spec)
    # R = A1,A2,A3,A4,A9 ; I = A1 (100/1000), A2 (0/500) ; A3 num_unknown, A4 den_unknown, A9 den_nonpositive
    assert res["included_ids"] == ["A1", "A2"]
    assert res["num"] == 100 and res["den"] == 1500
    assert close(res["value"], 100 / 1500 * 100)
    assert res["status"] == "partial"
    assert res["excluded"] == [{"id": "A3", "reason": "num_unknown"}, {"id": "A4", "reason": "den_unknown"},
                               {"id": "A9", "reason": "den_nonpositive"}]
    # num_known_total includes A4 (300) and A9 (0): 100 + 0 + 300 + 0
    assert res["num_known_total"] == 400
    assert close(res["coverage_num_share"], 100 / 400)
    # complete
    res = eng.compute(dict(spec, where={"eq": ["id", "A1"]}))
    assert res["status"] == "complete" and close(res["value"], 10.0)
    # only zero revenue => value 0 (zero included, not excluded)
    res = eng.compute(dict(spec, where={"eq": ["id", "A2"]}))
    assert res["value"] == 0 and res["status"] == "complete"
    # not_computable when I is empty
    res = eng.compute(dict(spec, where={"eq": ["id", "A3"]}))
    assert res["status"] == "not_computable" and res["value"] is None and res["n_included"] == 0
    # coverage_num_share null when num_known_total == 0
    res = eng.compute(dict(spec, where={"eq": ["id", "A2"]}))
    assert res["coverage_num_share"] is None


def test_diff_and_missing_expense():
    eng = engine()
    spec = {"id": "noi", "kind": "diff", "a": "{REV}", "b": "{EXP}", "unit": "SAR",
            "where": {"all": [{"eq": ["income_nature", "rent"]}, NOT_MFO]},
            "unavailable_reason": "لا مصروفات لهذه الفترة"}
    res = eng.compute(spec, {"period": "2025"})
    # I = A1 (100-10), A2 (0-5); A3 a_unknown; A4, A9 b_unknown
    assert res["value"] == 85 and res["status"] == "partial" and res["n_included"] == 2
    assert res["excluded"][0] == {"id": "A3", "reason": "a_unknown"}
    # exp placeholder null in 2024 -> not_computable with unavailable_reason
    res = eng.compute(spec, {"period": "2024"})
    assert res["status"] == "not_computable" and res["reason"] == "لا مصروفات لهذه الفترة"
    assert res["n_included"] == 0 and res["value"] is None
    # field defined but no record of R has it at all
    res = eng.compute(dict(spec, b="exp.none"), {"period": "2025"})
    assert res["status"] == "not_computable" and res["reason"] == me.NO_EXPENSE_REASON


def test_growth():
    eng = engine()
    spec = {"id": "g", "kind": "growth", "unit": "%", "where": {"all": [{"eq": ["income_nature", "rent"]}, NOT_MFO]}}
    res = eng.compute(spec, {"period": "2025"})
    # R=A1,A2,A3,A4,A9 ; I = A1 (80->100), A2 (0->0); only_curr = A4, A9 ; only_prev = A3
    assert res["included_ids"] == ["A1", "A2"]
    assert res["prev_sum"] == 80 and res["curr_sum"] == 100 and res["abs_change"] == 20
    assert close(res["value"], 25.0) and res["revenue_start"] is False
    assert res["only_curr"] == ["A4", "A9"] and res["only_prev"] == ["A3"]
    assert res["status"] == "partial"
    assert res["comparability"]["status"] == "unverified"
    # revenue start: prev 0, curr > 0
    recs = [rec("B1", rev_2025=tv(50), rev_2024=tv(0))]
    r2 = me.Engine(config(), recs).compute(spec, {"period": "2025"})
    assert r2["value"] is None and r2["revenue_start"] is True and r2["abs_change"] == 50
    assert r2["status"] == "complete"
    # not_comparable
    cfg = config()
    cfg["periods"][0]["comparable_with_prev"] = {"status": "not_comparable", "reason": "partial"}
    r3 = me.Engine(cfg, records()).compute(spec, {"period": "2025"})
    assert r3["status"] == "not_comparable" and r3["value"] is None and r3["n_included"] == 0
    # no previous period -> not_computable
    r4 = eng.compute(spec, {"period": "2024"})
    assert r4["status"] == "not_computable" and r4["value"] is None


def test_concentration_tiebreak_and_hhi():
    recs = [rec("C3", val_2024=tv(100)), rec("C1", val_2024=tv(100)), rec("C2", val_2024=tv(300)),
            rec("C4", val_2024=tv(None)), rec("C5", val_2024=tv(0, "zero_recorded")), rec("C6", val_2024=tv(100))]
    eng = me.Engine(config(), recs)
    spec = {"id": "vc", "kind": "concentration", "field": "{VAL}", "top": 3, "unit": "%", "where": ASSET}
    res = eng.compute(spec)
    assert res["top_ids"] == ["C2", "C1", "C3"]           # tie at 100 broken by id asc
    assert res["total"] == 600 and res["n_included"] == 4
    assert close(res["value"], 500 / 600 * 100)
    shares = [300 / 600, 100 / 600, 100 / 600, 100 / 600]
    assert close(res["hhi"], sum(x * x for x in shares) * 10000)
    assert res["status"] == "partial"                      # C4 unknown
    assert {"id": "C5", "reason": "nonpositive"} in res["excluded"]
    res = eng.compute(dict(spec, where={"eq": ["id", "C4"]}))
    assert res["status"] == "not_computable" and res["value"] is None
    res = eng.compute(dict(spec, where={"in": ["id", ["C1", "C2", "C5"]]}))
    assert res["status"] == "complete"                     # known zero does not make it partial


def test_coverage():
    eng = engine()
    spec = {"id": "cov", "kind": "coverage", "fields": ["{REV}", "{VAL}"], "unit": "%", "where": ASSET}
    res = eng.compute(spec)
    # R = A1..A6, A9 (7). REV known: A1,A2,A4,A6,A9 = 5 ; VAL known: A1,A2,A3,A5,A6,A9 = 6
    assert res["per_field"] == [{"field": "rev.2025", "n_known": 5, "n": 7}, {"field": "val.2024", "n_known": 6, "n": 7}]
    assert res["value"] == min(5 / 7, 6 / 7) * 100 and res["status"] == "complete"
    assert res["n_included"] == 7
    res = eng.compute(dict(spec, where={"eq": ["id", "nope"]}))
    assert res["status"] == "empty" and res["value"] is None


def test_unresolved_placeholder():
    eng = engine(valuation={"field": None, "year": None})
    spec = {"id": "rv", "kind": "sum", "field": "{VAL}", "unit": "SAR", "where": ASSET,
            "unavailable_reason": "لا يوجد تقييم"}
    res = eng.compute(spec)
    assert res["status"] == "not_computable" and res["reason"] == "لا يوجد تقييم" and res["value"] is None
    assert res["n_scope"] == 7  # |R| (base_where AND where); only the field placeholder is unresolved
    assert res["n_included"] == 0 and res["excluded"] == []
    # a placeholder inside where: n_scope falls back to base_where AND filters (8 in-scope records)
    res = eng.compute(dict(spec, field="rev.2025", where={"known": "{VAL}"}))
    assert res["status"] == "not_computable" and res["n_scope"] == 8
    res = eng.compute(dict(spec, unavailable_reason=None))
    assert res["reason"] == me.DEFAULT_UNAVAILABLE


def test_breakdown_list_dim_and_none():
    recs = records() + [rec("A10", own=(), atype=None, rev_2025=tv(20), val_2024=tv(10))]
    eng = me.Engine(config(), recs)
    spec = {"id": "v_own", "label": "x", "field": "{VAL}", "by": "ownership", "where": ASSET, "unit": "SAR"}
    res = eng.compute_breakdown(spec)
    keys = [r["key"] for r in res["rows"]]
    assert res["overlap"] is True and res["note"] == me.OVERLAP_NOTE
    owned = next(r for r in res["rows"] if r["key"] == "owned")
    # owned: A1 1000, A2 500, A3 2000, A5 700, A9 -1 known ; A4 unknown
    assert owned["value"] == 4199 and owned["n_known"] == 5 and owned["n_unknown"] == 1
    none = next(r for r in res["rows"] if r["key"] == me.NONE_KEY)
    assert none["value"] == 10
    assert res["total"] == 4199 + 999 + 10
    assert close(owned["share"], 4199 / res["total"])
    assert keys[0] == "owned"
    spec2 = dict(spec, by="asset_type", id="v_type")
    res2 = me.Engine(config(), recs).compute_breakdown(spec2)
    assert any(r["key"] == me.NONE_KEY for r in res2["rows"])
    land = next(r for r in res2["rows"] if r["key"] == "أرض")
    assert land["value"] is None and land["n_unknown"] == 1     # only unknown values -> null, last
    assert res2["rows"][-1]["key"] == "أرض"
    assert res2["status"] == "partial"


def test_trend():
    eng = engine()
    t = eng.compute_trend()
    # R = rent, not mfo, in scope: A1,A2,A3,A4,A9
    p24, p25 = t["points"]
    assert p24 == {"period": "2024", "field": "rev.2024", "value": 130, "n_known": 3, "n": 5}
    assert p25["value"] == 400 and p25["n_known"] == 4
    assert t["status"] == "partial" and t["not_comparable"] is False
    cfg = config()
    cfg["trend"]["comparable"] = {"status": "not_comparable", "reason": "partial year"}
    t2 = me.Engine(cfg, records()).compute_trend()
    assert t2["not_comparable"] is True and t2["points"][0]["value"] == 130   # values kept, flagged
    t3 = eng.compute_trend({"location": "nowhere"})
    assert t3["status"] == "empty"


def test_filters_and_compute_all():
    eng = engine()
    f = eng.normalize_filters({"use": "", "location": None, "asset_type": "عمارة"})
    assert f == {"period": "2025", "asset_type": "عمارة"}
    ids = [r["id"] for r in eng.filter_records({"location": "شمال"})]
    assert ids == ["A3", "A5"]
    ids = [r["id"] for r in eng.filter_records({"ownership": "usufruct"})]   # list dim -> has
    assert ids == ["A7"]
    allres = eng.compute_all({"period": "2024", "use": "investment"})
    assert allres["filters"] == {"period": "2024", "use": "investment"}
    assert "A5" not in allres["records"] and "A8" not in allres["records"]
    try:
        eng.normalize_filters({"period": "1999"})
        raise AssertionError("unknown period must raise")
    except ValueError:
        pass


def test_text_by_period():
    eng = engine()
    spec = {"id": "pr", "kind": "sum", "field": "{REV}", "unit": "SAR", "label": "إيراد 2025",
            "text_by_period": {"2024": {"label": "إيراد 2024"}}}
    assert eng.compute(spec, {"period": "2025"})["label"] == "إيراد 2025"
    assert eng.compute(spec, {"period": "2024"})["label"] == "إيراد 2024"


def test_determinism_independent_of_input_order():
    specs = [
        {"id": "s", "kind": "sum", "field": "{VAL}", "unit": "SAR", "where": ASSET},
        {"id": "r", "kind": "ratio", "num": "{REV}", "den": "{VAL}", "scale": 100, "unit": "%", "where": ASSET},
        {"id": "c", "kind": "concentration", "field": "{REV}", "top": 2, "unit": "%"},
        {"id": "g", "kind": "growth", "unit": "%"},
    ]
    base = me.Engine(config(metrics=specs), records()).compute_all()
    for seed in range(5):
        recs = copy.deepcopy(records())
        random.Random(seed).shuffle(recs)
        other = me.Engine(config(metrics=specs), recs).compute_all()
        for mid in base["metrics"]:
            a, b = base["metrics"][mid], other["metrics"][mid]
            assert a["value"] == b["value"] and a["included_ids"] == b["included_ids"] and a["excluded"] == b["excluded"]
        assert base["records"] == other["records"]


def test_check_records():
    recs = records()
    recs.append(copy.deepcopy(recs[0]))
    recs[-1]["fields"]["rev.2025"] = {"v": None, "s": "value"}
    errs = me.check_records(recs, {"rev.2025"})
    assert any("مكرر" in e for e in errs) and any("rev.2025" in e for e in errs)
    assert me.check_records(records(), {"rev.2025", "val.2024"}) == []


def _run_all():
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_") and callable(f)]
    failed = 0
    for name, fn in tests:
        try:
            fn()
            print(f"PASS {name}")
        except Exception as exc:  # noqa: BLE001
            failed += 1
            print(f"FAIL {name}: {type(exc).__name__}: {exc}")
    print(f"{len(tests) - failed}/{len(tests)} passed")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(_run_all())
