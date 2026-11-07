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

sys.dont_write_bytecode = True
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
    assert res["empty_cause"] is None and res["note"] is None
    # nothing qualifies in the whole scope: a count is a real zero (empty_as_zero defaults to true for count)
    res = eng.compute({"id": "c", "kind": "count", "where": {"eq": ["record_type", "unit"]}, "unit": "count"})
    assert res["value"] == 0 and res["status"] == "complete" and res["n_included"] == 0
    assert res["empty_cause"] == "scope" and res["note"] == me.EMPTY_SCOPE_REASON and res["reason"] is None
    # explicit empty_as_zero false keeps the empty status
    res = eng.compute({"id": "c", "kind": "count", "where": {"eq": ["record_type", "unit"]}, "unit": "count",
                       "empty_as_zero": False, "empty_text": "لا وحدات"})
    assert res["value"] == 0 and res["status"] == "empty" and res["reason"] == "لا وحدات"
    # a dim filter that empties a non-empty scope stays "empty" with cause filters
    res = eng.compute(spec, {"location": "nowhere"})
    assert res["value"] == 0 and res["status"] == "empty" and res["empty_cause"] == "filters"
    assert res["reason"] == me.EMPTY_FILTERS_REASON


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
    assert res["revenue_start"] is False and res["revenue_start_ids"] == [] and res["start_change"] is None
    # revenue start only: prev 0 (recorded), curr > 0 -> no %, «بدء إيراد» with its riyal change
    recs = [rec("B1", rev_2025=tv(50), rev_2024=tv(0))]
    r2 = me.Engine(config(), recs).compute(spec, {"period": "2025"})
    assert r2["value"] is None and r2["revenue_start"] is True and r2["revenue_start_ids"] == ["B1"]
    assert r2["start_change"] == 50 and r2["abs_change"] is None and r2["n_included"] == 0
    assert r2["status"] == "not_computable" and r2["reason"] == me.REVENUE_START_ONLY_REASON
    assert r2["excluded"] == [{"id": "B1", "reason": "revenue_start"}]
    # B9: a revenue start next to an existing asset stays outside the % (A 100->100, B 0->1000 is 0 %, not 1000 %)
    recs = [rec("A", rev_2025=tv(100), rev_2024=tv(100)), rec("B", rev_2025=tv(1000), rev_2024=tv(0, "zero_confirmed"))]
    r5 = me.Engine(config(), recs).compute(spec, {"period": "2025"})
    assert r5["value"] == 0 and r5["included_ids"] == ["A"] and r5["status"] == "partial"
    assert r5["revenue_start"] is True and r5["revenue_start_ids"] == ["B"] and r5["start_change"] == 1000
    assert r5["prev_sum"] == 100 and r5["curr_sum"] == 100 and r5["abs_change"] == 0
    # B19: prev_sum <= 0 without a revenue start -> not_computable with a reason (never complete + null)
    for recs in ([rec("N", rev_2025=tv(100), rev_2024=tv(-50))], [rec("Z", rev_2025=tv(0), rev_2024=tv(0))]):
        r6 = me.Engine(config(), recs).compute(spec, {"period": "2025"})
        assert r6["status"] == "not_computable" and r6["value"] is None and r6["reason"] == me.PREV_NONPOSITIVE_REASON
        assert r6["revenue_start"] is False and r6["n_included"] == 1 and r6["prev_sum"] is not None
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
    h = 0
    for v in (300, 100, 100, 100):      # ranked order, s*s (bit-identical with the JS twin)
        h = h + (v / 600) * (v / 600)
    assert res["hhi"] == h * 10000
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
    assert res["per_field"] == [{"field": "rev.2025", "n_known": 5, "n": 7, "n_not_applicable": 0},
                                {"field": "val.2024", "n_known": 6, "n": 7, "n_not_applicable": 0}]
    assert res["value"] == min(5 / 7, 6 / 7) * 100 and res["status"] == "complete"
    assert res["n_included"] == 7
    res = eng.compute(dict(spec, where={"eq": ["id", "nope"]}))
    assert res["status"] == "empty" and res["value"] is None
    # B13: «لا ينطبق» leaves that field's denominator (admin HQ without revenue is not missing data)
    recs = records()
    recs[4]["fields"]["rev.2025"] = tv(None, "not_applicable")          # A5 service asset
    res = me.Engine(config(), recs).compute(spec)
    assert res["per_field"][0] == {"field": "rev.2025", "n_known": 5, "n": 6, "n_not_applicable": 1}
    assert res["value"] == min(5 / 6, 6 / 7) * 100
    # every record not applicable for the only field -> not_computable with a reason
    res = me.Engine(config(), recs).compute(dict(spec, fields=["rev.2025"], where={"eq": ["id", "A5"]}))
    assert res["status"] == "not_computable" and res["value"] is None and res["reason"] == me.NO_COVERAGE_FIELD_REASON


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
    assert t3["status"] == "empty" and t3["empty_cause"] == "filters"
    assert t3["same_records"]["n"] == 0 and t3["n_known_varies"] is False


def test_trend_same_records():
    # A18: 2024 known A1 80, A2 0, A3 50 ; 2025 known A1 100, A2 0, A4 300, A9 0 -> the record set changes
    t = engine().compute_trend()
    assert t["n_known_varies"] is True
    sr = t["same_records"]
    assert sr["n"] == 2 and sr["ids"] == ["A1", "A2"]
    assert [p["value"] for p in sr["points"]] == [80, 100] and [p["n_known"] for p in sr["points"]] == [2, 2]
    # same records in every period -> n_known does not vary
    recs = [rec("S1", rev_2024=tv(10), rev_2025=tv(20)), rec("S2", rev_2024=tv(5), rev_2025=tv(0))]
    t2 = me.Engine(config(), recs).compute_trend()
    assert t2["n_known_varies"] is False and t2["same_records"]["ids"] == ["S1", "S2"]
    assert [p["value"] for p in t2["same_records"]["points"]] == [15, 20]
    # equal counts but different records still vary (membership, not only the count)
    recs = [rec("X1", rev_2024=tv(10)), rec("X2", rev_2025=tv(10))]
    t3 = me.Engine(config(), recs).compute_trend()
    assert [p["n_known"] for p in t3["points"]] == [1, 1] and t3["n_known_varies"] is True
    assert t3["same_records"]["n"] == 0 and t3["same_records"]["points"][0]["value"] is None


def test_valid_periods():
    eng = engine()
    spec = {"id": "gy", "kind": "ratio", "num": "{REV}", "den": "{VAL}", "scale": 100, "unit": "%", "where": ASSET,
            "valid_periods": ["2025"], "invalid_period_reason": "تقييم 2024 لا يقابل إيراد 2024 هنا"}
    ok = eng.compute(spec, {"period": "2025"})
    assert ok["status"] == "partial" and ok["value"] is not None
    bad = eng.compute(spec, {"period": "2024"})
    assert bad["status"] == "not_comparable" and bad["value"] is None and bad["reason"] == "تقييم 2024 لا يقابل إيراد 2024 هنا"
    assert bad["n_included"] == 0 and bad["included_ids"] == [] and bad["n_scope"] == 7
    # default reason; and the period check wins over an unresolved placeholder
    eng_nv = me.Engine(config(valuation={"field": None}), records())
    bad = eng_nv.compute(dict(spec, invalid_period_reason=None), {"period": "2024"})
    assert bad["status"] == "not_comparable" and bad["reason"] == me.INVALID_PERIOD_REASON and bad["n_scope"] == 7
    # placeholder unresolved inside where: n_scope falls back to base_where AND filters (8 in-scope records)
    bad = eng_nv.compute(dict(spec, where={"known": "{VAL}"}), {"period": "2024"})
    assert bad["status"] == "not_comparable" and bad["n_scope"] == 8
    assert eng_nv.compute(dict(spec, where={"known": "{VAL}"}), {"period": "2025"})["status"] == "not_computable"
    # no valid_periods key -> every period is valid
    assert eng.compute(dict(spec, valid_periods=None), {"period": "2024"})["status"] != "not_comparable"


def test_empty_cause_and_empty_as_zero():
    eng = engine()
    zr = {"id": "zr", "kind": "sum", "field": "{VAL}", "unit": "SAR", "empty_as_zero": True,
          "empty_text": "لا توجد أصول استثمارية بإيراد صفري مسجل",
          "where": {"all": [ASSET, {"eq": ["use", "admin"]}, {"status": ["{REV}", ["zero_recorded"]]}]}}
    res = eng.compute(zr)
    assert res["status"] == "complete" and res["value"] == 0 and res["n_included"] == 0
    assert res["empty_cause"] == "scope" and res["note"] == zr["empty_text"] and res["reason"] is None
    # sum without empty_as_zero: empty with the metric's own text
    res = eng.compute(dict(zr, empty_as_zero=False))
    assert res["status"] == "empty" and res["value"] is None and res["reason"] == zr["empty_text"]
    # a filter that empties a non-empty scope: never a zero
    op = {"id": "op", "kind": "sum", "field": "{REV}", "unit": "SAR", "empty_as_zero": True,
          "where": {"eq": ["income_nature", "operating"]}}
    res = eng.compute(op, {"location": "شمال"})
    assert res["status"] == "empty" and res["value"] is None and res["empty_cause"] == "filters"
    assert res["reason"] == me.EMPTY_FILTERS_REASON
    # ratio: empty_as_zero ignored, cause and text reported
    ratio = {"id": "r", "kind": "ratio", "num": "{REV}", "den": "{VAL}", "unit": "%", "empty_as_zero": True,
             "where": {"eq": ["record_type", "unit"]}}
    res = eng.compute(ratio)
    assert res["status"] == "empty" and res["empty_cause"] == "scope" and res["reason"] == me.EMPTY_SCOPE_REASON
    # breakdown
    bd = {"id": "b", "label": "b", "field": "{VAL}", "by": "use", "where": {"eq": ["record_type", "unit"]},
          "empty_text": "لا وحدات"}
    res = eng.compute_breakdown(bd)
    assert res["status"] == "empty" and res["empty_cause"] == "scope" and res["reason"] == "لا وحدات"
    res = eng.compute_breakdown(dict(bd, where=ASSET), {"location": "nowhere"})
    assert res["status"] == "empty" and res["empty_cause"] == "filters"


# ---------------------------------------------------------------------------
# compute_metrics.py config builder (init-config rules that feed the engine)
# ---------------------------------------------------------------------------
def _cm():
    import compute_metrics as cm  # imports normalize (openpyxl/jsonschema) — only for these tests
    return cm


def _rec_doc(records_list, meta):
    return {"contract_version": "1.0", "source": {"name": "t.xlsx"}, "field_meta": meta, "records": records_list}


def _meta_a():
    return {
        "rev.2025": {"label": "إيراد 2025 بدون ضريبة", "group": "revenue", "period": "2025", "vat": "excl",
                     "basis": "recorded", "completeness": "full_year"},
        "rev.2024": {"label": "إيراد 2024 شامل الضريبة", "group": "revenue", "period": "2024", "vat": "incl",
                     "basis": "recorded", "completeness": "full_year"},
        "rev.2022": {"label": "إيراد 2022", "group": "revenue", "period": "2022", "vat": "excl",
                     "basis": "recorded", "completeness": "full_year"},
        "rev.2023": {"label": "إيراد 2023 (ورقة معزولة)", "group": "revenue", "period": "2023", "vat": "excl"},
        "val.2024": {"label": "تقييم 2024", "group": "valuation", "year": "2024", "component": "total"},
    }


def _src(rid, sheet, rtype="asset", scope="in_scope", **kw):
    r = rec(rid, rtype=rtype, scope=scope, **kw)
    r["source"] = {"sheet": sheet, "row": 1}
    return r


def _doc_a():
    return _rec_doc([
        _src("T1", "السجل", rev_2025=tv(100), rev_2024=tv(110), rev_2022=tv(90), val_2024=tv(1000)),
        _src("T1.1", "السجل", rtype="unit", rev_2025=tv(60), rev_2024=tv(60)),   # unit of T1: never added again
        _src("T2", "السجل", loc="شمال", rev_2025=tv(50), rev_2024=tv(40), rev_2022=tv(40), val_2024=tv(500)),
        _src("I1", "التخصيص", scope="isolated", rev_2023=tv(999)),
    ], _meta_a())


def test_init_config_scope_rules():
    cm = _cm()
    doc = _doc_a()
    cfg = cm.init_config(doc)
    # C22: rev.2023 is known only on an isolated record -> no 2023 period
    assert [p["id"] for p in cfg["periods"]] == ["2025", "2024", "2022"]
    assert not cm.semantic_errors(cfg, doc["records"]) and not cm.comparability_errors(cfg, doc["field_meta"])
    M = {m["id"]: m for m in cfg["metrics"]}
    # A2: unit rows never add revenue to their asset
    res = me.Engine(cfg, doc["records"]).compute_all({"period": "2025"})["metrics"]
    assert res["property_revenue"]["value"] == 150 and "T1.1" not in res["property_revenue"]["included_ids"]
    assert res["revenue_concentration_top3"]["included_ids"] == ["T1", "T2"]
    # C1: source text names only the sheets of the records the metric reads
    assert "التخصيص" not in M["property_revenue"]["source_text"] and "«السجل»" in M["property_revenue"]["source_text"]
    # C2: revenue/valuation ratios only for the valuation year and the year after
    assert M["gross_yield_investment"]["valid_periods"] == ["2025", "2024"]
    r22 = me.Engine(cfg, doc["records"]).compute(M["gross_yield_investment"], {"period": "2022"})
    assert r22["status"] == "not_comparable" and "2024" in r22["reason"]
    assert any(w["code"] == "yield_periods_limited" for w in cfg["init_warnings"])
    # B1: 2025 (excl) vs 2024 (incl): both known, different -> not_comparable; 2024 vs 2022 excl/incl too
    assert cfg["periods"][0]["comparable_with_prev"]["status"] == "not_comparable"
    assert cfg["trend"]["comparable"]["status"] == "not_comparable"
    g = me.Engine(cfg, doc["records"]).compute(M["revenue_growth"], {"period": "2025"})
    assert g["status"] == "not_comparable" and g["value"] is None
    # B3: 2024 has only a VAT-inclusive field -> labels say so, and 04 carries a warning
    t24 = M["property_revenue"]["text_by_period"]["2024"]
    assert "شامل الضريبة" in t24["label"]
    assert "شامل الضريبة" in M["gross_yield_investment"]["text_by_period"]["2024"]["label"]
    assert "شامل الضريبة" not in M["gross_yield_investment"]["label"]
    assert [w["period"] for w in cfg["init_warnings"] if w["code"] == "revenue_incl_vat"] == ["2024"]
    # C3: standard empty texts / empty_as_zero
    assert M["zero_rev_investment_value"]["empty_as_zero"] is True
    assert M["zero_rev_investment_value"]["empty_text"] == "لا توجد أصول استثمارية بإيراد صفري مسجل"
    assert M["activity_revenue"]["empty_text"] == "لا توجد أنشطة تشغيلية ضمن النطاق"
    z = me.Engine(cfg, doc["records"]).compute(M["zero_rev_investment_value"], {"period": "2025"})
    assert z["status"] == "complete" and z["value"] == 0 and z["empty_cause"] == "scope"
    # C19: location has 2 values on assets (جنوب, شمال); use has one (investment) -> no use breakdown
    bids = [b["id"] for b in cfg["breakdowns"]]
    assert "value_by_location" in bids and "value_by_use" not in bids and "revenue_by_use" not in bids
    # C22: no in-scope record -> fields of isolated records never become periods or the valuation
    iso = _doc_a()
    for r in iso["records"]:
        r["scope"] = "isolated"
    cfg2 = cm.init_config(iso)
    assert cfg2["periods"] == [] and cfg2["valuation"]["field"] is None
    assert [w["code"] for w in cfg2["init_warnings"]] == ["no_in_scope_records"]
    schema = os.path.join(cm.SKILL_DIR, "schemas", "metrics_config.schema.json")
    assert cm.schema_errors(cfg, schema) == [] and cm.schema_errors(cfg2, schema) == []


def test_comparability_rules():
    cm = _cm()
    c = cm.comparability
    full = {"vat": "excl", "basis": "recorded", "completeness": "full_year"}
    assert c(full, dict(full), "2025", "2024")["status"] == "confirmed"
    assert c(full, dict(full, vat="incl"), "2025", "2024")["status"] == "not_comparable"
    assert c(full, dict(full, vat="none_evidenced"), "2025", "2024")["status"] == "unverified"   # same class
    assert c(full, dict(full, vat="unknown"), "2025", "2024")["status"] == "unverified"
    assert c(full, dict(full, basis="expected"), "2025", "2024")["status"] == "not_comparable"
    assert c(full, dict(full, basis="actual"), "2025", "2024")["status"] == "unverified"
    assert c(dict(full, cash_basis="collected"), dict(full, cash_basis="accrued"), "2025", "2024")["status"] == "not_comparable"
    assert c(full, dict(full, completeness="partial"), "2025", "2024")["status"] == "not_comparable"
    # the previous field belongs to the same year (e.g. 2025 incl used as "previous")
    d = c(full, dict(full), "2025", "2024", "2025", "2025")
    assert d["status"] == "not_comparable" and "2025" in d["reason"]
    # prev_id null (previous period trimmed from 04) stays valid
    assert c(full, dict(full), "2025", None, "2025", "2024")["status"] == "confirmed"


def test_semantic_errors_new_rules():
    cm = _cm()
    cfg = config(metrics=[
        {"id": "gyo", "kind": "ratio", "num": "{REV}", "den": "{VAL}", "unit": "%",
         "where": {"all": [ASSET, {"eq": ["ownership", "owned"]}]}},
        {"id": "inown", "kind": "count", "unit": "count", "where": {"in": ["ownership", ["owned", "partial"]]}},
        {"id": "vp", "kind": "count", "unit": "count", "valid_periods": ["2025", "2031"]},
        {"id": "ez", "kind": "ratio", "num": "{REV}", "den": "{VAL}", "unit": "%", "empty_as_zero": True},
    ])
    errs = cm.semantic_errors(cfg, records())
    joined = "\n".join(errs)
    assert "metrics/gyo/where/all/1" in joined and '{"has": ["ownership", "owned"]}' in joined   # B12
    assert "metrics/inown/where" in joined and '"any"' in joined
    assert "2031" in joined                                                                   # unknown period id
    assert "metrics/ez" in joined                                                             # empty_as_zero kind
    assert cm.semantic_errors(cfg) == [e for e in errs if "where" not in e]                   # records optional
    # prev_id pointing to a removed period: the message tells how to trim
    cfg2 = config()
    cfg2["periods"] = cfg2["periods"][:1]
    msg = "\n".join(cm.semantic_errors(cfg2))
    assert "prev_id" in msg and "null" in msg
    cfg2["periods"][0]["prev_id"] = None
    assert cm.semantic_errors(cfg2) == []


def test_comparability_errors_block_hand_set_status():
    cm = _cm()
    meta = {"rev.2025": {"vat": "excl", "basis": "recorded", "completeness": "full_year", "period": "2025"},
            "rev.2024": {"vat": "incl", "basis": "recorded", "completeness": "partial", "period": "2024"}}
    cfg = config()
    cfg["periods"][0]["comparable_with_prev"] = {"status": "confirmed", "reason": "متماثلتان"}
    errs = cm.comparability_errors(cfg, meta)
    assert errs and "periods/2025" in errs[0] and "not_comparable" in errs[0]
    assert any(e.startswith("• trend") for e in errs)
    cfg["periods"][0]["comparable_with_prev"] = {"status": "not_comparable", "reason": "x"}
    cfg["trend"]["comparable"] = {"status": "not_comparable", "reason": "x"}
    assert cm.comparability_errors(cfg, meta) == []


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
