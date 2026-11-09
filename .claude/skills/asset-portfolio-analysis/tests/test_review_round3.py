#!/usr/bin/env python3
"""Regression tests for the round-3 review fixes, on FICTIONAL data only (no real workbook value).

    python3 -I tests/test_review_round3.py        # exit 0 = all pass, 1 = a failure

Covers:
  N3   validate totals_reconciliation: a total row of a source with no in-scope record is compared with all its records
       (never «recorded − 0»), and the warning names the field by its Arabic label and the sheet by its name.
  N6   validate export_growth_pct: a growth % with unverified comparability in the supporting workbook is reported.
  V5   normalize issue texts: Arabic field labels and metric names, no engine keys, analyst instructions in dev_note,
  V8   and number–noun agreement («4 سجلات»).
  N5   build_dashboard: {{m:<growth>}} with unverified comparability renders the change in riyals, never a %.
  BP-03 facts.py: per-source sums / zero counts / largest-line share and per-section record counts of a statement.
"""
import json
import os
import re
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.environ.get("APA_SKILL_DIR") or os.path.dirname(HERE)
sys.path.insert(0, os.path.join(SKILL, "scripts"))
sys.dont_write_bytecode = True

import build_dashboard as bd  # noqa: E402
import facts  # noqa: E402
import normalize as nz  # noqa: E402
import validate_analysis as va  # noqa: E402

FAILS, PASSES = [], []
ENGINE_KEY_RE = re.compile(r"[a-z][a-z0-9]*(?:[._][a-z0-9]+)+")


def check(name, cond, info=""):
    (PASSES if cond else FAILS).append(name)
    print(("  PASS  " if cond else "  FAIL  ") + name + ("" if cond else f" — {info}"))


def tv(v, s="value", ref=None):
    return {"v": v, "s": s, "raw": None if v is None else str(v), "ref": ref}


def write(d, name, data):
    with open(os.path.join(d, name), "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)


META = {"rev.2025.stmt": {"label": "إيراد البيان التجريبي 2025", "group": "revenue", "unit": "SAR", "period": "2025"},
        "exp.2025.stmt": {"label": "مصروف البيان التجريبي 2025", "group": "expense", "unit": "SAR"},
        "val.2024.total": {"label": "تقييم 2024 التجريبي", "group": "valuation", "unit": "SAR"}}


def statement_run(d):
    """A fictional isolated statement «بيان وهمي»: section 1 (rows 2–3), total row 4; section 2 (row 5), total row 6;
    a typed grand total on row 7 that does not match; one zero line."""
    def rec(rid, row, rev, scope="isolated"):
        return {"id": rid, "name": f"بند وهمي {rid}", "record_type": "asset", "scope": scope, "dims": {},
                "source": {"source_id": "V", "sheet": "بيان وهمي", "row": row},
                "fields": {"rev.2025.stmt": tv(rev, "zero_recorded" if rev == 0 else "value", f"بيان وهمي!C{row}")}}
    records = [rec("V-r2", 2, 400), rec("V-r3", 3, 0), rec("V-r5", 5, 100),
               {"id": "R-1", "name": "أصل وهمي", "record_type": "asset", "scope": "in_scope", "dims": {},
                "source": {"source_id": "R", "sheet": "سجل وهمي", "row": 2}, "fields": {"rev.2025.stmt": tv(50)}}]
    excluded = [{"source_id": "V", "sheet": "بيان وهمي", "row": r, "reason": "صف إجمالي",
                 "values": {"rev.2025.stmt": tv(v, ref=f"بيان وهمي!C{r}")}} for r, v in ((4, 400), (6, 100), (7, 350))]
    write(d, "03_records.json", {"contract_version": "1.0", "field_meta": META, "records": records,
                                 "excluded_rows": excluded, "issues": []})


def test_reconciliation_isolated():
    print("[N3] مطابقة إجمالي مصدر معزول")
    with tempfile.TemporaryDirectory() as d:
        statement_run(d)
        ctx = va.Ctx(d)
        rows, warns = va.g4b_reconciliation(ctx)
        by = {r["row"]: r for r in rows}
        check("section totals of an isolated source match all its records, difference 0",
              by[4]["status"] == "match_section_all_records" and by[4]["difference"] == 0
              and by[6]["difference"] == 0, [(r["row"], r["status"], r["difference"]) for r in rows])
        check("grand total: difference 350 − 500 = −150 against all records (basis all_records)",
              by[7]["status"] == "mismatch" and by[7]["difference"] == -150 and by[7]["difference_basis"] == "all_records",
              by[7])
        txt = " ".join(w[0] for w in warns)
        check("warning says «لا سجلات ضمن النطاق» and gives −150, not «الفرق 350»",
              "لا سجلات ضمن النطاق" in txt and "-150" in txt and "الفرق 350" not in txt, txt)
        check("warning names the field by its Arabic label and the sheet by its name, not rev.2025.stmt / V",
              "«إيراد البيان التجريبي 2025»" in txt and "«بيان وهمي»" in txt and "rev.2025.stmt" not in txt, txt)


def test_facts_statement():
    print("[BP-03] حقائق وصفية لبيان معزول")
    with tempfile.TemporaryDirectory() as d:
        statement_run(d)
        f = facts.compute_facts(d)
        g = lambda k: (f.get(k) or {}).get("value")  # noqa: E731
        check("sources.V.<field>.sum / n_known / n_zero over every record of the source",
              g("sources.V.rev.2025.stmt.sum") == 500 and g("sources.V.rev.2025.stmt.n_known") == 3
              and g("sources.V.rev.2025.stmt.n_zero") == 1, {k: v["value"] for k, v in f.items() if k.startswith("sources.V")})
        check("largest line and its share (unit pct)",
              g("sources.V.rev.2025.stmt.max") == 400 and abs(g("sources.V.rev.2025.stmt.max_share") - 80) < 1e-9
              and f["sources.V.rev.2025.stmt.max_share"]["unit"] == "pct", "")
        check("records per section above each total row (an empty grand-total section has no fact)",
              g("sections.V.r4.records") == 2 and g("sections.V.r6.records") == 1 and "sections.V.r7.records" not in f,
              {k: v["value"] for k, v in f.items() if k.startswith("sections.")})
        tr = bd.TokenRenderer({}, {}, "SAR", {}, f)
        check("build renders a pct fact as a percentage", tr.render("{{f:sources.V.rev.2025.stmt.max_share}}") == "80.0%",
              tr.render("{{f:sources.V.rev.2025.stmt.max_share}}"))


def test_issue_wording():
    print("[V5/V8] صياغة مشكلات البيانات للمجلس")
    meta = dict(META)
    iss = nz.Issues()
    iss.add("valuation_total_mismatch", "high", "سجل وهمي!X2", ["R-1"],
            "إجمالي التقييم لا يطابق مكوّنيه في val.2024.total.",
            nz.impact_for("val.2024.total", "money", meta),
            "لم يُصحَّح تلقائيًا؛ إن ثبت خطأ الصيغة فأضف derived_fields صريحًا", field="val.2024.total")
    iss.add("zero_as_missing", "medium", "سجل وهمي!X2:X5", ["R-1"],
            "قيمة صفرية في الحقل " + nz.flabel("val.2024.total", meta) + " (" + nz.count_ar(4, "record") + ")",
            "x", "عومل الصفر كغير متاح (zero_as_missing) ولم يدخل في المجاميع", field="val.2024.total")
    items = iss.finalize({"R-1": {"scope": "in_scope"}}, meta)
    texts = " ".join(f"{i['detail']} {i['impact']} {i['treatment']}" for i in items)
    check("no engine keys (field ids, metric ids, zero_as_missing, derived_fields) in detail/impact/treatment",
          not ENGINE_KEY_RE.search(texts), ENGINE_KEY_RE.findall(texts))
    check("field shown by its label «تقييم 2024 التجريبي»; metrics by Arabic names", "«تقييم 2024 التجريبي»" in texts
          and "القيمة المسجلة" in texts and "العائد الإجمالي" in texts, texts)
    check("analyst instruction moved to dev_note", "derived_fields" in (items[0].get("dev_note") or ""), items[0])
    check("treatment keeps its board-facing clause", items[0]["treatment"].startswith("لم يُصحَّح تلقائيًا"), items[0]["treatment"])
    check("count agreement «4 سجلات» (not «4 سجل»)", "(4 سجلات)" in texts and "4 سجل)" not in texts, texts)
    check("«__original» field named «… كما ورد قبل الاشتقاق»",
          nz.flabel("val.2024.total__original", meta) == "«تقييم 2024 التجريبي كما ورد قبل الاشتقاق»",
          nz.flabel("val.2024.total__original", meta))


def growth_metrics(status="unverified"):
    return {"revenue_growth": {"id": "revenue_growth", "kind": "growth", "value": 25.0, "unit": "%", "status": "complete",
                               "abs_change": 2000, "comparability": {"status": status, "reason": "س"}}}


def test_growth_rendering():
    print("[N5] رمز النمو غير المتحقق في السرد")
    tr = bd.TokenRenderer(growth_metrics(), {}, "SAR", {}, {})
    out = tr.render("التغير {{m:revenue_growth}}")
    check("unverified growth token -> riyals, no %", "2,000" in out and "%" not in out, out)
    tr = bd.TokenRenderer(growth_metrics("confirmed"), {}, "SAR", {}, {})
    out = tr.render("{{m:revenue_growth}}")
    check("confirmed growth token keeps its %", out.endswith("%"), out)


def test_export_growth_check():
    print("[N6] فحص نسبة النمو في المرفق الداعم")
    import openpyxl
    with tempfile.TemporaryDirectory() as d:
        write(d, "05_metrics.json", {"metrics": growth_metrics()})
        path = os.path.join(d, va.EXPORT_NAME)
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "ملخص"
        ws["A1"], ws["B1"], ws["C1"] = "تغير الإيراد", 25.0, "النسبة المحسوبة 25.0% للمراجعة فقط، لا تُقرأ نموًا"
        ws["D1"], ws["E1"] = "حصة أكبر سطر 25.0% من المجموع", "تغير الإيراد 25.0% عن العام السابق"
        wb.save(path)
        out = va.c_export_growth_pct(va.Ctx(d))
        check("a % value cell, a «x.x%» growth text and a «للمراجعة فقط» note are reported; an unrelated share is not",
              len(out) == 1 and "ملخص!B1" in out[0][1] and "ملخص!E1" in out[0][1] and "ملخص!C1" in out[0][1]
              and "D1" not in out[0][1], out)
        wb = openpyxl.Workbook()
        wb.active["A1"], wb.active["B1"] = "تغير الإيراد", 2000
        wb.save(path)
        check("riyals only: nothing reported", va.c_export_growth_pct(va.Ctx(d)) == [], "")


def main():
    test_reconciliation_isolated()
    test_facts_statement()
    test_issue_wording()
    test_growth_rendering()
    test_export_growth_check()
    print(f"\nناجح: {len(PASSES)} · فاشل: {len(FAILS)}")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
