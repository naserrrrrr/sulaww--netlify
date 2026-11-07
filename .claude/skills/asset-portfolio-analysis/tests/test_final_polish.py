#!/usr/bin/env python3
"""Regression tests for the final-polish items, on FICTIONAL data only (no real workbook value).

    python3 -I tests/test_final_polish.py        # exit 0 = all pass, 1 = a failure

Covers:
  item 2   inspect: the personal-data column note has no engine word («pii»); export: issue types, gap scopes,
           evidence tokens and check ids by their Arabic labels — keys only in code-styled «المرجع التقني» cells.
  item 3   engine scope_ids: a not_computable metric that included and excluded none lists its scope (both engines are
           compared in tests/test_metric_engine_js.mjs); the export counts and lists the same records.
  item 4   export «ملخص»: top gaps and top actions are two separate tables (an action names the gaps it addresses).
  item 6   export «المطابقة مع الإجماليات»: a section total lists its own section's lines.
  item 7   normalize impact text: «حقل التقييم المستخدم في المؤشرات», never «التقييم المعتمد».
  item 8   inspect: creator and last_modified_by are both masked (only a software name is kept).
  item 10  validate exec_summary_hygiene (warning): Sheet!Cell / range / formula and «1 أصلًا واحدًا» in the summary.
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

import openpyxl  # noqa: E402

import apa_common  # noqa: E402
import export_workbook as ex  # noqa: E402
import inspect_workbook as iw  # noqa: E402
import metric_engine as me  # noqa: E402
import normalize as nz  # noqa: E402
import validate_analysis as va  # noqa: E402

FAILS, PASSES = [], []
RAW_KEY_RE = re.compile(r"(?<![A-Za-z0-9_.])(?:[a-z][a-z0-9]*(?:[._][a-z0-9*]+)+|pii|portfolio|zero_as_missing|info)(?![A-Za-z0-9_])")


def check(name, cond, info=""):
    (PASSES if cond else FAILS).append(name)
    print(("  PASS  " if cond else "  FAIL  ") + name + ("" if cond else f" — {info}"))


def tv(v, s="value", ref=None):
    return {"v": v, "s": s, "raw": None if v is None else str(v), "ref": ref}


def write(d, name, data):
    with open(os.path.join(d, name), "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)


META = {"rev.2025.demo": {"label": "إيراد تجريبي 2025", "group": "revenue", "unit": "SAR", "period": "2025"},
        "val.2024.demo": {"label": "تقييم تجريبي 2024", "group": "valuation", "unit": "SAR"}}


# ---------------------------------------------------------------------------------------------------------------
def test_engine_scope_ids():
    print("[3] المحرك: نطاق المؤشر المتعذر (scope_ids)")
    recs = [{"id": f"D-{i}", "scope": "in_scope", "record_type": "asset", "dims": {"use": "investment"},
             "fields": {"rev.2025.demo": tv(100 * i)}} for i in (1, 2)]
    cfg = {"base_where": {"eq": ["scope", "in_scope"]},
           "periods": [{"id": "2025", "rev_field": "rev.2025.demo", "exp_field": None}],
           "valuation": {"field": "val.2024.demo", "year": 2024}, "metrics": []}
    eng = me.Engine(cfg, recs)
    res = eng.compute({"id": "noi", "kind": "diff", "a": "{REV}", "b": "{EXP}", "unavailable_reason": "لا مصروفات"},
                      {"period": "2025"})
    check("unresolved {EXP}: not_computable with scope_ids = the in-scope records",
          res["status"] == "not_computable" and res["scope_ids"] == ["D-1", "D-2"] and res["excluded"] == [], res)
    out = ex.Export.excluded_of(res)
    check("export lists the same records with the reason (counts and list from one source)",
          [e["id"] for e in out] == ["D-1", "D-2"] and all(e["reason_text"] == "لا مصروفات" for e in out), out)


# ---------------------------------------------------------------------------------------------------------------
def demo_run(d):
    """A fictional run: one in-scope asset, an isolated statement with two section totals and a grand total."""
    def rec(rid, row, rev, scope="isolated", src="S"):
        return {"id": rid, "name": f"بند تجريبي {rid}", "record_type": "asset", "scope": scope, "dims": {},
                "source": {"source_id": src, "sheet": "بيان تجريبي", "row": row},
                "fields": {"rev.2025.demo": tv(rev, ref=f"بيان تجريبي!C{row}")}}
    records = [rec("S-r2", 2, 10), rec("S-r3", 3, 20), rec("S-r5", 5, 5),
               {"id": "A-1", "name": "مبنى تجريبي", "record_type": "asset", "scope": "in_scope", "dims": {},
                "source": {"source_id": "A", "sheet": "سجل تجريبي", "row": 2},
                "fields": {"rev.2025.demo": tv(50, ref="سجل تجريبي!C2"), "val.2024.demo": tv(1000, ref="سجل تجريبي!D2")}}]
    excluded = [{"source_id": "S", "sheet": "بيان تجريبي", "row": r, "reason": "صف إجمالي",
                 "values": {"rev.2025.demo": tv(v, ref=f"بيان تجريبي!C{r}")}} for r, v in ((4, 30), (6, 5), (7, 40))]
    issues = [{"id": "DQ-001", "type": "zero_as_missing", "severity": "medium", "where": "سجل تجريبي!D2",
               "record_ids": ["A-1"], "field": "val.2024.demo",
               "detail": "قيمة صفرية في الحقل val.2024.demo (zero_as_missing).",
               "impact": "يؤثر في recorded_value و gross_yield_*", "treatment": "عومل الصفر كغير متاح"}]
    write(d, "03_records.json", {"contract_version": "1.0", "source": {"file": "demo.xlsx"}, "field_meta": META,
                                 "records": records, "excluded_rows": excluded, "issues": issues, "stats": {"records": 4}})
    metrics = {
        "recorded_value": {"id": "recorded_value", "kind": "sum", "label": "القيمة المسجلة التجريبية",
                           "short_label": "القيمة المسجلة", "value": 1000, "unit": "SAR", "status": "complete",
                           "n_scope": 1, "n_included": 1, "included_ids": ["A-1"], "excluded": [],
                           "resolved_fields": {"field": "val.2024.demo"}},
        "gross_yield_investment": {"id": "gross_yield_investment", "kind": "ratio", "label": "نسبة الإيراد إلى التقييم",
                                   "value": 5.0, "unit": "%", "status": "complete", "n_scope": 1, "n_included": 1,
                                   "included_ids": ["A-1"], "excluded": [], "num": 50, "den": 1000,
                                   "resolved_fields": {"num": "rev.2025.demo", "den": "val.2024.demo"}},
        "noi": {"id": "noi", "kind": "diff", "label": "صافي الدخل التشغيلي", "value": None, "unit": "SAR",
                "status": "not_computable", "reason": "لا مصروفات في الملف", "n_scope": 1, "n_included": 0,
                "included_ids": [], "excluded": [], "scope_ids": ["A-1"], "unresolved": "{EXP}",
                "resolved_fields": {"a": "rev.2025.demo"}},
        "data_coverage": {"id": "data_coverage", "kind": "coverage", "label": "تغطية البيانات", "value": 100.0,
                          "unit": "%", "status": "complete", "n_scope": 1, "n_included": 1,
                          "per_field": [{"field": "rev.2025.demo", "n_known": 1, "n": 1},
                                        {"field": "val.2024.demo", "n_known": 1, "n": 1}]}}
    write(d, "05_metrics.json", {"metrics": metrics})
    write(d, "01_workbook_map.json", {"anomalies": [
        {"id": "WB-001", "type": "pii_column", "severity": "info", "where": "سجل تجريبي!F:F",
         "detail": "العمود F يحتوي بيانات شخصية. لا تُنقل قيمه إلى أي مخرج."}]})
    write(d, "07_validation.json", {"blocking_count": 0, "warning_count": 0, "acknowledged_count": 1,
                                    "generated_at": "2026-01-02T10:20:30",
                                    "checks": [{"id": "totals_reconciliation", "level": "acknowledged", "passed": False,
                                                "title": "مطابقة الإجماليات المدونة"}]})
    write(d, "06_narrative.json", {
        "executive_summary": "خلاصة تجريبية.",
        "gaps": [{"id": "G1", "title": "فجوة تجريبية أولى", "scope": "portfolio", "severity": "high",
                  "is_proven_problem": False, "evidence": [{"ref": "m:recorded_value", "kind": "derived"},
                                                           {"ref": "سجل تجريبي!C2", "kind": "fact"}],
                  "impact": {"text": "أثر تجريبي", "metric": "m:noi"}},
                 {"id": "G2", "title": "فجوة تجريبية ثانية", "scope": "A-1", "severity": "medium",
                  "is_proven_problem": True, "evidence": ["m:gross_yield_investment"], "impact": {"text": "أثر"}}],
        "recommendations": [{"id": "R1", "priority": 1, "gap_ids": ["G2"], "target": "A-1",
                             "first_action": "إجراء تجريبي للفجوة الثانية", "owner_role": "مدير تجريبي",
                             "evidence": ["m:gross_yield_investment"]},
                            {"id": "R2", "priority": 2, "gap_ids": ["G1"], "target": "المحفظة",
                             "first_action": "إجراء تجريبي للفجوة الأولى", "owner_role": "مدير مالي تجريبي"}],
        "top_gaps": ["G1", "G2"], "top_actions": ["R1", "R2"]})


def cell_rows(ws):
    return [list(r) for r in ws.iter_rows(values_only=True)]


def test_export_board_wording():
    print("[2/3/4/6] المرفق الداعم: صياغة المجلس والقائمتان والمطابقة")
    with tempfile.TemporaryDirectory() as d:
        demo_run(d)
        out = os.path.join(d, "supporting_workbook.xlsx")
        rc = ex.export(d, out)
        check("export succeeds", rc == 0 and os.path.isfile(out), rc)
        if rc != 0:
            return
        wb = openpyxl.load_workbook(out)
        # item 2: no engine key outside the code-styled technical cells
        leaks = []
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    if not isinstance(c.value, str) or (c.font is not None and c.font.name == ex.CODE_FONT):
                        continue
                    if c.value.startswith(ex.TECH):
                        continue
                    for k in RAW_KEY_RE.findall(c.value):
                        if not k.endswith((".xlsx", ".json")):
                            leaks.append(f"{ws.title}!{c.coordinate}: {k}")
        check("no field key, metric id, enum word or issue type outside «المرجع التقني» (code-styled) cells",
              not leaks, leaks[:8])
        ws = wb["مشكلات البيانات"]
        rows = {r[0]: r for r in cell_rows(ws) if r and r[0] in ("WB-001", "DQ-001")}
        check("issue type by its Arabic label, severity «معلوماتية», keys in the technical column",
              rows.get("WB-001", [None] * 10)[2] == apa_common.ANOMALY_TYPE_AR["pii_column"]
              and rows["WB-001"][3] == "معلوماتية" and rows["DQ-001"][9] == "zero_as_missing · val.2024.demo"
              and "«تقييم تجريبي 2024»" in rows["DQ-001"][6], rows)
        ws = wb["الفجوات والتوصيات"]
        g1 = next((r for r in cell_rows(ws) if r and r[0] == "G1"), None)
        check("gap scope «المحفظة كاملة»; evidence «مؤشر: القيمة المسجلة»; impact metric by its label",
              bool(g1) and g1[3] == "المحفظة كاملة" and "مؤشر: القيمة المسجلة" in str(g1[6])
              and str(g1[8]).startswith("مؤشر: صافي الدخل التشغيلي"), g1)
        # item 4: two separate lists
        ws = wb["ملخص"]
        rows = cell_rows(ws)
        heads = [r[0] for r in rows if r and r[0] in ("أهم ثلاث فجوات", "أهم ثلاثة إجراءات أو قرارات مطلوبة",
                                                        "أهم ثلاث فجوات وأهم ثلاثة إجراءات")]
        gap_rows = [r for r in rows if r and r[0] in ("فجوة تجريبية أولى", "فجوة تجريبية ثانية")]
        check("«ملخص» lists the top gaps and the top actions as two separate tables (no row pairs a gap with an action)",
              heads == ["أهم ثلاث فجوات", "أهم ثلاثة إجراءات أو قرارات مطلوبة"] and len(gap_rows) == 2
              and not any("إجراء" in str(c) for r in gap_rows for c in r), (heads, gap_rows))
        act = [r for r in rows if r and r[0] in ("إجراء تجريبي للفجوة الثانية", "إجراء تجريبي للفجوة الأولى")]
        check("each action names the gap it addresses («يعالج الفجوة»)",
              [(r[0], r[3]) for r in act] == [("إجراء تجريبي للفجوة الثانية", "G2"), ("إجراء تجريبي للفجوة الأولى", "G1")], act)
        vline = next((r[1] for r in rows if r and r[0] == "التحقق الآلي"), "")
        check("validation line names the acknowledged check by its Arabic title, not its id",
              "«مطابقة الإجماليات المدونة»" in str(vline) and "totals_reconciliation" not in str(vline), vline)
        cov = next((r for r in rows if r and r[0] == "تغطية البيانات"), None)
        check("coverage note per field by Arabic labels", bool(cov) and "«إيراد تجريبي 2025»: 1 من 1" in str(cov[8]), cov)
        # item 3: calculations — counts and list from the same source
        ws = wb["الحسابات"]
        noi = next((r for r in cell_rows(ws) if r and r[0] == "noi"), None)
        check("«الحسابات»: NOI in scope 1 · excluded 1, listed with its reason",
              bool(noi) and noi[6] == 1 and noi[7] == 0 and noi[8] == 1 and "A-1 (لا مصروفات في الملف)" == noi[21], noi)
        # item 6: a section total lists its own section's lines
        ws = wb["المطابقة مع الإجماليات"]
        by_ref = {r[1]: r for r in cell_rows(ws) if r and isinstance(r[1], str) and r[1].startswith("بيان تجريبي!")}
        check("section totals list their own section's lines; the grand total every line",
              by_ref.get("بيان تجريبي!C4", [None] * 13)[12] == "S-r2، S-r3" and by_ref["بيان تجريبي!C6"][12] == "S-r5"
              and by_ref["بيان تجريبي!C7"][12] == "S-r2، S-r3، S-r5", {k: v[12] for k, v in by_ref.items()})


# ---------------------------------------------------------------------------------------------------------------
def test_normalize_wording():
    print("[7] صياغة الأثر في مشكلات البيانات")
    t = nz.impact_for("val.2024.demo", "money", META) + " " + nz.impact_for("rev.2025.demo", "money", META)
    check("«حقل التقييم المستخدم في المؤشرات» and «حقل إيراد الفترة المستخدم في المؤشرات», never «المعتمد»",
          "حقل التقييم المستخدم في المؤشرات" in t and "حقل إيراد الفترة المستخدم في المؤشرات" in t and "المعتمد" not in t, t)


def test_inspect_author_and_pii_note():
    print("[2/8] فحص المصنف: حجب المؤلف وصياغة عمود البيانات الشخصية")
    with tempfile.TemporaryDirectory() as d:
        src = os.path.join(d, "demo.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "سجل تجريبي"
        ws.append(["الرقم", "اسم الأصل", "الإيراد 2025", "جوال المستأجر"])
        for i in range(1, 6):
            ws.append([i, f"عقار تجريبي {i}", 1000 * i, f"05{i}0000000"])
        wb.properties.creator = "س.ت"              # initials: a person
        wb.properties.lastModifiedBy = "مستخدم وهمي"
        wb.save(src)
        out = os.path.join(d, "run")
        os.makedirs(out)
        rc = iw.main([src, "--out", out]) if hasattr(iw, "main") else None
        wm = json.load(open(os.path.join(out, "01_workbook_map.json"), encoding="utf-8"))
        props = (wm.get("file") or {}).get("properties") or {}
        check("creator (initials) and last_modified_by both masked", props.get("creator") == apa_common.REDACTED
              and props.get("last_modified_by") == apa_common.REDACTED, (rc, props))
        md = open(os.path.join(out, "01_workbook_map.md"), encoding="utf-8").read()
        check("neither author value appears in 01", "س.ت" not in md and "مستخدم وهمي" not in md
              and "س.ت" not in json.dumps(wm, ensure_ascii=False), "")
        notes = [a.get("detail") or "" for a in wm.get("anomalies") or [] if a.get("type") == "pii_column"]
        check("the personal-data column note is Arabic only (no «pii»)",
              bool(notes) and not any(re.search(r"\bpii\b", n) for n in notes), notes)
    check("a software author is kept", iw.is_software_name("openpyxl") and iw.is_software_name("Microsoft Office User")
          and not iw.is_software_name("س.ت") and not iw.is_software_name("abc"), "")


# ---------------------------------------------------------------------------------------------------------------
def test_exec_summary_hygiene():
    print("[10] نظافة الخلاصة التنفيذية (تحذير)")
    metrics = {"asset_count": {"id": "asset_count", "kind": "count", "value": 1, "unit": "asset", "status": "complete"},
               "recorded_value": {"id": "recorded_value", "kind": "sum", "value": 1000, "unit": "SAR", "status": "complete"}}
    with tempfile.TemporaryDirectory() as d:
        write(d, "05_metrics.json", {"metrics": metrics})
        bad = ("يضم السجل التجريبي {{m:asset_count}} أصلًا واحدًا، ويحمل البيان التجريبي (بيان تجريبي!B7) رقمًا آخر، "
               "والإجمالي B50:C50 يُربط بالصيغة =B20+B35+B48. القيمة {{m:recorded_value}}.")
        write(d, "06_narrative.json", {"executive_summary": bad})
        out = va.c_nq_exec_summary_hygiene(va.Ctx(d))
        txt = " ".join(o[0] for o in out)
        check("cell reference, range and formula flagged", len(out) == 2 and "تجريبي!B7" in txt and "B50:C50" in txt
              and "=B20+B35+B48" in txt, out)
        check("«{{m:asset_count}} أصلًا واحدًا» (renders «1 أصلًا واحدًا») flagged as a doubled numeral", "1 اصلا واحدا" in txt, txt)
        good = ("يضم السجل التجريبي أصلًا واحدًا ضمن النطاق، ويحمل البيان التجريبي المخفي رقمًا آخر لإيراد الأصل، "
                "وقيمته المسجلة {{m:recorded_value}} من 2021 إلى 2025.")
        write(d, "06_narrative.json", {"executive_summary": good})
        check("a clean summary passes", va.c_nq_exec_summary_hygiene(va.Ctx(d)) == [], va.c_nq_exec_summary_hygiene(va.Ctx(d)))
        natural = ("ينفَّذ السجل الموحد خلال 90 يومًا على ثلاث مراحل، وتغطي البيانات من 2021 إلى 2025 خمس سنوات، "
                   "والقيمة المسجلة {{m:recorded_value}}.")
        write(d, "06_narrative.json", {"executive_summary": natural})
        out = va.c_nq_exec_summary_hygiene(va.Ctx(d))
        check("a different count after a numeral is natural phrasing, not flagged", out == [], out)
        write(d, "06_narrative.json", {"executive_summary": "يضم السجل 2 أصلان، و3 ثلاث وحدات."})
        txt = " ".join(o[0] for o in va.c_nq_exec_summary_hygiene(va.Ctx(d)))
        check("«2 أصلان» and «3 ثلاث» still flagged", "2 اصلان" in txt and "3 ثلاث" in txt, txt)
    lvl = dict((cid, lv) for cid, lv, _ in va.CHECKS).get("exec_summary_hygiene")
    check("registered as a warning with an Arabic title", lvl == "warning" and "الخلاصة" in va.CHECK_TITLES.get("exec_summary_hygiene", ""), lvl)


def main():
    test_engine_scope_ids()
    test_export_board_wording()
    test_normalize_wording()
    test_inspect_author_and_pii_note()
    test_exec_summary_hygiene()
    print(f"\nناجح: {len(PASSES)} · فاشل: {len(FAILS)}")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
