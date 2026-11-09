#!/usr/bin/env python3
"""Regression tests for the board-ready round (engine and text side), on FICTIONAL data only.

    python3 -I tests/test_engine_text.py        # exit 0 = all pass, 1 = a failure

Covers:
  C1  metric review flag: schema, both Python paths (engine result, compute_metrics 05), validate metric_review
      (unknown status, reason outside 6–25 words, digits, 05 out of date) — the JS twin is in test_metric_engine_js.mjs
  C2  {{f:rec.<record_id>.<field_id>}} facts: compute_facts lists the cited ones, record_fact reasons, money format in
      the build renderer, fact_tokens blocks a bad key, f: evidence refs resolve, facts.py prints the pattern once
  C4  top_actions_priority (blocking): ascending order, unique integer priorities; recommendation_quality warns when a
      lower priority number is left out of the top actions
  ar  Arabic number agreement: apa_common.ar_count / ar_n / ar_agree, normalize.count_ar, and every count phrase in the
      anomaly texts of a small fictional workbook («7 قيم»، «15 صفًا»، «105 صفوف»، «اسم واحد»، «لا نصوص»)
  rec validate 07 reconciliation: difference = recorded − compared sum, the matched sum apart (never 0 beside two
      different numbers)
  doc every $SKILL_DIR command quoted, no bare «python3 scripts/…», the 09 template without the QA paragraph,
      feminine severity labels, the zero-revenue card label
"""
import contextlib
import copy
import io
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

import apa_common as ac  # noqa: E402
import build_dashboard as bd  # noqa: E402
import compute_metrics as cm  # noqa: E402
import facts  # noqa: E402
import inspect_workbook as iw  # noqa: E402
import metric_engine as me  # noqa: E402
import normalize as nz  # noqa: E402
import validate_analysis as va  # noqa: E402

FAILS, PASSES = [], []


def check(name, cond, info=""):
    (PASSES if cond else FAILS).append(name)
    print(("  PASS  " if cond else "  FAIL  ") + name + ("" if cond else f" — {info}"))


def tv(v, s=None, ref=None, raw=None):
    if s is None:
        s = "blank" if v is None else ("zero_recorded" if v == 0 else "value")
    return {"v": v if s in ac.KNOWN_STATUSES else None, "s": s, "raw": raw if raw is not None else (None if v is None else str(v)),
            "ref": ref, "note": None}


def write(d, name, data):
    with open(os.path.join(d, name), "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)


META = {"rev.2025.demo": {"label": "إيراد تجريبي 2025", "group": "revenue", "unit": "SAR", "period": "2025"},
        "rev.2025.stmt": {"label": "إيراد البيان التجريبي 2025", "group": "revenue", "unit": "SAR", "period": "2025"},
        "val.2024.demo": {"label": "تقييم تجريبي 2024", "group": "valuation", "unit": "SAR"},
        "units": {"label": "عدد الوحدات", "group": "other"}}
REVIEW = {"status": "needs_verification", "reason": "إيراد الفترة له رقم بديل في بيان آخر من الملف لم يُحسم بعد"}


# ---------------------------------------------------------------------------------------------------------------
def test_ar_count():
    print("[ar] مطابقة العدد والمعدود")
    row = lambda n, **k: ac.ar_n(n, "row", **k)  # noqa: E731
    cases = {0: "لا صفوف", 1: "صف واحد", 2: "صفان", 3: "3 صفوف", 10: "10 صفوف", 11: "11 صفًا", 15: "15 صفًا",
             99: "99 صفًا", 100: "100 صف", 101: "101 صف", 102: "102 صف", 103: "103 صفوف", 105: "105 صفوف",
             111: "111 صفًا", 1000: "1,000 صف", 1015: "1,015 صفًا"}
    got = {n: row(n) for n in cases}
    check("rows: 0 / 1 / 2 / 3–10 / 11–99 / 100–102 / 103 / 105 / 111 / thousands", got == cases,
          {n: (got[n], cases[n]) for n in cases if got[n] != cases[n]})
    check("values (feminine): «قيمة واحدة»، «7 قيم»، «12 قيمة»، «لا قيم»",
          [ac.ar_n(n, "value") for n in (1, 7, 12, 0)] == ["قيمة واحدة", "7 قيم", "12 قيمة", "لا قيم"])
    check("numbers and texts: «7 أرقام»، «لا نصوص»، «15 رقمًا»، «نص واحد»",
          [ac.ar_n(7, "number"), ac.ar_n(0, "text"), ac.ar_n(15, "number"), ac.ar_n(1, "text")]
          == ["7 أرقام", "لا نصوص", "15 رقمًا", "نص واحد"])
    check("case acc/gen changes only 1 and 2: «صفًا واحدًا»، «صفين»، «سجلين»، «قيمتين»، «خطأين»",
          [row(1, case="acc"), row(2, case="gen"), ac.ar_n(2, "record", case="gen"), ac.ar_n(2, "value", case="acc"),
           ac.ar_n(2, "error", case="gen"), row(7, case="gen")]
          == ["صفًا واحدًا", "صفين", "سجلين", "قيمتين", "خطأين", "7 صفوف"])
    phrase = lambda n: ac.ar_count(n, "تاريخ هجري", "تاريخان هجريان", "تواريخ هجرية")  # noqa: E731
    check("noun + adjective: «تاريخ هجري واحد»، «تاريخان هجريان»، «4 تواريخ هجرية»، «12 تاريخًا هجريًا»",
          [phrase(n) for n in (1, 2, 4, 12)] == ["تاريخ هجري واحد", "تاريخان هجريان", "4 تواريخ هجرية", "12 تاريخًا هجريًا"])
    check("explicit one / acc / zero forms",
          ac.ar_count(1, "اسم متطابق", "اسمان متطابقان", "أسماء متطابقة", one="اسم واحد متطابق") == "اسم واحد متطابق"
          and ac.ar_count(13, "صف بيانات", "صفا بيانات", "صفوف بيانات", acc="صفًا من البيانات") == "13 صفًا من البيانات"
          and ac.ar_count(0, "سجل مشمول", "سجلان مشمولان", "سجلات مشمولة", zero="لا سجلات مشمولة") == "لا سجلات مشمولة")
    check("ar_agree: verb/adjective forms follow the count",
          [ac.ar_agree(n, "يقع", "يقعان", "تقع", "تقع") for n in (1, 2, 5, 12)] == ["يقع", "يقعان", "تقع", "تقع"])
    check("normalize.count_ar delegates (0 → «لا سجلات», gen dual «سجلين»)",
          nz.count_ar(0, "record") == "لا سجلات" and nz.count_ar(2, "record", "gen") == "سجلين"
          and nz.count_ar(12, "record") == "12 سجلًا")
    check("inspect helpers: «صف بياناته الوحيد يحمل نصًا»، «صف واحد من أصل 12 صفًا…»، «7 أرقام ولا نصوص»",
          iw._text_rows_ar(1, 1) == "صف بياناته الوحيد يحمل نصًا"
          and iw._text_rows_ar(1, 12) == "صف واحد من أصل 12 صفًا في بياناته يحمل نصًا"
          and iw._value_kinds_ar({"numeric": 7, "text": 0}) == "7 أرقام ولا نصوص"
          and iw._value_kinds_ar({"numeric": 3, "text": 3, "date": 1}) == "3 أرقام و3 نصوص وتاريخ واحد",
          (iw._text_rows_ar(1, 12), iw._value_kinds_ar({"numeric": 7, "text": 0})))


# A digit count followed by a counted noun must take the form ar_count gives for that number.
_FORMS = {k: v for k, v in ac.AR_NOUNS.items() if k in ("row", "value", "number", "text", "name", "date", "cell", "formula",
                                                          "sheet", "record")}
_NOUN_RE = re.compile(r"(?<![\d,.])(\d[\d,]*)\s+(" + "|".join(
    sorted({w for f in _FORMS.values() for w in (f[0], f[2], ac._acc_word(f[0]))}, key=len, reverse=True)) + r")(?![؀-ۿ])")


def agreement_errors(text: str) -> list:
    bad = []
    for m in _NOUN_RE.finditer(text):
        n = int(m.group(1).replace(",", ""))
        word = m.group(2)
        key = next(k for k, f in _FORMS.items() if word in (f[0], f[2], ac._acc_word(f[0])))
        f = _FORMS[key]
        r = n % 100
        want = f[2] if 3 <= r <= 10 else (ac._acc_word(f[0]) if r >= 11 else (f[0] if n >= 100 else None))
        if want is None or word != want:
            bad.append(m.group(0))
    return bad


def test_inspect_texts():
    print("[ar] نصوص فحص المصنف على مصنف وهمي")
    with tempfile.TemporaryDirectory() as d:
        p = os.path.join(d, "demo_register.xlsx")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "سجل تجريبي"
        ws.append(["م", "اسم العقار", "الإيراد 2025", "تاريخ الملكية", None, "صافي الأصول", "نوع العقار"])
        rows = [(1, "عمارة النموذج الأولى", 1000, "1410/5/12هـ", 300, "غير مقيد", "سكني"),
                (2, "عمارة النموذج الثانية", 2000, None, 400, "غير مقيد", "تجاري"),
                (3, "أرض النموذج", 1500, None, 500, "غير مقيد", "أرض"),
                (4, "مجمع النموذج", 2500, None, 600, "غير مقيد", "تجاري"),
                (5, "محلات النموذج", 900, None, 700, "غير مقيد", "تجاري"),
                (6, "فلة النموذج", 800, None, 800, "غير مقيد", "سكني"),
                (7, "مستودع النموذج", 700, None, 900, "غير مقيد", "تجاري")]
        for r in rows:
            ws.append(list(r))
        h1 = wb.create_sheet("بيان مخفي")
        h1.append(["البند", "المبلغ"])
        for i in range(1, 15):
            h1.append([f"بند تجريبي {i}", i * 100])
        h1.sheet_state = "hidden"
        h2 = wb.create_sheet("قائمة مخفية")
        for i in range(1, 9):
            h2.append([f"خيار {i}"])
        h2.sheet_state = "hidden"
        wb.save(p)
        out = os.path.join(d, "run")
        with contextlib.redirect_stdout(io.StringIO()):
            wm, sas = iw.inspect(p, out)
            iw.write_markdown(os.path.join(out, "01_workbook_map.md"), wm, sas)
        an = wm.get("anomalies") or []
        by_type = {}
        for a in an:
            by_type.setdefault(a["type"], []).append(a["detail"])
        hidden = " ".join(by_type.get("hidden_sheet", []))
        check("hidden sheets: «15 صفًا غير فارغ» and «8 صفوف غير فارغة» (never «8 صفًا»)",
              "وفيها 15 صفًا غير فارغ" in hidden and "وفيها 8 صفوف غير فارغة" in hidden and "8 صفًا" not in hidden, hidden)
        unl = " ".join(by_type.get("unlabeled_column", []))
        check("unlabeled column: «7 قيم … (7 أرقام ولا نصوص؛ …)», never «7 قيمة» / «رقمًا و0 نصًا»",
              "يحتوي 7 قيم في صفوف البيانات (7 أرقام ولا نصوص؛" in unl and "7 قيمة" not in unl and "0 نصًا" not in unl, unl)
        hij = " ".join(by_type.get("hijri_date_text", []))
        check("one Hijri date: «تاريخ هجري واحد», never «1 تاريخًا»", "تاريخ هجري واحد" in hij and "1 تاريخًا" not in hij, hij)
        tin = " ".join(by_type.get("text_in_numeric", []))
        check("money column with text in every row: «كل صفوف بياناته (7 صفوف) تحمل نصوصًا», never «في 7 من 7 صف»",
              "كل صفوف بياناته (7 صفوف) تحمل نصوصًا" in tin and "من 7 صف" not in tin, tin)
        all_text = " ".join(a["detail"] for a in an)
        check("every digit count phrase of 01 agrees with its noun", not agreement_errors(all_text), agreement_errors(all_text))
        with open(os.path.join(out, "01_workbook_map.md"), encoding="utf-8") as fh:
            md = fh.read()
        check("01_workbook_map.md count phrases agree", not agreement_errors(md), agreement_errors(md))


# ---------------------------------------------------------------------------------------------------------------
def demo_records():
    return [
        {"id": "DEMO-1", "name": "عمارة تجريبية", "record_type": "asset", "scope": "in_scope",
         "dims": {"use": "investment", "income_nature": "rent"}, "source": {"source_id": "DEMO", "sheet": "سجل تجريبي", "row": 2},
         "fields": {"rev.2025.demo": tv(1250, ref="سجل تجريبي!C2"), "val.2024.demo": tv(25000, ref="سجل تجريبي!D2"),
                    "units": tv(4)}},
        {"id": "STM-r7", "name": "بند البيان التجريبي", "record_type": "asset", "scope": "isolated", "dims": {},
         "source": {"source_id": "STM", "sheet": "بيان تجريبي", "row": 7},
         "fields": {"rev.2025.stmt": tv(900, ref="بيان تجريبي!C7"), "rev.2025.demo": tv(None, "text", raw="انظر الملاحظة")}},
        {"id": "U-1.2", "name": "وحدة تجريبية", "record_type": "unit", "scope": "excluded", "dims": {},
         "source": {"source_id": "DEMO", "sheet": "سجل تجريبي", "row": 3}, "fields": {"units": tv(2)}},
    ]


def demo_config(review=None):
    spec = {"id": "rev_demo", "label": "الإيراد التجريبي 2025", "kind": "sum", "field": "{REV}", "unit": "SAR",
            "where": {"eq": ["record_type", "asset"]}}
    if review is not None:
        spec["review"] = review
    return {"contract_version": "1.0", "currency": "SAR",
            "periods": [{"id": "2025", "label": "2025", "rev_field": "rev.2025.demo", "exp_field": None, "prev_id": None,
                         "prev_rev_field": None}],
            "valuation": {"field": "val.2024.demo", "year": "2024"},
            "base_where": {"all": [{"eq": ["scope", "in_scope"]}]}, "metrics": [spec], "breakdowns": [],
            "trend": {"enabled": False}, "filters": ["use"]}


def test_review_flag():
    print("[C1] علامة «يحتاج تحقق» على المؤشر")
    recs = demo_records()
    schema = os.path.join(SKILL, "schemas", "metrics_config.schema.json")
    check("schema accepts a review object", not cm.schema_errors(demo_config(REVIEW), schema),
          cm.schema_errors(demo_config(REVIEW), schema))
    check("schema rejects an unknown review status",
          bool(cm.schema_errors(demo_config({"status": "approved", "reason": REVIEW["reason"]}), schema)))
    eng = me.Engine(demo_config(REVIEW), recs)
    r1 = eng.compute(demo_config(REVIEW)["metrics"][0], {"period": "2025"})
    r2 = eng.compute(demo_config(REVIEW)["metrics"][0], {"period": "2025", "use": "service"})
    check("engine copies review unchanged under any filter (also when the filter empties the metric)",
          r1.get("review") == REVIEW and r2.get("review") == REVIEW and r2["status"] == "empty", (r1.get("review"), r2))
    plain = eng.compute(demo_config()["metrics"][0], {"period": "2025"})
    check("no review key without a spec review", "review" not in plain, plain.keys())
    out = cm.compute({"records": recs, "field_meta": META}, demo_config(REVIEW))
    check("compute_metrics writes review into 05 metrics[id] and every parity case keeps the spec",
          out["metrics"]["rev_demo"].get("review") == REVIEW, out["metrics"]["rev_demo"].get("review"))
    with tempfile.TemporaryDirectory() as d:
        def run(review, met_review="same"):
            write(d, "03_records.json", {"contract_version": "1.0", "field_meta": META, "records": recs, "issues": []})
            write(d, "04_metrics_config.json", demo_config(review))
            res = me.Engine(demo_config(review), recs).compute_all({"period": "2025"})["metrics"]
            if met_review != "same":
                res["rev_demo"].pop("review", None)
            write(d, "05_metrics.json", {"metrics": res})
            return va.c_metric_review(va.Ctx(d))
        check("valid review passes metric_review", run(REVIEW) == [], run(REVIEW))
        bad_status = run({"status": "checked", "reason": REVIEW["reason"]})
        check("unknown status blocks", any("غير معروفة" in x[0] for x in bad_status), bad_status)
        short = run({"status": "needs_verification", "reason": "يحتاج مراجعة"})
        check("a two-word reason blocks (6–25 words)", any("كلمتين" in x[0] or "6 إلى 25" in x[0] for x in short), short)
        long_r = run({"status": "needs_verification", "reason": " ".join(["كلمة"] * 26)})
        check("a 26-word reason blocks", bool(long_r), long_r)
        digits = run({"status": "needs_verification", "reason": "رقم بديل قدره ٩٠٠ ريال في بيان آخر من الملف"})
        check("digits in the reason (Arabic-Indic too) block", any("أرقام" in x[0] for x in digits), digits)
        stale = run(REVIEW, met_review="drop")
        check("05 without the review of 04 blocks (re-run compute)", any("05" in x[0] for x in stale), stale)
        check("metric_review is registered as blocking",
              ("metric_review", "blocking") in [(c[0], c[1]) for c in va.CHECKS])
    check("apa_common.review_errors: valid → []; reason with digits → error",
          ac.review_errors(REVIEW) == [] and ac.review_errors({"status": "needs_verification",
                                                                  "reason": "الإيراد أقل بمقدار 30 بالمئة في بيان آخر هنا"}))


# ---------------------------------------------------------------------------------------------------------------
def fact_run(d, narrative):
    write(d, "03_records.json", {"contract_version": "1.0", "field_meta": META, "records": demo_records(), "issues": []})
    write(d, "06_narrative.json", narrative)


def test_record_facts():
    print("[C2] حقائق قيم السجلات rec.<record_id>.<field_id>")
    nar = {"executive_summary": "الإيراد في البيان الآخر {{f:rec.STM-r7.rev.2025.stmt}} مقابل {{f:rec.DEMO-1.rev.2025.demo}}.",
           "gaps": [{"id": "G1", "evidence": [{"ref": "f:rec.U-1.2.units", "kind": "fact"}]}]}
    with tempfile.TemporaryDirectory() as d:
        fact_run(d, nar)
        f = facts.compute_facts(d)
        a, b, c = (f.get("rec.STM-r7.rev.2025.stmt"), f.get("rec.DEMO-1.rev.2025.demo"), f.get("rec.U-1.2.units"))
        check("cited rec.* keys are facts (an isolated and an in-scope record; a token and an f: evidence ref)",
              a and a["value"] == 900 and b and b["value"] == 1250 and c and c["value"] == 2, (a, b, c))
        check("money fields have unit SAR, a count field unit number; a dotted record id resolves",
              a["unit"] == "SAR" and b["unit"] == "SAR" and c["unit"] == "number", (a and a["unit"], c and c["unit"]))
        check("uncited record values are not enumerated", not any(k.startswith("rec.DEMO-1.val") for k in f), list(f)[:20])
        rd = json.load(open(os.path.join(d, "03_records.json"), encoding="utf-8"))
        reasons = [facts.record_fact(rd, k)[1] for k in ("rec.NOPE-9.units", "rec.DEMO-1.rev.2023.demo",
                                                         "rec.STM-r7.rev.2025.demo", "rec.DEMO-1.val.2024.demo")]
        check("record_fact reasons: unknown record, unknown field, not numeric; a known value resolves",
              reasons == ["unknown_record", "unknown_field", "not_numeric", None], reasons)
        tr = bd.TokenRenderer({}, {}, "SAR", facts=f)
        txt = tr.render(nar["executive_summary"])
        check("the build renders rec facts as money («900 ر.س»، «1,250 ر.س»)",
              "900 ر.س" in txt and "1,250 ر.س" in txt and not tr.unknown_facts, txt)
        ctx = va.Ctx(d)
        check("fact_tokens passes for resolvable rec tokens", va.c_fact_tokens(ctx) == [], va.c_fact_tokens(ctx))
        ok, note = ctx.resolver().resolve("f:rec.U-1.2.units")
        bad_ok, bad_note = ctx.resolver().resolve("f:rec.STM-r7.rev.2025.demo")
        check("f: evidence refs resolve through the same rule (and a text value does not)",
              ok and not bad_ok and "ليست رقمًا" in (bad_note or ""), (ok, note, bad_ok, bad_note))
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            facts.main(["--run-dir", d])
        lines = buf.getvalue().splitlines()
        rec_lines = [ln for ln in lines if ln.startswith("{{f:rec.")]
        check("facts.py --run-dir prints the rec.* pattern once, with one example",
              len(rec_lines) == 1 and rec_lines[0].startswith("{{f:rec.<record_id>.<field_id>}}") and "مثال: {{f:rec." in rec_lines[0],
              rec_lines)
    for bad_key, frag in (("rec.NOPE-9.units", "لا يوجد سجل"), ("rec.DEMO-1.rev.2023.demo", "الحقل غير موجود"),
                          ("rec.STM-r7.rev.2025.demo", "ليست رقمًا")):
        with tempfile.TemporaryDirectory() as d:
            fact_run(d, {"executive_summary": f"قيمة {{{{f:{bad_key}}}}}."})
            found = va.c_fact_tokens(va.Ctx(d))
            check(f"fact_tokens blocks {bad_key.split('.')[1]} ({frag})",
                  len(found) == 1 and frag in found[0][0] and "scripts/facts.py --run-dir RUN" not in found[0][0], found)
    with tempfile.TemporaryDirectory() as d:
        fact_run(d, {"executive_summary": "قيمة {{f:nope.key}}."})
        found = va.c_fact_tokens(va.Ctx(d))
        hint_path = os.path.join(os.path.realpath(SKILL), "scripts", "facts.py")
        check("an unknown plain fact key hints an absolute, quoted facts.py path (no bare scripts/ path)",
              found and f'python3 "{hint_path}" --run-dir "{d}"' in found[0][0], found)


# ---------------------------------------------------------------------------------------------------------------
def rec_item(rid, p, **rank):
    r = {"id": rid, "priority": p, "ranking": dict({"impact": "medium", "urgency": "medium", "evidence_strength": "moderate",
                                                   "feasibility": "medium", "cost_level": "medium"}, **rank)}
    return r


def prio_findings(recs, top):
    with tempfile.TemporaryDirectory() as d:
        write(d, "06_narrative.json", {"recommendations": recs, "top_actions": top})
        ctx = va.Ctx(d)
        return va.c_top_actions_priority(ctx), va.c_nq_recommendation_quality(ctx)


def test_top_actions_priority():
    print("[C4] ترتيب أهم الإجراءات بالأولوية")
    recs = [rec_item("R1", 1), rec_item("R2", 2), rec_item("R3", 3), rec_item("R4", 4)]
    ok, _ = prio_findings(recs, ["R1", "R2", "R3"])
    check("ascending top_actions with unique priorities pass", ok == [], ok)
    swapped, _ = prio_findings(recs, ["R1", "R3", "R2"])
    check("a swapped order blocks and names the ascending order", len(swapped) == 1 and "R1، R2، R3" in swapped[0][0], swapped)
    dup, _ = prio_findings([rec_item("R1", 1), rec_item("R2", 2), rec_item("R3", 2)], ["R1", "R2", "R3"])
    check("a duplicated priority blocks", any("مكررة" in x[0] for x in dup), dup)
    nonint, _ = prio_findings([rec_item("R1", 1), rec_item("R2", "2")], ["R1", "R2"])
    check("a non-integer priority blocks", any("ليست عددًا صحيحًا" in x[0] for x in nonint), nonint)
    skip_block, skip_warn = prio_findings(recs, ["R1", "R3", "R4"])
    check("skipping a lower priority number is ascending (no block) but warns in recommendation_quality",
          skip_block == [] and any("R2" in x[0] and "أصغر أرقام الأولوية" in x[0] for x in skip_warn), (skip_block, skip_warn))
    check("top_actions_priority is registered as blocking",
          ("top_actions_priority", "blocking") in [(c[0], c[1]) for c in va.CHECKS])
    syn = json.load(open(os.path.join(HERE, "synthetic", "06_narrative.json"), encoding="utf-8"))
    pr = {r["id"]: r["priority"] for r in syn["recommendations"]}
    check("the synthetic narrative lists top_actions in ascending priority",
          [pr[x] for x in syn["top_actions"]] == sorted(pr[x] for x in syn["top_actions"]), syn["top_actions"])


# ---------------------------------------------------------------------------------------------------------------
def test_reconciliation_difference():
    print("[rec] جدول المطابقة في 07: الفرق مقابل المجموع المقارن")
    with tempfile.TemporaryDirectory() as d:
        recs = [{"id": "A-1", "name": "أصل تجريبي", "record_type": "asset", "scope": "in_scope", "dims": {},
                 "source": {"source_id": "A", "sheet": "سجل تجريبي", "row": 2},
                 "fields": {"val.2024.demo": tv(800, "derived"), "val.2024.demo__original": tv(1000)}},
                {"id": "A-2", "name": "أصل تجريبي ثان", "record_type": "asset", "scope": "in_scope", "dims": {},
                 "source": {"source_id": "A", "sheet": "سجل تجريبي", "row": 3},
                 "fields": {"rev.2025.demo": tv(300), "val.2024.demo": tv(None)}},
                {"id": "A-3", "name": "أصل تجريبي ثالث", "record_type": "asset", "scope": "in_scope", "dims": {},
                 "source": {"source_id": "A", "sheet": "سجل تجريبي", "row": 4},
                 "fields": {"rev.2025.demo": dict(tv(200), raw="200", note="نص رقمي حُوِّل")}}]
        exc = [{"source_id": "A", "sheet": "سجل تجريبي", "row": 5, "reason": "صف إجمالي",
                "values": {"val.2024.demo": tv(1000, ref="سجل تجريبي!D5"), "rev.2025.demo": tv(300, ref="سجل تجريبي!C5")}}]
        write(d, "03_records.json", {"contract_version": "1.0", "field_meta": META, "records": recs, "excluded_rows": exc,
                                     "issues": []})
        rows, warns = va.g4b_reconciliation(va.Ctx(d))
        by = {r["field"]: r for r in rows}
        v = by["val.2024.demo"]
        check("match_original: difference = recorded − compared (in-scope) sum = 200, matched sum 1000 apart",
              v["status"] == "match_original" and v["compared_sum"] == 800 and v["difference"] == 200
              and v["matched_sum"] == 1000 and v["difference_basis"] == "in_scope", v)
        t = by["rev.2025.demo"]
        check("match_without_text_numbers: the difference is the numbers stored as text (−200), matched 300",
              t["status"] == "match_without_text_numbers" and t["compared_sum"] == 500 and t["difference"] == -200
              and t["matched_sum"] == 300, t)
        md = "\n".join(va.g4b_md_reconciliation({"reconciliation": rows}))
        line = next(ln for ln in md.splitlines() if "val.2024.demo" in ln)
        check("07.md: «المجموع المقارن» and «أساس المقارنة» columns; the row shows 800 and 200, not 0 beside 1,000/800",
              "المجموع المقارن" in md and "أساس المقارنة" in md and "| 800 | السجلات ضمن النطاق | 200 | 1,000 |" in line, line)
        check("no warning for explained differences", warns == [], warns)


# ---------------------------------------------------------------------------------------------------------------
def test_docs_and_labels():
    print("[doc] التوثيق والتسميات")
    docs = [os.path.join(SKILL, "SKILL.md")] + [os.path.join(SKILL, "references", f)
                                               for f in sorted(os.listdir(os.path.join(SKILL, "references")))
                                               if f.endswith(".md") and not f.startswith("00-")]
    unquoted, bare = [], []
    for p in docs:
        for i, line in enumerate(open(p, encoding="utf-8"), 1):
            if re.search(r'(?<!")\$SKILL_DIR/', line):
                unquoted.append(f"{os.path.basename(p)}:{i}")
            if re.search(r"(?:python3|node|bash)(?:\s+-I)?\s+(?:scripts|tests)/", line):
                bare.append(f"{os.path.basename(p)}:{i}")
    check("every $SKILL_DIR path in SKILL.md and references is quoted", not unquoted, unquoted)
    check("no bare «python3 scripts/…» command in SKILL.md and references", not bare, bare)
    skill = open(os.path.join(SKILL, "SKILL.md"), encoding="utf-8").read()
    tpl = skill.split("## قالب رسالة التسليم", 1)[1].split("## عند التعذر", 1)[0]
    check("09 template: one «المراجعة» sentence, no «نتيجة الفحوص» QA paragraph, «المخرجات» file names only",
          "نتيجة الفحوص" not in tpl and "**المراجعة:** روجعت الأرقام آليًا ويدويًا" in tpl and "تنبيهات معلَّلة" in tpl
          and "**المخرجات:** dashboard.html · supporting_workbook.xlsx\n" in tpl, tpl[:400])
    check("SKILL.md setup uses python3 -m pip with a venv fallback",
          "python3 -m pip install -r \"$SKILL_DIR/requirements.txt\"" in skill and "python3 -m venv .venv" in skill)
    check("SKILL.md rule 12 documents {{f:rec.<record_id>.<field_id>}} and step 6 the review flag",
          "{{f:rec.<record_id>.<field_id>}}" in skill and "\"review\": {\"status\": \"needs_verification\"" in skill)
    vocab = ac.load_vocab()
    check("severity labels are feminine (agree with «الخطورة»/«الأهمية»)",
          vocab["severity"] == {"high": "مرتفعة", "medium": "متوسطة", "low": "منخفضة"}, vocab["severity"])
    texts = cm.metric_texts("zero_rev_investment_value", {"id": "2025", "rev_field": "rev.2025.demo"},
                            {"field": "val.2024.demo", "year": "2024"}, META, "ملف تجريبي")
    check("zero-revenue card label «قيمة الأصول الاستثمارية بإيراد صفري»",
          texts["short_label"] == "قيمة الأصول الاستثمارية بإيراد صفري", texts["short_label"])


def main():
    for t in (test_ar_count, test_inspect_texts, test_review_flag, test_record_facts, test_top_actions_priority,
              test_reconciliation_difference, test_docs_and_labels):
        try:
            t()
        except Exception as exc:  # a crash is a failure, never a pass
            import traceback
            traceback.print_exc()
            check(f"{t.__name__} ran without an exception", False, f"{type(exc).__name__}: {exc}")
    print(f"\nناجح: {len(PASSES)} · فاشل: {len(FAILS)}")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
