#!/usr/bin/env python3
"""Regression tests for scripts/normalize.py on small FICTIONAL workbooks (no real data).

    python3 -I tests/test_normalize.py            # exit 0 = all pass, 1 = a failure

Covers: totals never become records (A3/A14), blank and skipped rows (A4/A8), cached 0 of blank inputs (A5),
controlled vocabulary (A6), joins through confirmed links (A7), merged name/id cells (A8), sheet decisions /
hidden / managed sheets (A9), overlapping sources (A10), required context copied to 03 (A13), percent text in
amounts (A16), «غير متاح» tokens in text columns (A19), scope-aware impact and ordering (C12), PII name columns;
round 2: named rows with weak total signs kept (possible_total_row), totals after data_rows kept for reconciliation,
ambiguous «12.500», foreign currency, area units, IBAN / labelled national ID masked as spans, percent scale mixed,
joins on a deed shared with excluded units, rollups of unit revenue, section labels.
"""
import copy
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.environ.get("APA_SKILL_DIR") or os.path.dirname(HERE)
NORMALIZE = os.path.join(SKILL, "scripts", "normalize.py")
sys.dont_write_bytecode = True

FAILS, PASSES = [], []


def check(name, cond, info=""):
    (PASSES if cond else FAILS).append(name)
    print(("  PASS  " if cond else "  FAIL  ") + name + ("" if cond else f" — {info}"))


# ---------------------------------------------------------------------------------------------------------------
# Fixture
# ---------------------------------------------------------------------------------------------------------------
def build_workbook(path):
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "سجل"
    ws.append(["م", "اسم الأصل", "الإيجار الشهري", "الإيراد 2024", "التقييم 2024", "الاستخدام", "الحي", "ملاحظات"])
    rows = {
        2: [1, "عمارة الريحان", 10000, 120000, 3000000, "investment", "النخيل", "إجمالي المساحة 500 م2"],
        3: [2, "عمارة السدرة", 8000, 96000, 2000000, "investment", "—", None],
        4: [None, "جملة", None, 216000, 5000000, None, None, None],
        5: [None, None, None, None, None, None, None, None],
        6: [" ", " ", None, " ", None, None, None, None],
        7: [5, None, None, 50000, None, None, None, None],          # name blank, revenue present
        8: [4, "مجمع الواحة", None, 300000, 9000000, "investment", None, None],
        9: [None, None, None, 200000, None, None, None, "مكاتب"],    # inside merged A8:A9 / B8:B9
        10: [6, "عمارة النرجس", None, "=C10*12", 4000000, "investment", None, None],   # C10 blank, cached 0
        11: [7, "محلات الورد", None, "100%", 1500000, "investment", None, None],
        12: [8, "أرض الياسمين", None, 0, 2500000, "استثماري", None, None],
        13: [9, "عمارة الفل", 0, "=C13*12", 1000000, "investment", None, None],       # C13 = 0 -> real zero
        14: [10, "مبنى الجوري", None, 70000, 1200000, "service", None, "Total"],      # label-like total in notes
    }
    for r, vals in rows.items():
        for c, v in enumerate(vals, 1):
            ws.cell(row=r, column=c, value=v)
    ws.merge_cells("A8:A9")
    ws.merge_cells("B8:B9")

    ex = wb.create_sheet("مصروفات")
    ex.sheet_state = "hidden"
    ex.append(["م", "الاسم", "مصروفات 2024", "الإيراد 2024"])
    ex.append([1, "عمارة الريحان", 15000, 999999])     # conflict on rev.2024 (REG has 120000)
    ex.append([99, "غير موجود", 4000, None])           # unmatched
    ex.append([4, "مجمع الواحة", 7000, None])          # ambiguous: two REG records carry id «4» (merged rows)

    mg = wb.create_sheet("مدار")
    mg.sheet_state = "hidden"
    mg.append(["م", "العقار", "اسم المالك", "الإيراد 2024", "التقييم 2024"])
    mg.append(["M1", "عمارة الغير", "مالك وهمي", 80000, 1000000])

    iso = wb.create_sheet("قديمة")
    iso.sheet_state = "hidden"
    iso.append(["م", "العقار", "الإيراد 2024"])
    iso.append([1, "عمارة الريحان", 110000])
    wb.save(path)
    # openpyxl writes formulas without cached values: inject Excel's cached 0 for D10 and D13.
    inject_cached(path, "xl/worksheets/sheet1.xml", {"D10": "0", "D13": "0"})


def inject_cached(path, member, cells):
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == member:
                xml = data.decode("utf-8")
                for coord, val in cells.items():
                    pat = re.compile(r'(<c r="%s"[^>]*>)(<f>[^<]*</f>)(<v\s*/>|<v></v>)?' % coord)
                    xml, n = pat.subn(lambda m: m.group(1) + m.group(2) + f"<v>{val}</v>", xml)
                    assert n == 1, f"cell {coord} not found in {member}"
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)


def base_mapping():
    meta_rev = {"label": "الإيراد 2024", "group": "revenue", "period": "2024", "vat": "unknown", "basis": "recorded"}
    return {
        "contract_version": "1.0",
        "source_file": "fixture.xlsx",
        "entity": {"name": None, "evidence": "اسم الجهة غير مذكور في الملف"},
        "data_period": {"label": "2024", "evidence": "سجل!D1"},
        "study_date": {"value": None, "basis": None, "evidence": "غير مذكور"},
        "currency": "SAR",
        "valuation_dates": [{"year": "2024", "basis": "غير موضح", "evidence": "سجل!E1"}],
        "sheets": [
            {"sheet": "سجل", "role": "asset_register", "decision": "primary", "reason": "السجل الرئيسي"},
            {"sheet": "مصروفات", "role": "other", "decision": "supplementary", "reason": "مطابقة بالمعرّف"},
            {"sheet": "مدار", "role": "managed_properties", "decision": "supplementary", "reason": "عقارات للغير"},
            {"sheet": "قديمة", "role": "other", "decision": "isolated", "reason": "نطاق غير متحقق"},
        ],
        "sources": [{
            "id": "REG", "sheet": "سجل", "header_row": 1, "data_rows": {"from": 2, "to": 14},
            "skip_if_blank": "B", "default_record_type": "asset", "default_scope": "in_scope",
            "id_rule": {"column": "A", "prefix": "REG-"}, "name_column": "B",
            "columns": {"C": {"field": "rent_monthly", "kind": "money"},
                        "D": {"field": "rev.2024.recorded", "kind": "money"},
                        "E": {"field": "val.2024.total", "kind": "money"},
                        "F": {"field": "use_text", "kind": "text", "dim": "use"},
                        "G": {"field": "district", "kind": "text", "dim": "district"},
                        "H": {"field": "notes", "kind": "text"}},
        }],
        "field_meta": {"rev.2024.recorded": meta_rev,
                       "val.2024.total": {"label": "التقييم 2024", "group": "valuation", "year": "2024"},
                       "exp.2024.opex": {"label": "مصروفات 2024", "group": "expense", "period": "2024"}},
        "records": {},
        "links": [{"id": "L-EXP", "status": "confirmed", "basis": "المعرّف م مشترك", "used": True, "reason": "اختبار"}],
    }


def build_round2(path):
    """Units sharing their parent's deed, a parent row summing them, a section label, a total after the data rows,
    ambiguous / foreign / unit-suffixed numbers and personal data inside notes."""
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "سجل"
    ws.append(["م", "اسم العقار", "رقم الصك", "الإيراد 2024", "الإيراد 2023", "التقييم 2024", "المساحة", "الحصة", "ملاحظات",
               "أساس الإيجار"])
    rows = [
        [1, "عمارة الفيصلية", "410112003344", "=SUM(D3:D5)", None, 15000000, 900, 100, None, "إجمالي"],
        [2, "عمارة الفيصلية - شقة 1", "410112003344", 100000, 290000, None, None, None, "هوية المستأجر 1023456789", "إجمالي"],
        [3, "عمارة الفيصلية - شقة 2", "410112003344", 100000, 290000, None, None, None, "الحوالة على SA0380000000608010167519", "صافي"],
        [4, "عمارة الفيصلية - معرض", "410112003344", 100000, 290000, None, None, None, None, "إجمالي"],
        ["Jeddah", None, None, None, None, None, None, None, None, None],
        [5, "أرض النسيم", "410112005566", "12.500", None, 4000000, "٨٠٠ م²", 0.5, None, "صافي"],
        [6, "محلات الوقف", "410112007788", "USD 100,000", 50000, 3000000, 1200, 50, "قيمة الأرض 500 000 000 ريال", "إجمالي"],
    ]
    for i, vals in enumerate(rows, start=2):
        for c, v in enumerate(vals, 1):
            ws.cell(row=i, column=c, value=v)
    ws.cell(row=9, column=2, value="الإجمالي")
    ws.cell(row=9, column=4, value=400000)
    ex = wb.create_sheet("مصروفات")
    ex.append(["رقم الصك", "العقار", "مصروفات 2024"])
    ex.append(["410112003344", "عمارة الفيصلية", 60000])
    wb.save(path)
    inject_cached(path, "xl/worksheets/sheet1.xml", {"D2": 300000})


def round2_mapping():
    m = base_mapping()
    m["sheets"] = [{"sheet": "سجل", "role": "asset_register", "decision": "primary", "reason": "السجل"},
                   {"sheet": "مصروفات", "role": "other", "decision": "supplementary", "reason": "مصروفات بالصك"}]
    m["sources"] = [{
        "id": "W", "sheet": "سجل", "header_row": 1, "data_rows": {"from": 2, "to": 8}, "skip_if_blank": "B",
        "default_record_type": "asset", "default_scope": "in_scope", "id_rule": {"column": "A", "prefix": "W-"},
        "name_column": "B",
        "columns": {"C": {"field": "deed_no", "kind": "text"}, "D": {"field": "rev.2024.recorded", "kind": "money"},
                    "E": {"field": "rev.2023.recorded", "kind": "money"}, "F": {"field": "val.2024.total", "kind": "money"},
                    "G": {"field": "area", "kind": "area"}, "H": {"field": "own_share", "kind": "percent"},
                    "I": {"field": "notes", "kind": "text"}, "J": {"field": "rent_basis", "kind": "text"}}},
        {"id": "EXP", "sheet": "مصروفات", "header_row": 1, "data_rows": {"from": 2, "to": 2}, "skip_if_blank": "B",
         "default_scope": "in_scope", "scope_reason": "مصروفات مربوطة برقم الصك عبر L-EXP",
         "id_rule": {"column": "A", "prefix": "EXP-"}, "name_column": "B",
         "columns": {"A": {"field": "deed_key", "kind": "text"}, "C": {"field": "exp.2024.opex", "kind": "money"}},
         "join": {"into": "W", "on": "deed", "key_column": "A", "link_id": "L-EXP"}}]
    m["field_meta"]["rev.2023.recorded"] = {"label": "الإيراد 2023", "group": "revenue", "period": "2023"}
    m["records"] = {k: {"record_type": "unit"} for k in ("W-2", "W-3", "W-4")}
    m["records"]["W-1"] = {"rollup": {"from": ["W-2", "W-3", "W-4"], "fields": ["rev.2023.recorded"],
                                      "evidence": "إيراد 2023 مسجل على الوحدات وحدها (E3:E5)"}}
    return m


def run_normalize(tmp, xlsx, mapping, name, wb_map=None):
    d = os.path.join(tmp, name)
    os.makedirs(d, exist_ok=True)
    mp = os.path.join(d, "02_mapping.json")
    json.dump(mapping, open(mp, "w", encoding="utf-8"), ensure_ascii=False)
    if wb_map is not None:
        json.dump(wb_map, open(os.path.join(d, "01_workbook_map.json"), "w", encoding="utf-8"), ensure_ascii=False)
    out = os.path.join(d, "03_records.json")
    p = subprocess.run([sys.executable, "-I", NORMALIZE, "--xlsx", xlsx, "--mapping", mp, "--out", out],
                       capture_output=True, text=True)
    doc = json.load(open(out, encoding="utf-8")) if p.returncode == 0 else None
    return p.returncode, p.stdout + p.stderr, doc


def issues_of(doc, typ):
    return [i for i in doc["issues"] if i["type"] == typ]


def rec(doc, rid):
    return next((r for r in doc["records"] if r["id"] == rid), None)


# ---------------------------------------------------------------------------------------------------------------
def main():
    tmp = tempfile.mkdtemp(prefix="apa_norm_test_")
    try:
        xlsx = os.path.join(tmp, "fixture.xlsx")
        build_workbook(xlsx)

        # ---- main run: REG only --------------------------------------------------------------------------
        print("[1] السجل الرئيسي")
        rc, log, d = run_normalize(tmp, xlsx, base_mapping(), "main")
        check("normalize exits 0", rc == 0, log[-400:])
        if d is None:
            return finish(tmp)
        ids = [r["id"] for r in d["records"]]
        # A3/A14: «جملة» (name) and «Total» (notes) rows are not records; kept in excluded_rows with their values
        check("A3 total row (name «جملة») is not a record", all(r["source"]["row"] != 4 for r in d["records"]), ids)
        # a named row with an id («10 مبنى الجوري») and «Total» in its notes is kept, flagged for review (round 2)
        r14 = next((r for r in d["records"] if r["source"]["row"] == 14), None)
        pt = issues_of(d, "possible_total_row")
        check("A3 named row with a total word in another cell is kept (possible_total_row, medium)",
              r14 is not None and any(i["record_ids"] == [r14["id"]] and i["severity"] == "medium" for i in pt), pt)
        check("A3 long note «إجمالي المساحة 500 م2» does not make a total", rec(d, "REG-1") is not None, ids)
        ex4 = next((e for e in d["excluded_rows"] if e["row"] == 4), {})
        check("A14 excluded total row keeps numeric values",
              (ex4.get("values") or {}).get("rev.2024.recorded", {}).get("v") == 216000 and
              (ex4.get("values") or {}).get("val.2024.total", {}).get("v") == 5000000, ex4)
        tl = issues_of(d, "total_like_row")
        check("A3 total_like_row issues are high", len(tl) == 1 and all(i["severity"] == "high" for i in tl), tl)
        # A4: blank and filler rows skipped and reported
        b = issues_of(d, "blank_rows_skipped")
        check("A4 blank/filler rows reported once (low, count 2)",
              len(b) == 1 and b[0]["severity"] == "low" and b[0].get("count") == 2, b)
        check("A4 no record from blank rows", all(r["source"]["row"] not in (5, 6) for r in d["records"]), ids)
        # A8: row 7 skipped by skip_if_blank but holds revenue -> reported + excluded_rows
        sk = issues_of(d, "row_skipped_with_content")
        check("A8 row skipped with content reported (medium)", len(sk) == 1 and sk[0]["severity"] == "medium"
              and sk[0].get("rows") == [7], sk)
        check("A8 skipped row listed in excluded_rows", any(e["row"] == 7 and "skip_if_blank" in e["reason"]
                                                              for e in d["excluded_rows"]), d["excluded_rows"])
        # A8: merged A8:A9/B8:B9 -> row 9 reads the anchor id/name with a merge note
        r9 = next((r for r in d["records"] if r["source"]["row"] == 9), None)
        check("A8 merged continuation row becomes a record named from the anchor",
              r9 is not None and r9["name"] == "مجمع الواحة" and "B8:B9" in (r9["source"].get("merge_note") or ""), r9)
        check("A8 merged continuation keeps its own revenue", r9 is not None and
              r9["fields"]["rev.2024.recorded"]["v"] == 200000, r9 and r9["fields"].get("rev.2024.recorded"))
        check("A8 merged duplicate id reported", any("REG-4" in i["record_ids"] for i in issues_of(d, "duplicate_id")),
              issues_of(d, "duplicate_id"))
        # A5: =C10*12 with C10 blank (cached 0) -> not_available; =C13*12 with C13 = 0 stays a recorded zero
        f10 = rec(d, "REG-6")["fields"]["rev.2024.recorded"]
        check("A5 cached 0 from blank inputs -> not_available", f10["s"] == "not_available" and f10["v"] is None, f10)
        f13 = rec(d, "REG-9")["fields"]["rev.2024.recorded"]
        check("A5 cached 0 from a written 0 stays zero_recorded", f13["s"] == "zero_recorded" and f13["v"] == 0, f13)
        z = issues_of(d, "zero_from_blank_inputs")
        check("A5 zero_from_blank_inputs (medium) lists REG-6 only",
              len(z) == 1 and z[0]["severity"] == "medium" and z[0]["record_ids"] == ["REG-6"], z)
        # A16: «100%» in a money column
        f11 = rec(d, "REG-7")["fields"]["rev.2024.recorded"]
        check("A16 percent text in money -> text, no number", f11["s"] == "text" and f11["v"] is None, f11)
        check("A16 percent_in_amount issue (medium)", [i["severity"] for i in issues_of(d, "percent_in_amount")] == ["medium"],
              issues_of(d, "percent_in_amount"))
        # A6: «استثماري» in a direct use column -> unknown + high issue; vocab code kept
        r8 = rec(d, "REG-8")
        check("A6 out-of-vocabulary dim stored as unknown", r8["dims"]["use"] == "unknown", r8["dims"])
        check("A6 vocab code kept", rec(d, "REG-1")["dims"]["use"] == "investment", rec(d, "REG-1")["dims"])
        oov = issues_of(d, "dim_out_of_vocabulary")
        check("A6 dim_out_of_vocabulary (high) cites the cell", len(oov) == 1 and oov[0]["severity"] == "high"
              and "سجل!F12" in oov[0]["where"] and oov[0]["record_ids"] == ["REG-8"], oov)
        # A19: «—» in a text dim column -> not_available, not a category
        f3 = rec(d, "REG-2")["fields"]["district"]
        check("A19 NA token in text column -> not_available", f3["s"] == "not_available" and f3["v"] is None, f3)
        check("A19 NA token is not a dim category", rec(d, "REG-2")["dims"].get("district") is None, rec(d, "REG-2")["dims"])
        # A13: context copied
        check("A13 context copied to 03", set(d.get("context") or {}) >= {"entity", "data_period", "study_date",
                                                                           "currency", "valuation_dates", "sheets"},
              list(d.get("context") or {}))
        # ids are sequential after sorting
        check("issue ids are DQ-001… in order", [i["id"] for i in d["issues"]] ==
              [f"DQ-{k:03d}" for k in range(1, len(d["issues"]) + 1)], [i["id"] for i in d["issues"]])

        # ---- overrides: a total-like row kept by an explicit override ---------------------------------------
        print("[2] تجاوز صريح لصف يشبه الإجمالي")
        m = base_mapping()
        m["records"]["REG-r4"] = {"record_type": "asset", "scope": "in_scope", "scope_reason": "ليس إجماليًا: مبنى قائم"}
        rc, log, d2 = run_normalize(tmp, xlsx, m, "override")
        tl = issues_of(d2, "total_like_row") if d2 else []
        check("A3 overridden total-like row kept as record", d2 is not None and rec(d2, "REG-r4") is not None, log[-300:])
        check("A3 override still reported (mapping_override)", any(i.get("mapping_override") and i["record_ids"] == ["REG-r4"]
                                                                    for i in tl), tl)

        # ---- 01 row kinds -------------------------------------------------------------------------------------
        print("[3] أنواع الصفوف من 01_workbook_map.json")
        from hashlib import sha256
        sha = sha256(open(xlsx, "rb").read()).hexdigest()
        wm = {"file": {"sha256": sha}, "sheets": [{"name": "سجل", "row_kinds": {"2": "data", "3": "total_label"}}]}
        rc, log, d3 = run_normalize(tmp, xlsx, base_mapping(), "rowkinds", wb_map=wm)
        check("A3 row kind total_label in 01 on a named row -> kept with possible_total_row",
              d3 is not None and rec(d3, "REG-2") is not None and
              any(i["record_ids"] == ["REG-2"] and "01" in i["detail"] for i in issues_of(d3, "possible_total_row")),
              d3 and [r["id"] for r in d3["records"]])
        wm_bad = copy.deepcopy(wm)
        wm_bad["file"]["sha256"] = "0" * 64
        rc, log, d3b = run_normalize(tmp, xlsx, base_mapping(), "rowkinds_other_file", wb_map=wm_bad)
        check("A3 01 of another file is ignored (with a warning)", d3b is not None and rec(d3b, "REG-2") is not None
              and "لا يخص الملف" in log, log[-300:])

        # ---- join, managed, isolated, overlap ----------------------------------------------------------------
        print("[4] الربط والأوراق المخفية والتداخل")
        m = base_mapping()
        m["sources"].append({
            "id": "EXP", "sheet": "مصروفات", "header_row": 1, "data_rows": {"from": 2, "to": 4},
            "default_scope": "in_scope", "scope_reason": "مصروفات مطابقة بالمعرّف عبر الرابط L-EXP",
            "id_rule": {"column": "A", "prefix": "EXP-"}, "name_column": "B",
            "columns": {"C": {"field": "exp.2024.opex", "kind": "money"},
                        "D": {"field": "rev.2024.recorded", "kind": "money"}},
            "join": {"into": "REG", "on": "id", "key_column": "A", "link_id": "L-EXP"}})
        m["sources"].append({
            "id": "MAN", "sheet": "مدار", "header_row": 1, "data_rows": {"from": 2, "to": 2},
            "default_scope": "in_scope", "scope_reason": "عقار للغير يعرض منفصلًا",
            "id_rule": {"column": "A", "prefix": "MAN-"}, "name_column": "C",
            "columns": {"D": {"field": "rev.2024.recorded", "kind": "money"},
                        "E": {"field": "val.2024.total", "kind": "money"}}})
        m["sources"].append({
            "id": "OLD", "sheet": "قديمة", "header_row": 1, "data_rows": {"from": 2, "to": 2},
            "id_rule": {"column": "A", "prefix": "OLD-"}, "name_column": "B",
            "columns": {"C": {"field": "rev.2024.recorded", "kind": "money"}}})
        m["sources"].append({
            "id": "DUP", "sheet": "سجل", "header_row": 1, "data_rows": {"from": 2, "to": 3}, "default_scope": "in_scope",
            "id_rule": {"column": "A", "prefix": "DUP-"}, "name_column": "B",
            "columns": {"D": {"field": "rev.2024.recorded", "kind": "money"}}})
        rc, log, d4 = run_normalize(tmp, xlsx, m, "join")
        check("join run exits 0", rc == 0, log[-500:])
        if d4 is not None:
            r1 = rec(d4, "REG-1")
            check("A7 joined field merged with its own ref and joined_from",
                  r1["fields"].get("exp.2024.opex", {}).get("v") == 15000 and
                  r1["fields"]["exp.2024.opex"].get("joined_from") == "EXP" and
                  r1["fields"]["exp.2024.opex"].get("ref") == "مصروفات!C2", r1["fields"].get("exp.2024.opex"))
            check("A7 conflicting field keeps the target value", r1["fields"]["rev.2024.recorded"]["v"] == 120000,
                  r1["fields"]["rev.2024.recorded"])
            check("A7 join_conflict (medium)", [i["severity"] for i in issues_of(d4, "join_conflict")] == ["medium"],
                  issues_of(d4, "join_conflict"))
            un = issues_of(d4, "join_unmatched")
            iso = [r for r in d4["records"] if r["source"]["source_id"] == "EXP"]
            check("A7 unmatched/ambiguous rows become isolated records only",
                  len(iso) == 2 and all(r["scope"] == "isolated" for r in iso), [(r["id"], r["scope"]) for r in iso])
            check("A7 join_unmatched (medium)", len(un) == 1 and un[0]["severity"] == "medium", un)
            amb = issues_of(d4, "join_ambiguous")
            check("A7 join_ambiguous (high) lists both targets", len(amb) == 1 and amb[0]["severity"] == "high"
                  and {"REG-4", "REG-4-r9"} <= set(amb[0]["record_ids"]), amb)
            check("A7 no in-scope record created by the join source",
                  not any(r["source"]["source_id"] == "EXP" and r["scope"] == "in_scope" for r in d4["records"]))
            # A9
            mu = issues_of(d4, "managed_unflagged")
            check("A9 managed_unflagged (high) for MAN-M1", len(mu) == 1 and mu[0]["record_ids"] == ["MAN-M1"], mu)
            old = rec(d4, "OLD-1")
            check("A9 hidden isolated sheet without default_scope -> isolated", old is not None and old["scope"] == "isolated",
                  old and old["scope"])
            ssm = [i for i in issues_of(d4, "sheet_scope_mismatch") if i.get("source_id") == "OLD"]
            check("A9 sheet_scope_mismatch (high) for the hidden source", len(ssm) == 1 and ssm[0]["severity"] == "high", ssm)
            # PII name column («اسم المالك») never becomes a record name
            check("PII name column not stored as record name", rec(d4, "MAN-M1")["name"] is None, rec(d4, "MAN-M1")["name"])
            check("PII name column reported", any("C1" in i["where"] for i in issues_of(d4, "pii_header_not_marked")),
                  issues_of(d4, "pii_header_not_marked"))
            # A10
            ov = issues_of(d4, "overlapping_source_rows")
            check("A10 overlapping_source_rows (high) lists both record ids",
                  len(ov) == 1 and ov[0]["severity"] == "high" and {"REG-1", "DUP-1", "REG-2", "DUP-2"} <= set(ov[0]["record_ids"]),
                  ov)
            # C12: isolated-only issues say so and sort after in-scope issues
            iso_only = [i for i in d4["issues"] if i["record_ids"] and
                        all(rec(d4, x)["scope"] != "in_scope" for x in i["record_ids"])]
            check("C12 isolated-only issues: impact «لا أثر على المؤشرات»",
                  iso_only and all(i["impact"].startswith("لا أثر على المؤشرات (") for i in iso_only),
                  [(i["type"], i["impact"]) for i in iso_only])
            first_iso = min(d4["issues"].index(i) for i in iso_only) if iso_only else 0
            check("C12 isolated-only issues come after in-scope issues",
                  all(d4["issues"].index(i) >= first_iso for i in iso_only) and
                  not any(i in iso_only for i in d4["issues"][:first_iso]) and
                  all(i not in iso_only for i in d4["issues"][:first_iso]),
                  [(i["id"], i["type"]) for i in d4["issues"]])
            sev = {"high": 0, "medium": 1, "low": 2}
            head = [sev[i["severity"]] for i in d4["issues"][:first_iso]]
            check("C12 severity sorted within the in-scope group", head == sorted(head), head)

        # ---- round 2: totals after data_rows, parsing, PII spans, joins on a shared deed, rollups ------------------
        print("[6] الإجماليات بعد النطاق والأرقام الملتبسة والربط بالصك والتجميع")
        x2 = os.path.join(tmp, "round2.xlsx")
        build_round2(x2)
        rc, log, d6 = run_normalize(tmp, x2, round2_mapping(), "round2")
        check("round2 run exits 0", rc == 0, log[-500:])
        if d6 is not None:
            after = [e for e in d6["excluded_rows"] if e.get("after_data_rows")]
            check("A14 total row after data_rows.to kept for reconciliation (with values)",
                  len(after) == 1 and after[0]["row"] == 9 and after[0]["values"]["rev.2024.recorded"]["v"] == 400000, after)
            check("named parent with a vertical SUM is kept (possible_total_row)", rec(d6, "W-1") is not None and
                  any(i["record_ids"] == ["W-1"] for i in issues_of(d6, "possible_total_row")),
                  [r["id"] for r in d6["records"]])
            w5 = rec(d6, "W-5")["fields"]
            check("A16 «12.500» in money -> text + ambiguous_thousands_separator (medium)",
                  w5["rev.2024.recorded"]["s"] == "text" and w5["rev.2024.recorded"]["v"] is None and
                  [i["severity"] for i in issues_of(d6, "ambiguous_thousands_separator")] == ["medium"], w5["rev.2024.recorded"])
            check("U-4.2 «٨٠٠ م²» in area -> 800 with the unit noted", w5["area"]["v"] == 800 and "المتر" in (w5["area"]["note"] or ""),
                  w5["area"])
            w6 = rec(d6, "W-6")["fields"]
            check("U-4.2 «USD 100,000» -> text + foreign_currency (high)",
                  w6["rev.2024.recorded"]["s"] == "text" and
                  [i["severity"] for i in issues_of(d6, "foreign_currency")] == ["high"], w6["rev.2024.recorded"])
            notes = {r["id"]: (r["fields"].get("notes") or {}) for r in d6["records"]}
            blob = json.dumps(d6, ensure_ascii=False)
            check("A12 IBAN and labelled national ID never stored in 03",
                  "SA0380000000608010167519" not in blob and "1023456789" not in blob, notes)
            check("A12 IBAN masked as a span (rest of the note kept)", notes["W-3"].get("v") == "الحوالة على <محجوب>", notes["W-3"])
            check("A12 «500 000 000» in a note is not a phone number", notes["W-6"].get("v") == "قيمة الأرض 500 000 000 ريال",
                  notes["W-6"])
            check("percent field mixing 0.5 and 50 -> percent_scale_mixed",
                  len(issues_of(d6, "percent_scale_mixed")) == 1, issues_of(d6, "percent_scale_mixed"))
            w1 = rec(d6, "W-1")["fields"]
            check("join on a deed shared with excluded units reaches the parent",
                  w1.get("exp.2024.opex", {}).get("v") == 60000 and not issues_of(d6, "join_ambiguous"),
                  (w1.get("exp.2024.opex"), issues_of(d6, "join_ambiguous")))
            check("rollup sums the units' revenue into the parent (derived)",
                  w1["rev.2023.recorded"]["v"] == 870000 and w1["rev.2023.recorded"]["s"] == "derived", w1.get("rev.2023.recorded"))
            check("rollup puts the units out of scope", all(rec(d6, x)["scope"] == "excluded" for x in ("W-2", "W-3", "W-4")),
                  [(x, rec(d6, x)["scope"]) for x in ("W-2", "W-3", "W-4")])
            sl = issues_of(d6, "section_label_skipped")
            check("section label in the id column is reported as such, not as a blank row",
                  len(sl) == 1 and sl[0]["rows"] == [6] and not any(6 in (i.get("rows") or []) for i in issues_of(d6, "blank_rows_skipped")),
                  (sl, issues_of(d6, "blank_rows_skipped")))

        # ---- mapping errors (exit 2) ------------------------------------------------------------------------------
        print("[5] أخطاء الربط (خروج 2)")
        cases = []
        m = base_mapping(); m["sources"][0]["data_rows"] = {"from": 9, "to": 3}; cases.append(("A4 from > to", m))
        m = base_mapping(); m["sources"][0]["data_rows"] = {"from": 2, "to": None}; cases.append(("A4 to null", m))
        m = base_mapping(); m["sources"][0]["data_rows"] = {"from": 5, "to": 6}; cases.append(("A4 zero records", m))
        m = base_mapping(); del m["entity"]; cases.append(("A13 entity missing", m))
        m = base_mapping(); m["data_period"] = {"label": "2024"}; cases.append(("A13 evidence missing", m))
        m = base_mapping(); m["dim_rules"] = [{"dim": "use", "from_field": "use_text", "map": {"استثماري": "استثماري"}}]
        cases.append(("A6 dim_rules value outside vocab", m))
        m = base_mapping(); m["records"]["REG-1"] = {"dims": {"income_nature": "إيجار"}}
        cases.append(("A6 records dims outside vocab", m))
        m = base_mapping()
        m["sources"].append({"id": "EXP", "sheet": "مصروفات", "header_row": 1, "data_rows": {"from": 2, "to": 4},
                             "default_scope": "in_scope", "scope_reason": "x", "columns": {"C": {"field": "exp.2024.opex", "kind": "money"}},
                             "join": {"into": "REG", "on": "id", "key_column": "A", "link_id": "L-EXP"}})
        m["links"][0]["status"] = "probable"
        cases.append(("A7 join through an unconfirmed link", m))
        for name, mp in cases:
            rc, log, _ = run_normalize(tmp, xlsx, mp, "err_" + re.sub(r"\W+", "_", name))
            check(f"{name} -> exit 2", rc == 2, f"rc={rc} {log[-300:]}")
    finally:
        pass
    return finish(tmp)


def finish(tmp):
    if os.environ.get("APA_KEEP"):
        print("kept:", tmp)
    else:
        shutil.rmtree(tmp, ignore_errors=True)
    print(f"\nناجح: {len(PASSES)} · فاشل: {len(FAILS)}")
    return 1 if FAILS else 0


if __name__ == "__main__":
    sys.exit(main())
