#!/usr/bin/env python3
"""Robustness regression tests for scripts/inspect_workbook.py on small FICTIONAL workbooks.

    python3 -I tests/robustness/test_inspect_robustness.py [--keep]

Every workbook is generated here with openpyxl (synthetic names and amounts only — never the real sample), inspected
with `python3 -I scripts/inspect_workbook.py` exactly as the skill runs it, and 01_workbook_map.json is checked.
Output goes to $APA_BUILD_DIR/robustness when APA_BUILD_DIR is set, else to a temporary folder (removed unless --keep).
Exit code: 0 = every check passed, 1 = at least one failure.

Cases (review finding in brackets):
  [A1]  header detection on plain ranges: a data row with a text cell in a numeric column is not a header;
        stacked sub-tables, a title line and a two-tier (merged) header still work
  [A3]  label-only total rows («المجموع» with #VALUE! cells, «جملة» in a later column) are total rows
  [A5]  a formula whose referenced cells are all blank but whose cached value is 0 -> formula_blank_inputs
  [A12] a personal-data column is never the name key; samples, notes and comments are masked
  [A15] Hijri examples come from Hijri cells only; mixed_calendar; ambiguous day/month note in parse_date
  [A17] «طبيعة/نوع/حالة …» headers are categories, not money (no false text_in_numeric)
  [A19] n/a tokens («—»، «غير متوفر»، «-») are counted as na, not as text categories
  [C18] cross-sheet name matches between two hidden sheets, with equal-amount evidence
"""
from __future__ import annotations

import sys

sys.dont_write_bytecode = True

import datetime as dt  # noqa: E402
import json  # noqa: E402
import os  # noqa: E402
import re  # noqa: E402
import shutil  # noqa: E402
import subprocess  # noqa: E402
import tempfile  # noqa: E402
import zipfile  # noqa: E402

HERE = os.path.dirname(os.path.abspath(__file__))
SKILL = os.path.dirname(os.path.dirname(HERE))
SCRIPTS = os.path.join(SKILL, "scripts")
sys.path.insert(0, SCRIPTS)

from apa_common import AMBIGUOUS_DATE_NOTE, REDACTED, is_pii_header, parse_date  # noqa: E402

try:
    import openpyxl
    from openpyxl.comments import Comment
except ImportError:  # pragma: no cover
    print("FAIL  openpyxl غير مثبتة")
    sys.exit(1)


# ------------------------------------------------------------------------------------------------
class Report:
    def __init__(self):
        self.passed, self.failed = 0, []

    def check(self, name: str, ok: bool, detail: str = "") -> bool:
        if ok:
            self.passed += 1
            print(f"  PASS  {name}")
        else:
            self.failed.append(name)
            print(f"  FAIL  {name}" + (f" — {detail}" if detail else ""))
        return ok


def inject_cached(path: str, sheet_no: int, values: dict) -> None:
    """openpyxl cannot store a formula's cached value; write <v> for the given cells as Excel would after a recalc."""
    name = f"xl/worksheets/sheet{sheet_no}.xml"
    tmp = path + ".tmp"
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == name:
                xml = data.decode("utf-8")
                for cell, v in values.items():
                    xml, n = re.subn(r'(<c r="%s"[^>]*>)<f>(.*?)</f><v\s*/?>(?:</v>)?' % cell,
                                     lambda m, v=v: f"{m.group(1)}<f>{m.group(2)}</f><v>{v}</v>", xml)
                    if n != 1:
                        raise RuntimeError(f"cached value not injected for {cell}")
                data = xml.encode("utf-8")
            zout.writestr(item, data)
    os.replace(tmp, path)


def run_inspect(xlsx: str, out: str) -> dict:
    proc = subprocess.run([sys.executable, "-I", os.path.join(SCRIPTS, "inspect_workbook.py"), xlsx, "--out", out,
                           "--quiet"], capture_output=True, text=True)
    if proc.returncode != 0:
        raise RuntimeError(f"inspect_workbook.py exit {proc.returncode}: {proc.stderr[-1500:]}")
    with open(os.path.join(out, "01_workbook_map.json"), encoding="utf-8") as fh:
        return json.load(fh)


def sheet(wm: dict, name: str) -> dict:
    return next(s for s in wm["sheets"] if s["name"] == name)


def anomalies(wm: dict, typ: str, sheet_name: str = None) -> list:
    return [a for a in wm["anomalies"] if a["type"] == typ
            and (sheet_name is None or a["where"].startswith(sheet_name + "!"))]


def col(sh: dict, letter: str) -> dict:
    return next((c for c in sh["columns"] if c["col"] == letter), {})


def put_rows(ws, rows, start=1):
    for r, row in enumerate(rows, start=start):
        for c, v in enumerate(row, start=1):
            if v is not None:
                ws.cell(row=r, column=c, value=v)


# ------------------------------------------------------------------------------------------------
REG_HDR = ["م", "اسم العقار", "الحي", "المدينة", "نوع العقار", "طريقة التملك", "الاستخدام", "عدد الوحدات",
           "الإيراد 2024", "التقييم 2024", "رقم الصك", "ملاحظات"]


def reg_rows(units_text: bool, deed_text: bool):
    return [
        [1, "عمارة السنابل", "حي الواحة", "مدينة الأمل", "عمارة", "شراء", "استثماري", 8, 240000, 4000000, "900100200301", "مؤجرة"],
        [2, "محلات الفجر", "حي النسيم", "مدينة الأمل", "محلات", "وقف", "استثماري",
         "4 محلات ومستودع" if units_text else 5, 180000, 3000000, "900100200302", None],
        [3, "فيلا الضياء", "حي الربوة", "مدينة الأمل", "فيلا", "هبة", "خدمي", 1, 0, 2500000,
         "صك قيد التحديث" if deed_text else "900100200303", "سكن مستفيدين"],
        [4, "أرض المرج", "حي المرج", "مدينة الأمل", "أرض", "شراء", "استثماري", 0, 0, 6000000, "900100200304", None],
        [5, "عمارة القمر", "حي القمر", "مدينة الأمل", "عمارة", "وقف", "استثماري", 12, 360000, 5500000, "900100200305", None],
        [6, "مستودع الصفا", "حي الصفا", "مدينة الأمل", "مستودع", "شراء", "استثماري", "غير محدد", 90000, 1200000,
         "900100200306", "عقد ينتهي 2025"],
    ]


def build_structure_book(path: str) -> None:
    wb = openpyxl.Workbook()
    # A1a: plain range, a units cell and a deed cell hold text
    ws = wb.active
    ws.title = "سجل أ"
    put_rows(ws, [REG_HDR] + reg_rows(True, True))
    # A1b: only the units cell holds text
    ws = wb.create_sheet("سجل ب")
    put_rows(ws, [REG_HDR] + reg_rows(True, False))
    # A1c: the first data row right under the header is mostly text-like values
    ws = wb.create_sheet("Assets")
    put_rows(ws, [
        ["ID", "Name", "Revenue 2024", "Value 2024", "Area", "Date acquired", "Ownership share", "Use"],
        ["X1", "Tower A", "12.500", "1.250.000", "500 m2", "03/04/2020", "50%", "investment"],
        ["X2", "Tower B", "1,5", "2,000,000", "1,200", "2020-01-13", 0.5, "investment"],
        ["X3", "Tower C", "100%", "3e6", "800 m2", "2019-12-25", "1/2", "investment"],
        ["X4", "Tower D", "20,000", "4,000,000", None, "2018-02-01", None, "investment"],
        ["X5", "Tower E", "1,200,000", " ", None, "2017-06-30", None, "investment"],
    ])
    # A1 guard: title + blank + header + data + total, then a section label and a second sub-table
    ws = wb.create_sheet("أقسام")
    ws["A1"] = "بيان الأصول الوهمية للاختبار"
    ws.merge_cells("A1:F1")
    put_rows(ws, [["م", "اسم العقار", "الحي", "الاستخدام", "الإيراد 2024", "التقييم 2024"],
                  [1, "عمارة السنابل", "حي الواحة", "استثماري", 240000, 4000000],
                  [2, "محلات الفجر", "حي النسيم", "استثماري", 180000, 3000000],
                  [3, "فيلا الضياء", "حي الربوة", "خدمي", 0, 2500000],
                  [4, "عمارة القمر", "حي القمر", "استثماري", 360000, 5500000],
                  [None, "الإجمالي", None, None, "=SUM(E4:E7)", "=SUM(F4:F7)"]], start=3)
    put_rows(ws, [["ثانيًا: الأراضي"],
                  ["م", "اسم الأرض", "الحي", "المساحة م2", "التقييم 2024"],
                  [1, "أرض المرج", "حي المرج", 1200, 6000000],
                  [2, "أرض الربوة", "حي الربوة", 900, 3600000],
                  [3, "أرض الساحل", "حي الساحل", 1500, 4500000]], start=10)
    # A1 guard: two-tier header (merged group labels over column labels)
    ws = wb.create_sheet("رأسان")
    for rng, label in (("A1:B1", "بيانات العقار"), ("C1:D1", "الإيرادات"), ("E1:F1", "التقييمات")):
        ws[rng.split(":")[0]] = label
        ws.merge_cells(rng)
    put_rows(ws, [["م", "اسم العقار", "الإيراد 2023", "الإيراد 2024", "التقييم 2023", "التقييم 2024"],
                  [1, "عمارة السنابل", 230000, 240000, 3900000, 4000000],
                  [2, "محلات الفجر", 170000, 180000, 2900000, 3000000],
                  [3, "عمارة القمر", 350000, 360000, 5400000, 5500000],
                  [4, "مستودع الصفا", 85000, 90000, 1150000, 1200000]], start=2)
    # A3: label-only total with Excel errors; «جملة» total labelled in a later column; «مجمع» is a name, not a total
    ws = wb.create_sheet("إجماليات")
    put_rows(ws, [["م", "اسم العقار", "البيان", "الإيراد 2024", "التقييم 2024"],
                  [1, "مجمع الواحة", "تجاري", 400000, 7000000],
                  [2, "عمارة السنابل", "سكني", 240000, 4000000],
                  [3, "محلات الفجر", "تجاري", 180000, 3000000],
                  [None, "المجموع", None, "#VALUE!", "#VALUE!"],
                  [4, "عمارة القمر", "سكني", 360000, 5500000],
                  [None, None, "جملة القسم", 360000, 5500000]])
    # round 2: a serial + a real name with «إجمالي» in a later attribute column (rent basis) is a data row
    ws = wb.create_sheet("أساس الإيجار")
    put_rows(ws, [["م", "اسم العقار", "الإيراد 2024", "التقييم 2024", "أساس الإيجار"],
                  [1, "عمارة السنابل", 240000, 4000000, "إجمالي"],
                  [2, "محلات الفجر", 180000, 3000000, "صافي"],
                  [3, "عمارة القمر", 360000, 5500000, "إجمالي"],
                  [None, "الإجمالي", 780000, 12500000, None]])
    # A15: Hijri text, Gregorian text and an Excel date cell in one column
    ws = wb.create_sheet("تواريخ")
    put_rows(ws, [["م", "اسم العقار", "تاريخ التملك"],
                  [1, "عمارة السنابل", "1440/05/10هـ"],
                  [2, "محلات الفجر", "15/03/2019"],
                  [3, "فيلا الضياء", dt.datetime(2018, 7, 1)],
                  [4, "عمارة القمر", "1445/02/01هـ"],
                  [5, "مستودع الصفا", "03/04/2021"]])
    # A17 + A19: classifier headers and n/a tokens
    ws = wb.create_sheet("تصنيفات")
    put_rows(ws, [["م", "اسم العقار", "طبيعة الدخل", "نوع الإيراد", "حالة التقييم", "الاستخدام", "الإيراد 2024"],
                  [1, "عمارة السنابل", "إيجار", "إيجار سنوي", "مقيّم", "استثماري", 240000],
                  [2, "محلات الفجر", "إيجار", "إيجار شهري", "مقيّم", "—", 180000],
                  [3, "مدرسة الفجر", "تشغيلي", "رسوم", "غير مقيّم", "غير متوفر", "تحت التحصيل"],
                  [4, "عمارة القمر", "إيجار", "إيجار سنوي", "مقيّم", "-", 360000],
                  [5, "مستودع الصفا", "أخرى", "أخرى", "مقيّم", "خدمي", 90000]])
    wb.save(path)


def build_formula_book(path: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "صيغ"
    put_rows(ws, [["م", "اسم العقار", "الإيجار الشهري", "الإيراد السنوي", "قيمة الأرض", "قيمة المبنى", "إجمالي التقييم",
                   "عدد الأرقام"],
                  [1, "عمارة السنابل", 10000, "=C2*12", 3000000, 1000000, "=E2+F2", "=COUNT(E2:F2)"],
                  [2, "محلات الفجر", None, "=C3*12", 2000000, 1000000, "=E3+F3", "=COUNT(E3:F3)"],
                  [3, "فيلا الضياء", 0, "=C4*12", 1500000, 500000, "=E4+F4", "=COUNT(E4:F4)"],
                  [4, "أرض المرج", 20000, "=C5*12", None, None, "=E5+F5", "=COUNT(E5:F5)"]])
    wb.save(path)
    inject_cached(path, 1, {"D2": 120000, "D3": 0, "D4": 0, "D5": 240000, "G2": 4000000, "G3": 3000000,
                            "G4": 2000000, "G5": 0, "H2": 2, "H3": 2, "H4": 2, "H5": 0})


PERSONS = ("سالم الوهمي", "نورة الافتراضية", "فهد التجريبي", "هند المثالية")
PHONE = "0551230000"
IBAN = "SA4420000001234567891234"
EMAIL = "tenant.test@example.com"


def build_pii_book(path: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "مستأجرون"
    put_rows(ws, [["م", "اسم المستأجر", "اسم العقار", "جوال المستأجر", "ملاحظات", "الإيراد 2024"],
                  [1, PERSONS[0], "عمارة السنابل", PHONE, f"للتواصل {PHONE}", 240000],
                  [2, PERSONS[1], "محلات الفجر", "0551230001", f"حوالة على {IBAN}", 180000],
                  [3, PERSONS[2], "عمارة القمر", "0551230002", "مؤجرة", 360000],
                  [4, PERSONS[3], "مستودع الصفا", "0551230003", None, 90000]])
    ws["E2"].comment = Comment(f"راسل {EMAIL} أو {PHONE}", "اختبار")
    # header spelled with tatweel / without hamza
    ws = wb.create_sheet("تهجئة")
    put_rows(ws, [["م", "اســم المستــأجر", "المستاجر الثاني", "اسم العقار", "الإيراد 2024"],
                  [1, PERSONS[0], PERSONS[3], "عمارة السنابل", 240000],
                  [2, PERSONS[1], PERSONS[2], "محلات الفجر", 180000],
                  [3, PERSONS[2], PERSONS[1], "عمارة القمر", 360000]])
    # the only text column names a person: no name key may be proposed
    ws = wb.create_sheet("ملاك")
    put_rows(ws, [["م", "اسم المالك", "الإيجار السنوي 2024", "التقييم 2024"],
                  [1, PERSONS[0], 300000, 5000000],
                  [2, PERSONS[1], 120000, 2000000],
                  [3, PERSONS[2], 90000, 1500000]])
    # national-ID-like values: flagged in a reference column, never in an amount column
    ws = wb.create_sheet("مراجع")
    put_rows(ws, [["م", "اسم العقار", "المرجع", "التقييم 2024"],
                  [1, "عمارة السنابل", "1023456789", 1500000000],
                  [2, "محلات الفجر", "1034567890", 1250000000],
                  [3, "عمارة القمر", "2045678901", 2100000000]])
    wb.save(path)


def build_cross_book(path: str) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "السجل"
    put_rows(ws, [["م", "اسم العقار", "الإيراد 2024", "التقييم 2024"],
                  [1, "برج الأفق", 500000, 9000000],
                  [2, "عمارة الندى", 210000, 3500000],
                  [3, "محلات الواحة", 150000, 2600000]])
    ws = wb.create_sheet("تخصيص")
    ws.sheet_state = "hidden"
    put_rows(ws, [["م", "اسم العقار", "الإيراد", "المصروفات"],
                  [1, "برج الأفق رقم ( 3 )", 500000, 40000],
                  [2, "عمارة السنابل وقف مدرسة الفجر", 480000, 25000],
                  [3, "عمارة الندى", 210000, 9000],
                  [4, "شقة الريحان", 60000, 1500]])
    ws = wb.create_sheet("أملاك الغير")
    ws.sheet_state = "hidden"
    put_rows(ws, [["م", "اسم العقار", "الإيجار السنوي 2024", "التقييم 2024"],
                  [1, "عمارة السنابل", 480000, 7000000],
                  [2, "فيلا الشروق", 95000, 1800000]])
    wb.save(path)


# ------------------------------------------------------------------------------------------------
def test_structure(rep: Report, wm: dict) -> None:
    print("[A1] رؤوس الجداول في النطاقات العادية")
    for name in ("سجل أ", "سجل ب"):
        sh = sheet(wm, name)
        rep.check(f"A1 «{name}»: header row 1", sh["primary_header_row"] == 1, f"got {sh['primary_header_row']}")
        rep.check(f"A1 «{name}»: data rows 2–7 (6)", sh["data_rows"] == {"from": 2, "to": 7, "count": 6},
                  f"got {sh['data_rows']}")
        rep.check(f"A1 «{name}»: row kinds", sh["row_kinds"] == {"1": "header", "2-7": "data"}, f"got {sh['row_kinds']}")
        rep.check(f"A1 «{name}»: name key = B (اسم العقار)", sh["key_columns"].get("name") == "B",
                  f"got {sh['key_columns']}")
        rep.check(f"A1 «{name}»: role asset_register", sh["role_hint"] == "asset_register", f"got {sh['role_hint']}")
    sh = sheet(wm, "Assets")
    rep.check("A1 «Assets»: text-like first data row is not a header",
              sh["primary_header_row"] == 1 and sh["data_rows"] == {"from": 2, "to": 6, "count": 5},
              f"header {sh['primary_header_row']}, data {sh['data_rows']}")
    sh = sheet(wm, "أقسام")
    rk = sh["row_kinds"]
    rep.check("A1 guard «أقسام»: both sub-table headers kept (rows 3 and 11)",
              rk.get("3") == "header" and rk.get("11") == "header", f"row_kinds {rk}")
    rep.check("A1 guard «أقسام»: total row 8 and section label 10",
              rk.get("8") in ("total_formula", "total_label") and rk.get("10") == "section_label", f"row_kinds {rk}")
    rep.check("A1 guard «أقسام»: data rows of both blocks",
              rk.get("4-7") == "data" and rk.get("12-14") == "data", f"row_kinds {rk}")
    sh = sheet(wm, "رأسان")
    rep.check("A1 guard «رأسان»: two-tier header -> column labels on row 2, data 3–6",
              sh["primary_header_row"] == 2 and sh["data_rows"] == {"from": 3, "to": 6, "count": 4},
              f"header {sh['primary_header_row']}, data {sh['data_rows']}")
    rep.check("A1 guard «رأسان»: group label prefixed to the column header",
              (col(sh, "D").get("header") or "").startswith("الإيرادات /"), f"got {col(sh, 'D').get('header')!r}")

    print("[A3] صفوف الإجمالي")
    sh = sheet(wm, "إجماليات")
    rk = sh["row_kinds"]
    rep.check("A3 «المجموع» with #VALUE! cells is a total row (row 5)",
              rk.get("5") in ("total_label", "total_formula"), f"row_kinds {rk}")
    rep.check("A3 «جملة القسم» in column C is a total row (row 7)",
              rk.get("7") in ("total_label", "total_formula"), f"row_kinds {rk}")
    rep.check("A3 «مجمع الواحة» stays a data row (row 2)", rk.get("2-4") == "data", f"row_kinds {rk}")
    rep.check("A3 data rows exclude the totals", sh["data_rows"]["count"] == 4, f"got {sh['data_rows']}")
    sh = sheet(wm, "أساس الإيجار")
    rk = sh["row_kinds"]
    rep.check("A3 named rows with «إجمالي» in a later attribute column stay data (rows 2–4)",
              rk.get("2-4") == "data" and rk.get("5") in ("total_label", "total_formula"), f"row_kinds {rk}")

    print("[A15] التقويم")
    hj = anomalies(wm, "hijri_date_text", "تواريخ")
    raws = [e["raw"] for a in hj for e in a["evidence"]["examples"]]
    rep.check("A15 hijri_date_text examples are Hijri cells only",
              bool(raws) and all(re.search(r"14\d\d", x) for x in raws) and "15/03/2019" not in raws, f"examples {raws}")
    mc = anomalies(wm, "mixed_calendar", "تواريخ")
    rep.check("A15 mixed_calendar on column C", len(mc) == 1 and mc[0]["evidence"]["col"] == "C"
              and mc[0]["evidence"]["hijri_count"] == 2 and mc[0]["evidence"]["gregorian_count"] == 3,
              f"got {[a['evidence'] for a in mc]}")
    rep.check("A15 mixed_calendar severity medium", bool(mc) and mc[0]["severity"] == "medium")
    notes = " ".join(col(sheet(wm, "تواريخ"), "C").get("notes") or [])
    rep.check("A15 ambiguous day/month noted on the column", "ملتبس" in notes, f"notes {notes!r}")
    amb = parse_date("03/04/2021")
    rep.check("A15 parse_date 03/04/2021 -> 2021-04-03 with ambiguity note",
              amb["v"] == "2021-04-03" and amb["note"] == AMBIGUOUS_DATE_NOTE, f"got {amb}")
    for txt in ("13/04/2021", "05/05/2021", "2021-04-03"):
        d = parse_date(txt)
        rep.check(f"A15 parse_date {txt} has no ambiguity note", not (d.get("note") or "").startswith("تاريخ ملتبس"),
                  f"got {d}")
    hj_amb = parse_date("10/05/1440هـ")
    rep.check("A15 Hijri d/m/y also carries the ambiguity note",
              hj_amb["calendar"] == "hijri" and (hj_amb["note"] or "").startswith(AMBIGUOUS_DATE_NOTE), f"got {hj_amb}")

    print("[A17] عناوين التصنيف")
    sh = sheet(wm, "تصنيفات")
    tin = anomalies(wm, "text_in_numeric", "تصنيفات")
    cols_flagged = sorted(a["evidence"]["col"] for a in tin)
    rep.check("A17 no text_in_numeric on «طبيعة الدخل»/«نوع الإيراد»/«حالة التقييم»",
              not set(cols_flagged) & {"C", "D", "E"}, f"flagged {cols_flagged}")
    rep.check("A17 control: text in «الإيراد 2024» still flagged", "G" in cols_flagged, f"flagged {cols_flagged}")
    for letter in ("C", "D", "E"):
        h = col(sh, letter).get("hints") or []
        rep.check(f"A17 column {letter} hints: category, no money/revenue/valuation",
                  "category" in h and not {"money", "revenue", "valuation"} & set(h), f"hints {h}")

    print("[A19] رموز غير المتاح")
    prof = col(sh, "F").get("profile") or {}
    rep.check("A19 «—»، «غير متوفر»، «-» counted as na (3), real values as text (2)",
              prof.get("na") == 3 and prof.get("text") == 2, f"profile {prof}")


def test_formulas(rep: Report, wm: dict) -> None:
    print("[A5] صفر محسوب من خانات فارغة")
    fb = anomalies(wm, "formula_blank_inputs", "صيغ")
    cells = sorted(e["ref"] for a in fb for e in a["evidence"]["examples"])
    rep.check("A5 D3 (=C3*12, C3 blank) and G5 (=E5+F5, both blank) flagged", cells == ["D3", "G5"], f"got {cells}")
    rep.check("A5 D4 (=C4*12 with a recorded 0) not flagged", "D4" not in cells)
    rep.check("A5 COUNT of blanks (H5) not flagged", "H5" not in cells)
    rep.check("A5 severity medium", bool(fb) and all(a["severity"] == "medium" for a in fb))
    ex = next((e for a in fb for e in a["evidence"]["examples"] if e["ref"] == "G5"), {})
    rep.check("A5 evidence lists the blank inputs", ex.get("inputs") == ["E5", "F5"], f"got {ex}")


def test_pii(rep: Report, wm: dict, raw_json: str, raw_md: str) -> None:
    print("[A12] البيانات الشخصية")
    sh = sheet(wm, "مستأجرون")
    rep.check("A12 name key is «اسم العقار» (C), never «اسم المستأجر»", sh["key_columns"].get("name") == "C",
              f"got {sh['key_columns']}")
    pii = {(p["sheet"], p["col"]) for p in wm["pii_columns"]}
    rep.check("A12 tenant name and mobile columns are PII", {("مستأجرون", "B"), ("مستأجرون", "D")} <= pii, f"got {pii}")
    rep.check("A12 PII samples redacted", col(sh, "B").get("samples") == [REDACTED], f"got {col(sh, 'B').get('samples')}")
    sh2 = sheet(wm, "تهجئة")
    rep.check("A12 headers with tatweel / without hamza are PII",
              {("تهجئة", "B"), ("تهجئة", "C")} <= pii and sh2["key_columns"].get("name") == "D",
              f"pii {sorted(p for p in pii if p[0] == 'تهجئة')}, keys {sh2['key_columns']}")
    sh3 = sheet(wm, "ملاك")
    rep.check("A12 a person-name column is never proposed as the name key",
              sh3["key_columns"].get("name") != "B" and ("ملاك", "B") in pii, f"keys {sh3['key_columns']}")
    rep.check("A12 national-ID-like reference column flagged by pattern",
              any(p["sheet"] == "مراجع" and p["col"] == "C" and p["reason"] == "pattern" for p in wm["pii_columns"]),
              f"got {[p for p in wm['pii_columns'] if p['sheet'] == 'مراجع']}")
    rep.check("A12 10-digit valuations are not PII", ("مراجع", "D") not in pii)
    for label, text in (("01_workbook_map.json", raw_json), ("01_workbook_map.md", raw_md)):
        leaked = [x for x in PERSONS + (PHONE, IBAN, EMAIL, "0551230001") if x in text]
        rep.check(f"A12 no person name / phone / IBAN / e-mail in {label}", not leaked, f"leaked {leaked}")
    notes_samples = col(sh, "E").get("samples") or []
    rep.check("A12 notes column keeps its text with the personal data masked",
              any(REDACTED in s for s in notes_samples), f"samples {notes_samples}")
    comments = " ".join(c["text"] for c in sh.get("comments") or [])
    rep.check("A12 cell comment masked", REDACTED in comments and EMAIL not in comments, f"comments {comments!r}")
    rep.check("A12 is_pii_header variants",
              is_pii_header("اســم المستــأجر") and is_pii_header("المستاجر") and is_pii_header("Tenant Name")
              and not is_pii_header("اسم العقار") and not is_pii_header("طبيعة الدخل"))


def test_cross(rep: Report, wm: dict) -> None:
    print("[C18] المطابقة بين الأوراق المخفية")
    cs = [a for a in wm["anomalies"] if a["type"] == "cross_sheet_name_match"]
    pair = [a for a in cs if {a["evidence"]["sheet_a"], a["evidence"]["sheet_b"]} == {"تخصيص", "أملاك الغير"}]
    rep.check("C18 hidden ↔ hidden pair compared (allocation ↔ managed)", len(pair) == 1,
              f"pairs {[(a['evidence']['sheet_a'], a['evidence']['sheet_b']) for a in cs]}")
    m = [x for a in pair for x in a["evidence"]["matches"]]
    hit = [x for x in m if "السنابل" in x["a_name"] + x["b_name"]]
    rep.check("C18 one-word name contained in the allocation line is matched (token_subset)",
              bool(hit) and hit[0]["match"] == "token_subset", f"matches {m}")
    eq = hit[0].get("equal_amounts") if hit else None
    rep.check("C18 equal amount recorded as evidence (480,000)", bool(eq) and eq[0]["value"] == 480000, f"got {eq}")
    rep.check("C18 detail mentions the equal amount", bool(pair) and "بمبلغ متساوٍ" in pair[0]["detail"])
    reg = [a for a in cs if {a["evidence"]["sheet_a"], a["evidence"]["sheet_b"]} == {"السجل", "تخصيص"}]
    rep.check("C18 visible register ↔ hidden allocation still compared", len(reg) == 1)
    rep.check("C18 severity stays info", all(a["severity"] == "info" for a in cs))


# ------------------------------------------------------------------------------------------------
def main(argv=None) -> int:
    keep = "--keep" in (argv or sys.argv[1:])
    base = os.environ.get("APA_BUILD_DIR")
    if base:
        work = os.path.join(base, "robustness")
        shutil.rmtree(work, ignore_errors=True)
        os.makedirs(work)
    else:
        work = tempfile.mkdtemp(prefix="apa_robustness_")
    print(f"== اختبارات متانة فحص المصنف (بيانات وهمية) ==\n   المجلد: {work}")
    rep = Report()
    try:
        cases = (("structure", build_structure_book, test_structure),
                 ("formulas", build_formula_book, test_formulas),
                 ("pii", build_pii_book, test_pii),
                 ("cross", build_cross_book, test_cross))
        for name, build, test in cases:
            xlsx = os.path.join(work, f"{name}.xlsx")
            out = os.path.join(work, f"run_{name}")
            build(xlsx)
            try:
                wm = run_inspect(xlsx, out)
            except RuntimeError as exc:
                rep.check(f"inspect_workbook on {name}.xlsx", False, str(exc))
                continue
            if name == "pii":
                with open(os.path.join(out, "01_workbook_map.json"), encoding="utf-8") as fh:
                    raw_json = fh.read()
                with open(os.path.join(out, "01_workbook_map.md"), encoding="utf-8") as fh:
                    raw_md = fh.read()
                test(rep, wm, raw_json, raw_md)
            else:
                test(rep, wm)
    finally:
        if not base and not keep:
            shutil.rmtree(work, ignore_errors=True)
    print(f"\n== الملخص: ناجح {rep.passed} · فاشل {len(rep.failed)} ==")
    for f in rep.failed:
        print(f"  ✗ {f}")
    print("RESULT: " + ("FAIL" if rep.failed else "PASS"))
    return 1 if rep.failed else 0


if __name__ == "__main__":
    sys.exit(main())
