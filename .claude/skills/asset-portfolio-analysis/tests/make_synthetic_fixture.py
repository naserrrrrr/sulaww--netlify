#!/usr/bin/env python3
"""Build the FICTIONAL regression workbook used by tests/run_tests.sh, and its hand-computed expectations.

Usage:
    python3 tests/make_synthetic_fixture.py --out PATH.xlsx [--write-expected PATH.json] [--check-expected PATH.json]

Every name, number, deed and phone in this file is invented («جمعية تجريبية لرعاية الأيتام (بيانات وهمية)»).
The workbook deliberately reproduces every trap listed in references/00-original-brief.md:
  * visible register (Excel table + SUBTOTAL totals row + NBSP filler rows), title «لعام 2025م»;
  * «صافي الأصول» holding text labels, mixed «عدد الوحدات», area held as text, numbers typed as Arabic-Indic text;
  * VAT incl ≠ excl for commercial rows, incl = excl for residential rows, «لا ينطبق» in a revenue cell;
  * 2023 totals as formulas WITHOUT cached values (file not recalculated) + one hard-coded wrong 2023 total;
  * 2024 totals as formulas WITH cached values (injected into the XML, since openpyxl cannot compute caches);
  * investment land with a recorded zero, a villa with revenue but a blank 2024 valuation (older 2023 one exists),
    service HQ (school), admin HQ, development land, two parcels with the same name and different deeds,
    two rows sharing one deed, a complex on leased land (land rent only in a note), an operating activity;
  * a serial gap (no «م» = 13), Hijri text dates and one Gregorian date cell, personal-data columns (fake);
  * hidden «تخصيص الإيرادات» (two sections with SUM totals, hard-coded wrong grand total and item count,
    a name duplicated across sections, names that only partially match the register);
  * hidden «إدارة أملاك» (managed for others, partial-period header «ايراد حتى 2024/10م», a row total
    ≠ land + building, zero valuations meaning «not valued», garbled header, personal-data columns);
  * hidden «مصروفات 2025 » (trailing space; operating expenses keyed by the register serial «م»);
  * hidden «المنسدلات» (lookup lists, referenced by data validations).

Expected metric values are computed below from the same data lists with plain, explicit Python
(no import of the skill's engine), then cross-checked against literal hand-computed numbers.
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import re
import sys
import tempfile
import zipfile
from collections import OrderedDict

sys.dont_write_bytecode = True

ENTITY = "جمعية تجريبية لرعاية الأيتام (بيانات وهمية)"
FILE_NAME = "synthetic_portfolio.xlsx"
REG = "سجل الأصول"
ALC = "تخصيص الإيرادات"
MAN = "إدارة أملاك"
EXP = "مصروفات 2025 "      # trailing space on purpose (exact sheet names must be kept)
LKP = "المنسدلات"
FIXED_TS = dt.datetime(2025, 12, 15, 9, 30, 0)


# ---------------------------------------------------------------------------
# Cell value wrappers
# ---------------------------------------------------------------------------
class F:
    """A formula written WITH a cached value (injected into the sheet XML after saving)."""

    def __init__(self, formula: str, cached):
        self.formula, self.cached = formula, cached


class FNC:
    """A formula written WITHOUT a cached value (as if the file was saved by a tool that does not recalculate)."""

    def __init__(self, formula: str):
        self.formula = formula


class NumText:
    """A number stored as text (e.g. Arabic-Indic digits); Excel's SUM/SUBTOTAL ignore it."""

    def __init__(self, text: str, value):
        self.text, self.value = text, value


NA = "لا ينطبق"
NBSP = " "

# ---------------------------------------------------------------------------
# Register «سجل الأصول»
# ---------------------------------------------------------------------------
REG_COLS = [
    ("A", "م", "serial"), ("B", "اسم العقار", "name"),
    ("C", "الايراد بدون ضريبة 2025", "rev25_ex"), ("D", "الايراد مع الضريبة 2025", "rev25_in"),
    ("E", "الحي", "district"), ("F", "الموقع", "location"), ("G", "تخصيص الريع", "rent_alloc"),
    ("H", "صافي الأصول", "restriction"), ("I", "طريقة التملك", "own_method"), ("J", "نوع العقار", "atype"),
    ("K", "التخصيص", "alloc"), ("L", "عدد الوحدات التأجيرية", "units"), ("M", "مساحة الأرض", "area"),
    ("N", "الإيراد الفعلي عام 2024", "rev24"), ("O", "مصروفات التشغيل 2025", "opex25"),
    ("P", "تقييم أرض 2023", "l23"), ("Q", "تقييم مبنى 2023", "b23"), ("R", "الإجمالي وفق تقييم 2023", "t23"),
    ("S", "تقييم أرض 2024", "l24"), ("T", "تقييم مبنى 2024", "b24"), ("U", "الإجمالي وفق تقييم 2024", "t24"),
    ("V", "تاريخ الملكية", "date"), ("W", "رقم صك الملكية", "deed"), ("X", "اسم المشرف", "sup"),
    ("Y", "رقم التواصل", "phone"), ("Z", "ملاحظات", "notes"), ("AA", "مقترحات", "proposals"),
]
COL_OF = {key: col for col, _, key in REG_COLS}
REG_HEADER_ROW = 2
REG_FIRST = 3
T23 = "t23_formula"   # sentinel: R = P + Q written as a formula without a cache
T24 = "t24_formula"   # sentinel: U = S + T written as a formula with a cache

# Classification = the analyst's truth used for the expectations (mirrors tests/synthetic/02_mapping.json).
#   rt: record_type, use, inc: income_nature, own: ownership list, waqf
REG_ROWS = [
    dict(serial=1, name="عمارة النور التجارية", rev25_ex=420000, rev25_in=483000, district="العليا",
         location="شمال الرياض", rent_alloc="كفالة الأيتام", restriction="غير مقيد", own_method="شراء",
         atype="عمارة تجارية", alloc="استثماري", units="6 محلات + 4 مكاتب", area=1200, rev24=400000, opex25=60000,
         l23=3500000, b23=2000000, t23=T23, l24=3800000, b24=2100000, t24=T24, date="1410/5/12هـ",
         deed="310112345678",
         _rt="asset", _use="investment", _inc="rent", _own=["owned"], _waqf="unknown"),
    dict(serial=2, name="مركز الريان التجاري", rev25_ex=300000, rev25_in=345000, district="الملز",
         location="شرق الرياض", rent_alloc="كفالة الأيتام", restriction="مقيد", own_method="وقف",
         atype="عمارة تجارية", alloc="استثماري", units="8 محلات", area=950, rev24=310000, opex25=None,
         l23=2400000, b23=1400000, t23=4100000,  # hard-coded wrong total (land + building = 3,800,000)
         l24=2500000, b24=1500000, t24=T24, date="1415/9/3هـ", deed="310112345679",
         _rt="asset", _use="investment", _inc="rent", _own=["owned"], _waqf="yes"),
    dict(serial=3, name="عمارة السلام السكنية", rev25_ex=180000, rev25_in=180000, district="السويدي",
         location="غرب الرياض", rent_alloc="كفالة الأيتام", restriction="غير مقيد", own_method="شراء",
         atype="عمارة سكنية", alloc="استثماري", units=12, area=800, rev24=175000, opex25=25000,
         l23=1100000, b23=750000, t23=T23, l24=1200000, b24=800000, t24=T24,
         date=dt.datetime(1998, 3, 15),  # the only Gregorian Excel date cell in a Hijri text column
         deed="410200000001",
         _rt="asset", _use="investment", _inc="rent", _own=["owned"], _waqf="unknown"),
    dict(serial=4, name="عمارة الفجر السكنية", rev25_ex=150000, rev25_in=150000, district="الشفا",
         location="جنوب الرياض", rent_alloc="برامج الأيتام التعليمية", restriction="غير مقيد", own_method="هبة",
         atype="عمارة سكنية", alloc="استثماري", units="4 محلات + 2 شقة", area=700, rev24=150000, opex25=None,
         l23=950000, b23=650000, t23=T23, l24=1000000, b24=700000, t24=T24, date="1420/1/20هـ",
         deed="410200000002",
         _rt="asset", _use="investment", _inc="rent", _own=["owned"], _waqf="unknown"),
    dict(serial=5, name="عمارة الهدى السكنية", rev25_ex=96000, rev25_in=96000, district="الشفا",
         location="جنوب الرياض", rent_alloc="كفالة الأيتام", restriction="مقيد", own_method="وقف",
         atype="عمارة سكنية", alloc="استثماري", units=8, area=600, rev24=None, opex25=12000,
         l23=850000, b23=450000, t23=T23, l24=900000, b24=500000, t24=T24, date="1425/7/1هـ",
         deed="410300000555",  # shared with serial 6
         _rt="asset", _use="investment", _inc="rent", _own=["owned"], _waqf="yes"),
    dict(serial=6, name="عمارة الأمل السكنية", rev25_ex=120000, rev25_in=120000, district="الشفا",
         location="جنوب الرياض", rent_alloc="كفالة الأيتام", restriction="غير مقيد", own_method="شراء",
         atype="عمارة سكنية", alloc="استثماري", units=10, area=650, rev24=110000, opex25=None,
         l23=1000000, b23=550000, t23=T23, l24=1100000, b24=600000, t24=T24, date="1426/2/10هـ",
         deed="410300000555",  # shared with serial 5
         _rt="asset", _use="investment", _inc="rent", _own=["owned"], _waqf="unknown"),
    dict(serial=7, name="أرض استثمارية بحي النرجس", rev25_ex=0, rev25_in=0, district="النرجس",
         location="شمال الرياض", rent_alloc="كفالة الأيتام", restriction="غير مقيد", own_method="شراء",
         atype="أرض", alloc="استثماري", units=None, area=2500, rev24=0, opex25=None,
         l23=2000000, b23=None, t23=T23, l24=2200000, b24=None, t24=T24, date="1430/11/5هـ",
         deed="510100000777", notes="الأرض غير مؤجرة خلال 2025", proposals="بيع",
         _rt="land_parcel", _use="investment", _inc="rent", _own=["owned"], _waqf="unknown"),
    dict(serial=8, name="فيلا الروضة", rev25_ex=85000, rev25_in=85000, district="الروضة",
         location="شرق الرياض", rent_alloc="كفالة الأيتام", restriction="غير مقيد", own_method="شراء",
         atype="فيلا", alloc="استثماري", units=1, area=500, rev24=80000, opex25=None,
         l23=700000, b23=400000, t23=T23, l24=None, b24=None, t24=None,  # blank 2024 valuation
         date="1418/4/22هـ", deed="510100000888", notes="لم يُحدَّث تقييم الفيلا في 2024", proposals="ترميم",
         _rt="asset", _use="investment", _inc="rent", _own=["owned"], _waqf="unknown"),
    dict(serial=9, name="مقر مدرسة تحفيظ القرآن", rev25_ex=0, rev25_in=0, district="النسيم",
         location="شرق الرياض", rent_alloc="برامج الأيتام التعليمية", restriction="مقيد", own_method="وقف",
         atype="مقر", alloc="غير استثماري", units=None, area=1500, rev24=0, opex25=None,
         l23=1400000, b23=850000, t23=T23, l24=1500000, b24=900000, t24=T24, date="1408/8/8هـ",
         deed="510100000999", notes="يُستخدم مقرًا لمدرسة تحفيظ تخدم الأيتام دون إيجار",
         _rt="asset", _use="service", _inc="none", _own=["owned"], _waqf="yes"),
    dict(serial=10, name="المقر الإداري للجمعية", rev25_ex=NA, rev25_in=NA, district="الورود",
         location="شمال الرياض", rent_alloc=None, restriction="غير مقيد", own_method="شراء",
         atype="مقر", alloc="اداري", units=None, area=900, rev24=None, opex25=None,
         l23=1200000, b23=650000, t23=T23, l24=1300000, b24=700000, t24=T24, date="1412/3/14هـ",
         deed="510200000111",
         _rt="asset", _use="admin", _inc="none", _own=["owned"], _waqf="unknown"),
    dict(serial=11, name="أرض قيد التطوير بحي الملقا", rev25_ex=None, rev25_in=None, district="الملقا",
         location="شمال الرياض", rent_alloc="كفالة الأيتام", restriction="غير مقيد", own_method="هبة",
         atype="أرض", alloc="استثماري", units=None, area="بانتظار الرفع المساحي", rev24=None, opex25=None,
         l23=None, b23=None, t23=None, l24=None, b24=None, t24=None, date="1440/6/1هـ",
         deed="لم يفرغ باسم الجمعية بعد", notes="المشروع قيد التطوير ولم يكتمل البناء؛ لا يوجد تقييم",
         _rt="land_parcel", _use="development", _inc="none", _own=["owned"], _waqf="unknown"),
    dict(serial=12, name="أرض المنار", rev25_ex=NumText("٣٠٬٠٠٠", 30000), rev25_in=30000, district="المنار",
         location="شرق الرياض", rent_alloc="كفالة الأيتام", restriction="غير مقيد", own_method="شراء",
         atype="أرض", alloc="استثماري", units=None, area=1000, rev24=30000, opex25=None,
         l23=None, b23=None, t23=None, l24=600000, b24=None, t24=T24, date="1435/2/2هـ",
         deed="610100000001", notes="مؤجرة مواقف سيارات",
         _rt="land_parcel", _use="investment", _inc="rent", _own=["owned"], _waqf="unknown"),
    # serial 13 is deliberately missing (serial gap)
    dict(serial=14, name="أرض المنار", rev25_ex=None, rev25_in=None, district="المنار",
         location="شرق الرياض", rent_alloc="كفالة الأيتام", restriction="غير مقيد", own_method="شراء",
         atype="أرض", alloc="استثماري", units=None, area=1100, rev24=None, opex25=None,
         l23=None, b23=None, t23=None, l24=650000, b24=None, t24=T24, date="1435/2/2هـ",
         deed="610100000002",
         _rt="land_parcel", _use="investment", _inc="rent", _own=["owned"], _waqf="unknown"),
    dict(serial=15, name="مجمع الواحة التجاري على أرض مستأجرة", rev25_ex=260000, rev25_in=299000,
         district="الحزم", location="جنوب الرياض", rent_alloc="كفالة الأيتام", restriction="غير مقيد",
         own_method="إيجار أرض", atype="مجمع تجاري", alloc="استثماري", units="10 محلات + مبنى مدرسة", area=None,
         rev24=240000, opex25=None, l23=None, b23=None, t23=None, l24=None, b24=None, t24=None, date=None,
         deed=None,
         notes="مقام على أرض مستأجرة بعقد طويل؛ أجرة الأرض السنوية 90,000 ريال حسب العقد؛ يضم محلات ومبنى مدرسة",
         _rt="asset", _use="investment", _inc="rent", _own=["leased_land_project"], _waqf="unknown"),
    dict(serial=16, name="نشاط مدارس أهلية", rev25_ex=1250000, rev25_in=1250000, district="الحزم",
         location="جنوب الرياض", rent_alloc="برامج الأيتام التعليمية", restriction="غير مقيد",
         own_method="تشغيل ذاتي", atype="نشاط تعليمي", alloc="استثماري", units=None, area=None, rev24=1100000,
         opex25=None, l23=None, b23=None, t23=None, l24=None, b24=None, t24=None, date=None, deed=None,
         notes="إيراد رسوم دراسية لنشاط تشغيلي؛ تكاليف النشاط غير مدرجة في الملف",
         _rt="operating_activity", _use="investment", _inc="operating", _own=["unknown"], _waqf="unknown"),
]
for _i, _r in enumerate(REG_ROWS):
    _r["_row"] = REG_FIRST + _i
    _r["_id"] = f"REG-{_r['serial']}"
    _r["sup"] = f"مشرف تجريبي {_i + 1}"
    _r["phone"] = f"05{_i + 1:08d}"
    _r.setdefault("notes", None)
    _r.setdefault("proposals", None)
REG_LAST = REG_FIRST + len(REG_ROWS) - 1          # 17
REG_TOTAL_ROW = REG_LAST + 1                       # 18
REG_FILLER = list(range(REG_TOTAL_ROW + 1, REG_TOTAL_ROW + 11))  # 19..28
REG_LAST_COL = "AA"
REG_TOTALS_COLS = ["C", "D", "N", "O", "S", "T", "U"]

# ---------------------------------------------------------------------------
# Hidden «إدارة أملاك» (managed for others)
# ---------------------------------------------------------------------------
MAN_HEADERS = ["م", "اسم العقار", "الحي", "الموقع", "نوع العقار", "الايراد بدون ضريبة 2025", "ايراد حتى 2024/10م",
               "نسبة أتعاب الإدارة", "تقييم أرض 2024", "تقييم مبنى 2024", "الإجمالي وفق تقييم 2024",
               "ملاحظات، ملاحظات المالك", "اسم المشرف", "رقم التواصل"]
MAN_ROWS = [
    dict(serial=1, name="عمارة الياسمين (ملك للغير)", district="الياسمين", location="شمال الرياض",
         atype="عمارة سكنية", rev25=240000, rev_oct=190000, fee=0.05, land=0, bld=0, total=F("I3+J3", 0),
         notes="الإيراد المدون كامل إيراد العقار؛ حصة الجمعية أتعاب الإدارة فقط", sup="مشرف تجريبي 16",
         phone="0500000016"),
    dict(serial=2, name="مركز الشفا التجاري (ملك للغير)", district="الشفا", location="جنوب الرياض",
         atype="عمارة تجارية", rev25=360000, rev_oct=300000, fee=0.07, land=1200000, bld=800000,
         total=1800000,  # hard-coded; land + building = 2,000,000
         notes="لم يُحدَّث الإجمالي بعد التقييم", sup="مشرف تجريبي 17", phone="0500000017"),
]
for _i, _r in enumerate(MAN_ROWS):
    _r["_row"] = 3 + _i
    _r["_id"] = f"MAN-{_r['serial']}"

# ---------------------------------------------------------------------------
# Hidden «تخصيص الإيرادات» (allocation; isolated)
# ---------------------------------------------------------------------------
ALC_SEC1 = [(1, "وقف عمارة النور", 390000, 45000), (2, "وقف عمارة السلام بحي السويدي", 170000, 20000),
            (3, "وقف مستودعات الشمال", 210000, 30000), (4, "وقف محلات الربوة", 95000, 8000),
            (5, "وقف أرض المنار", 28000, 0)]
ALC_SEC2 = [(1, "وقف عمارة الفجر", 140000, 15000), (2, "وقف محلات الربوة", 95000, 8000),
            (3, "وقف مدارس الحزم", 1200000, 900000), (4, "وقف شقق العزيزية", 60000, 5000)]
ALC_GRAND = dict(count=10, rev=2150000, exp=980000)  # hard-coded and wrong on purpose
ALC_ROWS_SEC1 = list(range(4, 9))     # 4..8, section total row 9
ALC_ROWS_SEC2 = list(range(13, 17))   # 13..16, section total row 17

# ---------------------------------------------------------------------------
# Hidden «مصروفات 2025 » (cross-check source; isolated)
# ---------------------------------------------------------------------------
EXP_ROWS = [(1, "عمارة النور التجارية", 40000, 20000), (3, "عمارة السلام السكنية", 15000, 10000),
            (5, "عمارة الهدى السكنية", 8000, 4000)]

LOOKUPS = OrderedDict([
    ("أنواع العقارات", ["عمارة سكنية", "عمارة تجارية", "أرض", "فيلا", "مقر", "مجمع تجاري", "نشاط تعليمي"]),
    ("التخصيص", ["استثماري", "غير استثماري", "اداري"]),
    ("طرق التملك", ["شراء", "وقف", "هبة", "إيجار أرض", "تشغيل ذاتي"]),
])


# ===========================================================================
# Workbook writer
# ===========================================================================
def _cell_write(ws, coord, value, caches, sheet_title):
    if value is None:
        return
    if isinstance(value, F):
        ws[coord] = "=" + value.formula
        caches[(sheet_title, coord)] = value.cached
    elif isinstance(value, FNC):
        ws[coord] = "=" + value.formula
    elif isinstance(value, NumText):
        ws[coord] = value.text
        ws[coord].data_type = "s"
    else:
        ws[coord] = value


def _reg_cell_value(r: dict, key: str):
    row = r["_row"]
    v = r.get(key)
    if key == "t23" and v == T23:
        return FNC(f"P{row}+Q{row}")
    if key == "t24" and v == T24:
        return F(f"S{row}+T{row}", (r["l24"] or 0) + (r["b24"] or 0))
    return v


def excel_sum(values) -> float:
    """What Excel's SUM/SUBTOTAL(109) return: numbers only (text, NumText and blanks ignored)."""
    tot = 0
    for v in values:
        if isinstance(v, F):
            v = v.cached
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            tot += v
    return tot


def build_workbook(path: str) -> dict:
    import openpyxl
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.filters import AutoFilter
    from openpyxl.worksheet.table import Table, TableColumn, TableStyleInfo

    caches = {}
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = REG
    ws.sheet_view.rightToLeft = True

    # --- register ---------------------------------------------------------------
    ws["A1"] = f"بيان أصول {ENTITY} لعام 2025م"
    ws.merge_cells(f"A1:{REG_LAST_COL}1")
    for col, header, _ in REG_COLS:
        ws[f"{col}{REG_HEADER_ROW}"] = header
    for r in REG_ROWS:
        for col, _, key in REG_COLS:
            v = r["serial"] if key == "serial" else _reg_cell_value(r, key)
            _cell_write(ws, f"{col}{r['_row']}", v, caches, REG)
        if isinstance(r.get("date"), dt.datetime):
            ws[f"{COL_OF['date']}{r['_row']}"].number_format = "yyyy-mm-dd"
    tr = REG_TOTAL_ROW
    ws[f"B{tr}"] = "الإجمالي"
    for col in REG_TOTALS_COLS:
        key = next(k for c, _, k in REG_COLS if c == col)
        vals = [_reg_cell_value(r, key) for r in REG_ROWS]
        _cell_write(ws, f"{col}{tr}", F(f"SUBTOTAL(109,{col}{REG_FIRST}:{col}{REG_LAST})", excel_sum(vals)),
                    caches, REG)
    for row in REG_FILLER:
        for col in ("A", "B", "C", "D", "E", "F"):
            ws[f"{col}{row}"] = NBSP
    table = Table(displayName="جدول_سجل_الأصول", ref=f"A{REG_HEADER_ROW}:{REG_LAST_COL}{tr}")
    table.totalsRowCount = 1
    table.autoFilter = AutoFilter(ref=f"A{REG_HEADER_ROW}:{REG_LAST_COL}{REG_LAST}")
    cols = []
    for i, (col, header, _) in enumerate(REG_COLS, start=1):
        tc = TableColumn(id=i, name=header)
        if col == "B":
            tc.totalsRowLabel = "الإجمالي"
        elif col in REG_TOTALS_COLS:
            tc.totalsRowFunction = "sum"
        cols.append(tc)
    table.tableColumns = cols
    table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(table)
    for key, src_col, n in (("atype", "A", len(LOOKUPS["أنواع العقارات"])),
                            ("alloc", "B", len(LOOKUPS["التخصيص"])),
                            ("own_method", "C", len(LOOKUPS["طرق التملك"]))):
        dv = DataValidation(type="list", formula1=f"='{LKP}'!${src_col}$2:${src_col}${n + 1}", allow_blank=True)
        dv.add(f"{COL_OF[key]}{REG_FIRST}:{COL_OF[key]}{REG_LAST}")
        ws.add_data_validation(dv)

    # --- allocation (hidden) -----------------------------------------------------
    wa = wb.create_sheet(ALC)
    wa.sheet_state = "hidden"
    wa.sheet_view.rightToLeft = True
    wa["A1"] = "توزيع إيرادات الأوقاف على المصارف (بيانات وهمية)"
    hdr = ["م", "اسم الوقف", "الإيراد السنوي", "المصروفات", "المصرف"]
    wa["A2"] = "أولا: أوقاف كفالة الأيتام"
    for j, h in enumerate(hdr):
        wa.cell(3, j + 1, h)
    for row, (s, name, rev, exp) in zip(ALC_ROWS_SEC1, ALC_SEC1):
        for j, v in enumerate((s, name, rev, exp, "كفالة الأيتام")):
            wa.cell(row, j + 1, v)
    wa["B9"] = "المجموع"
    _cell_write(wa, "C9", F("SUM(C4:C8)", sum(x[2] for x in ALC_SEC1)), caches, ALC)
    _cell_write(wa, "D9", F("SUM(D4:D8)", sum(x[3] for x in ALC_SEC1)), caches, ALC)
    wa["A11"] = "ثانيا: أوقاف التعليم"
    for j, h in enumerate(hdr):
        wa.cell(12, j + 1, h)
    for row, (s, name, rev, exp) in zip(ALC_ROWS_SEC2, ALC_SEC2):
        for j, v in enumerate((s, name, rev, exp, "برامج الأيتام التعليمية")):
            wa.cell(row, j + 1, v)
    wa["B17"] = "المجموع"
    _cell_write(wa, "C17", F("SUM(C13:C16)", sum(x[2] for x in ALC_SEC2)), caches, ALC)
    _cell_write(wa, "D17", F("SUM(D13:D16)", sum(x[3] for x in ALC_SEC2)), caches, ALC)
    wa["A19"] = ALC_GRAND["count"]
    wa["B19"] = "المجموع الكلي"
    wa["C19"] = ALC_GRAND["rev"]
    wa["D19"] = ALC_GRAND["exp"]

    # --- managed for others (hidden) ----------------------------------------------
    wm = wb.create_sheet(MAN)
    wm.sheet_state = "hidden"
    wm.sheet_view.rightToLeft = True
    wm["A1"] = "عقارات تديرها الجمعية لصالح ملاك آخرين (بيانات وهمية)"
    for j, h in enumerate(MAN_HEADERS):
        wm.cell(2, j + 1, h)
    for r in MAN_ROWS:
        vals = [r["serial"], r["name"], r["district"], r["location"], r["atype"], r["rev25"], r["rev_oct"], r["fee"],
                r["land"], r["bld"], r["total"], r["notes"], r["sup"], r["phone"]]
        for j, v in enumerate(vals):
            coord = f"{openpyxl.utils.get_column_letter(j + 1)}{r['_row']}"
            _cell_write(wm, coord, v, caches, MAN)
        wm[f"H{r['_row']}"].number_format = "0%"
    last_col = openpyxl.utils.get_column_letter(len(MAN_HEADERS))
    t2 = Table(displayName="جدول_إدارة_الأملاك", ref=f"A2:{last_col}{MAN_ROWS[-1]['_row']}")
    t2.tableStyleInfo = TableStyleInfo(name="TableStyleLight9", showRowStripes=True)
    wm.add_table(t2)

    # --- expenses (hidden, trailing space in the name) -----------------------------
    we = wb.create_sheet(EXP)
    we.sheet_state = "hidden"
    we.sheet_view.rightToLeft = True
    we["A1"] = "مصروفات تشغيل العقارات لعام 2025 (بيانات وهمية)"
    for j, h in enumerate(["م", "اسم العقار", "مصروفات الصيانة", "مصروفات الخدمات", "إجمالي المصروفات"]):
        we.cell(2, j + 1, h)
    for i, (s, name, a, b) in enumerate(EXP_ROWS):
        row = 3 + i
        we.cell(row, 1, s)
        we.cell(row, 2, name)
        we.cell(row, 3, a)
        we.cell(row, 4, b)
        _cell_write(we, f"E{row}", F(f"C{row}+D{row}", a + b), caches, EXP)
    tot = 3 + len(EXP_ROWS)
    we[f"B{tot}"] = "الإجمالي"
    for col, idx in (("C", 2), ("D", 3)):
        _cell_write(we, f"{col}{tot}", F(f"SUM({col}3:{col}{tot - 1})", sum(x[idx] for x in EXP_ROWS)), caches, EXP)
    _cell_write(we, f"E{tot}", F(f"SUM(E3:E{tot - 1})", sum(x[2] + x[3] for x in EXP_ROWS)), caches, EXP)

    # --- lookups (hidden) ---------------------------------------------------------------
    wl = wb.create_sheet(LKP)
    wl.sheet_state = "hidden"
    wl.sheet_view.rightToLeft = True
    for j, (title, items) in enumerate(LOOKUPS.items()):
        wl.cell(1, j + 1, title)
        for i, it in enumerate(items):
            wl.cell(2 + i, j + 1, it)

    wb.properties.creator = "مولّد بيانات تجريبية"
    wb.properties.lastModifiedBy = "مستخدم تجريبي"
    wb.properties.title = "دراسة أصول تجريبية (بيانات وهمية)"
    wb.properties.created = FIXED_TS
    wb.properties.modified = FIXED_TS
    wb.active = 0

    fd, tmp = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        wb.save(tmp)
        inject_caches(tmp, path, caches)
    finally:
        os.unlink(tmp)
    return caches


def _sheet_xml_paths(zf: zipfile.ZipFile) -> dict:
    wbx = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid_target = dict(re.findall(r'<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    rid_target.update({a: b for b, a in re.findall(r'<Relationship[^>]*Target="([^"]+)"[^>]*Id="([^"]+)"', rels)})
    out = {}
    for name, rid in re.findall(r'<sheet [^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbx):
        import html
        tgt = rid_target[rid].lstrip("/")
        out[html.unescape(name)] = tgt if tgt.startswith("xl/") else "xl/" + tgt
    return out


def _fmt_cache(v) -> str:
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return repr(v) if isinstance(v, float) else str(v)


def inject_caches(src: str, dst: str, caches: dict) -> None:
    """Write cached values for formula cells (openpyxl leaves <v></v> empty) and normalise timestamps."""
    with zipfile.ZipFile(src) as zin:
        paths = _sheet_xml_paths(zin)
        by_part = {}
        for (sheet, coord), v in caches.items():
            by_part.setdefault(paths[sheet], {})[coord] = v
        items = []
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename in by_part:
                xml = data.decode("utf-8")
                for coord, v in by_part[info.filename].items():
                    pat = re.compile(r'(<c r="' + coord + r'"[^>]*>\s*<f>[^<]*</f>)\s*(?:<v\s*/>|<v>\s*</v>)')
                    xml, n = pat.subn(lambda m: m.group(1) + "<v>" + _fmt_cache(v) + "</v>", xml)
                    if n != 1:
                        raise SystemExit(f"internal error: cannot inject cached value at {info.filename}!{coord}")
                data = xml.encode("utf-8")
            elif info.filename == "docProps/core.xml":
                xml = data.decode("utf-8")
                stamp = FIXED_TS.strftime("%Y-%m-%dT%H:%M:%SZ")
                xml = re.sub(r"(<dcterms:(?:created|modified)[^>]*>)[^<]*(</dcterms:)", r"\g<1>" + stamp + r"\2", xml)
                data = xml.encode("utf-8")
            items.append((info.filename, data))
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in items:
            zi = zipfile.ZipInfo(name, date_time=FIXED_TS.timetuple()[:6])
            zi.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(zi, data)


# ===========================================================================
# Expectations (independent of the skill's engine)
# ===========================================================================
KNOWN = ("value", "derived", "zero_recorded", "zero_confirmed")


def tv(raw, zero_policy="recorded"):
    """Expected typed value {s, v} for a raw cell, per data-contract §2 (explicit cases only)."""
    if isinstance(raw, F):
        raw = raw.cached
    if raw is None:
        return {"s": "blank", "v": None}
    if isinstance(raw, FNC):
        return {"s": "blank", "v": None}  # formula without cached value -> no value read
    if isinstance(raw, NumText):
        return {"s": "value", "v": raw.value}
    if raw == NA:
        return {"s": "not_applicable", "v": None}
    if isinstance(raw, (int, float)):
        if raw == 0:
            return {"s": "zero_as_missing", "v": None} if zero_policy == "missing" else {"s": "zero_recorded", "v": 0}
        return {"s": "value", "v": raw}
    raise ValueError(f"no expectation rule for {raw!r}")


def build_truth() -> list:
    """In-scope + isolated records with the fields the metrics use (period 2025 / valuation 2024)."""
    recs = []
    for r in REG_ROWS:
        f = OrderedDict()
        f["rev.2025.excl_vat"] = tv(r["rev25_ex"])
        f["rev.2025.incl_vat"] = tv(r["rev25_in"])
        f["rev.2024.actual"] = tv(r["rev24"])
        f["exp.2025.opex"] = tv(r["opex25"])
        f["val.2024.total"] = tv(_reg_cell_value(r, "t24"), "missing")
        # val.2023.total: R formulas have no cache -> derived (land + building) when both are known
        t23 = _reg_cell_value(r, "t23")
        if isinstance(t23, FNC):
            if r["l23"] is not None and r["b23"] is not None:
                f["val.2023.total"] = {"s": "derived", "v": r["l23"] + r["b23"]}
            else:
                f["val.2023.total"] = {"s": "blank", "v": None}
        else:
            f["val.2023.total"] = tv(t23, "missing")
        if r["serial"] == 9:  # documented override in 02_mapping.json
            f["rev.2025.excl_vat"] = {"s": "zero_confirmed", "v": 0}
        recs.append(OrderedDict([
            ("id", r["_id"]), ("row", r["_row"]), ("source_id", "REG"), ("name", r["name"]),
            ("record_type", r["_rt"]), ("scope", "in_scope"),
            ("dims", OrderedDict([("use", r["_use"]), ("income_nature", r["_inc"]), ("ownership", r["_own"]),
                                  ("waqf", r["_waqf"]), ("managed_for_others", False),
                                  ("asset_type", r["atype"]), ("location", r["location"])])),
            ("fields", f)]))
    for r in MAN_ROWS:
        f = OrderedDict()
        f["rev.2025.excl_vat"] = tv(r["rev25"])
        f["rev.2024.to_oct"] = tv(r["rev_oct"])
        f["val.2024.total"] = tv(r["total"], "missing")
        recs.append(OrderedDict([
            ("id", r["_id"]), ("row", r["_row"]), ("source_id", "MAN"), ("name", r["name"]),
            ("record_type", "asset"), ("scope", "in_scope"),
            ("dims", OrderedDict([("use", "investment"), ("income_nature", "rent"),
                                  ("ownership", ["managed_for_others"]), ("waqf", "unknown"),
                                  ("managed_for_others", True), ("asset_type", r["atype"]),
                                  ("location", r["location"])])),
            ("fields", f)]))
    for row, (s, name, rev, exp) in list(zip(ALC_ROWS_SEC1, ALC_SEC1)) + list(zip(ALC_ROWS_SEC2, ALC_SEC2)):
        recs.append(OrderedDict([("id", f"ALC-r{row}"), ("row", row), ("source_id", "ALC"), ("name", name),
                                 ("record_type", "other"), ("scope", "isolated"), ("dims", {}),
                                 ("fields", OrderedDict([("alloc.rev", tv(rev))]))]))
    for i, (s, name, a, b) in enumerate(EXP_ROWS):
        recs.append(OrderedDict([("id", f"EXP-{s}"), ("row", 3 + i), ("source_id", "EXP"), ("name", name),
                                 ("record_type", "other"), ("scope", "isolated"), ("dims", {}),
                                 ("fields", OrderedDict([("xchk.total", tv(a + b))]))]))
    return recs


# --- explicit metric definitions (the standard ids of compute_metrics.py --init-config) ---------------
ASSET_TYPES = ("asset", "land_parcel")


def _known(rec, field):
    t = rec["fields"].get(field)
    return t is not None and t["s"] in KNOWN


def _v(rec, field):
    return rec["fields"][field]["v"]


def _status(rec, field):
    t = rec["fields"].get(field)
    return None if t is None else t["s"]


def _d(rec, k):
    return rec["dims"].get(k)


def is_asset(r):
    return r["record_type"] in ASSET_TYPES


def not_managed(r):
    return _d(r, "managed_for_others") is not True


def m_count(R):
    return OrderedDict([("value", len(R)), ("status", "complete" if R else "empty"), ("n_scope", len(R)),
                        ("n_included", len(R)), ("included_ids", sorted(r["id"] for r in R))])


def m_sum(R, field):
    K = [r for r in R if _known(r, field)]
    U = [r for r in R if not _known(r, field)]
    st = "empty" if not R else ("not_computable" if not K else ("partial" if U else "complete"))
    return OrderedDict([("value", sum(_v(r, field) for r in K) if (R and K) else None), ("status", st),
                        ("n_scope", len(R)), ("n_included", len(K)), ("included_ids", sorted(r["id"] for r in K)),
                        ("excluded", sorted(({"id": r["id"], "reason": _status(r, field) or "absent"} for r in U),
                                            key=lambda x: x["id"]))])


def m_ratio(R, num, den, scale=100):
    I, ex = [], []
    for r in R:
        if not _known(r, num):
            ex.append({"id": r["id"], "reason": "num_unknown"})
        elif not _known(r, den):
            ex.append({"id": r["id"], "reason": "den_unknown"})
        elif not _v(r, den) > 0:
            ex.append({"id": r["id"], "reason": "den_nonpositive"})
        else:
            I.append(r)
    ns = sum(_v(r, num) for r in I)
    ds = sum(_v(r, den) for r in I)
    nkt = sum(_v(r, num) for r in R if _known(r, num))
    st = "empty" if not R else ("not_computable" if not I else ("partial" if ex else "complete"))
    return OrderedDict([("value", ns / ds * scale if I else None), ("status", st), ("n_scope", len(R)),
                        ("n_included", len(I)), ("included_ids", sorted(r["id"] for r in I)),
                        ("excluded", sorted(ex, key=lambda x: x["id"])), ("num", ns), ("den", ds),
                        ("num_known_total", nkt), ("coverage_num_share", ns / nkt if nkt else None)])


def m_diff(R, a, b):
    if not R:
        return OrderedDict([("value", None), ("status", "empty"), ("n_scope", 0), ("n_included", 0)])
    if b is None or not any(b in r["fields"] for r in R):
        return OrderedDict([("value", None), ("status", "not_computable"), ("n_scope", len(R)), ("n_included", 0)])
    I, ex = [], []
    for r in R:
        if not _known(r, a):
            ex.append({"id": r["id"], "reason": "a_unknown"})
        elif not _known(r, b):
            ex.append({"id": r["id"], "reason": "b_unknown"})
        else:
            I.append(r)
    st = "not_computable" if not I else ("partial" if ex else "complete")
    return OrderedDict([("value", sum(_v(r, a) - _v(r, b) for r in I) if I else None), ("status", st),
                        ("n_scope", len(R)), ("n_included", len(I)), ("included_ids", sorted(r["id"] for r in I)),
                        ("excluded", sorted(ex, key=lambda x: x["id"])),
                        ("a_sum", sum(_v(r, a) for r in I)), ("b_sum", sum(_v(r, b) for r in I))])


def m_growth(R, prev, curr, comparability):
    """Growth over records known in both periods; a revenue start (prev 0, curr > 0) is reported apart, outside the %."""
    if prev is None:
        return OrderedDict([("value", None), ("status", "not_computable"), ("n_scope", len(R)), ("n_included", 0)])
    if comparability == "not_comparable":
        return OrderedDict([("value", None), ("status", "not_comparable"), ("n_scope", len(R)), ("n_included", 0)])
    I, ex, oc, op, starts = [], [], [], [], []
    for r in R:
        kp, kc = _known(r, prev), _known(r, curr)
        if kp and kc and _v(r, prev) == 0 and _v(r, curr) > 0:
            starts.append(r)
            ex.append({"id": r["id"], "reason": "revenue_start"})
        elif kp and kc:
            I.append(r)
        elif kc:
            oc.append(r["id"])
            ex.append({"id": r["id"], "reason": "prev_unknown"})
        elif kp:
            op.append(r["id"])
            ex.append({"id": r["id"], "reason": "curr_unknown"})
        else:
            ex.append({"id": r["id"], "reason": "both_unknown"})
    ps = sum(_v(r, prev) for r in I)
    cs = sum(_v(r, curr) for r in I)
    pct = (cs - ps) / ps * 100 if (I and ps > 0) else None
    st = "empty" if not R else ("not_computable" if not I else ("partial" if ex else "complete"))
    if I and not ps > 0:
        st = "not_computable"   # previous total zero or negative: no growth % (never "complete" with no value)
    return OrderedDict([("value", pct), ("status", st), ("n_scope", len(R)), ("n_included", len(I)),
                        ("included_ids", sorted(r["id"] for r in I)), ("excluded", sorted(ex, key=lambda x: x["id"])),
                        ("prev_sum", ps if I else None), ("curr_sum", cs if I else None),
                        ("abs_change", cs - ps if I else None), ("revenue_start", bool(starts)),
                        ("only_curr", sorted(oc)), ("only_prev", sorted(op)),
                        ("comparability_status", comparability)])


def m_concentration(R, field, top=3):
    K = [r for r in R if _known(r, field) and _v(r, field) > 0]
    any_unknown = any(not _known(r, field) for r in R)
    total = sum(_v(r, field) for r in K)
    ranked = sorted(K, key=lambda r: (-_v(r, field), r["id"]))
    top_recs = ranked[:top]
    if not R:
        st = "empty"
    elif not K or total == 0:
        st = "not_computable"
    else:
        st = "partial" if any_unknown else "complete"
    val = sum(_v(r, field) for r in top_recs) / total * 100 if (K and total) else None
    hhi = sum((_v(r, field) / total) ** 2 for r in ranked) * 10000 if (K and total) else None
    return OrderedDict([("value", val), ("status", st), ("n_scope", len(R)), ("n_included", len(K)),
                        ("included_ids", sorted(r["id"] for r in K)), ("top_ids", [r["id"] for r in top_recs]),
                        ("total", total), ("hhi", hhi)])


def m_coverage(R, fields):
    """Per field: known / (records where the field applies); «لا ينطبق» is not missing data."""
    n = len(R)
    per = [OrderedDict([("field", f), ("n_known", sum(1 for r in R if _known(r, f))),
                        ("n", n - sum(1 for r in R if _status(r, f) == "not_applicable"))]) for f in fields]
    ratios = [p["n_known"] / p["n"] for p in per if p["n"] > 0]
    val = min(ratios) * 100 if (n and ratios) else None
    st = "empty" if not n else ("complete" if ratios else "not_computable")
    return OrderedDict([("value", val), ("status", st), ("n_scope", n), ("n_included", n), ("per_field", per)])


def as_zero_if_scope_empty(res, R_unfiltered):
    """A sum/count whose whole scope (before dashboard filters) is empty is a real 0, status complete."""
    if res["status"] == "empty" and not R_unfiltered:
        res["value"] = 0
        res["status"] = "complete"
    return res


PERIODS = OrderedDict([
    ("2025", dict(REV="rev.2025.excl_vat", EXP="exp.2025.opex", REV_PREV="rev.2024.actual", comparability="unverified")),
    ("2024", dict(REV="rev.2024.actual", EXP=None, REV_PREV=None, comparability="unverified")),
])
VAL = "val.2024.total"
# revenue/valuation ratios exist only for the valuation year (2024) and the year after
VALID_YIELD_PERIODS = ("2025", "2024")
# revenue is read from whole assets, land parcels and operating activities only (never from unit/contract rows)
REVENUE_TYPES = ("asset", "land_parcel", "operating_activity")


def compute_metrics(truth: list, filters: dict) -> OrderedDict:
    filters = dict(filters)
    period = filters.pop("period", None) or "2025"
    P = PERIODS[period]
    REV, EXP_, PREV = P["REV"], P["EXP"], P["REV_PREV"]
    scope = [r for r in truth if r["scope"] == "in_scope"]
    base = scope
    for k, v in filters.items():
        base = [r for r in base if ((v in _d(r, k)) if isinstance(_d(r, k), list) else _d(r, k) == v)]

    def rev_rec(r):
        return r["record_type"] in REVENUE_TYPES

    def zero_rev_set(recs):
        return [r for r in recs if _d(r, "use") == "investment" and is_asset(r) and not_managed(r)
                and _status(r, REV) in ("zero_recorded", "zero_confirmed")]

    def assets(recs):
        return [r for r in recs if is_asset(r) and not_managed(r)]

    def operating(recs):
        return [r for r in recs if r["record_type"] == "operating_activity"]

    def managed(recs):
        return [r for r in recs if not not_managed(r)]

    inv = [r for r in base if _d(r, "use") == "investment" and is_asset(r)
           and _d(r, "income_nature") != "operating" and not_managed(r)]
    rent = [r for r in base if _d(r, "income_nature") == "rent" and not_managed(r) and rev_rec(r)]
    out = OrderedDict()
    out["asset_count"] = as_zero_if_scope_empty(m_count(assets(base)), assets(scope))
    out["operating_records_count"] = as_zero_if_scope_empty(m_count(operating(base)), operating(scope))
    out["managed_records_count"] = as_zero_if_scope_empty(m_count(managed(base)), managed(scope))
    out["recorded_value"] = m_sum(assets(base), VAL)
    out["property_revenue"] = m_sum(rent, REV)
    out["activity_revenue"] = m_sum([r for r in base if _d(r, "income_nature") == "operating" and rev_rec(r)], REV)
    out["gross_yield_investment"] = m_ratio(inv, REV, VAL)
    out["gross_yield_owned"] = m_ratio([r for r in base if is_asset(r) and "owned" in (_d(r, "ownership") or [])
                                        and _d(r, "income_nature") != "operating" and not_managed(r)], REV, VAL)
    out["gross_yield_income_producing"] = m_ratio([r for r in inv if _known(r, REV) and _v(r, REV) > 0], REV, VAL)
    if period not in VALID_YIELD_PERIODS:
        for mid in ("gross_yield_investment", "gross_yield_owned", "gross_yield_income_producing"):
            out[mid] = OrderedDict([("value", None), ("status", "not_comparable"), ("n_scope", out[mid]["n_scope"]),
                                    ("n_included", 0)])
    out["noi"] = m_diff(rent, REV, EXP_)
    out["zero_rev_investment_value"] = as_zero_if_scope_empty(m_sum(zero_rev_set(base), VAL), zero_rev_set(scope))
    out["data_coverage"] = m_coverage(assets(base), [REV, VAL])
    out["revenue_growth"] = m_growth(rent, PREV, REV, P["comparability"])
    out["value_concentration_top3"] = m_concentration(assets(base), VAL)
    out["revenue_concentration_top3"] = m_concentration([r for r in base if not_managed(r) and rev_rec(r)], REV)
    return out


# Literal, hand-computed numbers (see the arithmetic in tests/README.md) — the generator refuses to write
# expectations that disagree with them.
HAND = {
    ("2025", "asset_count"): 14,
    ("2025", "managed_records_count"): 2,
    ("2025", "operating_records_count"): 1,
    ("2025", "recorded_value"): 24_550_000,
    ("2025", "property_revenue"): 1_641_000,
    ("2025", "activity_revenue"): 1_250_000,
    ("2025", "gross_yield_investment"): 1_296_000 / 19_500_000 * 100,
    ("2025", "gross_yield_owned"): 1_296_000 / 21_900_000 * 100,
    ("2025", "gross_yield_income_producing"): 1_296_000 / 17_300_000 * 100,
    ("2025", "noi"): 599_000,
    ("2025", "zero_rev_investment_value"): 2_200_000,
    ("2025", "data_coverage"): 11 / 14 * 100,
    ("2025", "revenue_growth"): 50_000 / 1_495_000 * 100,
    ("2025", "value_concentration_top3"): 12_300_000 / 24_550_000 * 100,
    ("2025", "revenue_concentration_top3"): 1_970_000 / 2_891_000 * 100,
    ("2024", "property_revenue"): 1_495_000,
    ("2024", "gross_yield_investment"): 1_175_000 / 18_100_000 * 100,
    ("2024", "data_coverage"): 10 / 14 * 100,
}


def expected_document(truth: list) -> OrderedDict:
    default = compute_metrics(truth, {"period": "2025"})
    for (p, mid), v in HAND.items():
        got = (default if p == "2025" else compute_metrics(truth, {"period": p}))[mid]["value"]
        if got is None or abs(got - v) > 1e-9 * max(1.0, abs(v)):
            raise SystemExit(f"internal error: hand value for {p}/{mid} = {v} but data lists give {got}")
    cases = []
    for flt in ({"period": "2024"}, {"use": "investment"}, {"location": "جنوب الرياض"}, {"use": "service"}):
        full = dict(flt)
        full.setdefault("period", "2025")
        res = compute_metrics(truth, full)
        cases.append(OrderedDict([("filters", flt), ("metrics", OrderedDict(
            (mid, OrderedDict([("value", m["value"]), ("status", m["status"]), ("n_included", m["n_included"])]))
            for mid, m in res.items()))]))

    in_scope = [r for r in truth if r["scope"] == "in_scope"]
    by_type, by_scope, by_source = {}, {}, {}
    for r in truth:
        by_type[r["record_type"]] = by_type.get(r["record_type"], 0) + 1
        by_scope[r["scope"]] = by_scope.get(r["scope"], 0) + 1
        by_source[r["source_id"]] = by_source.get(r["source_id"], 0) + 1

    # trend (rent records, not managed; one point per period)
    rent = [r for r in in_scope if _d(r, "income_nature") == "rent" and not_managed(r)]
    trend = []
    for p in ("2024", "2025"):
        f = PERIODS[p]["REV"]
        k = [r for r in rent if _known(r, f)]
        trend.append(OrderedDict([("period", p), ("value", sum(_v(r, f) for r in k)), ("n_known", len(k)),
                                  ("n", len(rent))]))

    doc = OrderedDict()
    doc["contract_version"] = "1.0"
    doc["description"] = ("Hand-computed expectations for tests/_build/synthetic_portfolio.xlsx (fictional data). "
                          "Generated by tests/make_synthetic_fixture.py --write-expected and reviewed by hand; "
                          "default filters: period 2025, valuation val.2024.total.")
    doc["fixture"] = OrderedDict([("file_name", FILE_NAME), ("entity", ENTITY)])
    doc["workbook"] = OrderedDict([
        ("sheets", OrderedDict([(REG, "visible"), (ALC, "hidden"), (MAN, "hidden"), (EXP, "hidden"), (LKP, "hidden")])),
        ("tables", [OrderedDict([("sheet", REG), ("header_row", REG_HEADER_ROW),
                                 ("data_rows", [REG_FIRST, REG_LAST]), ("totals_row", REG_TOTAL_ROW)]),
                    OrderedDict([("sheet", MAN), ("header_row", 2), ("data_rows", [3, 4]), ("totals_row", None)])]),
        ("anomaly_types", sorted([
            "hidden_sheet", "filler_block", "text_in_numeric", "mixed_types", "duplicate_name", "shared_deed",
            "pii_column", "partial_period_header", "row_total_mismatch", "total_mismatch", "cached_missing",
            "count_mismatch", "hardcoded_total", "serial_gap", "garbled_header", "zero_valuation",
            "hijri_date_text", "numeric_text", "mixed_calendar"])),
        ("anomalies_at", [
            {"type": "row_total_mismatch", "sheet": REG, "cell": f"R{REG_FIRST + 1}"},
            {"type": "row_total_mismatch", "sheet": MAN, "cell": "K4"},
            {"type": "total_mismatch", "sheet": ALC, "cell": "C19"},
            {"type": "total_mismatch", "sheet": ALC, "cell": "D19"},
            {"type": "count_mismatch", "sheet": ALC, "cell": "A19"},
            {"type": "cached_missing", "sheet": REG, "cell": None},
            {"type": "partial_period_header", "sheet": MAN, "cell": "G2"},
            {"type": "text_in_numeric", "sheet": REG, "cell": None},
            {"type": "mixed_types", "sheet": REG, "cell": None},
            {"type": "duplicate_name", "sheet": REG, "cell": None},
            {"type": "duplicate_name", "sheet": ALC, "cell": None},
            {"type": "shared_deed", "sheet": REG, "cell": None},
            {"type": "serial_gap", "sheet": REG, "cell": None},
            {"type": "filler_block", "sheet": REG, "cell": None},
            {"type": "garbled_header", "sheet": MAN, "cell": "L2"},
            {"type": "zero_valuation", "sheet": MAN, "cell": None},
            {"type": "hidden_sheet", "sheet": ALC, "cell": None},
            {"type": "hidden_sheet", "sheet": MAN, "cell": None},
            {"type": "hidden_sheet", "sheet": EXP, "cell": None},
            {"type": "hidden_sheet", "sheet": LKP, "cell": None},
            {"type": "mixed_calendar", "sheet": REG, "cell": None},   # column V: Hijri text plus one Gregorian date
        ]),
        ("pii_columns", [{"sheet": REG, "col": COL_OF["sup"]}, {"sheet": REG, "col": COL_OF["phone"]},
                         {"sheet": MAN, "col": "M"}, {"sheet": MAN, "col": "N"}]),
        ("formulas_cached_missing_min", sum(1 for r in REG_ROWS if r.get("t23") == T23)),
    ])
    doc["pii_values"] = sorted({r["sup"] for r in REG_ROWS} | {r["phone"] for r in REG_ROWS}
                               | {r["sup"] for r in MAN_ROWS} | {r["phone"] for r in MAN_ROWS})
    doc["records"] = OrderedDict([
        ("count", len(truth)), ("by_type", dict(sorted(by_type.items()))), ("by_scope", dict(sorted(by_scope.items()))),
        ("by_source", dict(sorted(by_source.items()))), ("pii_columns_dropped", 4),
        ("items", OrderedDict((r["id"], OrderedDict([
            ("source_id", r["source_id"]), ("row", r["row"]), ("record_type", r["record_type"]),
            ("scope", r["scope"]), ("dims", r["dims"]), ("fields", r["fields"])])) for r in truth)),
    ])
    doc["issue_types"] = sorted([
        "valuation_total_mismatch", "vat_equal", "text_in_numeric", "zero_as_missing", "count_mixed_text",
        "hijri_conversion", "possible_duplicate_name", "shared_deed", "derived_field", "formula_no_cache",
        "field_override", "numeric_text_converted", "blank_rows_skipped"])
    doc["issues_absent"] = ["duplicate_id", "total_like_row", "missing_classification", "pii_header_not_marked",
                            "pii_redacted", "vat_incl_lt_excl", "mapping_unknown_record", "mapping_sheet_name"]
    doc["config"] = OrderedDict([
        ("default_period", "2025"), ("periods", ["2025", "2024"]),
        ("rev_field", {"2025": "rev.2025.excl_vat", "2024": "rev.2024.actual"}),
        ("exp_field", {"2025": "exp.2025.opex", "2024": None}),
        ("prev_rev_field", {"2025": "rev.2024.actual", "2024": None}),
        ("comparable_with_prev", {"2025": "unverified"}),
        ("valuation_field", VAL), ("filters", ["asset_type", "location", "use"]),
        ("trend_comparable", "unverified"),
    ])
    doc["metrics"] = default
    doc["trend"] = trend
    doc["filtered_cases"] = cases
    doc["validation"] = OrderedDict([("blocking_count", 0),
                                     # MAN-1/2 come from the hidden managed sheet, kept in scope with a reason
                                     ("expected_warnings", ["growth_comparability", "hidden_sheet_in_scope"])])
    doc["version_comparison_self"] = OrderedDict([("total_changes", 0), ("only_prev", 0), ("only_curr", 0)])
    return doc


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="يبني مصنفًا تجريبيًا وهميًا لاختبارات المهارة، مع النتائج المتوقعة المحسوبة يدويًا")
    ap.add_argument("--out", help="مسار ملف xlsx الناتج")
    ap.add_argument("--write-expected", help="كتابة النتائج المتوقعة (expected.json) إلى هذا المسار")
    ap.add_argument("--check-expected", help="مقارنة النتائج المتوقعة المولدة بملف expected.json محفوظ (خروج 1 عند الاختلاف)")
    args = ap.parse_args(argv)
    if not (args.out or args.write_expected or args.check_expected):
        ap.error("حدّد --out أو --write-expected أو --check-expected")
    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
        caches = build_workbook(args.out)
        print(f"تم إنشاء المصنف التجريبي: {args.out} ({len(caches)} صيغة بقيم مخزنة)")
    doc = json.loads(json.dumps(expected_document(build_truth()), ensure_ascii=False))
    if args.write_expected:
        with open(args.write_expected, "w", encoding="utf-8") as fh:
            json.dump(doc, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
        print(f"تم إنشاء النتائج المتوقعة: {args.write_expected}")
    if args.check_expected:
        with open(args.check_expected, encoding="utf-8") as fh:
            saved = json.load(fh)
        if saved != doc:
            diff = [k for k in doc if saved.get(k) != doc.get(k)] + [k for k in saved if k not in doc]
            print(f"خطأ: expected.json المحفوظ لا يطابق بيانات المولّد (أقسام مختلفة: {', '.join(diff)}). "
                  "أعد توليده بـ --write-expected بعد مراجعة التغيير يدويًا.", file=sys.stderr)
            return 1
        print("expected.json المحفوظ مطابق لبيانات المولّد")
    return 0


if __name__ == "__main__":
    sys.exit(main())
