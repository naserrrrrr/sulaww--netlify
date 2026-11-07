#!/usr/bin/env python3
"""Map every sheet of an Excel workbook (hidden ones included) and flag anomalies.

Writes RUN/01_workbook_map.json and RUN/01_workbook_map.md (data-contract.md §1).

Usage:
    python3 scripts/inspect_workbook.py <file.xlsx|.xlsm|.xls|.ods|.xlsb> --out RUN

Design rules:
  * Generic: no sheet names, row numbers or amounts are hard-coded.
  * The source file is opened read-only and never modified; .xls/.ods/.xlsb inputs are
    converted from a COPY with LibreOffice into RUN/_converted and the copy is inspected.
  * Personal data (PII) values are never written: only the location of PII columns.
  * The map describes; it does not decide scope. Claude reads it to author 02_mapping.json.
"""
from __future__ import annotations

import os
import sys

sys.dont_write_bytecode = True  # never leave __pycache__ inside the skill folder
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argparse  # noqa: E402
import bisect  # noqa: E402
import datetime as dt  # noqa: E402
import html  # noqa: E402
import re  # noqa: E402
import shutil  # noqa: E402
import subprocess  # noqa: E402
import tempfile  # noqa: E402
import time  # noqa: E402
import warnings  # noqa: E402
import zipfile  # noqa: E402
from collections import Counter, defaultdict  # noqa: E402

from apa_common import (  # noqa: E402
    AMBIGUOUS_DATE_NOTE, CONTRACT_VERSION, EXCEL_ERRORS, NA_TOKENS, NOT_APPLICABLE_TOKENS, REDACTED,
    clean_text, extract_years, is_filler, is_pii_header, is_total_label, looks_like_phone, mask_pii, name_key,
    normalize_arabic, parse_date, parse_number, pii_value_hits, ref, save_json, sha256_file,
)
import formula_eval as fe  # noqa: E402
from functools import lru_cache  # noqa: E402

# memoised wrappers (the same strings are normalised many times on large sheets)
normalize_arabic = lru_cache(maxsize=262144)(normalize_arabic)
name_key = lru_cache(maxsize=262144)(name_key)

SCRIPT_VERSION = "1.1.0"
TOL = fe.TOLERANCE
SEV_RANK = {"high": 0, "medium": 1, "low": 2, "info": 3}
SEV_AR = {"high": "مرتفعة", "medium": "متوسطة", "low": "منخفضة", "info": "معلوماتية"}
STATE_AR = {"visible": "ظاهرة", "hidden": "مخفية", "veryHidden": "مخفية بعمق (veryHidden)"}
DIRECT_EXT = {".xlsx", ".xlsm", ".xltx", ".xltm"}
CONVERT_EXT = {".xls", ".ods", ".xlsb", ".fods"}

TYPE_AR = {
    "formula_year_mismatch": "صيغة تجمع أعمدة سنة تختلف عن سنة عنوان عمودها",
    "cached_value_mismatch": "القيمة المخزنة لا تطابق إعادة حساب الصيغة",
    "cached_missing": "صيغ بلا قيم محسوبة مخزنة",
    "excel_error": "خطأ Excel",
    "total_mismatch": "إجمالي لا يطابق مجموع تفاصيله",
    "hardcoded_total": "إجمالي مدوّن يدويًا بلا صيغة",
    "count_mismatch": "عدد البنود في صف الإجمالي لا يطابق العدد الفعلي",
    "row_total_mismatch": "إجمالي الصف لا يساوي الأرض + المبنى للسنة نفسها",
    "derived_from_total": "مكوّن مشتق بالطرح من الإجمالي",
    "mixed_types": "أنواع بيانات مختلطة في عمود",
    "text_in_numeric": "نص في عمود عنوانه مالي/رقمي",
    "numeric_text": "أرقام مخزنة كنص",
    "vat_identical": "الإيراد بدون الضريبة يساوي الإيراد مع الضريبة",
    "vat_inconsistent": "إيراد مع الضريبة أقل من بدونها",
    "duplicate_name": "اسم مكرر داخل الورقة",
    "cross_sheet_name_match": "أسماء متطابقة بين أوراق مختلفة (للربط)",
    "shared_deed": "رقم صك مشترك بين أكثر من صف",
    "serial_gap": "فجوة أو تكرار في عمود التسلسل",
    "pii_column": "عمود بيانات شخصية",
    "partial_period_header": "عنوان عمود يدل على فترة جزئية",
    "garbled_header": "عنوان عمود مشوّه أو مكرر النص",
    "filler_block": "صفوف حشو (مسافات غير مرئية)",
    "hidden_sheet": "ورقة مخفية",
    "hijri_date_text": "تاريخ هجري مخزن كنص",
    "zero_valuation": "تقييمات مسجلة بصفر",
    "unlabeled_column": "قيم في عمود بلا عنوان",
    "outside_table_rows": "بيانات خارج نطاق جدول Excel",
    "hidden_rows": "صفوف بيانات مخفية",
    "hidden_columns": "أعمدة مخفية تحتوي بيانات",
    "mixed_calendar": "عمود تواريخ يخلط التقويمين الهجري والميلادي",
    "formula_blank_inputs": "صيغة قيمتها المخزنة صفر ومدخلاتها فارغة",
}

ROLE_AR = {
    "asset_register": "سجل أصول",
    "activity_register": "سجل نشاط تشغيلي",
    "allocation": "تخصيص/توزيع إيرادات",
    "managed_properties": "إدارة أملاك (للغير/مدارة)",
    "lookup": "قوائم مرجعية (منسدلات)",
    "notes": "ملاحظات",
    "other": "أخرى",
}

KIND_AR = {"data": "بيانات", "header": "عناوين", "title": "عنوان/ملاحظة", "section_label": "عنوان قسم",
           "total_formula": "إجمالي بصيغة", "total_label": "إجمالي مدوّن", "filler": "حشو", "blank": "فارغ"}

# --- regexes applied to normalize_arabic() text -----------------------------------------
# Total/subtotal row labels: apa_common.TOTAL_LABEL_RE / is_total_label (shared with normalize.py).
TOTAL_HDR_RE = re.compile(r"(اجمالي|مجموع|\btotal\b)")
SECTION_RE = re.compile(r"^(اولا|ثانيا|ثالثا|رابعا|خامسا|سادسا|سابعا|ثامنا|تاسعا|عاشرا|الحادي عشر|الثاني عشر"
                        r"|القسم|الجزء|المجموعه|الفئه|section|part)\b")
SERIAL_EXACT = {"م", "م.", "مسلسل", "المسلسل", "تسلسل", "التسلسل", "الرقم", "رقم", "ر", "#", "no", "no.",
                "s/n", "sn", "serial", "ت", "الترقيم", "ترقيم", "رقم تسلسلي", "الرقم التسلسلي"}
SERIAL_PREFIX = ("الترقيم", "ترقيم", "تسلسل", "التسلسل", "الرقم التسلسلي", "رقم تسلسلي", "م ")
NAME_HDR_RE = re.compile(r"(اسم|البيان|الاصل|العقار|الوصف|المسمي|\bname\b|asset|property|description)")
NOT_NAME_HDR_RE = re.compile(r"(نوع|تصنيف|الحي|الموقع|المدينه|حاله|طريقه|تخصيص|تقييد|نشاط|رقم|تاريخ|رابط|ملاحظ"
                             r"|تقييم|قيمه|ايراد|ايجار|مساحه|مسطح|عدد|اجمالي|مجموع|صك)")
MONEY_HDR_RE = re.compile(r"(صافي|قيمه|ايراد|تقييم|مبلغ|مصروف|مصاريف|ايجار|اجمالي|مجموع|سعر|تكلفه|تكاليف|دخل|ريال"
                          r"|رصيد|amount|value|revenue|income|cost|price|rent)")
COUNT_HDR_RE = re.compile(r"^(ال)?عدد")
AREA_HDR_RE = re.compile(r"(مساحه|مسطح|م2|متر مربع|\barea\b)")
VALUATION_HDR_RE = re.compile(r"(تقييم|القيمه السوقيه|قيمه العقار|valuation)")
REVENUE_HDR_RE = re.compile(r"(ايراد|ايجار|دخل|revenue|income|rent)")
# «طبيعة الدخل»، «نوع الإيراد»، «تصنيف الأصل»، «حالة التقييم»: a classification of an amount, not an amount.
CLASSIFIER_HDR_RE = re.compile(r"(^|[\s/(\-])(ال)?(طبيعه|نوع|تصنيف|حاله)($|[\s/)\-])"
                               r"|\b(type|category|status|nature)\b")
EXPENSE_HDR_RE = re.compile(r"(مصروف|مصاريف|تكلفه|تكاليف|expense|cost)")
DATE_HDR_RE = re.compile(r"(تاريخ|\bdate\b)")
VAT_EXCL_RE = re.compile(r"((بدون|غير شامل|قبل|بلا|دون|غير متضمن)\s*(ال)?ضريب|excl)")
VAT_INCL_RE = re.compile(r"((مع|شامل|بعد|متضمن|شامله|بال)\s*(ال)?ضريب|incl)")
LAND_RE = re.compile(r"(ارض|اراضي|\bland\b)")
BUILDING_RE = re.compile(r"(مبني|مباني|building)")
MONTHS = ("يناير|فبراير|مارس|ابريل|مايو|يونيو|يوليو|اغسطس|سبتمبر|اكتوبر|نوفمبر|ديسمبر|محرم|صفر|ربيع|جمادي"
          "|رجب|شعبان|رمضان|شوال|القعده|الحجه")
PARTIAL_RE = re.compile(r"((حتي|لغايه|الي تاريخ|until|to date)\s*(شهر\s*)?(\d|" + MONTHS + r"))|جزيي|جزئي|\bytd\b"
                        r"|year to date|نصف سنوي|ربع سنوي|الربع (الاول|الثاني|الثالث)|\bq[1-4]\b")
DATE_TEXT_RE = re.compile(r"^\d{1,4}\s*[/\-.]\s*\d{1,2}\s*[/\-.]\s*\d{1,4}\s*(هـ|ه|م|ھ)?\.?$")
URL_RE = re.compile(r"^(https?://|www\.)", re.IGNORECASE)
LOOKUP_NAME_RE = re.compile(r"(منسدل|قوايم|قايمه|lookup|list|ترميز|اعدادات|settings|مرجع|تصنيفات)")
MANAGED_NAME_RE = re.compile(r"(اداره املاك|اداره الاملاك|املاك الغير|للغير|managed)")
ALLOC_NAME_RE = re.compile(r"(تخصيص|توزيع|allocation)")
ACTIVITY_RE = re.compile(r"(نشاط|مدارس|مدرسه|تشغيل|activity)")

# A header candidate below an accepted header (a new sub-table) needs this share of its labels above numeric/date
# columns, unless it scores at least as high as the header above it (see analyze_structure).
LATER_HEADER_MIN_CONTRAST = 0.5

_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
_NA_SET = {normalize_arabic(t) for t in NA_TOKENS} | {normalize_arabic(t) for t in NOT_APPLICABLE_TOKENS}
_EXCEL_ERRORS_UP = {e.upper() for e in EXCEL_ERRORS}
_NUMERIC_KINDS = ("N", "Z", "X")


# =========================================================================================
# Small helpers
# =========================================================================================
def die(msg: str, code: int = 1) -> None:
    sys.stderr.write(msg.rstrip() + "\n")
    sys.exit(code)


def col_letter(c: int) -> str:
    return fe.idx_to_col(c)


def fmt_num(x) -> str:
    if x is None:
        return "—"
    try:
        xf = float(x)
    except (TypeError, ValueError):
        return str(x)
    if abs(xf - round(xf)) < 1e-9:
        return f"{int(round(xf)):,}"
    return f"{xf:,.2f}"


def fmt_diff(x: float) -> str:
    return ("+" if x >= 0 else "−") + fmt_num(abs(x))


def jnum(v):
    """JSON-friendly scalar."""
    if v is None or isinstance(v, bool):
        return v
    if isinstance(v, fe.XLError):
        return v.code
    if isinstance(v, (int, float)):
        f = float(v)
        if f != f:
            return None
        return int(f) if f.is_integer() and abs(f) < 1e15 else round(f, 6)
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        return v.isoformat()
    return clean_text(v)


def redact_text(s: str) -> str:
    """Mask personal data inside free text (mobile/landline numbers, e-mails, IBANs — the block-level patterns of
    assets/pii_patterns.json). Arabic-Indic digits are matched too; the rest of the text is kept as written."""
    if not s:
        return s
    return mask_pii(s, ("block",))[0]


def short(s: str, n: int = 80) -> str:
    """Trimmed display text; phone numbers / e-mails inside free text are always masked."""
    s = redact_text(clean_text(s))
    return s if len(s) <= n else s[: n - 1] + "…"


def compress_ints(nums) -> list:
    """[1,2,3,7,9,10] -> ["1-3", 7, "9-10"]."""
    nums = sorted(set(nums))
    out, i = [], 0
    while i < len(nums):
        j = i
        while j + 1 < len(nums) and nums[j + 1] == nums[j] + 1:
            j += 1
        out.append(nums[i] if i == j else f"{nums[i]}-{nums[j]}")
        i = j + 1
    return out


def rows_text(rows, limit: int = 12) -> str:
    parts = [str(x) for x in compress_ints(rows)]
    if len(parts) > limit:
        parts = parts[:limit] + ["…"]
    return "، ".join(parts)


def area_ref(col: int, rows) -> str:
    rows = sorted(rows)
    if not rows:
        return ""
    if len(rows) == 1:
        return f"{col_letter(col)}{rows[0]}"
    return f"{col_letter(col)}{rows[0]}:{col_letter(col)}{rows[-1]}"


def sheet_key(name: str) -> str:
    return clean_text(name).lower()


def looks_like_person_name(value: str) -> bool:
    s = clean_text(value)
    if not s or not re.search(r"[^\W\d_]", s):
        return False
    n = normalize_arabic(s)
    generic = ("جمعيه", "مؤسسه", "شركه", "مكتب", "microsoft", "openpyxl", "apache", "libreoffice", "admin",
               "user", "owner", "windows", "office", "unknown", "pc", "laptop", "computer", "مستخدم")
    return not any(g in n for g in generic)


# =========================================================================================
# Cell classification
# =========================================================================================
_STR_MEMO: dict = {}


def classify_str(s: str):
    """Return (kind, numeric_value). Kinds: F filler, E error, X numeric text, H date text,
    A n/a token, T text."""
    hit = _STR_MEMO.get(s)
    if hit is not None:
        return hit
    if is_filler(s):
        res = ("F", None)
    else:
        c = clean_text(s)
        up = c.upper()
        if up in _EXCEL_ERRORS_UP or up.startswith("#REF"):
            res = ("E", None)
        else:
            t = c.translate(_AR_DIGITS)
            if any(ch.isdigit() for ch in t):
                num, st, _note = parse_number(c)
                if num is not None:
                    res = ("X", num)
                elif DATE_TEXT_RE.match(t):
                    res = ("H", None)
                elif st in ("not_available", "not_applicable"):
                    res = ("A", None)
                else:
                    res = ("T", None)
            else:
                res = ("A", None) if normalize_arabic(c) in _NA_SET else ("T", None)
    if len(_STR_MEMO) < 500_000:
        _STR_MEMO[s] = res
    return res


def classify_value(v):
    if v is None:
        return "B", None
    if isinstance(v, bool):
        return "L", None
    if isinstance(v, (int, float)):
        if v != v:
            return "E", None
        return ("Z" if v == 0 else "N"), None
    if isinstance(v, (dt.datetime, dt.date, dt.time)):
        return "D", None
    if isinstance(v, str):
        return classify_str(v)
    return "T", None


def display_value(v) -> str:
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        f = float(v)
        return str(int(f)) if f.is_integer() and abs(f) < 1e15 else repr(round(f, 6))
    if isinstance(v, dt.datetime):
        return v.date().isoformat() if v.time() == dt.time(0) else v.isoformat()
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    return redact_text(short(str(v), 60))


# =========================================================================================
# Workbook model (also implements the protocol expected by formula_eval)
# =========================================================================================
class TableMeta:
    def __init__(self, name, sheet, ref_text, min_row, max_row, min_col, max_col, header_rows, totals_rows,
                 columns, calc):
        self.name, self.sheet, self.ref = name, sheet, ref_text
        self.min_row, self.max_row, self.min_col, self.max_col = min_row, max_row, min_col, max_col
        self.header_rows, self.totals_rows = header_rows, totals_rows
        self.columns, self.calc = columns, calc

    @property
    def first_data(self):
        return self.min_row + self.header_rows

    @property
    def last_data(self):
        return self.max_row - self.totals_rows


class SheetView:
    def __init__(self, name: str, index: int, state: str):
        self.name, self.index, self.state = name, index, state
        self.sheet_type = "worksheet"
        self.rows: dict = {}          # r -> {c: raw value (formula text for formulas)}
        self.formulas: dict = {}      # (r, c) -> "=..."
        self.cached: dict = {}        # (r, c) -> cached value of formula cells
        self.kinds: dict = {}         # r -> {c: kind char}
        self.numval: dict = {}        # (r, c) -> number parsed from numeric text
        self.max_row = 0
        self.max_col = 0
        self.dimensions = None
        self.hidden_rows: set = set()
        self.hidden_cols: set = set()
        self.has_filter = False
        self.autofilter = None
        self.tables: list = []
        self.merged: list = []        # (min_row, min_col, max_row, max_col)
        self.comments: list = []
        self.validations: list = []

    # formula_eval protocol ------------------------------------------------------------
    def value(self, r, c):
        if (r, c) in self.formulas:
            return self.cached.get((r, c))
        row = self.rows.get(r)
        return row.get(c) if row else None

    def formula(self, r, c):
        return self.formulas.get((r, c))

    def is_hidden_row(self, r):
        return r in self.hidden_rows

    # helpers ----------------------------------------------------------------------------
    def kind(self, r, c) -> str:
        row = self.kinds.get(r)
        return row.get(c, "B") if row else "B"

    def num(self, r, c):
        k = self.kind(r, c)
        if k in ("N", "Z"):
            return float(self.value(r, c))
        if k == "X":
            return self.numval.get((r, c))
        return None

    def text(self, r, c) -> str:
        v = self.value(r, c)
        return "" if v is None else clean_text(v if not isinstance(v, (int, float)) else display_value(v))

    def merged_top_left(self, r, c):
        for (r1, c1, r2, c2) in self.merged:
            if r1 <= r <= r2 and c1 <= c <= c2:
                return r1, c1, r2, c2
        return None


class Book:
    def __init__(self):
        self.sheets: dict = {}
        self.order: list = []
        self._lower: dict = {}
        self.tables: dict = {}

    def add(self, sv: SheetView):
        self.sheets[sv.name] = sv
        self.order.append(sv)
        self._lower.setdefault(sheet_key(sv.name), sv)

    def sheet(self, name):
        if name is None:
            return None
        s = self.sheets.get(name)
        if s is None:
            s = self._lower.get(sheet_key(name))
        return s if (s is not None and s.sheet_type == "worksheet") else None

    def table(self, name):
        return self.tables.get(str(name).lower()) if name else None

    def table_at(self, sheet, r, c):
        sv = self.sheet(sheet)
        if sv is None:
            return None
        for t in sv.tables:
            if t.min_row <= r <= t.max_row and t.min_col <= c <= t.max_col:
                return t
        return None


# =========================================================================================
# Input preparation and loading
# =========================================================================================
def prepare_input(src: str, out_dir: str):
    ext = os.path.splitext(src)[1].lower()
    if ext in DIRECT_EXT:
        return src, None
    if ext not in CONVERT_EXT:
        die(f"صيغة الملف غير مدعومة ({ext}). الصيغ المدعومة: xlsx, xlsm, xls, ods, xlsb.")
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        die("الملف بصيغة قديمة/مختلفة ويتطلب تحويله إلى xlsx، لكن LibreOffice (soffice) غير متاح. "
            "حوّل نسخة من الملف إلى xlsx يدويًا ثم أعد التشغيل.")
    conv_dir = os.path.join(out_dir, "_converted")
    copy_dir = os.path.join(conv_dir, "_source_copy")
    os.makedirs(copy_dir, exist_ok=True)
    copy_path = os.path.join(copy_dir, os.path.basename(src))
    shutil.copy2(src, copy_path)
    profile = tempfile.mkdtemp(prefix="apa_lo_profile_")
    cmd = [soffice, f"-env:UserInstallation=file://{profile}", "--headless", "--convert-to", "xlsx",
           "--outdir", conv_dir, copy_path]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
    except subprocess.TimeoutExpired:
        die("تعذر تحويل الملف: انتهت مهلة LibreOffice.")
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    target = os.path.join(conv_dir, os.path.splitext(os.path.basename(src))[0] + ".xlsx")
    if not os.path.exists(target):
        die("تعذر تحويل الملف إلى xlsx عبر LibreOffice: " + (proc.stderr or proc.stdout or "").strip()[:400])
    return target, {
        "tool": "soffice",
        "command": "soffice --headless --convert-to xlsx --outdir <RUN>/_converted <نسخة من الملف>",
        "source_copy": os.path.abspath(copy_path),
        "path": os.path.abspath(target),
        "sha256": sha256_file(target),
        "note": "فُحصت نسخة محوّلة؛ قد يعيد LibreOffice حساب الصيغ عند التحويل. الأصل لم يُعدَّل.",
    }


def read_raw_parts(path: str) -> dict:
    """Read what openpyxl drops: x14 data validations, threaded comments, external links."""
    out = {"x14_validations": {}, "threaded_comments": False, "external_links": 0, "pivot_tables": 0,
           "charts": 0, "images": 0, "vba": False}
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            out["threaded_comments"] = any(n.startswith("xl/threadedComments/") for n in names)
            out["external_links"] = sum(1 for n in names if n.startswith("xl/externalLinks/externalLink"))
            out["pivot_tables"] = sum(1 for n in names if n.startswith("xl/pivotTables/pivotTable"))
            out["charts"] = sum(1 for n in names if n.startswith("xl/charts/chart"))
            out["images"] = sum(1 for n in names if n.startswith("xl/media/"))
            out["vba"] = "xl/vbaProject.bin" in names
            wbxml = z.read("xl/workbook.xml").decode("utf-8", "replace")
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8", "replace")
            rid_target = {}
            for m in re.finditer(r"<Relationship\b[^>]*>", rels):
                tag = m.group(0)
                i = re.search(r'\bId="([^"]+)"', tag)
                t = re.search(r'\bTarget="([^"]+)"', tag)
                if i and t:
                    rid_target[i.group(1)] = t.group(1)
            for m in re.finditer(r"<(?:\w+:)?sheet\b[^>]*>", wbxml):
                tag = m.group(0)
                nm = re.search(r'\bname="([^"]*)"', tag)
                rid = re.search(r'\br:id="([^"]+)"', tag) or re.search(r'\b\w+:id="([^"]+)"', tag)
                if not nm or not rid or rid.group(1) not in rid_target:
                    continue
                target = rid_target[rid.group(1)].lstrip("/")
                if not target.startswith("xl/"):
                    target = "xl/" + target
                if target not in names:
                    continue
                data = z.read(target)
                pos = data.rfind(b"<extLst")
                if pos < 0:
                    continue
                tail = data[pos:].decode("utf-8", "replace")
                items = []
                for dv in re.finditer(r"<x14:dataValidation\b([^>]*)>(.*?)</x14:dataValidation>", tail, re.S):
                    attrs, body = dv.group(1), dv.group(2)
                    typ = re.search(r'\btype="([^"]+)"', attrs)
                    f1 = re.search(r"<x14:formula1>\s*<xm:f>(.*?)</xm:f>", body, re.S)
                    sq = re.search(r"<xm:sqref>(.*?)</xm:sqref>", body, re.S)
                    items.append({"sqref": sq.group(1).strip() if sq else None,
                                  "type": typ.group(1) if typ else None,
                                  "formula1": html.unescape(f1.group(1).strip()) if f1 else None,
                                  "source": "x14"})
                if items:
                    out["x14_validations"][html.unescape(nm.group(1))] = items
    except (zipfile.BadZipFile, KeyError, OSError):
        pass
    return out


def load_book(path: str, notes: list):
    try:
        import openpyxl
        from openpyxl.utils.cell import range_boundaries
        from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula
        from openpyxl.worksheet.worksheet import Worksheet
    except ImportError:
        die("مكتبة openpyxl غير مثبتة؛ لا يمكن قراءة ملف Excel.")
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        try:
            wb_f = openpyxl.load_workbook(path, data_only=False, keep_links=False)
            wb_v = openpyxl.load_workbook(path, data_only=True, read_only=True, keep_links=False)
        except Exception as exc:  # corrupt / encrypted / not a workbook
            die(f"تعذر فتح المصنف: {exc}")
    seen = set()
    for w in caught:
        msg = str(w.message)
        if msg in seen:
            continue
        seen.add(msg)
        ar = "تحذير من مكتبة القراءة"
        if "Data Validation" in msg:
            ar = ("قواعد التحقق/القوائم المنسدلة الموسّعة لا تقرؤها openpyxl؛ قُرئت من XML الخام حيث أمكن "
                  "(انظر data_validations لكل ورقة)")
        elif "Conditional Formatting" in msg:
            ar = "تنسيق شرطي موسّع غير مقروء (لا يؤثر على القيم)"
        elif "extension" in msg.lower():
            ar = "امتداد Excel غير مدعوم في مكتبة القراءة (قد تكون بعض الخصائص غير مقروءة)"
        notes.append({"source": "openpyxl", "message": msg, "note": ar})

    book = Book()
    for idx, ws in enumerate(wb_f._sheets):
        state = getattr(ws, "sheet_state", "visible") or "visible"
        sv = SheetView(ws.title, idx, state)
        if not isinstance(ws, Worksheet):
            sv.sheet_type = "chartsheet"
            book.add(sv)
            continue
        maxr = maxc = 0
        for (r, c), cell in ws._cells.items():
            v = cell.value
            if v is None:
                continue
            if cell.data_type == "f" or isinstance(v, (ArrayFormula, DataTableFormula)):
                text = v if isinstance(v, str) else getattr(v, "text", None)
                if not text:
                    text = "=__DATA_TABLE__()"
                if not text.startswith("="):
                    text = "=" + text
                sv.formulas[(r, c)] = text
            sv.rows.setdefault(r, {})[c] = v
            if r > maxr:
                maxr = r
            if c > maxc:
                maxc = c
            if cell.comment is not None and len(sv.comments) < 200:
                sv.comments.append({"ref": cell.coordinate, "text": redact_text(short(cell.comment.text or "", 300))})
        for (r, c), cell in ws._cells.items():
            if cell.comment is not None and cell.value is None and len(sv.comments) < 200:
                sv.comments.append({"ref": cell.coordinate, "text": redact_text(short(cell.comment.text or "", 300))})
        sv.max_row, sv.max_col = maxr, maxc
        try:
            sv.dimensions = ws.calculate_dimension()
        except Exception:
            sv.dimensions = None
        for r, d in ws.row_dimensions.items():
            if d.hidden:
                sv.hidden_rows.add(r)
        for key, d in ws.column_dimensions.items():
            if d.hidden:
                lo = d.min or fe.col_to_idx(key)
                hi = d.max or lo
                for c in range(lo, min(hi, max(maxc, lo)) + 1):
                    sv.hidden_cols.add(c)
        sv.autofilter = ws.auto_filter.ref if ws.auto_filter is not None else None
        sv.has_filter = bool(sv.autofilter)
        for rng in ws.merged_cells.ranges:
            sv.merged.append((rng.min_row, rng.min_col, rng.max_row, rng.max_col))
        for t in ws.tables.values():
            try:
                min_col, min_row, max_col, max_row = range_boundaries(t.ref)
            except Exception:
                continue
            hdr = 1 if t.headerRowCount is None else int(t.headerRowCount)
            tot = int(t.totalsRowCount or 0)
            cols, calc = [], {}
            for i, tc in enumerate(t.tableColumns):
                cols.append(tc.name)
                f = tc.calculatedColumnFormula
                if f is not None:
                    calc[min_col + i] = getattr(f, "attr_text", None) or str(f)
            tm = TableMeta(t.displayName or t.name, sv.name, t.ref, min_row, max_row, min_col, max_col,
                           hdr, tot, cols, calc)
            sv.tables.append(tm)
            book.tables[str(tm.name).lower()] = tm
            if t.name and t.name != tm.name:
                book.tables[str(t.name).lower()] = tm
        if ws.data_validations is not None:
            for dv in ws.data_validations.dataValidation:
                sv.validations.append({"sqref": str(dv.sqref), "type": dv.type, "formula1": dv.formula1,
                                       "source": "standard"})
        # cached values of formula cells (read-only data_only pass)
        if sv.formulas:
            by_row = defaultdict(list)
            for (r, c) in sv.formulas:
                by_row[r].append(c)
            last_r = max(by_row)
            last_c = max(c for (_, c) in sv.formulas)
            try:
                wsv = wb_v[ws.title]
                for r, vals in enumerate(wsv.iter_rows(min_row=1, max_row=last_r, max_col=last_c,
                                                       values_only=True), start=1):
                    cols_here = by_row.get(r)
                    if not cols_here:
                        continue
                    for c in cols_here:
                        if c - 1 < len(vals):
                            sv.cached[(r, c)] = vals[c - 1]
            except KeyError:
                pass
        # cell kinds
        for r, row in sv.rows.items():
            kr = {}
            for c, raw in row.items():
                if (r, c) in sv.formulas:
                    cv = sv.cached.get((r, c))
                    if cv is None:
                        kr[c] = "M"
                        continue
                    k, nv = classify_value(cv)
                else:
                    k, nv = classify_value(raw)
                kr[c] = k
                if k == "X":
                    sv.numval[(r, c)] = nv
            sv.kinds[r] = kr
        book.add(sv)

    props = {}
    p = wb_f.properties
    for key, attr in (("creator", "creator"), ("last_modified_by", "lastModifiedBy")):
        val = getattr(p, attr, None)
        if val:
            props[key] = REDACTED if looks_like_person_name(val) else clean_text(val)
    for key, attr in (("created", "created"), ("modified", "modified"), ("last_printed", "lastPrinted")):
        val = getattr(p, attr, None)
        if val:
            props[key] = val.isoformat() if hasattr(val, "isoformat") else str(val)
    for key, attr in (("title", "title"), ("subject", "subject"), ("keywords", "keywords"),
                      ("category", "category"), ("description", "description")):
        val = getattr(p, attr, None)
        if val:
            props[key] = redact_text(short(str(val), 200))

    defined = []
    try:
        for name, dn in wb_f.defined_names.items():
            defined.append({"name": name, "refers_to": short(str(dn.attr_text), 200), "scope": "workbook"})
        for ws in wb_f.worksheets:
            for name, dn in getattr(ws, "defined_names", {}).items():
                defined.append({"name": name, "refers_to": short(str(dn.attr_text), 200), "scope": ws.title})
    except Exception:
        pass
    try:
        wb_v.close()
    except Exception:
        pass
    return book, props, defined[:100]


# =========================================================================================
# Structure analysis (row kinds, header rows, regions, column profiles)
# =========================================================================================
class Region:
    def __init__(self, header_row, start, end, tables=None, after_table=None):
        self.header_row, self.start, self.end = header_row, start, end
        self.tables = tables or []
        self.after_table = after_table
        self.family = None
        self.header = {}


class SheetAnalysis:
    def __init__(self, sv: SheetView):
        self.sv = sv
        self.row_kind: dict = {}
        self.candidates: list = []           # (row, score, source)
        self.header_source: dict = {}        # header row -> table|scored|repeat
        self.regions: list = []
        self.region_starts: list = []
        self.families: dict = {}             # family id -> list of regions
        self.family_rows: dict = {}          # family id -> sorted data rows
        self.family_header: dict = {}        # family id -> {col: header}
        self.primary = None
        self.profiles: dict = {}             # family id -> {col: profile dict}
        self.key_cols: dict = {}             # family id -> {"name": c, "serial": c, "deed": c}
        self.total_rows: list = []
        self.title_cells: list = []
        self.pii_cols: dict = {}             # (family, col) -> reason
        self.role = "other"
        self.role_evidence: list = []

    # header lookup for any row --------------------------------------------------------------
    def region_of(self, r):
        i = bisect.bisect_right(self.region_starts, r) - 1
        while i >= 0:
            reg = self.regions[i]
            if reg.start <= r <= reg.end:
                return reg
            if reg.header_row == r:
                return reg
            i -= 1
        return None

    def header_for(self, r, c):
        if self.row_kind.get(r) == "header":
            for reg in self.regions:
                if reg.header_row == r:
                    return reg.header.get(c)
        reg = self.region_of(r)
        if reg is not None and reg.header:
            return reg.header.get(c)
        if self.primary is not None:
            return self.family_header.get(self.primary, {}).get(c)
        return None

    def family_of_row(self, r):
        reg = self.region_of(r)
        return reg.family if reg is not None else None


def _row_features(sv: SheetView, r: int):
    ks = sv.kinds.get(r, {})
    content = sorted(c for c, k in ks.items() if k not in ("B", "F"))
    texts = [c for c in content if ks[c] == "T"]
    nums = [c for c in content if ks[c] in ("N", "Z", "X", "M")]
    return content, texts, nums


def analyze_structure(sv: SheetView, book: Book) -> SheetAnalysis:
    sa = SheetAnalysis(sv)
    rows_sorted = sorted(sv.rows)
    feats = {}
    for r in rows_sorted:
        content, texts, nums = _row_features(sv, r)
        if not content:
            sa.row_kind[r] = "filler" if sv.kinds.get(r) else "blank"
            continue
        feats[r] = (content, texts, nums)
    content_rows = [r for r in rows_sorted if r in feats]
    pos = {r: i for i, r in enumerate(content_rows)}

    table_hdr = defaultdict(list)
    table_tot = {}
    table_data = []
    for t in sv.tables:
        if t.header_rows:
            table_hdr[t.min_row].append(t)
        if t.totals_rows:
            table_tot[t.max_row] = t
        table_data.append((t.first_data, t.last_data, t))

    def in_table_data(r):
        return any(a <= r <= b for a, b, _ in table_data)

    def total_label(r):
        """Column of the row's total label. A label that follows a serial/id and a non-total name (e.g. a rent-basis
        column reading «إجمالي» on a named asset row) describes an attribute, not the row: it does not count."""
        content, texts, _ = feats[r]
        for c in texts:
            v = sv.value(r, c)
            if len(clean_text(v)) <= 45 and is_total_label(v):
                before = [c2 for c2 in content if c2 < c]
                names = [c2 for c2 in before if c2 in texts and not is_total_label(sv.value(r, c2))
                         and len(clean_text(sv.value(r, c2))) >= 3]
                serial = [c2 for c2 in before if c2 not in texts or re.fullmatch(r"[A-Za-z\u0600-\u06FF]{0,4}[-_ ]?\d{1,6}",
                                                                                   clean_text(sv.value(r, c2)))]
                if names and (serial or len(names) >= 2):
                    continue
                return c
        return None

    def label_only_total(r):
        """A total/subtotal label whose other cells are empty, Excel errors (#VALUE!…) or n/a marks — e.g. «المجموع»
        with a broken SUBTOTAL. It has no number, yet it is a total row, never an asset."""
        content, texts, nums = feats[r]
        if nums or len(texts) != 1:
            return False
        lab = total_label(r)
        if lab is None:
            return False
        return all(sv.kind(r, c) in ("E", "A") for c in content if c != lab)

    def agg_cells(r):
        """Numeric cells holding vertical aggregates of their own column above r."""
        out = []
        for c in feats[r][2]:
            f = sv.formulas.get((r, c))
            if not f:
                continue
            try:
                ast = fe.parse(f)
            except fe.NotEvaluable:
                continue
            ag = fe.aggregate_kind(ast)
            nodes = ag[1] if ag else None
            if nodes is None:
                terms = fe.additive_terms(ast)
                if terms and len(terms) >= 2 and all(s > 0 for s, _ in terms):
                    nodes = [n for _, n in terms]
                    ok = True
                    for n in nodes:
                        if n[0] != "ref" or n[1]["sheet"] is not None or n[1]["c1"] != c or n[1]["c2"] != c \
                                or n[1]["r1"] is None or n[1]["r2"] >= r:
                            ok = False
                    if ok:
                        out.append(c)
                continue
            for n in nodes:
                if n[0] == "struct":
                    out.append(c)
                    break
                spec = n[1]
                if spec["sheet"] is None and spec["c1"] == c and spec["c2"] == c and spec["r1"] is not None \
                        and spec["r2"] < r and spec["r2"] - spec["r1"] >= 1:
                    out.append(c)
                    break
        return out

    def total_row(r):
        """A total/subtotal row: a table totals row, a total label next to numbers, a label-only total, or (outside
        Excel tables) a row of vertical aggregates of its own columns."""
        if r in table_tot:
            return True
        _content, _texts, nums = feats[r]
        if not nums:
            return label_only_total(r)
        if total_label(r) is not None:
            return True
        if any((r, c) in sv.formulas for c in nums):
            ag = agg_cells(r)
            if ag and not in_table_data(r) and (len(ag) >= 2 or len(ag) >= 0.5 * len(nums)):
                return True
        return False

    # --- header candidates -----------------------------------------------------------------
    used_cols = {c for r in content_rows for c in feats[r][0]}
    width_norm = max(2, min(6, len(used_cols)))
    min_needed = 2 if len(used_cols) <= 3 else 3
    cand = []
    for r in content_rows:
        content, texts, nums = feats[r]
        if len(texts) < 2:
            continue
        tv = [normalize_arabic(sv.value(r, c)) for c in texts]
        distinct = len(set(tv))
        if distinct < min_needed or len(texts) / len(content) < 0.6:
            continue
        if nums and total_label(r) is not None:
            continue
        nxt = content_rows[pos[r] + 1: pos[r] + 9]
        fill = contrast = 0.0
        if nxt:
            hdr_cols = set(texts)
            fills = []
            num_hits, hits = Counter(), Counter()
            for r2 in nxt:
                c2, _t2, n2 = feats[r2]
                inter = hdr_cols.intersection(c2)
                fills.append(len(inter) / len(hdr_cols))
                n2s = set(n2)
                for c in inter:
                    hits[c] += 1
                    if c in n2s or sv.kind(r2, c) in ("D", "H"):
                        num_hits[c] += 1
            fill = sum(fills) / len(fills)
            contrast = sum(1 for c in hdr_cols if hits[c] and num_hits[c] / hits[c] >= 0.5) / len(hdr_cols)
        mean_len = sum(len(x) for x in tv) / len(tv)
        score = 0.35 * min(1.0, distinct / width_norm) + 0.15 * (distinct / len(texts)) + 0.25 * fill + 0.25 * contrast
        if mean_len > 50:
            score -= 0.3
        score -= 0.2 * (len(nums) / len(content))
        cand.append((r, round(max(0.0, min(1.0, score)), 3), round(contrast, 3)))

    # --- choose header rows -------------------------------------------------------------------
    # Walk down the sheet. The first candidate is the header. A later candidate is a header only when it opens a new
    # block: it is separated from the rows above it (a blank/filler row, or a title / section label / total row right
    # above it) and looks like a header (contrast ≥ LATER_HEADER_MIN_CONTRAST, or scores at least as high as the header
    # above). A text-heavy row inside a block of data rows (e.g. «4 محلات ومستودع» in a numeric column) stays data.
    selected = {}
    for r, ts in table_hdr.items():
        selected[r] = "table"
    table_end = {r: max(t.max_row for t in ts) for r, ts in table_hdr.items()}
    info = {r: (s, contrast) for r, s, contrast in cand}
    sa.candidate_contrast = {r: contrast for r, _s, contrast in cand}

    def boundary_row(p):
        """Right above a new sub-table header: a total row or a short text line (title / section label / note)."""
        content, _texts, nums = feats[p]
        return total_row(p) or (not nums and len(content) <= 2)

    def group_header(h):
        return any(r1 == h and c2 > c1 for (r1, c1, _r2, c2) in sv.merged)

    def separator_row(p):
        """Strictly not data: a total row, a section label («ثانيًا: …»، «…:») or a title merged across columns."""
        content, texts, nums = feats[p]
        if total_row(p):
            return True
        if nums or len(content) > 2 or not texts:
            return False
        t = normalize_arabic(sv.value(p, texts[0]))
        return bool(SECTION_RE.match(t)) or t.endswith(":") or group_header(p)

    eligible = sorted(r for r, s, _ in cand
                      if s >= 0.6 and r not in selected and not in_table_data(r) and r not in table_tot)
    for r in eligible:
        s, contrast = info[r]
        above = [h for h in selected if h < r]
        if not above:
            selected[r] = "scored"
            continue
        h = max(above)
        if contrast == 0:
            continue
        i_from = bisect.bisect_right(content_rows, h)      # first content row below the header above
        if not any(not separator_row(content_rows[k]) for k in range(i_from, pos[r])):
            # no data row since the header above (only totals / section labels / merged titles, or nothing):
            # both rows compete for the same block of data
            if selected[h] == "table":
                if r > table_end[h]:
                    selected[r] = "scored"
            elif i_from == pos[r] and r - 1 == h and group_header(h):
                selected[r] = "scored"              # two-tier header: merged group labels over column labels
            elif s > info.get(h, (0.0, 0.0))[0]:
                del selected[h]                     # the lower row is the better header; the upper one is a title line
                selected[r] = "scored"
            continue
        separated = (r - 1) not in feats or boundary_row(content_rows[pos[r] - 1])
        s_above = info[h][0] if h in info and selected[h] == "scored" else 0.99
        if separated and (contrast >= LATER_HEADER_MIN_CONTRAST or s >= s_above):
            selected[r] = "scored"
    # repeated headers (same texts as an accepted header)
    hdr_sets = {h: {normalize_arabic(sv.value(h, c)) for c in feats[h][1]} for h in selected if h in feats}
    for r in content_rows:
        if r in selected or in_table_data(r):
            continue
        content, texts, nums = feats[r]
        if len(texts) < 2 or nums:
            continue
        s = {normalize_arabic(sv.value(r, c)) for c in texts}
        for h, hs in hdr_sets.items():
            if hs and len(s & hs) / len(s | hs) >= 0.7:
                selected[r] = "repeat"
                break
    sa.header_source = selected
    score_of = {r: s for r, s, _ in cand}
    cands_out = []
    for r in selected:
        cands_out.append((r, 0.99 if selected[r] == "table" else score_of.get(r, 0.7), selected[r]))
    for r, s, _ in cand:
        if r not in selected and not in_table_data(r):
            cands_out.append((r, s, "scored"))
    cands_out.sort(key=lambda x: (-x[1], x[0]))
    sa.candidates = cands_out

    # --- row kinds -------------------------------------------------------------------------
    for r in content_rows:
        content, texts, nums = feats[r]
        if r in selected:
            sa.row_kind[r] = "header"
            continue
        if total_row(r):
            sa.row_kind[r] = "total_formula" if any((r, c) in sv.formulas for c in content) else "total_label"
            sa.total_rows.append(r)
            continue
        if not nums and len(content) <= 2 and texts:
            t = normalize_arabic(sv.value(r, texts[0]))
            if SECTION_RE.match(t):
                sa.row_kind[r] = "section_label"
                continue
            nxt = content_rows[pos[r] + 1] if pos[r] + 1 < len(content_rows) else None
            if t.endswith(":") and nxt in selected and any(h < r for h in selected):
                sa.row_kind[r] = "section_label"
                continue
        sa.row_kind[r] = "data"

    first_hdr = min(selected) if selected else None
    for r in content_rows:
        if sa.row_kind[r] != "data":
            continue
        content, texts, nums = feats[r]
        if nums or len(content) > 3:
            continue
        before_header = first_hdr is not None and r < first_hdr
        merged_title = False
        for c in content:
            m = sv.merged_top_left(r, c)
            if m and m[0] == r and m[1] == c and (m[3] - m[1]) >= 2:
                merged_title = True
        lone_txt = sv.text(r, content[0]) if len(content) == 1 else ""
        lone_long = bool(lone_txt) and first_hdr is not None and (
            len(lone_txt) > 60 or bool(re.match(r"^(ملاحظ|ملاحظه|\*|note|المصدر|اعداد)", normalize_arabic(lone_txt))))
        if before_header or merged_title or (first_hdr is None and r == content_rows[0] and len(content) == 1) \
                or (lone_long and not in_table_data(r)):
            sa.row_kind[r] = "title"
    for r in content_rows:
        if sa.row_kind[r] == "title":
            for c in feats[r][0]:
                if len(sa.title_cells) < 30:
                    sa.title_cells.append({"ref": f"{col_letter(c)}{r}", "text": redact_text(short(sv.text(r, c), 200))})

    # --- regions -------------------------------------------------------------------------
    hdr_rows = sorted(selected)
    regions = []
    if not hdr_rows:
        regions.append(Region(None, 1, max(1, sv.max_row)))
    else:
        if any(sa.row_kind.get(r) == "data" for r in range(1, hdr_rows[0])):
            regions.append(Region(None, 1, hdr_rows[0] - 1))
        for i, h in enumerate(hdr_rows):
            nxt = hdr_rows[i + 1] if i + 1 < len(hdr_rows) else sv.max_row + 1
            end = nxt - 1
            ts = table_hdr.get(h, [])
            if ts:
                tmax = max(t.max_row for t in ts)
                if tmax < end:
                    regions.append(Region(h, h + 1, tmax, tables=ts))
                    if any(sa.row_kind.get(r) == "data" for r in range(tmax + 1, nxt)):
                        regions.append(Region(None, tmax + 1, end, after_table=ts[0]))
                    continue
            regions.append(Region(h, h + 1, end, tables=ts))
    regions.sort(key=lambda g: (g.start if g.header_row is None else g.header_row))
    sa.regions = regions
    sa.region_starts = [g.start if g.header_row is None else g.header_row for g in regions]

    # header maps and families
    def header_map(h):
        m = {}
        span_cols = [c for c in range(1, sv.max_col + 1) if sv.kind(h, c) not in ("B", "F")]
        for c in range(1, sv.max_col + 1):
            k = sv.kind(h, c)
            if k in ("B", "F"):
                mt = sv.merged_top_left(h, c)
                if mt and (mt[0], mt[1]) != (h, c) and sv.kind(mt[0], mt[1]) not in ("B", "F"):
                    m[c] = sv.text(mt[0], mt[1])
                continue
            txt = sv.text(h, c)
            if txt:
                m[c] = txt
        if span_cols and m:
            width = max(m) - min(m) + 1
            up = h - 1
            if up >= 1 and sa.row_kind.get(up) not in ("title", "section_label", None, "blank", "filler"):
                for (r1, c1, r2, c2) in sv.merged:
                    if r1 == up and r2 == up and 2 <= (c2 - c1 + 1) <= 0.6 * width:
                        sup = sv.text(r1, c1)
                        if not sup:
                            continue
                        for c in range(c1, c2 + 1):
                            if c in m and normalize_arabic(sup) not in normalize_arabic(m[c]):
                                m[c] = f"{sup} / {m[c]}"
        return m

    fam_sets = {}
    for g in regions:
        if g.header_row is None:
            g.family = None
            continue
        g.header = header_map(g.header_row)
        s = {normalize_arabic(v) for v in g.header.values()}
        fam = None
        for fid, fs in fam_sets.items():
            if s and fs and len(s & fs) / len(s | fs) >= 0.7:
                fam = fid
                break
        if fam is None:
            fam = g.header_row
            fam_sets[fam] = s
        g.family = fam
    for g in regions:
        sa.families.setdefault(g.family, []).append(g)
        rows = sa.family_rows.setdefault(g.family, [])
        for r in range(g.start, g.end + 1):
            if sa.row_kind.get(r) == "data":
                rows.append(r)
    for fam, regs in sa.families.items():
        sa.family_rows[fam] = sorted(set(sa.family_rows.get(fam, [])))
        merged_hdr = {}
        for g in regs:
            for c, h in g.header.items():
                merged_hdr.setdefault(c, h)
        sa.family_header[fam] = merged_hdr

    # primary family: most data rows; tie -> table, score, earlier
    best = None
    for fam, rows in sa.family_rows.items():
        if fam is None:
            continue
        is_table = any(g.tables for g in sa.families[fam])
        multi_col = len(sa.family_header.get(fam, {})) >= 2
        key = (len(rows), multi_col, is_table, score_of.get(fam, 0.99 if is_table else 0), -fam)
        if best is None or key > best[0]:
            best = (key, fam)
    sa.primary = best[1] if best else None
    return sa


# =========================================================================================
# Column profiling
# =========================================================================================
_NUMERIC_MEANING_HINTS = {"serial", "valuation", "revenue", "expense", "money", "count", "area", "date", "deed",
                          "total", "land", "building"}


def _is_national_id_value(v) -> bool:
    """The whole cell is one 10-digit number starting with 1 or 2 (shared «national_id» pattern, review level)."""
    if isinstance(v, bool) or v is None:
        return False
    if isinstance(v, float):
        if not v.is_integer():
            return False
        v = int(v)
    s = str(v).translate(_AR_DIGITS).strip()
    return len(re.sub(r"\D", "", s)) == 10 and "national_id" in pii_value_hits(s, ("review",))


def column_hints(hn: str, profile: dict, url_like: int, digit_share: float = 0.0) -> list:
    hints = []
    if hn in SERIAL_EXACT or hn.startswith(SERIAL_PREFIX):
        hints.append("serial")
    if NAME_HDR_RE.search(hn) and not NOT_NAME_HDR_RE.search(hn):
        hints.append("name")
    # «طبيعة الدخل»/«نوع الإيراد»/«حالة التقييم» classify an amount; the column holds categories, not money
    classifier = bool(CLASSIFIER_HDR_RE.search(hn))
    if classifier:
        hints.append("category")
    if VALUATION_HDR_RE.search(hn) and not classifier:
        hints.append("valuation")
    if REVENUE_HDR_RE.search(hn) and not classifier:
        hints.append("revenue")
    if EXPENSE_HDR_RE.search(hn) and not classifier:
        hints.append("expense")
    if MONEY_HDR_RE.search(hn) and not classifier:
        hints.append("money")
    if COUNT_HDR_RE.search(hn):
        hints.append("count")
    if AREA_HDR_RE.search(hn):
        hints.append("area")
    if DATE_HDR_RE.search(hn) and "serial" not in hints:
        hints.append("date")
    if "صك" in hn and not COUNT_HDR_RE.search(hn) and not DATE_HDR_RE.search(hn) and \
            ("رقم" in hn or digit_share >= 0.5) and "serial" not in hints:
        hints.append("deed")
    if VAT_EXCL_RE.search(hn):
        hints.append("vat_excl")
    elif VAT_INCL_RE.search(hn):
        hints.append("vat_incl")
    if TOTAL_HDR_RE.search(hn):
        hints.append("total")
    if LAND_RE.search(hn):
        hints.append("land")
    if BUILDING_RE.search(hn):
        hints.append("building")
    if PARTIAL_RE.search(hn):
        hints.append("partial_period")
    if url_like and url_like >= 0.5 * max(1, profile.get("text", 0) + profile.get("numeric", 0)):
        hints.append("url")
    return hints


def profile_columns(sa: SheetAnalysis, fam, rows: list, header: dict) -> dict:
    sv = sa.sv
    cols = set(header)
    for r in rows:
        for c, k in sv.kinds.get(r, {}).items():
            if k != "B":
                cols.add(c)
    out = {}
    for c in sorted(cols):
        hdr = header.get(c)
        prof = Counter()
        samples, seen = [], set()
        text_keys = {}
        phone_like = nid_like = url_like = digit_cells = 0
        hijri_rows, greg_rows = [], []
        nonempty = 0
        num_vals = []
        text_examples = []
        for r in rows:
            k = sv.kind(r, c)
            if (r, c) in sv.formulas:
                prof["formula"] += 1
            if k == "B":
                prof["blank"] += 1
                continue
            if k == "F":
                prof["filler"] += 1
                continue
            v = sv.value(r, c)
            nonempty += 1
            if k in ("N", "Z"):
                prof["numeric"] += 1
                if k == "Z":
                    prof["zero"] += 1
                num_vals.append((r, float(v)))
            elif k == "X":
                prof["numeric"] += 1
                prof["numeric_text"] += 1
                nv = sv.numval.get((r, c))
                if nv == 0:
                    prof["zero"] += 1
                num_vals.append((r, nv))
            elif k == "T":
                prof["text"] += 1
                if len(text_keys) < 20000:
                    text_keys.setdefault(name_key(v), []).append(r)
                tx = redact_text(clean_text(v))
                if len(text_examples) < 5 and tx not in text_examples:
                    text_examples.append(tx)
                if URL_RE.match(clean_text(v)):
                    url_like += 1
            elif k == "D":
                prof["date"] += 1
                if isinstance(v, (dt.datetime, dt.date)):
                    greg_rows.append(r)       # an Excel date cell is always Gregorian
            elif k == "H":
                prof["date"] += 1
                prof["date_text"] += 1
                cal = parse_date(v).get("calendar")
                if cal == "hijri":
                    hijri_rows.append(r)
                elif cal == "gregorian":
                    greg_rows.append(r)
            elif k == "A":
                prof["na"] += 1
            elif k == "E":
                prof["error"] += 1
            elif k == "M":
                prof["uncached"] += 1
            elif k == "L":
                prof["bool"] += 1
            if looks_like_phone(v) or ({"email", "iban"} & set(pii_value_hits(v, ("block",)))):
                phone_like += 1
            elif (k == "X" or (k == "N" and hdr)) and _is_national_id_value(v):
                nid_like += 1         # a bare number in an unlabeled column may be an amount: text form only
            if k in ("N", "Z", "X") or (isinstance(v, str) and re.search(r"\d", v.translate(_AR_DIGITS))):
                digit_cells += 1
            if len(samples) < 3 and k not in ("M",):
                d = display_value(v)
                if d not in seen:
                    seen.add(d)
                    samples.append(d)
        hn = normalize_arabic(hdr) if hdr else ""
        profile = {key: prof.get(key, 0) for key in ("numeric", "zero", "text", "blank", "filler", "error",
                                                     "formula", "date")}
        for extra in ("numeric_text", "date_text", "na", "uncached", "bool"):
            if prof.get(extra):
                profile[extra] = prof[extra]
        groups = sum(1 for g in (profile["numeric"], profile["text"], profile["date"]) if g)
        hints = column_hints(hn, profile, url_like, digit_cells / nonempty if nonempty else 0.0)
        pii_reason = None
        if hdr and is_pii_header(hdr):
            pii_reason = "header"
        elif nonempty and phone_like >= 0.5 * nonempty:
            pii_reason = "pattern"
        elif nonempty and nid_like >= 0.5 * nonempty and not (set(hints) & _NUMERIC_MEANING_HINTS):
            # 10-digit numbers starting with 1/2 (Saudi national ID / iqama) in a column whose header gives them
            # no other meaning (not an amount, area, serial, deed or date)
            pii_reason = "pattern"
        if pii_reason:
            hints = ["pii"] + [h for h in hints if h not in ("name",)]
        hijri = len(hijri_rows)
        notes = []
        if "count" in hints and profile["numeric"] and not profile["text"]:
            notes.append("عمود عددي («عدد…»): تحقّق من تعريف العدد قبل جمعه أو استنتاج الإشغال منه "
                         "(مثل الوحدات مقابل الأدوار أو الصكوك)")
        if hijri:
            notes.append(f"تواريخ هجرية مخزنة كنص ({hijri}) — تُحوَّل بتقويم أم القرى مع حفظ النص الأصلي")
        if hijri and greg_rows:
            notes.append(f"العمود يخلط التقويمين (هجري {hijri}، ميلادي {len(greg_rows)}) — يُحدَّد التقويم لكل قيمة "
                         "ولا يُترك auto")
        if "url" in hints:
            notes.append("روابط مواقع (ليست بيانات مالية)")
        out[c] = {
            "col": col_letter(c),
            "header": hdr,
            "years_in_header": list(dict.fromkeys(extract_years(hdr))) if hdr else [],
            "profile": profile,
            "samples": ([REDACTED] if samples else []) if pii_reason else samples,
            "mixed_types": groups >= 2,
            "pii": bool(pii_reason),
            "hints": hints,
            "notes": notes,
            # internal (stripped before writing)
            "_pii_reason": pii_reason,
            "_text_keys": text_keys,
            "_text_examples": text_examples,
            "_num_vals": num_vals,
            "_hijri": hijri,
            "_hijri_rows": hijri_rows,
            "_greg_rows": greg_rows,
            "_hn": hn,
        }
    return out


def detect_key_columns(sa: SheetAnalysis, fam, profiles: dict) -> dict:
    sv = sa.sv
    rows = sa.family_rows.get(fam, [])
    keys = {}
    best = None
    for c, p in profiles.items():
        if p["pii"]:
            continue
        pr = p["profile"]
        nonblank = pr["numeric"] + pr["text"] + pr["date"]
        if pr["text"] == 0 or pr["text"] < 0.6 * nonblank:
            continue
        hn = p["_hn"]
        distinct = len(p["_text_keys"])
        avg_len = sum(len(x) for x in p["_text_examples"]) / max(1, len(p["_text_examples"]))
        score = 2.0 * distinct / max(1, pr["text"]) + min(1.0, avg_len / 25.0)
        if "اسم" in hn:
            score += 3
        elif NAME_HDR_RE.search(hn):
            score += 1.5
        if NOT_NAME_HDR_RE.search(hn) and "اسم" not in hn:
            score -= 2
        if "url" in p["hints"]:
            score -= 3
        if best is None or score > best[0]:
            best = (score, c)
    if best is not None and best[0] > 0.8:
        keys["name"] = best[1]
    for c, p in profiles.items():
        hn = p["_hn"]
        if not hn or p["pii"]:
            continue
        if "serial" in p["hints"] and "serial" not in keys:
            vals = [v for _, v in p["_num_vals"] if v is not None]
            nonblank = p["profile"]["numeric"] + p["profile"]["text"]
            if vals and sum(1 for v in vals if float(v).is_integer()) >= 0.6 * max(1, nonblank):
                keys["serial"] = c
        if "deed" in p["hints"] and "deed" not in keys and not p["pii"]:
            keys["deed"] = c
    if "serial" not in keys and rows:
        # fallback: leftmost column whose header is a serial word even if values are sparse
        for c, p in sorted(profiles.items()):
            if p["_hn"] in SERIAL_EXACT and not p["pii"]:
                keys["serial"] = c
                break
    _ = sv
    return keys


# =========================================================================================
# Anomaly collection
# =========================================================================================
class Anomalies:
    def __init__(self):
        self.items = []

    def add(self, typ, severity, sheet, where, detail, evidence, order=(0, 0)):
        self.items.append({"type": typ, "severity": severity, "where": where, "detail": detail,
                           "evidence": evidence, "_sheet": sheet, "_order": order})

    def finalize(self, book: Book):
        idx = {sv.name: sv.index for sv in book.order}
        self.items.sort(key=lambda a: (SEV_RANK.get(a["severity"], 9), idx.get(a["_sheet"], 999), a["_order"],
                                       a["type"]))
        out = []
        for i, a in enumerate(self.items, start=1):
            out.append({"id": f"WB-{i:03d}", "type": a["type"], "severity": a["severity"], "where": a["where"],
                        "detail": a["detail"], "evidence": a["evidence"]})
        return out


# ---- formulas ----------------------------------------------------------------------------
_COUNT_FUNCS = {"COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "ROWS", "COLUMNS"}
_BLANK_INPUT_KINDS = ("B", "F", "A")      # empty, whitespace-only, n/a marks («-»، «غير متوفر»)
_BLANK_SCAN_CELLS = 5000                  # larger precedent ranges are not scanned (not flagged)


def blank_input_areas(book: Book, ev: fe.Evaluator, sheet: str, r: int, c: int, memo: dict, depth: int = 0):
    """Referenced areas of the formula at sheet!(r, c) when EVERY referenced cell is blank / filler / n-a, or is
    itself a formula of that kind with a cached 0 (or no cached value); else None. Counting formulas (COUNT…) are
    excluded: a count of nothing is a true zero."""
    key = (sheet, r, c)
    if key in memo:
        return memo[key]
    memo[key] = None                      # cycle guard
    sv = book.sheet(sheet)
    text = sv.formulas.get((r, c)) if sv is not None else None
    if not text or depth > 4:
        return None
    try:
        ast = ev.parse(text)
    except fe.NotEvaluable:
        return None
    ag = fe.aggregate_kind(ast)
    if {f.upper() for f in fe.functions_used(ast)} & _COUNT_FUNCS or (ag and ag[0] in ("count", "counta")):
        return None
    nodes = fe.iter_ref_nodes(ast)
    if not nodes:
        return None
    budget = _BLANK_SCAN_CELLS
    areas = []
    for node in nodes:
        try:
            area = ev.resolve(node, sheet, r, c)
        except fe.NotEvaluable:
            return None
        osv = book.sheet(area.sheet)
        if osv is None:
            return None
        budget -= (area.r2 - area.r1 + 1) * (area.c2 - area.c1 + 1)
        if budget < 0:
            return None
        for rr in range(area.r1, area.r2 + 1):
            for cc in range(area.c1, area.c2 + 1):
                if (osv.name, rr, cc) == key:
                    continue
                k = osv.kind(rr, cc)
                if k in _BLANK_INPUT_KINDS:
                    continue
                if (rr, cc) in osv.formulas and k in ("Z", "M") and \
                        blank_input_areas(book, ev, osv.name, rr, cc, memo, depth + 1) is not None:
                    continue
                return None
        a1 = f"{col_letter(area.c1)}{area.r1}" + ("" if area.is_single() else f":{col_letter(area.c2)}{area.r2}")
        areas.append(a1 if area.sheet == sheet else ref(area.sheet, a1))
    memo[key] = areas
    return areas


def check_formulas(sa: SheetAnalysis, sas: dict, book: Book, ev: fe.Evaluator, an: Anomalies, pii_cols: set):
    sv = sa.sv
    stats = {"count": len(sv.formulas), "cached_missing": 0, "errors": 0, "evaluated": 0, "not_evaluated": 0,
             "cached_mismatch": 0}
    not_eval_reasons = Counter()
    missing, errors, mismatches, broken = [], [], [], []
    year_groups = defaultdict(list)
    derived_groups = defaultdict(list)
    blank_groups = defaultdict(list)
    blank_memo = {}
    year_flagged = set()
    for (r, c) in sorted(sv.formulas):
        text = sv.formulas[(r, c)]
        cached = sv.cached.get((r, c))
        is_pii = (sv.name, c) in pii_cols
        if cached is None:
            stats["cached_missing"] += 1
            missing.append((r, c))
        elif sv.kind(r, c) == "E":
            stats["errors"] += 1
            errors.append((r, c, clean_text(cached), text))
        if "#REF!" in text.upper() and not (cached is not None and sv.kind(r, c) == "E"):
            broken.append((r, c, text, cached))
        st, val = ev.evaluate(text, sv.name, r, c)
        if st == "ok":
            stats["evaluated"] += 1
            if cached is not None and not fe.values_equal(val, cached):
                stats["cached_mismatch"] += 1
                mismatches.append((r, c, text, cached, val, is_pii))
        else:
            stats["not_evaluated"] += 1
            not_eval_reasons[val[:40]] += 1
        # structural checks need the AST (never on PII columns)
        if is_pii:
            continue
        try:
            ast = ev.parse(text)
        except fe.NotEvaluable:
            continue
        hdr = sa.header_for(r, c)
        # --- formula_blank_inputs: a cached 0 computed only from empty cells is unknown, not zero
        if sv.kind(r, c) == "Z":
            areas = blank_input_areas(book, ev, sv.name, r, c, blank_memo)
            if areas is not None:
                blank_groups[c].append((r, text, areas))
        # --- formula_year_mismatch
        if hdr and sa.row_kind.get(r) == "data":
            years_h = set(extract_years(hdr))
            if years_h:
                refs_info = []
                for node in fe.iter_ref_nodes(ast):
                    try:
                        area = ev.resolve(node, sv.name, r, c)
                    except fe.NotEvaluable:
                        continue
                    other = sas.get(area.sheet)
                    for k, cc in enumerate(range(area.c1, min(area.c2, area.c1 + 30) + 1)):
                        if area.sheet == sv.name and cc == c:
                            continue
                        if area.struct_cols and k < len(area.struct_cols):
                            rh = clean_text(area.struct_cols[k])
                        elif other is not None:
                            rh = other.header_for(area.r1, cc)
                        else:
                            rh = None
                        refs_info.append({"sheet": area.sheet, "c": cc, "col": col_letter(cc), "header": rh,
                                          "years": extract_years(rh) if rh else []})
                year_refs = [x for x in refs_info if x["years"]]
                disjoint = [x for x in year_refs if not set(x["years"]) & years_h]
                hn = normalize_arabic(hdr)
                terms = fe.additive_terms(ast)
                flag = bool(year_refs) and (len(disjoint) == len(year_refs) or
                                            (bool(disjoint) and TOTAL_HDR_RE.search(hn) and terms is not None))
                if flag:
                    year_groups[c].append((r, text, cached, refs_info, disjoint, years_h))
                    year_flagged.add((r, c))
        # --- derived_from_total
        terms = fe.additive_terms(ast)
        if hdr and terms and len(terms) >= 2 and terms[0][0] > 0 and any(s < 0 for s, _ in terms):
            first = terms[0][1]
            try:
                a0 = ev.resolve(first, sv.name, r, c)
            except fe.NotEvaluable:
                a0 = None
            if a0 is not None and a0.is_single():
                other = sas.get(a0.sheet)
                th = (clean_text(a0.struct_cols[0]) if a0.struct_cols else
                      (other.header_for(a0.r1, a0.c1) if other else None))
                if th and TOTAL_HDR_RE.search(normalize_arabic(th)) and not TOTAL_HDR_RE.search(normalize_arabic(hdr)):
                    subs = []
                    for s, n in terms[1:]:
                        try:
                            a = ev.resolve(n, sv.name, r, c)
                        except fe.NotEvaluable:
                            continue
                        oh = (clean_text(a.struct_cols[0]) if a.struct_cols else
                              (sas[a.sheet].header_for(a.r1, a.c1) if a.sheet in sas else None))
                        subs.append({"col": col_letter(a.c1), "header": oh, "value": jnum(sv.value(a.r1, a.c1))
                                     if a.sheet == sv.name else None, "sign": s})
                    derived_groups[c].append((r, text, cached, {"col": col_letter(a0.c1), "header": th,
                                                                "value": jnum(book.sheet(a0.sheet).value(a0.r1, a0.c1))},
                                              subs))
    if not_eval_reasons:
        stats["not_evaluated_reasons"] = dict(not_eval_reasons.most_common(8))
    sa.formula_stats = stats
    sa.year_flagged = year_flagged
    sa.cached_mismatch_cells = {(m[0], m[1]) for m in mismatches}

    if broken:
        r0, c0, t0, v0 = broken[0]
        an.add("excel_error", "high", sv.name, ref(sv.name, f"{col_letter(c0)}{r0}"),
               f"{len(broken)} صيغة تحتوي مرجعًا مكسورًا #REF! (مثل {col_letter(c0)}{r0}: {short(t0, 80)})"
               + (f"؛ قيمتها المخزنة {fmt_num(jnum(v0)) if isinstance(jnum(v0), (int, float)) else jnum(v0)} لا يُعتمد عليها"
                  if v0 is not None else "") + ". المرجع الأصلي حُذف أو نُقل؛ تُصحح الصيغة قبل استخدام قيمتها.",
               {"count": len(broken), "examples": [{"ref": f"{col_letter(c)}{r}", "formula": t,
                                                   "cached": REDACTED if (sv.name, c) in pii_cols else jnum(v)}
                                                  for r, c, t, v in broken[:10]]},
               order=(r0, c0))

    # cached_missing (aggregated per sheet)
    if missing:
        recomputed = []
        for (r, c) in missing[:10]:
            st, val = ev.evaluate(sv.formulas[(r, c)], sv.name, r, c)
            recomputed.append({"ref": f"{col_letter(c)}{r}", "formula": sv.formulas[(r, c)],
                               "recomputed": (REDACTED if (sv.name, c) in pii_cols else jnum(val)) if st == "ok"
                               else None})
        share = len(missing) / max(1, len(sv.formulas))
        an.add("cached_missing", "high", sv.name, ref(sv.name, f"{col_letter(missing[0][1])}{missing[0][0]}"),
               f"{len(missing)} من {len(sv.formulas)} صيغة في الورقة بلا قيمة محسوبة مخزنة "
               f"(مثل {', '.join(x['ref'] for x in recomputed[:5])}). "
               + ("يبدو أن الملف حُفظ دون إعادة حساب (مثلًا بواسطة برنامج آلي). " if share > 0.9 else "")
               + "يلزم فتح الملف وإعادة حسابه في Excel أو LibreOffice قبل الاعتماد على قيم هذه الخلايا؛ "
                 "القيم المعاد حسابها في الدليل تقديرية من المقيِّم المصغّر.",
               {"count": len(missing), "formulas_total": len(sv.formulas), "examples": recomputed},
               order=(missing[0][0], missing[0][1]))

    # excel errors (formula results + literal error cells)
    literal = [(r, c, clean_text(sv.value(r, c)), None) for r, kr in sv.kinds.items() for c, k in kr.items()
               if k == "E" and (r, c) not in sv.formulas]
    all_err = sorted(errors + literal)
    by_col = defaultdict(list)
    for e in all_err:
        by_col[e[1]].append(e)
    for c, items in sorted(by_col.items()):
        r0, _, code0, f0 = items[0]
        hdr = sa.header_for(r0, c)
        codes = Counter(x[2] for x in items)
        an.add("excel_error", "high", sv.name, ref(sv.name, area_ref(c, [x[0] for x in items])),
               f"{len(items)} خلية في العمود {col_letter(c)}" + (f" «{short(hdr, 60)}»" if hdr else "")
               + f" تحمل خطأ Excel ({'، '.join(f'{k}×{v}' for k, v in codes.items())})"
               + (f"، مثل {col_letter(c)}{r0} = {code0}" + (f" من الصيغة {f0}" if f0 else "")) +
               ". القيم غير متاحة ولا تُعامل كصفر.",
               {"rows": compress_ints([x[0] for x in items])[:30], "codes": dict(codes),
                "examples": [{"ref": f"{col_letter(c)}{x[0]}", "value": x[2], "formula": x[3]} for x in items[:5]]},
               order=(r0, c))

    # cached_value_mismatch (aggregated per column)
    mm_col = defaultdict(list)
    for m in mismatches:
        mm_col[m[1]].append(m)
    for c, items in sorted(mm_col.items()):
        r0, _, text0, cached0, val0, is_pii = items[0]
        hdr = sa.header_for(r0, c)
        if is_pii:
            cached_s = val_s = REDACTED
        else:
            cached_s, val_s = jnum(cached0), jnum(val0)
        diff = None
        if not is_pii and isinstance(cached_s, (int, float)) and isinstance(val_s, (int, float)):
            diff = float(cached_s) - float(val_s)
        an.add("cached_value_mismatch", "high", sv.name, ref(sv.name, area_ref(c, [x[0] for x in items])),
               f"القيمة المخزنة في {col_letter(c)}{r0}" + (f" «{short(hdr, 60)}»" if hdr else "")
               + f" = {fmt_num(cached_s) if diff is not None else cached_s} لا تطابق إعادة حساب صيغتها {text0} = "
               + f"{fmt_num(val_s) if diff is not None else val_s}"
               + (f" (الفرق {fmt_diff(diff)})" if diff is not None else "")
               + (f"؛ ويتكرر ذلك في {len(items)} خلية بالعمود" if len(items) > 1 else "")
               + ". قد يكون المصنف لم يُعد حسابه بعد آخر تعديل؛ أعد الحساب في Excel وتحقق.",
               {"formula": text0, "cached": cached_s, "recomputed": val_s, "count": len(items),
                "examples": [{"ref": f"{col_letter(c)}{x[0]}", "formula": x[2],
                              "cached": REDACTED if x[5] else jnum(x[3]),
                              "recomputed": REDACTED if x[5] else jnum(x[4])} for x in items[:8]]},
               order=(r0, c))

    # formula_year_mismatch (aggregated per column)
    for c, items in sorted(year_groups.items()):
        hdr = sa.header_for(items[0][0], c)
        years_h = sorted(items[0][5])
        y_target = years_h[0] if len(years_h) == 1 else None
        examples = []
        tot_cached = tot_re = 0.0
        all_re = True
        remap_desc = {}
        comp_desc = None
        for (r, text, cached, refs_info, disjoint, _yh) in items:
            remap = {}
            reg = sa.region_of(r)
            hmap = reg.header if reg is not None else {}
            if y_target is not None:
                for x in disjoint:
                    if x["sheet"] != sv.name or not x["header"]:
                        continue
                    target = normalize_arabic(x["header"])
                    for y in x["years"]:
                        target = target.replace(str(y), str(y_target))
                    for c2, h2 in hmap.items():
                        if c2 != x["c"] and normalize_arabic(h2) == target:
                            remap[(sv.name, x["c"])] = c2
                            break
            recomputed = None
            if remap and len(remap) == len({x["c"] for x in disjoint if x["sheet"] == sv.name}):
                st, val = ev.evaluate(text, sv.name, r, c, remap=remap)
                if st == "ok" and isinstance(val, (int, float)) and not isinstance(val, bool):
                    recomputed = float(val)
            components = None
            if y_target is not None and TOTAL_HDR_RE.search(normalize_arabic(hdr)):
                comp_cols = [c2 for c2, h2 in hmap.items() if c2 != c and y_target in extract_years(h2)
                             and not TOTAL_HDR_RE.search(normalize_arabic(h2))
                             and (LAND_RE.search(normalize_arabic(h2)) or BUILDING_RE.search(normalize_arabic(h2)))]
                vals = {col_letter(c2): sv.num(r, c2) for c2 in comp_cols}
                if comp_cols and all(v is not None for v in vals.values()):
                    components = dict(vals)
                    components["sum"] = jnum(sum(vals.values()))
                    components = {k: jnum(v) for k, v in components.items()}
            if recomputed is None and components is not None:
                recomputed = float(components["sum"])
            cnum = jnum(cached) if isinstance(cached, (int, float)) and not isinstance(cached, bool) else None
            if cnum is None or recomputed is None:
                all_re = False
            else:
                tot_cached += float(cnum)
                tot_re += recomputed
            remap_desc = {col_letter(k[1]): col_letter(v) for k, v in remap.items()} or remap_desc
            comp_desc = comp_desc or components
            if len(examples) < 10:
                examples.append({"ref": f"{col_letter(c)}{r}", "formula": text, "cached": jnum(cached),
                                 "recomputed": jnum(recomputed) if recomputed is not None else None})
        r0, text0, cached0, refs0, disjoint0, _ = items[0]
        ref_years = sorted({y for x in disjoint0 for y in x["years"]})
        refs_txt = "، ".join(f"{x['col']} «{short(x['header'] or '', 50)}»" for x in disjoint0[:6])
        ex0 = examples[0]
        detail = (f"عنوان العمود {col_letter(c)} «{short(hdr, 70)}» يخص سنة {'/'.join(map(str, years_h))}، "
                  f"لكن الصيغة في {col_letter(c)}{r0} ({text0}) تشير إلى أعمدة سنة {'/'.join(map(str, ref_years))}: "
                  f"{refs_txt}.")
        if ex0["cached"] is not None:
            detail += f" القيمة المخزنة {fmt_num(ex0['cached'])}"
            if ex0["recomputed"] is not None:
                d = float(ex0["cached"]) - float(ex0["recomputed"])
                same = "، ".join(f"{k}" for k in (remap_desc.values() if remap_desc else []))
                detail += (f" بينما المقابل لسنة {y_target}" + (f" ({same})" if same else "")
                           + f" = {fmt_num(ex0['recomputed'])} (الفرق {fmt_diff(d)}).")
            else:
                detail += "."
        if len(items) > 1:
            detail += f" تتكرر المشكلة في {len(items)} صفًا"
            if all_re:
                detail += f" (مجموع المخزن {fmt_num(tot_cached)} مقابل {fmt_num(tot_re)})"
            detail += "."
        detail += " لا تُستخدم قيم هذا العمود قبل تصحيح الصيغة أو اعتماد اشتقاق صريح وموثق."
        tcalc = None
        for t in sv.tables:
            if c in t.calc and t.min_row <= r0 <= t.max_row:
                tcalc = t.calc[c]
        evidence = {"formula": text0, "cached": ex0["cached"], "recomputed": ex0["recomputed"],
                    "recomputed_basis": (f"الصيغة نفسها بأعمدة سنة {y_target}" if remap_desc else
                                         (f"أرض + مبنى لسنة {y_target}" if comp_desc else None)),
                    "header": hdr, "header_years": years_h,
                    "referenced": [{"col": x["col"], "header": x["header"], "years": x["years"]} for x in refs0[:10]],
                    "same_year_columns": remap_desc or None, "same_year_components": comp_desc,
                    "rows": compress_ints([x[0] for x in items])[:30], "count": len(items), "examples": examples}
        if tcalc:
            evidence["table_calculated_column_formula"] = tcalc
        if len(items) > 1 and all_re:
            evidence["cached_total"] = jnum(tot_cached)
            evidence["recomputed_total"] = jnum(tot_re)
        an.add("formula_year_mismatch", "high", sv.name, ref(sv.name, area_ref(c, [x[0] for x in items])),
               detail, evidence, order=(r0, c))

    # derived_from_total
    for c, items in sorted(derived_groups.items()):
        r0, text0, cached0, total0, subs0 = items[0]
        hdr = sa.header_for(r0, c)
        subs_txt = " و".join(f"{s['col']} «{short(s['header'] or '', 40)}»" for s in subs0)
        an.add("derived_from_total", "medium", sv.name, ref(sv.name, area_ref(c, [x[0] for x in items])),
               f"قيمة العمود {col_letter(c)} «{short(hdr, 60)}» في {col_letter(c)}{r0} مشتقة بالطرح من الإجمالي: "
               f"{text0} = {fmt_num(jnum(cached0))} (الإجمالي {total0['col']} = {fmt_num(total0['value'])} ناقص "
               f"{subs_txt}). ليست قيمة مُدخلة مستقلة؛ لا تُجمع مع الإجمالي ولا تُعرض كتقييم مستقل"
               + (f"؛ النمط نفسه في {len(items)} صفًا" if len(items) > 1 else "") + ".",
               {"formula": text0, "cached": jnum(cached0), "total": total0, "subtracted": subs0,
                "rows": compress_ints([x[0] for x in items])[:30], "count": len(items)},
               order=(r0, c))

    # formula_blank_inputs (aggregated per column)
    for c, items in sorted(blank_groups.items()):
        r0, text0, areas0 = items[0]
        hdr = sa.header_for(r0, c)
        rows_ = [x[0] for x in items]
        an.add("formula_blank_inputs", "medium", sv.name, ref(sv.name, area_ref(c, rows_)),
               f"صيغ في العمود {col_letter(c)}" + (f" «{short(hdr, 60)}»" if hdr else "")
               + f" (عددها {len(items)}) قيمتها المخزنة صفر وكل الخلايا التي تقرؤها فارغة أو غير متاحة: "
               f"{col_letter(c)}{r0} بالصيغة {text0} ومدخلاتها {'، '.join(areas0[:4])}"
               + (f"؛ والصفوف {rows_text(rows_, 8)}" if len(items) > 1 else "")
               + ". الصفر هنا ناتج عن خانات فارغة لا عن قيمة مسجلة؛ يُعامل كغير متاح ولا يدخل المجاميع صفرًا.",
               {"col": col_letter(c), "header": hdr, "count": len(items), "rows": compress_ints(rows_)[:30],
                "examples": [{"ref": f"{col_letter(c)}{x[0]}", "formula": x[1], "cached": 0, "inputs": x[2][:6]}
                             for x in items[:8]]},
               order=(r0, c))


# ---- totals -----------------------------------------------------------------------------
def check_totals(sa: SheetAnalysis, an: Anomalies, ev: fe.Evaluator, pii_cols: set):
    sv = sa.sv
    boundary = {"header", "total_formula", "total_label", "section_label", "title"}
    table_tot = {t.max_row: t for t in sv.tables if t.totals_rows}
    blocks = {}
    grand_children = {}
    last_grand = 0
    for t_row in sorted(sa.total_rows):
        if t_row in table_tot:
            t = table_tot[t_row]
            block = [r for r in range(t.first_data, t.last_data + 1) if sa.row_kind.get(r) == "data"]
        else:
            block = []
            r = t_row - 1
            while r >= 1:
                k = sa.row_kind.get(r, "blank")
                if k in boundary:
                    break
                if k == "data":
                    block.append(r)
                r -= 1
            block.reverse()
        blocks[t_row] = block
        if not block:
            children = [x for x in sorted(sa.total_rows) if last_grand < x < t_row and x not in grand_children
                        and blocks.get(x)]
            if children:
                grand_children[t_row] = children
                last_grand = t_row

    for t_row in sorted(sa.total_rows):
        block = blocks[t_row]
        children = grand_children.get(t_row)
        if not block and not children:
            continue
        content, texts, nums = _row_features(sv, t_row)
        fam = sa.family_of_row(t_row)
        keys = sa.key_cols.get(fam, {})
        serial_col = keys.get("serial")
        label_c = None
        for c in texts:
            if is_total_label(sv.value(t_row, c)):
                label_c = c
                break
        label = re.sub("ـ+", "", sv.text(t_row, label_c)) if label_c else ("صف إجمالي الجدول" if t_row in table_tot else "")
        hard = []
        n_items = len(block) if block else sum(len(blocks[x]) for x in children)
        for c in nums:
            if (sv.name, c) in pii_cols:
                continue
            is_f = (t_row, c) in sv.formulas
            actual = sv.num(t_row, c)
            if actual is None:
                continue
            kind = "sum"
            if is_f:
                try:
                    ast = ev.parse(sv.formulas[(t_row, c)])
                except fe.NotEvaluable:
                    continue
                ag = fe.aggregate_kind(ast)
                if ag:
                    kind = ag[0]
                else:
                    terms = fe.additive_terms(ast)
                    if not terms or any(s < 0 for s, _ in terms):
                        continue
            else:
                hard.append(c)
            hdr = sa.header_for(t_row, c)
            cell = f"{col_letter(c)}{t_row}"
            if c == serial_col:
                if not is_f or kind in ("count", "counta"):
                    if abs(actual - n_items) > TOL:
                        breakdown = " + ".join(str(len(blocks[x])) for x in children) if children else str(n_items)
                        an.add("count_mismatch", "medium", sv.name, ref(sv.name, cell),
                               f"{cell} في صف الإجمالي «{short(label, 40)}» = {fmt_num(actual)} بينما عدد البنود الفعلي "
                               f"= {n_items}" + (f" ({breakdown} في الأقسام)" if children else "")
                               + f" (الفرق {fmt_diff(actual - n_items)}). لا يُعتمد على هذا الرقم ولا على آخر رقم "
                                 "تسلسلي لعدد الأصول؛ العد يكون من السجلات المؤهلة.",
                               {"cell": cell, "value": jnum(actual), "item_count": n_items,
                                "sections": [{"total_row": x, "items": len(blocks[x])} for x in children]
                                if children else None, "formula": sv.formulas.get((t_row, c))},
                               order=(t_row, c))
                    if not is_f:
                        hard.remove(c)
                continue
            if block:
                vals = [sv.num(r, c) for r in block if sv.kind(r, c) in ("N", "Z")]
                n_numtext = sum(1 for r in block if sv.kind(r, c) == "X")
                n_text = sum(1 for r in block if sv.kind(r, c) == "T")
                counta = sum(1 for r in block if sv.kind(r, c) not in ("B", "F"))
                if kind == "sum":
                    expected = sum(vals)
                elif kind == "counta":
                    expected = float(counta)
                elif kind == "count":
                    expected = float(len(vals))
                elif kind == "average":
                    expected = sum(vals) / len(vals) if vals else None
                elif kind == "min":
                    expected = min(vals) if vals else 0.0
                else:
                    expected = max(vals) if vals else 0.0
                if expected is None:
                    continue
                if not is_f and kind == "sum" and abs(actual - expected) > TOL and abs(actual - counta) <= TOL:
                    continue  # hard-coded count of items, consistent
                if abs(actual - expected) > TOL and (t_row, c) not in getattr(sa, "cached_mismatch_cells", set()):
                    basis = {"sum": "مجموع", "counta": "عدد الخلايا غير الفارغة في", "count": "عدد الأرقام في",
                             "average": "متوسط", "min": "أصغر قيمة في", "max": "أكبر قيمة في"}[kind]
                    an.add("total_mismatch", "high", sv.name, ref(sv.name, cell),
                           f"{cell}" + (f" «{short(hdr, 50)}»" if hdr else "") + f" في صف الإجمالي «{short(label, 40)}» = "
                           f"{fmt_num(actual)}" + (" (رقم مدوّن يدويًا)" if not is_f else f" (الصيغة {sv.formulas[(t_row, c)]})")
                           + f" بينما {basis} الصفوف {block[0]}–{block[-1]} ({len(block)} صفًا) = {fmt_num(expected)} "
                           f"(الفرق {fmt_diff(actual - expected)})."
                           + (f" يوجد {n_numtext} رقم مخزن كنص لا تجمعه Excel." if n_numtext else "")
                           + (f" يوجد {n_text} خلية نصية في التفاصيل." if n_text else ""),
                           {"cell": cell, "value": jnum(actual), "recomputed": jnum(expected), "kind": kind,
                            "block_rows": [block[0], block[-1]], "n_rows": len(block), "formula": sv.formulas.get((t_row, c)),
                            "numeric_text_cells": n_numtext, "text_cells": n_text},
                           order=(t_row, c))
            else:
                child_vals = [sv.num(x, c) for x in children]
                if any(v is None for v in child_vals) or kind not in ("sum",):
                    continue
                expected = sum(child_vals)
                data_sum = sum(sv.num(r, c) or 0.0 for x in children for r in blocks[x] if sv.kind(r, c) in ("N", "Z"))
                if abs(actual - expected) > TOL:
                    parts = " + ".join(f"{col_letter(c)}{x}" for x in children)
                    an.add("total_mismatch", "high", sv.name, ref(sv.name, cell),
                           f"{cell}" + (f" «{short(hdr, 50)}»" if hdr else "") + f" في صف «{short(label, 40)}» = {fmt_num(actual)}"
                           + (" (رقم مدوّن يدويًا)" if not is_f else f" (الصيغة {sv.formulas[(t_row, c)]})")
                           + f" بينما مجموع إجماليات الأقسام ({parts}) = {fmt_num(expected)} (الفرق {fmt_diff(actual - expected)})"
                           + f"، ومجموع التفاصيل نفسها = {fmt_num(data_sum)}.",
                           {"cell": cell, "value": jnum(actual), "recomputed": jnum(expected),
                            "recomputed_from_details": jnum(data_sum), "kind": "grand_total",
                            "section_totals": [{"cell": f"{col_letter(c)}{x}", "value": jnum(sv.num(x, c))} for x in children],
                            "formula": sv.formulas.get((t_row, c))},
                           order=(t_row, c))
        if hard:
            span = f"{col_letter(min(hard))}{t_row}" + (f":{col_letter(max(hard))}{t_row}" if len(hard) > 1 else "")
            an.add("hardcoded_total", "medium", sv.name, ref(sv.name, span),
                   f"صف الإجمالي {t_row} «{short(label, 40)}» يحتوي أرقامًا مدوّنة يدويًا بلا صيغة في "
                   f"{'، '.join(f'{col_letter(c)}{t_row} = {fmt_num(sv.num(t_row, c))}' for c in hard[:8])} — "
                   "لا تتحدث تلقائيًا مع التفاصيل وتجب مطابقتها بإعادة الجمع.",
                   {"row": t_row, "cells": {f"{col_letter(c)}{t_row}": jnum(sv.num(t_row, c)) for c in hard},
                    "grand_total_of": children or None},
                   order=(t_row, min(hard)))


# ---- row-level totals (land + building) ---------------------------------------------------
def check_row_totals(sa: SheetAnalysis, an: Anomalies):
    sv = sa.sv
    for fam, header in sa.family_header.items():
        if fam is None:
            continue
        rows = sa.family_rows.get(fam, [])
        norm = {c: normalize_arabic(h) for c, h in header.items()}
        for tc, th in norm.items():
            if not TOTAL_HDR_RE.search(th):
                continue
            years = extract_years(header[tc])
            if not years:
                continue
            ys = set(years)
            land = [c for c, h in norm.items() if c != tc and not TOTAL_HDR_RE.search(h) and LAND_RE.search(h)
                    and set(extract_years(header[c])) == ys]
            bld = [c for c, h in norm.items() if c != tc and not TOTAL_HDR_RE.search(h) and BUILDING_RE.search(h)
                   and set(extract_years(header[c])) == ys]
            if len(land) != 1 or len(bld) != 1:
                continue
            lc, bc = land[0], bld[0]
            bad = []
            for r in rows:
                if (r, tc) in getattr(sa, "year_flagged", set()):
                    continue
                t, a, b = sv.num(r, tc), sv.num(r, lc), sv.num(r, bc)
                if t is None or a is None or b is None:
                    continue
                if abs(t - (a + b)) > TOL:
                    bad.append((r, t, a, b))
            if not bad:
                continue
            r0, t0, a0, b0 = bad[0]
            an.add("row_total_mismatch", "high", sv.name, ref(sv.name, area_ref(tc, [x[0] for x in bad])),
                   f"{col_letter(tc)}{r0} «{short(header[tc], 60)}» = {fmt_num(t0)} بينما {col_letter(lc)} (أرض) + "
                   f"{col_letter(bc)} (مبنى) لسنة {'/'.join(map(str, years))} = {fmt_num(a0)} + {fmt_num(b0)} = "
                   f"{fmt_num(a0 + b0)} (الفرق {fmt_diff(t0 - (a0 + b0))})"
                   + (f"؛ ويتكرر في {len(bad)} صفًا" if len(bad) > 1 else "") +
                   ". لا يُعتمد الإجمالي أو مكوّناته قبل التحقق من المصدر.",
                   {"total_col": col_letter(tc), "land_col": col_letter(lc), "building_col": col_letter(bc),
                    "year": years, "count": len(bad),
                    "examples": [{"row": r, "total": jnum(t), "land": jnum(a), "building": jnum(b),
                                  "land_plus_building": jnum(a + b), "diff": jnum(t - a - b)} for r, t, a, b in bad[:10]]},
                   order=(r0, tc))


# ---- column-level checks ------------------------------------------------------------------
def check_columns(sa: SheetAnalysis, an: Anomalies):
    sv = sa.sv
    seen_headers = set()
    for fam, profiles in sa.profiles.items():
        rows = sa.family_rows.get(fam, [])
        header = sa.family_header.get(fam, {})
        hdr_row = fam if fam is not None else None
        fam_label = "" if fam == sa.primary else (f" (الجدول الفرعي بعنوان الصف {fam})" if fam else " (صفوف بلا عنوان)")
        zero_val = {}
        for c, p in sorted(profiles.items()):
            pr = p["profile"]
            hdr = p["header"]
            hn = p["_hn"]
            L = col_letter(c)
            hdr_txt = f" «{short(hdr, 60)}»" if hdr else ""
            n = len(rows)
            first_row = rows[0] if rows else (hdr_row or 1)
            if p["pii"]:
                continue
            text_rows = [r for r in rows if sv.kind(r, c) == "T"]
            used_rows = [r for r in rows if sv.kind(r, c) not in ("B", "F")]
            # text in a money-like column
            if hdr and "money" in p["hints"] and pr["text"] > 0 and "count" not in p["hints"]:
                ex = "، ".join(f"«{short(x, 30)}»" for x in p["_text_examples"][:3])
                an.add("text_in_numeric", "high", sv.name, ref(sv.name, area_ref(c, text_rows) if text_rows else f"{L}:{L}"),
                       f"العمود {L}{hdr_txt}{fam_label} عنوانه يوحي بمبلغ، لكنه يحتوي نصوصًا في {pr['text']} من {n} صف بيانات "
                       f"(مثل {ex})" + (f" مقابل {pr['numeric']} قيمة رقمية" if pr["numeric"] else " ولا يحتوي أي رقم")
                       + ". النص ليس مبلغًا؛ لا يُستخدم هذا العمود قيمةً أو مقامًا ماليًا، ويُربط كحقل نصي إن لزم.",
                       {"col": L, "header": hdr, "text_cells": pr["text"], "numeric_cells": pr["numeric"],
                        "data_rows": n, "text_samples": [short(x, 40) for x in p["_text_examples"][:3]]},
                       order=(first_row, c))
            elif pr["numeric"] > 0 and pr["text"] > 0 and (hdr or not header):
                ex = "، ".join(f"«{short(x, 30)}»" for x in p["_text_examples"][:3])
                is_count = "count" in p["hints"]
                sev = "medium" if (is_count or pr["numeric"] >= pr["text"] or "money" in p["hints"]) else "low"
                an.add("mixed_types", sev, sv.name, ref(sv.name, area_ref(c, used_rows) if used_rows else f"{L}:{L}"),
                       f"العمود {L}{hdr_txt}{fam_label} يخلط أرقامًا ({pr['numeric']}) ونصوصًا ({pr['text']}) مثل {ex}."
                       + (" لا تُجمع الأرقام وحدها وتُعرض إجماليًا مكتملًا، ولا يُستنتج منها الإشغال." if is_count else
                          " تُعامل النصوص كقيم غير معروفة ولا تُستبدل بصفر."),
                       {"col": L, "header": hdr, "numeric_cells": pr["numeric"], "text_cells": pr["text"],
                        "text_samples": [short(x, 40) for x in p["_text_examples"][:3]]},
                       order=(first_row, c))
            elif pr["date"] > 0 and (pr["numeric"] > 0 or pr["text"] > 0):
                an.add("mixed_types", "low", sv.name, ref(sv.name, area_ref(c, used_rows) if used_rows else f"{L}:{L}"),
                       f"العمود {L}{hdr_txt}{fam_label} يخلط تواريخ ({pr['date']}) مع أرقام ({pr['numeric']}) أو نصوص "
                       f"({pr['text']}). تحقّق من صيغة التاريخ وتقويمه قبل التحويل.",
                       {"col": L, "header": hdr, "date_cells": pr["date"], "numeric_cells": pr["numeric"],
                        "text_cells": pr["text"]}, order=(first_row, c))
            if pr.get("numeric_text"):
                an.add("numeric_text", "low", sv.name, ref(sv.name, area_ref(c, [r for r in rows if sv.kind(r, c) == "X"]) or f"{L}:{L}"),
                       f"{pr['numeric_text']} خلية في العمود {L}{hdr_txt}{fam_label} تحمل أرقامًا مخزنة كنص — "
                       "لا تجمعها SUM في Excel؛ تُحوَّل عند التطبيع مع حفظ النص الأصلي.",
                       {"col": L, "header": hdr, "cells": pr["numeric_text"]}, order=(first_row, c))
            if p["_hijri"]:
                ex_rows = p["_hijri_rows"][:3]       # Hijri cells only (never a Gregorian «15/03/2019» as example)
                conv = [_date_example(sv, r, c) for r in ex_rows]
                an.add("hijri_date_text", "low", sv.name, ref(sv.name, f"{L}{ex_rows[0]}" if ex_rows else f"{L}:{L}"),
                       f"العمود {L}{hdr_txt} يحتوي {p['_hijri']} تاريخًا هجريًا مخزنًا كنص (مثل «{conv[0]['raw']}» ≈ "
                       f"{conv[0]['gregorian'] or 'غير قابل للتحويل'} ميلادي). يُحوَّل بتقويم أم القرى مع حفظ الأصل، "
                       "ولا يُقارن بتواريخ ميلادية قبل التحويل." if conv else
                       f"العمود {L}{hdr_txt} يحتوي تواريخ هجرية نصية.",
                       {"col": L, "header": hdr, "count": p["_hijri"], "examples": conv},
                       order=(ex_rows[0] if ex_rows else first_row, c))
            if p["_hijri_rows"] and p["_greg_rows"]:
                h_rows, g_rows = p["_hijri_rows"], p["_greg_rows"]
                g_ex = [_date_example(sv, r, c) for r in g_rows[:3]]
                h_ex = [_date_example(sv, r, c) for r in h_rows[:3]]
                both = sorted(h_rows + g_rows)
                an.add("mixed_calendar", "medium", sv.name, ref(sv.name, area_ref(c, both)),
                       f"العمود {L}{hdr_txt}{fam_label} يخلط التقويمين: تواريخ هجرية (عددها {len(h_rows)}، مثل "
                       f"«{h_ex[0]['raw']}» في {h_ex[0]['ref']}) وتواريخ ميلادية (عددها {len(g_rows)}، مثل "
                       f"«{g_ex[0]['raw']}» في {g_ex[0]['ref']}). لا يُربط العمود بتقويم auto ولا تُقارن قيمه قبل تحديد "
                       "تقويم كل خلية؛ ويُسجل الخلط مشكلة بيانات.",
                       {"col": L, "header": hdr, "hijri_count": len(h_rows), "gregorian_count": len(g_rows),
                        "hijri_rows": compress_ints(h_rows)[:20], "gregorian_rows": compress_ints(g_rows)[:20],
                        "hijri_examples": h_ex, "gregorian_examples": g_ex},
                       order=(both[0], c))
            ambiguous = [r for r in rows if sv.kind(r, c) == "H"
                         and (parse_date(sv.value(r, c)).get("note") or "").startswith(AMBIGUOUS_DATE_NOTE)]
            if ambiguous:
                p["notes"].append(f"{len(ambiguous)} تاريخًا نصيًا ملتبسًا (اليوم والشهر ≤ 12، مثل "
                                  f"«{clean_text(sv.value(ambiguous[0], c))}» في {L}{ambiguous[0]}) — قُرئ يوم/شهر؛ "
                                  "يُتحقق من ترتيب اليوم والشهر في المصدر")
            if "valuation" in p["hints"] and pr["zero"]:
                zero_val[c] = (pr["zero"], pr["numeric"],
                               [r for r in rows if sv.kind(r, c) == "Z" or (sv.kind(r, c) == "X" and sv.numval.get((r, c)) == 0)])
            if not hdr and header and len(header) >= 2:
                content_n = pr["numeric"] + pr["text"] + pr["date"]
                if content_n:
                    vrows = [r for r in rows if sv.kind(r, c) not in ("B", "F")]
                    ex = [display_value(sv.value(r, c)) for r in vrows[:3]]
                    an.add("unlabeled_column", "medium" if pr["numeric"] else "low", sv.name,
                           ref(sv.name, area_ref(c, vrows) if vrows else f"{L}:{L}"),
                           f"العمود {L}{fam_label} بلا عنوان لكنه يحتوي {content_n} قيمة في صفوف البيانات "
                           f"({pr['numeric']} رقمًا و{pr['text']} نصًا؛ مثل {'، '.join(ex)} في الصفوف {rows_text(vrows, 6)}). "
                           "معناها غير موثق؛ لا تُستخدم في الحسابات قبل الاستيضاح.",
                           {"col": L, "rows": compress_ints(vrows)[:20], "numeric_cells": pr["numeric"],
                            "text_cells": pr["text"], "samples": ex}, order=(vrows[0] if vrows else first_row, c))
        if zero_val:
            items = sorted(zero_val.items())
            c0 = items[0][0]
            zrows = sorted({r for _, (_, _, zr) in items for r in zr})
            desc = "، ".join(f"{col_letter(c)} ({z} من {nn})" for c, (z, nn, _) in items)
            an.add("zero_valuation", "medium", sv.name,
                   ref(sv.name, f"{col_letter(items[0][0])}{zrows[0]}:{col_letter(items[-1][0])}{zrows[-1]}"),
                   f"أعمدة تقييم تحتوي أصفارًا{fam_label}: {desc}. الصفر في التقييم يعني غالبًا «غير مقيّم» لا قيمة صفرية؛ "
                   "يُعامل كغير متاح ما لم يثبت خلافه.",
                   {"columns": {col_letter(c): {"zeros": z, "numeric": nn, "header": header.get(c),
                                                "rows": compress_ints(zr)[:20]} for c, (z, nn, zr) in items}},
                   order=(zrows[0], c0))

        # header-level checks (each header cell once)
        for g in sa.families.get(fam, []):
            if g.header_row is None:
                continue
            for c, h in g.header.items():
                key = (g.header_row, c)
                if key in seen_headers:
                    continue
                seen_headers.add(key)
                hn = normalize_arabic(h)
                cell = f"{col_letter(c)}{g.header_row}"
                if PARTIAL_RE.search(hn):
                    an.add("partial_period_header", "medium", sv.name, ref(sv.name, cell),
                           f"عنوان العمود {cell} «{short(h, 70)}» يدل على فترة جزئية (ليست سنة كاملة). لا يُقارن بسنة كاملة "
                           "ولا يُحوَّل إلى أداء سنوي محقق؛ أي تقدير سنوي يُعرض منفصلًا بافتراضاته.",
                           {"cell": cell, "header": h, "years": extract_years(h)}, order=(g.header_row, c))
                garbled = _garbled(h)
                if garbled:
                    an.add("garbled_header", "low", sv.name, ref(sv.name, cell),
                           f"عنوان العمود {cell} «{short(h, 90)}» مشوّه بتكرار نصي ({garbled['reason']}). المعنى المرجّح: "
                           f"«{short(garbled['suggested'], 60)}» — يُتحقق منه قبل ربط العمود.",
                           {"cell": cell, "header": h, **garbled}, order=(g.header_row, c))

        # vat pairs
        excl = [c for c, p in profiles.items() if "vat_excl" in p["hints"] and not p["pii"]]
        incl = [c for c, p in profiles.items() if "vat_incl" in p["hints"] and not p["pii"]]
        for ce in excl:
            ye = set(profiles[ce]["years_in_header"])
            for ci in incl:
                if set(profiles[ci]["years_in_header"]) != ye:
                    continue
                pairs = [(r, sv.num(r, ce), sv.num(r, ci)) for r in rows]
                pairs = [(r, a, b) for r, a, b in pairs if a is not None and b is not None]
                if not pairs:
                    continue
                equal = [x for x in pairs if abs(x[1] - x[2]) <= TOL]
                lower = [x for x in pairs if x[2] < x[1] - TOL]
                ratios = [x[2] / x[1] for x in pairs if x[1] > 0 and abs(x[1] - x[2]) > TOL]
                where = ref(sv.name, f"{col_letter(min(ce, ci))}{pairs[0][0]}:{col_letter(max(ce, ci))}{pairs[-1][0]}")
                if len(equal) == len(pairs):
                    an.add("vat_identical", "info", sv.name, where,
                           f"العمودان {col_letter(ce)} «{short(profiles[ce]['header'], 40)}» و{col_letter(ci)} "
                           f"«{short(profiles[ci]['header'], 40)}» متطابقان في جميع الصفوف ذات القيم ({len(pairs)}، "
                           f"مجموع {fmt_num(sum(x[1] for x in pairs))}). قد يعني ذلك عدم خضوع الإيراد للضريبة أو نسخ "
                           "القيمة؛ لا يُستنتج أساس الضريبة دون مستند.",
                           {"excl_col": col_letter(ce), "incl_col": col_letter(ci), "rows_compared": len(pairs),
                            "identical_rows": len(equal), "sum": jnum(sum(x[1] for x in pairs))},
                           order=(pairs[0][0], ce))
                else:
                    ev_ = {"excl_col": col_letter(ce), "incl_col": col_letter(ci), "rows_compared": len(pairs),
                           "identical_rows": len(equal), "incl_lower_rows": [x[0] for x in lower][:20]}
                    if ratios:
                        ev_["ratio_min"] = round(min(ratios), 4)
                        ev_["ratio_max"] = round(max(ratios), 4)
                    if lower:
                        r0, a0, b0 = lower[0]
                        an.add("vat_inconsistent", "medium", sv.name, ref(sv.name, f"{col_letter(ci)}{r0}"),
                               f"الإيراد مع الضريبة أقل من بدونها في {len(lower)} صفًا (مثل الصف {r0}: "
                               f"{fmt_num(b0)} مقابل {fmt_num(a0)}). يُتحقق من أساس الضريبة قبل الاستخدام.",
                               ev_, order=(r0, ci))

        keys = sa.key_cols.get(fam, {})
        # serial gaps
        sc = keys.get("serial")
        if sc is not None and rows:
            _check_serial(sa, fam, sc, rows, an, fam_label)
        # duplicate names
        nc = keys.get("name")
        if nc is not None:
            p = profiles.get(nc)
            if p:
                for k, rws in p["_text_keys"].items():
                    if not k or len(rws) < 2:
                        continue
                    numeric_cols = [c for c, pp in profiles.items() if pp["profile"]["numeric"] and not pp["pii"]
                                    and c != nc and "serial" not in pp["hints"]]
                    same, diff = [], []
                    for c in numeric_cols:
                        vals = [sv.num(r, c) for r in rws]
                        if all(v is not None for v in vals):
                            (same if max(vals) - min(vals) <= TOL else diff).append(col_letter(c))
                    names = [sv.text(r, nc) for r in rws]
                    an.add("duplicate_name", "medium", sv.name, ref(sv.name, f"{col_letter(nc)}{rws[0]}"),
                           f"الاسم «{short(names[0], 60)}» مكرر في الصفوف {rows_text(rws)}{fam_label}"
                           + (f"، والقيم متطابقة في الأعمدة {', '.join(same[:8])}" if same else "")
                           + (f" ومختلفة في {', '.join(diff[:8])}" if diff else "")
                           + ". قد يكون أصلًا واحدًا مكررًا أو أصولًا مستقلة؛ لا يُدمج ولا يُحذف قبل مطابقة المعرّف والصك "
                             "والموقع والمساحة.",
                           {"col": col_letter(nc), "rows": rws[:20], "refs": [ref(sv.name, f"{col_letter(nc)}{r}") for r in rws[:20]], "names": [short(x, 80) for x in names[:6]],
                            "key": k, "identical_numeric_columns": same, "different_numeric_columns": diff},
                           order=(rws[0], nc))
        # shared deed
        dc = keys.get("deed")
        if dc is not None:
            groups = defaultdict(list)
            for r in rows:
                v = sv.value(r, dc)
                if v is None or sv.kind(r, dc) in ("B", "F", "A"):
                    continue
                t = clean_text(display_value(v)).translate(_AR_DIGITS)
                if not re.search(r"\d", t):
                    continue
                groups[re.sub(r"\s+", "", t)].append(r)
            for k, rws in groups.items():
                if len(rws) < 2:
                    continue
                nc2 = keys.get("name")
                an.add("shared_deed", "medium", sv.name, ref(sv.name, f"{col_letter(dc)}{rws[0]}"),
                       f"رقم الصك «{short(k, 40)}» مشترك بين {len(rws)} صفوف ({rows_text(rws)})"
                       + (f": {'، '.join('«' + short(sv.text(r, nc2), 40) + '»' for r in rws[:4])}" if nc2 else "")
                       + ". قد يمثل أصلًا واحدًا بعدة سجلات أو صكًا يغطي عدة قطع؛ يحتاج فصلًا أو تحققًا قبل العد والجمع.",
                       {"col": col_letter(dc), "deed": redact_text(k), "rows": rws[:20], "refs": [ref(sv.name, f"{col_letter(dc)}{r}") for r in rws[:20]]}, order=(rws[0], dc))


def _date_example(sv: SheetView, r: int, c: int) -> dict:
    v = sv.value(r, c)
    d = parse_date(v)
    return {"ref": f"{col_letter(c)}{r}", "raw": clean_text(v) if isinstance(v, str) else display_value(v),
            "calendar": d.get("calendar"), "hijri": d.get("hijri"), "gregorian": d.get("v"), "note": d.get("note")}


def _garbled(h: str):
    parts = [clean_text(p) for p in re.split(r"[,،;؛|]", h) if clean_text(p)]
    if len(parts) >= 2:
        np_ = [normalize_arabic(p) for p in parts]
        for i, a in enumerate(np_):
            for j, b in enumerate(np_):
                if i != j and len(a) >= 3 and a in b:
                    return {"reason": "أجزاء مكررة مفصولة بفواصل", "parts": parts, "suggested": parts[0] if len(np_[0]) >= max(len(x) for x in np_) else max(parts, key=len)}
    toks = normalize_arabic(h).split()
    for i in range(len(toks) - 1):
        if toks[i] == toks[i + 1] and len(toks[i]) >= 3:
            return {"reason": f"كلمة مكررة متتالية «{toks[i]}»", "parts": [h], "suggested": " ".join(toks[:i + 1] + toks[i + 2:])}
    for nlen in (3, 2):
        grams = Counter(tuple(toks[i:i + nlen]) for i in range(len(toks) - nlen + 1))
        rep = [g for g, k in grams.items() if k >= 2 and sum(len(x) for x in g) >= 6]
        if rep:
            return {"reason": f"عبارة مكررة «{' '.join(rep[0])}»", "parts": [h], "suggested": h}
    return None


def _check_serial(sa: SheetAnalysis, fam, sc: int, rows: list, an: Anomalies, fam_label: str):
    sv = sa.sv
    # split rows into sections at any non-data boundary
    sections, cur = [], []
    prev = None
    for r in rows:
        if prev is not None and any(sa.row_kind.get(x) in ("header", "total_formula", "total_label", "section_label")
                                    for x in range(prev + 1, r)):
            sections.append(cur)
            cur = []
        cur.append(r)
        prev = r
    if cur:
        sections.append(cur)
    issues = []
    sec_info = []
    for sec in sections:
        vals = [(r, sv.num(r, sc)) for r in sec]
        vals = [(r, v) for r, v in vals if v is not None and float(v).is_integer()]
        if not vals:
            continue
        seq = [int(v) for _, v in vals]
        sec_info.append({"rows": [sec[0], sec[-1]], "first": seq[0], "last": seq[-1], "n_rows": len(sec),
                         "n_serials": len(seq)})
        counts = Counter(seq)
        dupes = sorted(v for v, k in counts.items() if k > 1)
        missing = sorted(set(range(min(seq), max(seq) + 1)) - set(seq))
        backwards = [vals[i][0] for i in range(1, len(seq)) if seq[i] < seq[i - 1]]
        if dupes or missing or backwards:
            issues.append({"rows": [sec[0], sec[-1]], "duplicates": dupes[:20], "missing": compress_ints(missing)[:20],
                           "out_of_order_rows": backwards[:20], "first": seq[0], "last": seq[-1],
                           "n_rows": len(sec)})
    sa.serial_sections = sa.__dict__.get("serial_sections", {})
    sa.serial_sections[fam] = sec_info
    for it in issues:
        parts = []
        if it["missing"]:
            parts.append(f"أرقام مفقودة {', '.join(map(str, it['missing'][:8]))}")
        if it["duplicates"]:
            parts.append(f"أرقام مكررة {', '.join(map(str, it['duplicates'][:8]))}")
        if it["out_of_order_rows"]:
            parts.append(f"تسلسل غير مرتب في الصفوف {rows_text(it['out_of_order_rows'], 6)}")
        an.add("serial_gap", "low", sv.name, ref(sv.name, f"{col_letter(sc)}{it['rows'][0]}:{col_letter(sc)}{it['rows'][1]}"),
               f"عمود التسلسل {col_letter(sc)}{fam_label} في الصفوف {it['rows'][0]}–{it['rows'][1]}: {'؛ '.join(parts)} "
               f"(من {it['first']} إلى {it['last']} مقابل {it['n_rows']} صف بيانات). لا يُستنتج عدد الأصول من آخر رقم تسلسلي.",
               {"col": col_letter(sc), **it}, order=(it["rows"][0], sc))


def check_sheet_level(sa: SheetAnalysis, an: Anomalies):
    sv = sa.sv
    nonempty = sum(1 for r, k in sa.row_kind.items() if k not in ("blank", "filler"))
    if sv.state in ("hidden", "veryHidden"):
        an.add("hidden_sheet", "info", sv.name, ref(sv.name, sv.dimensions or "A1"),
               f"الورقة «{sv.name.strip()}» {STATE_AR.get(sv.state, sv.state)} وتحتوي {nonempty} صفًا غير فارغ. "
               "يجب التحقق من وظيفتها ونطاقها وانتمائها للمحفظة قبل استخدام أي من بياناتها.",
               {"state": sv.state, "non_empty_rows": nonempty}, order=(0, 0))
    # filler blocks
    run = []
    runs = []
    for r in range(1, sv.max_row + 2):
        if sa.row_kind.get(r) == "filler":
            run.append(r)
        else:
            if len(run) >= 3:
                runs.append(run)
            run = []
    for rr in runs:
        cells = sum(1 for r in rr for k in sv.kinds.get(r, {}).values() if k == "F")
        cols = sorted({c for r in rr for c, k in sv.kinds.get(r, {}).items() if k == "F"})
        an.add("filler_block", "low", sv.name, ref(sv.name, f"{col_letter(cols[0])}{rr[0]}:{col_letter(cols[-1])}{rr[-1]}"),
               f"الصفوف {rr[0]}–{rr[-1]} ({len(rr)} صفًا) لا تحتوي إلا مسافات أو محارف غير مرئية (مثل NBSP) في {cells} خلية "
               f"({col_letter(cols[0])}–{col_letter(cols[-1])}). ليست بيانات وتُستبعد؛ لكنها توسّع نطاق الورقة وقد تُحسب في COUNTA.",
               {"rows": [rr[0], rr[-1]], "n_rows": len(rr), "filler_cells": cells,
                "cols": [col_letter(cols[0]), col_letter(cols[-1])]}, order=(rr[0], 0))
    # data outside tables
    for t in sv.tables:
        nxt_hdr = min([g.header_row for g in sa.regions if g.header_row and g.header_row > t.max_row] or [sv.max_row + 1])
        extra = [r for r in range(t.max_row + 1, nxt_hdr) if sa.row_kind.get(r) == "data"
                 and any(t.min_col <= c <= t.max_col for c, k in sv.kinds.get(r, {}).items() if k not in ("B", "F"))]
        if extra:
            an.add("outside_table_rows", "medium", sv.name, ref(sv.name, f"A{extra[0]}:{col_letter(t.max_col)}{extra[-1]}"),
                   f"{len(extra)} صف بيانات يقع أسفل جدول Excel «{t.name}» ({t.ref}) خارج نطاقه (الصفوف {rows_text(extra)}). "
                   "إجماليات الجدول وصيغه المحسوبة لا تشملها؛ تُراجع قبل الاعتماد على الإجمالي.",
                   {"table": t.name, "table_ref": t.ref, "rows": compress_ints(extra)[:30]}, order=(extra[0], 0))
    # hidden rows / columns with data
    data_rows = {r for r, k in sa.row_kind.items() if k == "data"}
    hr = sorted(sv.hidden_rows & data_rows)
    if hr:
        an.add("hidden_rows", "low", sv.name, ref(sv.name, f"A{hr[0]}" if len(hr) == 1 else f"A{hr[0]}:A{hr[-1]}"),
               f"{len(hr)} صف بيانات مخفي (الصفوف {rows_text(hr)}) — قد تكون مصفاة أو مخفية يدويًا؛ "
               "دوال SUBTOTAL(109) تستبعدها. تُراجع قبل الاستبعاد أو الإدراج.",
               {"rows": compress_ints(hr)[:30], "autofilter": sv.autofilter}, order=(hr[0], 0))
    hc = sorted(c for c in sv.hidden_cols if any(sv.kind(r, c) not in ("B", "F") for r in data_rows))
    if hc:
        an.add("hidden_columns", "low", sv.name, ref(sv.name, f"{col_letter(hc[0])}:{col_letter(hc[-1] if len(hc) > 1 else hc[0])}"),
               f"أعمدة مخفية تحتوي بيانات: {', '.join(col_letter(c) for c in hc[:15])}. تُفحص كغيرها ولا تُهمل.",
               {"cols": [col_letter(c) for c in hc]}, order=(0, hc[0]))


def check_pii(sa: SheetAnalysis, an: Anomalies, pii_out: list):
    sv = sa.sv
    seen = set()
    for fam, profiles in sa.profiles.items():
        for c, p in sorted(profiles.items()):
            if not p["pii"] or c in seen:
                continue
            seen.add(c)
            pii_out.append({"sheet": sv.name, "col": col_letter(c), "header": p["header"], "reason": p["_pii_reason"]})
            an.add("pii_column", "info", sv.name, ref(sv.name, f"{col_letter(c)}:{col_letter(c)}"),
                   f"العمود {col_letter(c)}" + (f" «{short(p['header'], 40)}»" if p["header"] else "")
                   + " يحتوي بيانات شخصية (" + ("عنوان العمود" if p["_pii_reason"] == "header"
                                               else "نمط القيم: أرقام اتصال أو بريد أو آيبان أو هوية")
                   + "). لا تُنقل قيمه إلى أي مخرج، ويُربط في التعيين بنوع pii ليُسقط.",
                   {"col": col_letter(c), "header": p["header"], "reason": p["_pii_reason"]}, order=(0, c))


GENERIC_NAME_TOKENS = {normalize_arabic(t) for t in (
    "عمارة", "عماره", "فيلا", "فلل", "أرض", "ارض", "شقة", "شقه", "محل", "محلات", "وقف", "مبنى", "مبني", "مركز",
    "مدرسة", "مدرسه", "برج", "مجمع", "استراحة", "دور", "أرضي", "اول", "مستودع", "معرض", "مكتب", "سكن", "حي")}


def _loose_tokens(s: str) -> list:
    toks = []
    for t in name_key(s).split():
        if t.isdigit():
            continue
        if t.startswith("ب") and len(t) > 4 and not t.startswith("بيت"):
            t = t[1:]
        toks.append(t)
    return toks


_AMOUNT_SKIP_HINTS = {"serial", "count", "area", "date", "deed", "url", "pii", "category"}


def _row_amounts(sa: SheetAnalysis) -> dict:
    """{row: [(col, value)]} for amount-like cells of the sheet's data rows (|value| ≥ 1000, not a year, not a
    serial/count/area/date/deed/personal-data column). Used only as matching evidence between sheets."""
    sv = sa.sv
    cols = {}
    for fam, profs in sa.profiles.items():
        for c, p in profs.items():
            if p["pii"] or not p["profile"]["numeric"] or set(p["hints"]) & _AMOUNT_SKIP_HINTS:
                continue
            cols.setdefault(c, p["header"])
    out = {}
    for rows in sa.family_rows.values():
        for r in rows:
            vals = []
            for c in cols:
                v = sv.num(r, c)
                if v is None or abs(v) < 1000 or (float(v).is_integer() and (1300 <= v <= 1500 or 1900 <= v <= 2100)):
                    continue
                vals.append((c, v))
            if vals:
                out[r] = vals
    return out


def check_cross_sheet_names(sas: dict, an: Anomalies):
    """Flag names that appear (after normalisation) in the name columns of two different sheets — every pair of
    non-lookup sheets, hidden ones included, and every sub-table that has a name column.

    Three levels: exact_key (name_key equal), loose_key (digits / leading «ب» dropped),
    token_subset (every word of the shorter name — at least one of them distinctive — appears in the longer one).
    When the two rows also hold the same amount (e.g. the annual rent of a managed property and an allocation line),
    the match carries `equal_amounts`. Linking candidates only - never a proof that two rows are the same asset."""
    index = []
    for sa in sas.values():
        if sa.role == "lookup":
            continue
        entries = []
        seen_rows = set()
        fams = sorted(sa.family_rows, key=lambda f: (f != sa.primary, f is None, f or 0))
        for fam2 in fams:
            nc = sa.key_cols.get(fam2, {}).get("name") if fam2 is not None else \
                sa.key_cols.get(sa.primary, {}).get("name")
            if nc is None:
                continue
            for r in sa.family_rows[fam2]:
                if r in seen_rows or sa.sv.kind(r, nc) != "T":
                    continue
                seen_rows.add(r)
                raw = sa.sv.text(r, nc)
                toks = _loose_tokens(raw)
                entries.append((r, nc, name_key(raw), " ".join(toks), frozenset(toks)))
        if entries:
            main_col = Counter(e[1] for e in entries).most_common(1)[0][0]
            index.append((sa, main_col, entries, _row_amounts(sa)))
    for i in range(len(index)):
        for j in range(i + 1, len(index)):
            a, ac, ea, amt_a = index[i]
            b, bc, eb, amt_b = index[j]
            by_exact, by_loose, by_tok = defaultdict(list), defaultdict(list), defaultdict(list)
            for e in eb:
                by_exact[e[2]].append(e)
                if len(e[4]) >= 2:
                    by_loose[e[3]].append(e)
                for t in e[4]:
                    if t not in GENERIC_NAME_TOKENS:
                        by_tok[t].append(e)
            matches, used = [], set()

            def add(x, xc, e, kind):
                y, yc = e[0], e[1]
                if (x, y) in used or len(matches) >= 300:
                    return
                used.add((x, y))
                m = {"a": f"{col_letter(xc)}{x}", "b": f"{col_letter(yc)}{y}",
                     "a_name": short(a.sv.text(x, xc), 70), "b_name": short(b.sv.text(y, yc), 70), "match": kind}
                eq = [{"a": f"{col_letter(c1)}{x}", "b": f"{col_letter(c2)}{y}", "value": jnum(v1)}
                      for c1, v1 in amt_a.get(x, []) for c2, v2 in amt_b.get(y, []) if abs(v1 - v2) <= TOL]
                if eq:
                    m["equal_amounts"] = eq[:3]
                matches.append(m)

            def token_subset(tx, ty):
                small, big = (tx, ty) if len(tx) <= len(ty) else (ty, tx)
                return len(small) >= 2 and small <= big and any(t not in GENERIC_NAME_TOKENS for t in small)

            for (x, xc, kx, lx, tx) in ea:
                if kx:
                    for e in by_exact.get(kx, [])[:3]:
                        add(x, xc, e, "exact_key")
                if len(tx) >= 2:
                    for e in by_loose.get(lx, [])[:3]:
                        add(x, xc, e, "loose_key")
                cands = {}
                for t in tx:
                    if t not in GENERIC_NAME_TOKENS:
                        for e in by_tok.get(t, [])[:50]:
                            cands[e[0]] = e
                for e in sorted(cands.values()):
                    if token_subset(tx, e[4]):
                        add(x, xc, e, "token_subset")
            if not matches:
                continue
            matches.sort(key=lambda m: (0 if m.get("equal_amounts") else 1,
                                        {"exact_key": 0, "loose_key": 1, "token_subset": 2}[m["match"]]))
            cnt = Counter(m["match"] for m in matches)
            n_amt = sum(1 for m in matches if m.get("equal_amounts"))
            ex = matches[0]
            amt_txt = ""
            if n_amt:
                q = ex["equal_amounts"][0]
                amt_txt = (f" منها {n_amt} بمبلغ متساوٍ في الصفين (مثل {q['a']} = {q['b']} = {fmt_num(q['value'])})؛ "
                           "تساوي المبلغ قرينة إضافية على الربط لا دليل كافٍ.")
            an.add("cross_sheet_name_match", "info", a.sv.name, ref(a.sv.name, f"{col_letter(ac)}:{col_letter(ac)}"),
                   f"{len(matches)} اسمًا متطابقًا أو متقاربًا بين «{a.sv.name.strip()}» ({col_letter(ac)}) و«{b.sv.name.strip()}» "
                   f"({col_letter(bc)}): تطابق بعد التطبيع {cnt.get('exact_key', 0)}، تقارب {cnt.get('loose_key', 0)}، "
                   f"احتواء كلمات {cnt.get('token_subset', 0)} (مثل «{ex['a_name']}» ↔ «{ex['b_name']}»)." + amt_txt +
                   " تشابه الاسم وحده لا يثبت أنه الأصل نفسه؛ يُستخدم مرشحًا للربط بعد مطابقة المعرّف والصك والموقع.",
                   {"sheet_a": a.sv.name, "col_a": col_letter(ac), "sheet_b": b.sv.name, "col_b": col_letter(bc),
                    "count": len(matches), "by_level": dict(cnt), "with_equal_amounts": n_amt,
                    "matches": matches[:100]}, order=(0, 0))


# =========================================================================================
# Roles and register candidates
# =========================================================================================
def assign_roles(sas: dict, book: Book, raw: dict):
    validation_sources = defaultdict(set)
    for sa in sas.values():
        for dv in sa.sv.validations:
            f1 = dv.get("formula1") or ""
            m = re.match(r"^=?'?([^'!]+)'?!", f1)
            if m:
                validation_sources[sheet_key(m.group(1))].add(sa.sv.name)
    for sa in sas.values():
        sv = sa.sv
        ev_ = []
        name_n = normalize_arabic(sv.name)
        titles_n = " ".join(normalize_arabic(t["text"]) for t in sa.title_cells[:5])
        fam = sa.primary
        prof = sa.profiles.get(fam, {})
        hints = Counter(h for p in prof.values() for h in p["hints"])
        n_rows = len(sa.family_rows.get(fam, []))
        n_num = sum(p["profile"]["numeric"] for p in prof.values())
        role = "other"
        srcs = validation_sources.get(sheet_key(sv.name))
        if not any(k not in ("blank", "filler") for k in sa.row_kind.values()):
            role = "other"
            ev_.append("ورقة فارغة (لا محتوى)")
        elif srcs:
            role = "lookup"
            ev_.append(f"مصدر قوائم منسدلة في: {'، '.join(s.strip() for s in sorted(srcs))}")
        elif (n_num == 0 and n_rows <= 300 and LOOKUP_NAME_RE.search(name_n)) or \
                (n_num == 0 and len(sv.tables) >= 2 and all(t.max_col == t.min_col for t in sv.tables)):
            role = "lookup"
            ev_.append("قوائم نصية قصيرة بلا أرقام")
        elif MANAGED_NAME_RE.search(name_n) or MANAGED_NAME_RE.search(titles_n):
            role = "managed_properties"
            ev_.append("اسم الورقة/عنوانها يشير إلى إدارة أملاك")
        elif (ALLOC_NAME_RE.search(name_n) or ALLOC_NAME_RE.search(titles_n)) and \
                (len(sa.total_rows) >= 2 or any(k == "section_label" for k in sa.row_kind.values())):
            role = "allocation"
            ev_.append("اسم/عنوان يشير إلى تخصيص أو توزيع، مع أقسام وإجماليات فرعية")
        elif hints.get("valuation") and sa.key_cols.get(fam, {}).get("name") is not None:
            role = "asset_register"
            ev_.append(f"أعمدة تقييم ({hints['valuation']}) وعمود اسم أصل")
        elif hints.get("revenue") and (ACTIVITY_RE.search(name_n) or ACTIVITY_RE.search(titles_n)):
            role = "activity_register"
            ev_.append("أعمدة إيراد مع اسم/عنوان يشير إلى نشاط تشغيلي")
        elif ALLOC_NAME_RE.search(name_n) or ALLOC_NAME_RE.search(titles_n):
            role = "allocation"
            ev_.append("اسم/عنوان يشير إلى تخصيص أو توزيع")
        elif n_rows <= 5 and n_num == 0:
            role = "notes"
            ev_.append("نصوص قليلة بلا أرقام")
        if hints.get("revenue"):
            ev_.append(f"أعمدة إيراد: {hints['revenue']}")
        if hints.get("valuation") and role != "asset_register":
            ev_.append(f"أعمدة تقييم: {hints['valuation']}")
        if sa.title_cells:
            ev_.append(f"العنوان: «{short(sa.title_cells[0]['text'], 70)}»")
        sa.role, sa.role_evidence = role, ev_


def register_candidates(sas: dict) -> list:
    out = []
    for sa in sas.values():
        if sa.role in ("lookup", "notes"):
            continue
        fam = sa.primary
        prof = sa.profiles.get(fam, {})
        hints = Counter(h for p in prof.values() for h in p["hints"])
        n_rows = len(sa.family_rows.get(fam, []))
        if n_rows == 0:
            continue
        score = 0.0
        reasons = []
        if sa.role == "asset_register":
            score += 100
            reasons.append("دور مقترح: سجل أصول")
        if hints.get("valuation"):
            score += 30
            reasons.append(f"{hints['valuation']} عمود تقييم")
        if hints.get("revenue"):
            score += 15
            reasons.append(f"{hints['revenue']} عمود إيراد")
        if sa.key_cols.get(fam, {}).get("name") is not None:
            score += 10
            reasons.append("عمود اسم أصل")
        if sa.key_cols.get(fam, {}).get("deed") is not None:
            score += 10
            reasons.append("عمود رقم صك")
        if sa.sv.state == "visible":
            score += 10
            reasons.append("ورقة ظاهرة")
        else:
            reasons.append("ورقة مخفية — نطاقها يحتاج تحققًا")
        if any(g.tables for g in sa.families.get(fam, [])):
            score += 5
            reasons.append("جدول Excel منظم")
        score += min(20.0, n_rows / 10.0)
        reasons.append(f"{n_rows} صف بيانات")
        out.append({"sheet": sa.sv.name, "role_hint": sa.role, "score": round(score, 1), "header_row": fam,
                    "data_rows": {"from": sa.family_rows[fam][0], "to": sa.family_rows[fam][-1], "count": n_rows},
                    "reasons": reasons})
    out.sort(key=lambda x: -x["score"])
    return out


# =========================================================================================
# Output assembly
# =========================================================================================
def row_kinds_runs(sa: SheetAnalysis) -> dict:
    sv = sa.sv
    out = {}
    if sv.max_row == 0:
        return out
    cur_kind, start = None, 1
    for r in range(1, sv.max_row + 2):
        k = sa.row_kind.get(r, "blank") if r <= sv.max_row else None
        if k != cur_kind:
            if cur_kind is not None:
                key = str(start) if start == r - 1 else f"{start}-{r - 1}"
                out[key] = cur_kind
            cur_kind, start = k, r
    return out


def public_columns(profiles: dict) -> list:
    out = []
    for c, p in sorted(profiles.items()):
        item = {k: v for k, v in p.items() if not k.startswith("_")}
        if not item["notes"]:
            item.pop("notes")
        out.append(item)
    return out


def build_sheet_json(sa: SheetAnalysis, raw: dict) -> dict:
    sv = sa.sv
    if sv.sheet_type != "worksheet":
        return {"name": sv.name, "index": sv.index, "state": sv.state, "kind": sv.sheet_type}
    fam = sa.primary
    rows = sa.family_rows.get(fam, [])
    tables = []
    for t in sv.tables:
        item = {"name": t.name, "ref": t.ref, "header_row": t.min_row if t.header_rows else None,
                "data_rows": [t.first_data, t.last_data], "totals_row": t.max_row if t.totals_rows else None,
                "columns": len(t.columns)}
        if t.calc:
            item["calculated_columns"] = [{"col": col_letter(c), "name": t.columns[c - t.min_col], "formula": f}
                                          for c, f in sorted(t.calc.items())]
        tables.append(item)
    cands = []
    for i, (r, s, src) in enumerate(sa.candidates[:6]):
        item = {"row": r, "score": s, "source": src}
        if i < 4:
            hm = None
            for g in sa.regions:
                if g.header_row == r:
                    hm = g.header
            if hm is None:
                hm = {c: sv.text(r, c) for c in sorted(sv.kinds.get(r, {})) if sv.kind(r, c) == "T"}
            item["headers"] = {col_letter(c): short(h, 80) for c, h in sorted(hm.items())}
        cands.append(item)
    regions = []
    for g in sa.regions:
        drows = [r for r in range(g.start, g.end + 1) if sa.row_kind.get(r) == "data"]
        regions.append({"header_row": g.header_row, "family": g.family, "rows": [g.start, g.end],
                        "data_rows": len(drows), "data_range": [drows[0], drows[-1]] if drows else None,
                        "total_rows": [r for r in range(g.start, g.end + 1) if r in sa.total_rows],
                        "section_labels": [{"row": r, "text": short(sv.text(r, min(c for c, k in sv.kinds[r].items() if k == "T")), 80)}
                                           for r in range(max(1, g.start - 1), g.end + 1)
                                           if sa.row_kind.get(r) == "section_label"],
                        "tables": [t.name for t in g.tables] or None,
                        "after_table": g.after_table.name if g.after_table else None,
                        "primary": g.family == fam and g.family is not None})
    secondary = []
    for f2, profs in sa.profiles.items():
        if f2 == fam:
            continue
        r2 = sa.family_rows.get(f2, [])
        if not r2:
            continue
        secondary.append({"header_row": f2, "data_rows": {"from": r2[0], "to": r2[-1], "count": len(r2)},
                          "key_columns": {k: col_letter(v) for k, v in sa.key_cols.get(f2, {}).items()},
                          "columns": public_columns(profs)})
    hidden_cols = sorted(sv.hidden_cols)
    out = {
        "name": sv.name,
        "index": sv.index,
        "state": sv.state,
        "kind": "worksheet",
        "dimensions": sv.dimensions,
        "max_row": sv.max_row,
        "max_col": sv.max_col,
        "role_hint": sa.role,
        "role_evidence": sa.role_evidence,
        "merged": [f"{col_letter(c1)}{r1}:{col_letter(c2)}{r2}" for (r1, c1, r2, c2) in sv.merged[:200]],
        "hidden_rows": sorted(sv.hidden_rows)[:500],
        "hidden_cols": [col_letter(c) for c in hidden_cols[:200]],
        "autofilter": sv.autofilter,
        "tables": tables,
        "header_candidates": cands,
        "primary_header_row": fam,
        "data_rows": {"from": rows[0], "to": rows[-1], "count": len(rows)} if rows else {"from": None, "to": None, "count": 0},
        "key_columns": {k: col_letter(v) for k, v in sa.key_cols.get(fam, {}).items()},
        "title_cells": sa.title_cells,
        "columns": public_columns(sa.profiles.get(fam, {})),
        "row_kinds": row_kinds_runs(sa),
        "regions": regions,
        "secondary_tables": secondary,
        "formulas": sa.__dict__.get("formula_stats", {"count": 0, "cached_missing": 0, "errors": 0}),
        "comments": sv.comments,
        "data_validations": sv.validations[:100],
    }
    ser = sa.__dict__.get("serial_sections", {}).get(fam)
    if ser:
        out["serial_sections"] = ser
    years = sorted({y for c in out["columns"] for y in c["years_in_header"]})
    title_years = sorted({y for t in sa.title_cells for y in extract_years(t["text"])})
    out["years"] = {"in_headers": years, "in_titles": title_years, "in_sheet_name": extract_years(sv.name)}
    return out


def scrub(obj, pii_values: set):
    """Final safety net: remove any exact PII value that slipped into strings."""
    if not pii_values:
        return obj
    skip_keys = {"sha256", "path", "source_copy"}

    def walk(o, key=None):
        if isinstance(o, dict):
            return {k: (v if k in skip_keys else walk(v, k)) for k, v in o.items()}
        if isinstance(o, list):
            return [walk(x) for x in o]
        if isinstance(o, str):
            s = o
            for p in pii_values:
                if p in s:
                    s = s.replace(p, REDACTED)
            return s
        return o

    return walk(obj)


def collect_pii_values(sas: dict, pii_cols: set) -> set:
    vals = set()
    for (sheet, c) in pii_cols:
        sa = sas.get(sheet)
        if sa is None:
            continue
        for r, kr in sa.sv.kinds.items():
            if kr.get(c, "B") in ("B", "F") or sa.row_kind.get(r) != "data":
                continue
            v = sa.sv.value(r, c)
            s = display_value(v) if isinstance(v, (int, float)) else clean_text(v)
            if not s or s == REDACTED:
                continue
            digits = re.sub(r"\D", "", s.translate(_AR_DIGITS))
            if len(digits) >= 7 or len(s.split()) >= 2 or len(s) >= 7:
                vals.add(s)
    return vals


# =========================================================================================
# Markdown
# =========================================================================================
def md_escape(s) -> str:
    return str(s if s is not None else "").replace("|", "\\|").replace("\n", " ")


def write_markdown(path: str, wm: dict, sas: dict):
    L = []
    f = wm["file"]
    L.append(f"# خريطة المصنف — {md_escape(f['name'])}")
    L.append("")
    L.append("> مخرج آلي من `inspect_workbook.py`. يصف الملف ولا يقرر النطاق؛ تُراجع قبل إعداد `02_mapping.json`. "
             "قيم البيانات الشخصية لا تُكتب هنا (الموقع فقط).")
    L.append("")
    L.append("## الملف")
    L.append(f"- الاسم: `{md_escape(f['name'])}` — الحجم: {fmt_num(f['size'])} بايت — SHA-256: `{f['sha256'][:16]}…`")
    p = f.get("properties", {})
    if p:
        bits = []
        if p.get("created"):
            bits.append(f"أُنشئ {p['created']}")
        if p.get("modified"):
            bits.append(f"آخر تعديل {p['modified']}")
        if "last_modified_by" in p:
            bits.append("آخر من عدّل: " + ("مسجل (محجوب)" if p["last_modified_by"] == REDACTED else md_escape(p["last_modified_by"])))
        if "creator" in p:
            bits.append("المنشئ: " + ("مسجل (محجوب)" if p["creator"] == REDACTED else md_escape(p["creator"])))
        L.append("- خصائص المصنف: " + " — ".join(bits))
    if f.get("converted"):
        L.append(f"- تحويل: {f['converted']['note']}")
    L.append(f"- الأصل لم يتغير بعد الفحص: {'نعم' if f.get('unchanged_after_inspection') else 'تعذر التحقق'}")
    if wm.get("notes"):
        for n in wm["notes"]:
            L.append(f"- ملاحظة قراءة: {md_escape(n['note'])}")
    L.append("")
    s = wm["summary"]
    L.append("## ملخص")
    L.append(f"- عدد الأوراق: {s['sheets']} (المخفية: {s['hidden_sheets']})")
    sev = s["anomalies_by_severity"]
    L.append(f"- الملاحظات: {s['anomalies']} — مرتفعة {sev.get('high', 0)}، متوسطة {sev.get('medium', 0)}، "
             f"منخفضة {sev.get('low', 0)}، معلوماتية {sev.get('info', 0)}")
    L.append(f"- أعمدة بيانات شخصية: {len(wm['pii_columns'])} (لن تُنقل قيمها لأي مخرج)")
    L.append(f"- الصيغ: {s['formulas']} — غير مقيّمة آليًا: {s['formulas_not_evaluated']} — بلا قيمة مخزنة: {s['formulas_cached_missing']}")
    L.append("")
    L.append("## الأوراق")
    L.append("| # | الورقة | الحالة | النطاق | الجداول | صف العناوين | صفوف البيانات | صفوف الإجمالي | الدور المقترح |")
    L.append("|---|---|---|---|---|---|---|---|---|")
    for sh in wm["sheets"]:
        if sh.get("kind") != "worksheet":
            L.append(f"| {sh['index'] + 1} | «{md_escape(sh['name'])}» | {STATE_AR.get(sh['state'], sh['state'])} | — | — | — | — | — | ورقة رسم |")
            continue
        tabs = "، ".join(f"{t['name']} ({t['ref']})" for t in sh["tables"]) or "—"
        dr = sh["data_rows"]
        dr_txt = f"{dr['count']} ({dr['from']}–{dr['to']})" if dr["count"] else "0"
        tot = sum(len(g["total_rows"]) for g in sh["regions"])
        L.append(f"| {sh['index'] + 1} | «{md_escape(sh['name'])}» | {STATE_AR.get(sh['state'], sh['state'])} | "
                 f"{md_escape(sh['dimensions'])} | {md_escape(tabs)} | {sh['primary_header_row'] or '—'} | {dr_txt} | {tot} | "
                 f"{ROLE_AR.get(sh['role_hint'], sh['role_hint'])} |")
    L.append("")
    for sh in wm["sheets"]:
        if sh.get("kind") != "worksheet":
            continue
        L.append(f"### الورقة «{md_escape(sh['name'])}»")
        if sh["title_cells"]:
            L.append(f"- العنوان: «{md_escape(sh['title_cells'][0]['text'])}» ({sh['title_cells'][0]['ref']})")
        if sh["role_evidence"]:
            L.append(f"- دلائل الدور المقترح ({ROLE_AR.get(sh['role_hint'])}): " + "؛ ".join(md_escape(x) for x in sh["role_evidence"]))
        rk = "، ".join(f"{k}: {KIND_AR.get(v, v)}" for k, v in sh["row_kinds"].items() if v != "blank")
        L.append(f"- أنواع الصفوف (عدا الفارغة): {md_escape(rk) or '—'}")
        kc = sh.get("key_columns") or {}
        if kc:
            names = {"name": "الاسم", "serial": "التسلسل", "deed": "الصك"}
            L.append("- أعمدة مفتاحية مقترحة: " + "، ".join(f"{names.get(k, k)} = {v}" for k, v in kc.items()))
        fs = sh["formulas"]
        if fs.get("count"):
            L.append(f"- الصيغ: {fs['count']} (قيّمت آليًا {fs.get('evaluated', 0)}، تعذر تقييم {fs.get('not_evaluated', 0)}، "
                     f"بلا قيمة مخزنة {fs.get('cached_missing', 0)}، أخطاء {fs.get('errors', 0)})")
        if sh["data_validations"]:
            L.append("- قوائم تحقق/منسدلات: " + "، ".join(f"{md_escape(d['sqref'])} ← {md_escape(d.get('formula1'))}"
                                                     for d in sh["data_validations"][:6]))
        if sh["comments"]:
            L.append(f"- تعليقات الخلايا: {len(sh['comments'])}")
        for sec in sh["secondary_tables"]:
            L.append(f"- جدول فرعي بعنوان الصف {sec['header_row']}: {sec['data_rows']['count']} صف بيانات "
                     f"({sec['data_rows']['from']}–{sec['data_rows']['to']})")
        if sh["columns"]:
            L.append("")
            L.append("| العمود | العنوان | سنوات | أرقام | نصوص | فارغ/حشو | صيغ | تلميحات |")
            L.append("|---|---|---|---|---|---|---|---|")
            for c in sh["columns"]:
                pr = c["profile"]
                hints = ", ".join(c["hints"])
                L.append(f"| {c['col']} | {md_escape(short(c['header'] or '—', 50))} | "
                         f"{','.join(map(str, c['years_in_header'])) or ''} | {pr['numeric']} | {pr['text']} | "
                         f"{pr['blank']}/{pr['filler']} | {pr['formula']} | {md_escape(hints)} |")
        L.append("")
    L.append("## الملاحظات حسب الخطورة")
    for sv_ in ("high", "medium", "low", "info"):
        items = [a for a in wm["anomalies"] if a["severity"] == sv_]
        if not items:
            continue
        L.append(f"### {SEV_AR[sv_]} ({len(items)})")
        for a in items:
            L.append(f"- **{a['id']}** — {TYPE_AR.get(a['type'], a['type'])} — `{md_escape(a['where'])}`: {md_escape(a['detail'])}")
        L.append("")
    L.append("## أعمدة البيانات الشخصية (الموقع فقط)")
    if wm["pii_columns"]:
        L.append("| الورقة | العمود | العنوان | سبب الاكتشاف |")
        L.append("|---|---|---|---|")
        for p_ in wm["pii_columns"]:
            L.append(f"| «{md_escape(p_['sheet'])}» | {p_['col']} | {md_escape(p_['header'] or '—')} | "
                     f"{'عنوان العمود' if p_['reason'] == 'header' else 'نمط القيم'} |")
    else:
        L.append("- لم تُكتشف أعمدة بيانات شخصية آليًا (يُتحقق يدويًا).")
    L.append("")
    L.append("## مرشحات السجل الرئيسي")
    if wm["register_candidates"]:
        for i, rc in enumerate(wm["register_candidates"], start=1):
            L.append(f"{i}. «{md_escape(rc['sheet'])}» — {ROLE_AR.get(rc['role_hint'], rc['role_hint'])} — "
                     f"صف العناوين {rc['header_row']}، البيانات {rc['data_rows']['from']}–{rc['data_rows']['to']} — "
                     + "؛ ".join(md_escape(x) for x in rc["reasons"]))
    else:
        L.append("- لا يوجد مرشح واضح؛ يلزم فحص يدوي.")
    L.append("")
    L.append("## تذكير بالقواعد")
    L.append("- لا تُجمع الأوراق تلقائيًا؛ كل ورقة مخفية أو فرعية تُطابق بالمعرّف والخصائص قبل استخدامها.")
    L.append("- صفوف الإجمالي والحشو والعناوين لا تُعد أصولًا؛ الإجماليات تُعاد من السجلات المؤهلة.")
    L.append("- النص في عمود مالي ليس مبلغًا، والصفر في التقييم يعني غالبًا «غير مقيّم».")
    L.append("")
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))


# =========================================================================================
# Main
# =========================================================================================
def inspect(src: str, out_dir: str) -> dict:
    t0 = time.time()
    src = os.path.abspath(src)
    if not os.path.isfile(src):
        die(f"الملف غير موجود: {src}")
    os.makedirs(out_dir, exist_ok=True)
    sha_before = sha256_file(src)
    size = os.path.getsize(src)
    work_path, converted = prepare_input(src, out_dir)
    notes = []
    book, props, defined = load_book(work_path, notes)
    raw = read_raw_parts(work_path)
    for sv in book.order:
        for dv in raw["x14_validations"].get(sv.name, []):
            sv.validations.append(dv)

    sas = {}
    for sv in book.order:
        if sv.sheet_type != "worksheet":
            continue
        sas[sv.name] = analyze_structure(sv, book)
    # profiles, then PII propagation, then key columns (a personal-data column is never proposed as a key)
    for sa in sas.values():
        for fam, rows in sa.family_rows.items():
            sa.profiles[fam] = profile_columns(sa, fam, rows, sa.family_header.get(fam, {}))
    pii_cols = set()
    for sa in sas.values():
        for fam, profs in sa.profiles.items():
            for c, p in profs.items():
                if p["pii"]:
                    pii_cols.add((sa.sv.name, c))
    for sv_name, sa in sas.items():
        for fam, profs in sa.profiles.items():
            for c, p in profs.items():
                if (sv_name, c) in pii_cols and not p["pii"]:
                    # same column flagged elsewhere in the sheet -> treat as PII everywhere
                    p["pii"] = True
                    p["_pii_reason"] = p["_pii_reason"] or "header"
                    p["samples"] = [REDACTED] if p["samples"] else []
                    p["hints"] = ["pii"] + [h for h in p["hints"] if h not in ("pii", "name")]
    for sa in sas.values():
        for fam in sa.family_rows:
            sa.key_cols[fam] = detect_key_columns(sa, fam, sa.profiles[fam])
    assign_roles(sas, book, raw)

    an = Anomalies()
    ev = fe.Evaluator(book)
    pii_out = []
    for sa in sas.values():
        check_sheet_level(sa, an)
        check_formulas(sa, sas, book, ev, an, pii_cols)
        check_totals(sa, an, ev, pii_cols)
        check_row_totals(sa, an)
        check_columns(sa, an)
        check_pii(sa, an, pii_out)
    check_cross_sheet_names(sas, an)
    anomalies = an.finalize(book)

    sheets_json = []
    for sv in book.order:
        if sv.sheet_type != "worksheet":
            sheets_json.append(build_sheet_json(None or SheetAnalysis(sv), raw))
            continue
        sheets_json.append(build_sheet_json(sas[sv.name], raw))

    sha_after = sha256_file(src)
    sev_counts = Counter(a["severity"] for a in anomalies)
    total_formulas = sum(s.get("formulas", {}).get("count", 0) for s in sheets_json)
    wm = {
        "contract_version": CONTRACT_VERSION,
        "generator": {"script": "inspect_workbook.py", "version": SCRIPT_VERSION,
                      "generated_at": dt.datetime.now().replace(microsecond=0).isoformat(),
                      "duration_s": None},
        "file": {"path": src, "name": os.path.basename(src), "sha256": sha_before, "size": size,
                 "properties": props, "unchanged_after_inspection": sha_after == sha_before},
        "summary": {
            "sheets": len(book.order),
            "hidden_sheets": sum(1 for s in book.order if s.state in ("hidden", "veryHidden")),
            "anomalies": len(anomalies),
            "anomalies_by_severity": {k: sev_counts.get(k, 0) for k in ("high", "medium", "low", "info")},
            "anomalies_by_type": dict(Counter(a["type"] for a in anomalies)),
            "formulas": total_formulas,
            "formulas_not_evaluated": sum(s.get("formulas", {}).get("not_evaluated", 0) for s in sheets_json),
            "formulas_cached_missing": sum(s.get("formulas", {}).get("cached_missing", 0) for s in sheets_json),
        },
        "sheets": sheets_json,
        "anomalies": anomalies,
        "pii_columns": pii_out,
        "register_candidates": register_candidates(sas),
        "defined_names": defined,
        "workbook_features": {k: raw[k] for k in ("threaded_comments", "external_links", "pivot_tables", "charts",
                                                  "images", "vba")},
        "notes": notes,
    }
    if converted:
        wm["file"]["converted"] = converted
    if raw["external_links"]:
        notes.append({"source": "inspect", "message": "external links",
                      "note": f"المصنف يحتوي {raw['external_links']} رابطًا خارجيًا؛ القيم المرتبطة بملفات أخرى لا يمكن التحقق منها"})
    if raw["threaded_comments"]:
        notes.append({"source": "inspect", "message": "threaded comments",
                      "note": "توجد تعليقات حوارية (threaded comments) قد لا يظهر نصها كاملًا"})
    wm = scrub(wm, collect_pii_values(sas, pii_cols))
    wm["generator"]["duration_s"] = round(time.time() - t0, 2)
    return wm, sas


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Inspect an Excel workbook (all sheets) -> 01_workbook_map.json/.md")
    ap.add_argument("xlsx", help="path to the workbook (.xlsx/.xlsm; .xls/.ods/.xlsb are converted from a copy)")
    ap.add_argument("--out", required=True, help="run folder (created if missing)")
    ap.add_argument("--quiet", action="store_true", help="print only the output paths")
    args = ap.parse_args(argv)
    sys.setrecursionlimit(max(sys.getrecursionlimit(), 5000))
    wm, _sas = inspect(args.xlsx, args.out)
    jpath = os.path.join(args.out, "01_workbook_map.json")
    mpath = os.path.join(args.out, "01_workbook_map.md")
    save_json(jpath, wm)
    write_markdown(mpath, wm, _sas)
    if not args.quiet:
        s = wm["summary"]
        sev = s["anomalies_by_severity"]
        print(f"تم فحص المصنف «{wm['file']['name']}»: {s['sheets']} أوراق (المخفية {s['hidden_sheets']})، "
              f"{s['anomalies']} ملاحظة (مرتفعة {sev['high']}، متوسطة {sev['medium']}، منخفضة {sev['low']}، "
              f"معلوماتية {sev['info']})، أعمدة بيانات شخصية {len(wm['pii_columns'])}، "
              f"خلال {wm['generator']['duration_s']} ث.")
        if wm["register_candidates"]:
            rc = wm["register_candidates"][0]
            print(f"المرشح الأول للسجل الرئيسي: «{rc['sheet']}» (صف العناوين {rc['header_row'] or '—'}، "
                  f"{rc['data_rows']['count']} صف بيانات).")
    print(jpath)
    print(mpath)
    return 0


if __name__ == "__main__":
    sys.exit(main())
