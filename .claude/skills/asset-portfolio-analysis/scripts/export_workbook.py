#!/usr/bin/env python3
"""Export the reviewable supporting workbook of a run (data-contract.md §0: supporting_workbook.xlsx).

Usage:
  python3 scripts/export_workbook.py --run-dir RUN [--out RUN/supporting_workbook.xlsx]

Reads 03_records.json and 05_metrics.json (required) plus 01, 02, 04, 06, 07, analysis.json and
version_comparison.json when present.
Sheets: «ملخص»، «الأصول»، «الحسابات»، «المستبعدات»، «المطابقة مع الإجماليات» (when excluded total rows kept their
values)، «مشكلات البيانات»، «الفجوات والتوصيات»، «الخطة التشغيلية»، «متابعة التوصيات» and «مقارنة الإصدارات» (when
06 has recommendation_followup / version_comparison)، «الافتراضات والمصادر». Right-to-left, navy headers, filters on
every table, gray = not available (never a zero; the original text of a non-numeric cell is shown next to its
status), red = a proven problem only. Values only, except SUM totals in the calculation detail of «الحسابات».
Board wording: engine keys (fields, metrics, issue types, check ids, gap scopes, evidence tokens) are shown by their
Arabic labels; the keys themselves only in code-styled «المرجع التقني» columns/rows.

Personal data: fields whose key, label or mapping kind is personal (assets/pii_patterns.json via apa_common) are
dropped; block-level patterns (mobile, landline, e-mail, IBAN; Arabic-Indic digits included) are masked in every text
cell; review-level hits (a 10-digit number that may be a national ID or a deed) are kept and reported. After saving,
every cell is scanned again: a block-level hit deletes the file and exits 2.
Exit 0 on success, 1 on unusable input, 2 when the privacy self-check fails. `--self-test` runs built-in cases.
"""
from __future__ import annotations

import os
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argparse  # noqa: E402
import datetime as _dt  # noqa: E402
import re  # noqa: E402
from collections import OrderedDict, defaultdict  # noqa: E402

from apa_common import (  # noqa: E402
    KNOWN_STATUSES, PII_KEY_RE, REDACTED, clean_text, is_pii_header, issue_type_ar, load_json, load_vocab, mask_pii,
    pii_value_hits,
)

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table
except ImportError:  # pragma: no cover
    print("خطأ: مكتبة openpyxl غير مثبتة (pip install openpyxl)", file=sys.stderr)
    sys.exit(1)

NAVY = "1F3864"
FONT = "Arial"
CODE_FONT = "Consolas"
TECH = "المرجع التقني"
GRAY_FILL, GRAY_TXT = "E7E6E6", "595959"
RED_FILL, RED_TXT = "FFC7CE", "9C0006"
AMBER_FILL, AMBER_TXT = "FFF2CC", "7F6000"
GREEN_FILL, GREEN_TXT = "E2EFDA", "375623"
BORDER = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
FMT_INT = "#,##0"
FMT_DEC = "#,##0.00"
FMT_PCT_UNIT = '#,##0.00" %"'      # metric values already multiplied by 100
FMT_SHARE = "0.0%"                 # shares stored in 0..1
PLAN_STATUSES = ["لم يبدأ", "قيد التنفيذ", "منجز", "متعثر"]  # plan items and recommendation follow-up
MAX_WIDTH, MIN_WIDTH = 60, 9
PII_KIND_AR = {"mobile": "رقم جوال", "landline": "رقم هاتف", "email": "بريد إلكتروني", "iban": "رقم حساب بنكي (IBAN)",
               "national_id": "رقم هوية أو رقم من 10 خانات قد يكون هوية",
               "mobile_unprefixed": "رقم من 9 خانات يبدأ بـ5 قد يكون جوالًا"}
_ORPHAN_BRACKET_RE = re.compile(r"\(\s*" + re.escape(REDACTED))
TOKEN_RE = re.compile(r"\{\{\s*m:([A-Za-z0-9_\-]+)((?:\.[A-Za-z0-9_]+)*)\s*\}\}")
MONEY_FIELDS = {"num", "den", "num_known_total", "prev_sum", "curr_sum", "abs_change", "total", "top_sum",
                "a_sum", "b_sum", "value"}
COUNT_FIELDS = {"n_scope", "n_included", "n_unknown", "top"}
METRIC_STATUS_STYLE = {"complete": (GREEN_FILL, GREEN_TXT), "partial": (AMBER_FILL, AMBER_TXT),
                       "not_computable": (GRAY_FILL, GRAY_TXT), "not_comparable": (GRAY_FILL, GRAY_TXT),
                       "empty": (GRAY_FILL, GRAY_TXT)}
SEVERITY_STYLE = {"high": (AMBER_FILL, AMBER_TXT), "medium": ("FFF9E5", AMBER_TXT), "low": (None, None)}
EXCL_REASON_AR = {
    "num_unknown": "البسط (الإيراد) غير متاح", "den_unknown": "المقام (التقييم) غير متاح",
    "den_nonpositive": "المقام صفر أو سالب", "a_unknown": "الإيراد غير متاح", "b_unknown": "المصروف غير متاح",
    "prev_unknown": "إيراد الفترة السابقة غير متاح", "curr_unknown": "إيراد الفترة الحالية غير متاح",
    "both_unknown": "الإيراد غير متاح في الفترتين", "nonpositive": "قيمة صفرية أو سالبة (لا تدخل في التركز)",
    "absent": "الحقل غير موجود لهذا السجل",
    "revenue_start": "بدء إيراد (الإيراد السابق صفر)؛ خارج نسبة النمو",
}
EMPTY_FILTERS_AR = "لا تطابق التصفية الحالية أي سجل في نطاق هذا المؤشر"
EMPTY_SCOPE_AR = "لا توجد سجلات مؤهلة لهذا المؤشر في الملف"


def empty_text(res: dict) -> str:
    """Wording of an empty result by its cause (engine empty_cause / reason), never a fixed «no records»."""
    if res.get("reason"):
        return res["reason"]
    return EMPTY_FILTERS_AR if res.get("empty_cause") == "filters" else (res.get("empty_text") or EMPTY_SCOPE_AR)


def growth_riyal_only(res: dict) -> bool:
    """A growth result whose comparability is not confirmed: shown as its change in riyals, never as a % (metrics.md §5,
    brief §6). The percentage appears only in the note, marked for review."""
    comp = res.get("comparability") if isinstance(res.get("comparability"), dict) else {}
    return res.get("kind") == "growth" and comp.get("status") != "confirmed"


def growth_comp_ar(res: dict) -> str:
    comp = res.get("comparability") if isinstance(res.get("comparability"), dict) else {}
    return "غير صالحة" if comp.get("status") == "not_comparable" else "غير متحققة"


def growth_status_cell(res: dict):
    return C(f"قابلية المقارنة {growth_comp_ar(res)} — لا نسبة", fill=AMBER_FILL, color=AMBER_TXT)


def growth_pct_note(res: dict) -> str:
    v = res.get("value")
    return (f"النسبة المحسوبة {v:,.1f}% للمراجعة فقط، لا تُقرأ نموًا (قابلية المقارنة {growth_comp_ar(res)})"
            if is_num(v) else "")


def metric_value_cell(res: dict, missing_text: str):
    """(value cell, unit label): an unconfirmed growth gives its change in riyals."""
    unit = res.get("unit")
    if growth_riyal_only(res) and is_num(res.get("value")):
        ac = res.get("abs_change")
        if is_num(ac):
            return C(ac, fmt=FMT_INT), UNIT_AR["SAR"]
        return C("التغير بالريال غير متاح", fill=GRAY_FILL, color=GRAY_TXT), "—"
    if is_num(res.get("value")):
        return C(res.get("value"), fmt=unit_fmt(unit)), UNIT_AR.get(unit, unit or "—")
    return C(missing_text, fill=GRAY_FILL, color=GRAY_TXT), UNIT_AR.get(unit, unit or "—")


def result_notes(res: dict, flabel=None) -> str:
    """Reason, engine note (e.g. empty_as_zero), coverage per field (by its Arabic label: not-applicable outside the
    denominator) and revenue starts kept outside the growth %."""
    flabel = flabel or (lambda f: f)
    parts = [res.get("reason"), res.get("note")]
    if res.get("kind") == "coverage":
        pf = [p for p in res.get("per_field") or [] if isinstance(p, dict)]
        if pf:
            parts.append("التغطية لكل حقل: " + "؛ ".join(
                f"«{flabel(p.get('field'))}»: {p.get('n_known', 0)} من {p.get('n', 0)}"
                + (f" (لا ينطبق على {p['n_not_applicable']} خارج المقام)" if p.get("n_not_applicable") else "")
                for p in pf))
    if res.get("kind") == "growth" and res.get("revenue_start_ids"):
        sc = res.get("start_change")
        parts.append(f"بدء إيراد (إيراد سابق صفر) خارج نسبة النمو: {'، '.join(res['revenue_start_ids'][:20])}"
                     + (f" — التغير {sc:,.0f}" if is_num(sc) else ""))
    if growth_riyal_only(res):
        parts.append(growth_pct_note(res))
    out, seen = [], set()
    for x in parts:
        if isinstance(x, str) and x.strip() and x not in seen:
            seen.add(x)
            out.append(x.strip())
    return "؛ ".join(out)


# engine words that may appear inside a free text (normalize.ENUM_AR mirrors them for the issue log)
ENUM_TEXT_AR = [
    (re.compile(r"\s*\((?:zero_as_missing|not_available|not_applicable|zero_recorded|zero_confirmed|isolated|excluded|in_scope|raw، hijri|raw|hijri)\)"), ""),
    (re.compile(r"\bzero_as_missing\b"), "غير متاح (صفر لا يُعد قيمة)"), (re.compile(r"\bnot_available\b"), "غير متاح"),
    (re.compile(r"\bisolated\b"), "معزولة"), (re.compile(r"\bin_scope\b"), "ضمن النطاق"),
    (re.compile(r"(?<![A-Za-z_])pii(?![A-Za-z_])"), "بيانات شخصية"), (re.compile(r"\bportfolio\b"), "المحفظة كاملة"),
]
_KEY_IN_TEXT_RE = re.compile(r"[A-Za-z][A-Za-z0-9_]*(?:\.[A-Za-z0-9_]+)*\*?")
FIELD_PART_AR = {"abs_change": "التغير بالريال", "num": "البسط", "den": "المقام", "n_included": "المشمول",
                 "n_scope": "سجلات النطاق", "coverage_num_share": "حصة الإيراد المعروف", "prev_sum": "مجموع الفترة السابقة",
                 "curr_sum": "مجموع الفترة الحالية", "start_change": "تغير بدء الإيراد", "total": "المجموع",
                 "top_sum": "مجموع الأكبر", "a_sum": "مجموع الإيراد", "b_sum": "مجموع المصروف", "hhi": "مؤشر هيرفندال"}
SEVERITY_AR_EXTRA = {"info": "معلوماتية"}
KIND_AR = {"count": "عدّ", "sum": "مجموع", "ratio": "نسبة من مجاميع متطابقة", "diff": "فرق", "growth": "تغير بين فترتين",
           "concentration": "تركّز", "coverage": "تغطية بيانات"}
SOURCE_TYPE_AR = {"file": "ملف", "web": "مصدر ويب", "user": "إفادة المستخدم"}
LEVEL_AR = {"high": "مرتفع", "medium": "متوسط", "low": "منخفض", "strong": "قوي", "moderate": "متوسط", "weak": "ضعيف"}
# Reconciliation of recorded totals (brief §4/§5.8) — same statuses and wording as validate_analysis.py (07.reconciliation)
RECON_STATUS_AR = {
    "match": "مطابق لمجموع السجلات ضمن النطاق",
    "match_section": "مطابق لمجموع سجلات قسمه ضمن النطاق (إجمالي فرعي)",
    "match_all_records": "يطابق مجموع كل سجلات المصدر؛ الفرق يعود إلى سجلات معزولة أو مستبعدة",
    "match_section_all_records": "يطابق مجموع سجلات قسمه كلها؛ الفرق يعود إلى سجلات معزولة أو مستبعدة",
    "match_without_text_numbers": "يطابق المجموع دون أرقام مخزنة كنص في الملف (لا تجمعها صيغ Excel)؛ الإجمالي المدون ناقص",
    "match_original": "يطابق القيم الأصلية قبل الاشتقاق الموثق في التصنيف؛ الفرق سببه الاشتقاق لا نقص في السجلات",
    "mismatch": "فرق غير مفسَّر",
    "recorded_unknown": "القيمة المدونة في صف الإجمالي غير معروفة (فارغة أو نص أو خطأ)",
}
# Version comparison (brief §12) — compare_versions.py output
VC_CATEGORIES = [("new_period_performance", "أداء فترة جديدة"), ("scope_change", "تغير نطاق المحفظة"),
                 ("valuation_change", "تغير التقييم"), ("data_correction", "تصحيح بيانات للفترة نفسها"),
                 ("definition_change", "تغير تعريف أو تصنيف")]
VC_SUBTYPE_AR = {
    "new_period_value": "قيمة لفترة جديدة", "period_absent_in_new": "فترة غير موجودة في النسخة الجديدة",
    "field_added": "قيمة أضيفت", "field_removed": "قيمة حُذفت", "status_changed": "تغيرت حالة القيمة",
    "value_changed": "تغيرت القيمة", "name_changed": "تغير الاسم", "classification": "تغير التصنيف",
    "absent_in_new": "غير موجود في النسخة الجديدة", "new_in_file": "جديد في الملف",
    "new_valuation_year": "سنة تقييم جديدة", "valuation_added": "تقييم أضيف", "valuation_absent_in_new": "تقييم غير موجود",
    "valuation_value_changed": "تغير تقييم السنة نفسها", "metric_added": "مؤشر جديد", "metric_removed": "مؤشر حُذف",
    "metric_label": "تغير وصف مؤشر", "metric_definition": "تغير تعريف مؤشر", "config": "تغير إعداد عام",
    "period_field": "تغير حقل الفترة المستخدم في المؤشرات", "period_comparability": "تغير حكم قابلية المقارنة",
    "field_meta": "تغير وصف حقل",
}
VC_RELATION_AR = {"same_period": "الفترة نفسها — رفع نسخة أحدث ليس نموًا زمنيًا", "new_period": "فترة جديدة",
                  "different": "فترتان افتراضيتان مختلفتان"}
VC_CONF_AR = {"exact": "تامة", "deed": "بالصك", "probable": "محتملة"}
FOLLOW_STYLE = {"منجز": (GREEN_FILL, GREEN_TXT), "قيد التنفيذ": (AMBER_FILL, AMBER_TXT), "متعثر": (AMBER_FILL, AMBER_TXT)}
DIM_AR_DEFAULT = {"income_generation": "توليد الدخل", "income_stability": "استقرار الدخل",
                  "asset_utilization": "كفاءة استخدام الأصول", "data_quality": "جودة البيانات"}


# ---------------------------------------------------------------------------
# Cell spec and sheet writer
# ---------------------------------------------------------------------------
class C:
    """A cell value with optional formatting. `code`: a technical reference (field key, metric id, issue type) in a
    monospace gray font — the only place an engine key appears in the workbook (column «المرجع التقني»)."""
    __slots__ = ("v", "fmt", "fill", "color", "bold", "wrap", "formula", "raw_ok", "code")

    def __init__(self, v, fmt=None, fill=None, color=None, bold=False, wrap=True, formula=False, raw_ok=False, code=False):
        self.v, self.fmt, self.fill, self.color, self.bold, self.wrap = v, fmt, fill, color, bold, wrap
        self.code = code
        self.formula = formula  # only deliberate formulas (SUM totals) are written unescaped
        self.raw_ok = raw_ok    # a deed number: review-level digits are kept (block-level patterns are still masked)


def is_num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


class Privacy:
    """Counters of what the export masked or dropped (printed after the build; values are never printed)."""
    def __init__(self):
        self.masked = 0
        self.kinds = set()
        self.fields_dropped = []

    def reset(self):
        self.__init__()


PRIVACY = Privacy()


def safe_text(x, review=True):
    """Text for a cell. Personal-data patterns of assets/pii_patterns.json are replaced by REDACTED: block level (mobile
    and landline numbers in any separator style, e-mail, IBAN, a labelled national ID; Arabic-Indic digits are matched
    after translation) always, and review level (a bare 10-digit number that may be a national ID) too unless review=False
    — deed-number cells (C(..., raw_ok=True)) keep their digits; privacy_self_check reports what is left."""
    if x is None:
        return None
    if isinstance(x, (list, tuple)):
        x = "\n".join(str(i) for i in x if i is not None and str(i) != "")
    elif isinstance(x, dict):
        x = "؛ ".join(f"{k}: {v}" for k, v in x.items())
    s, kinds = mask_pii(str(x), ("block", "review") if review else ("block",))
    if kinds:
        PRIVACY.masked += 1
        PRIVACY.kinds.update(kinds)
        s = _ORPHAN_BRACKET_RE.sub(REDACTED, s)  # "(050) 123-4567" is matched from its first digit
    if s.startswith("="):
        s = "'" + s  # never let text become a formula
    return s


class Sheet:
    def __init__(self, wb, name, title, subtitle=None):
        self.ws = wb.create_sheet(name)
        self.ws.sheet_view.rightToLeft = True
        self.ws.sheet_view.showGridLines = False
        self.row = 1
        self.widths = defaultdict(int)
        self.merged_notes = []
        self.fixed_heights = set()
        self.print_header_row = None
        self.wb = wb
        c = self.ws.cell(row=1, column=1, value=title)
        c.font = Font(name=FONT, size=14, bold=True, color=NAVY)
        c.alignment = Alignment(horizontal="right", readingOrder=2)
        self.row = 2
        if subtitle:
            c = self.ws.cell(row=2, column=1, value=safe_text(subtitle))
            c.font = Font(name=FONT, size=9, italic=True, color=GRAY_TXT)
            c.alignment = Alignment(horizontal="right", readingOrder=2)
        self.row = 4

    def _put(self, r, col, spec, header=False):
        if not isinstance(spec, C):
            spec = C(spec)
        v = spec.v
        if spec.formula:
            pass
        elif isinstance(v, str) or isinstance(v, (list, tuple, dict)):
            v = safe_text(v, review=not spec.raw_ok)
        elif v is not None and not is_num(v) and not isinstance(v, (_dt.date, _dt.datetime)):
            v = safe_text(v, review=not spec.raw_ok)
        cell = self.ws.cell(row=r, column=col, value=v)
        cell.border = BORDER
        if header:
            cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, readingOrder=2)
        else:
            cell.font = (Font(name=CODE_FONT, size=9, color=GRAY_TXT) if spec.code
                         else Font(name=FONT, size=10, bold=spec.bold, color=spec.color))
            if spec.fill:
                cell.fill = PatternFill("solid", fgColor=spec.fill)
            if spec.fmt:
                cell.number_format = spec.fmt
            elif is_num(v):
                cell.number_format = FMT_INT if float(v).is_integer() else FMT_DEC
            cell.alignment = Alignment(vertical="top", wrap_text=spec.wrap, readingOrder=2,
                                       horizontal=None if is_num(v) else "right")
        text = ("0" * 12) if spec.formula else (v if isinstance(v, str) else (f"{v:,.2f}" if is_num(v) else ""))
        longest = max((len(t) for t in str(text).split("\n")), default=0)
        self.widths[col] = max(self.widths[col], min(longest + 2, MAX_WIDTH) if not header else min(len(str(v)) + 4, 40))
        return cell

    def heading(self, text):
        c = self.ws.cell(row=self.row, column=1, value=text)
        c.font = Font(name=FONT, size=12, bold=True, color=NAVY)
        c.alignment = Alignment(horizontal="right", readingOrder=2)
        self.row += 1

    def note(self, text, ncols=8):
        c = self.ws.cell(row=self.row, column=1, value=safe_text(text))
        c.font = Font(name=FONT, size=10)
        c.alignment = Alignment(horizontal="right", vertical="top", wrap_text=True, readingOrder=2)
        if ncols > 1:
            self.ws.merge_cells(start_row=self.row, start_column=1, end_row=self.row, end_column=ncols)
        self.merged_notes.append((self.row, ncols, str(text)))
        self.row += 1

    def tech_row(self, keys, label_col=1):
        """A row of technical references (field keys) above a header: «المرجع التقني» then each column's key."""
        r = self.row
        c = self.ws.cell(row=r, column=label_col, value=TECH + ":")
        c.font = Font(name=FONT, size=9, italic=True, color=GRAY_TXT)
        c.alignment = Alignment(horizontal="right", readingOrder=2)
        for j, k in enumerate(keys, 1):
            if k and j != label_col:
                self._put(r, j, C(k, code=True))
        self.fixed_heights.add(r)
        self.row += 1

    def table(self, name, headers, rows, freeze=None, filters=True, tech=None):
        """Write a table with a navy header row (and, with `tech`, a row of technical references above it); returns
        (header_row, first_data_row, last_row)."""
        headers = _unique_headers(headers)
        if not rows:
            rows = [[C("لا توجد بنود", color=GRAY_TXT)] + [None] * (len(headers) - 1)]
        if tech:
            self.tech_row(tech)
        hr = self.row
        for j, h in enumerate(headers, 1):
            self._put(hr, j, h, header=True)
        self.ws.row_dimensions[hr].height = 30
        self.fixed_heights.add(hr)
        if self.print_header_row is None:
            self.print_header_row = hr
        for i, row in enumerate(rows, 1):
            for j in range(len(headers)):
                self._put(hr + i, j + 1, row[j] if j < len(row) else None)
        last = hr + len(rows)
        if filters:
            ref = f"A{hr}:{get_column_letter(len(headers))}{last}"
            t = Table(displayName=_table_name(self.wb, name), ref=ref)
            self.ws.add_table(t)
        if freeze:
            self.ws.freeze_panes = freeze if isinstance(freeze, str) else f"A{hr + 1}"
        self.row = last + 2
        return hr, hr + 1, last

    def finish(self, fixed=None, print_titles=True, fit_width=None):
        final = {}
        for col, w in self.widths.items():
            width = max(MIN_WIDTH, min(MAX_WIDTH, w))
            if fixed and col in fixed:
                width = fixed[col]
            final[col] = width
            self.ws.column_dimensions[get_column_letter(col)].width = width
        merged_rows = {r for r, _, _ in self.merged_notes}
        # Row heights: neither Excel nor LibreOffice reliably auto-fit wrapped rows written by openpyxl.
        for r in range(4, self.ws.max_row + 1):
            if r in self.fixed_heights or r in merged_rows:
                continue
            lines = 1
            for c in self.ws[r]:
                if c.value is None or not isinstance(c.value, str) or c.value.startswith("="):
                    continue
                w = final.get(c.column, 10)
                lines = max(lines, _lines(c.value, w))
            if lines > 1:
                self.ws.row_dimensions[r].height = min(13.5 * lines + 3, 409)
        for r, ncols, text in self.merged_notes:
            w = sum(final.get(c, 10) for c in range(1, ncols + 1))
            self.ws.row_dimensions[r].height = min(14 * _lines(text, w) + 6, 409)
        ps = self.ws.page_setup
        ps.orientation = "landscape"
        ps.paperSize = self.ws.PAPERSIZE_A4
        total_width = sum(final.values()) or 1
        ps.fitToWidth = fit_width or min(4, max(1, -(-int(total_width) // 230)))  # wide sheets span pages, not shrink
        ps.fitToHeight = 0
        self.ws.sheet_properties.pageSetUpPr.fitToPage = True
        self.ws.print_options.horizontalCentered = True
        self.ws.page_margins.left = self.ws.page_margins.right = 0.4
        if print_titles and self.print_header_row:
            self.ws.print_title_rows = f"{self.print_header_row}:{self.print_header_row}"
        self.ws.oddFooter.center.text = "&P / &N"


def _lines(text, width):
    """Estimated wrapped line count of text in a column of the given width (Arabic ~1 char per unit)."""
    per = max(int(width * 1.15) - 1, 4)
    return sum(max(1, -(-len(part) // per)) for part in str(text).split("\n"))


def _unique_headers(headers):
    seen, out = set(), []
    for h in headers:
        h = str(h)
        base, k = h, 2
        while h.lower() in seen:
            h = f"{base} ({k})"
            k += 1
        seen.add(h.lower())
        out.append(h)
    return out


def _table_name(wb, name):
    used = {t for ws in wb.worksheets for t in ws.tables}
    base = re.sub(r"[^A-Za-z0-9_]", "_", name) or "tbl"
    if not re.match(r"[A-Za-z_]", base):
        base = "t_" + base
    cand, k = base, 2
    while cand in used:
        cand = f"{base}_{k}"
        k += 1
    return cand


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------
UNIT_AR = {"SAR": "ريال", "%": "%", "count": "عدد", "asset": "أصل"}
ORIGINAL_SUFFIX = "__original"   # normalize derived_fields (when: always) keeps the replaced stored value here


def unit_fmt(unit):
    if unit == "%":
        return FMT_PCT_UNIT
    if unit in ("count", "asset"):
        return FMT_INT
    return FMT_INT


def fmt_value_text(v, unit):
    if v is None:
        return "غير متاح"
    if not is_num(v):
        return str(v)
    if unit == "%":
        return f"{v:,.1f}%"
    if unit in ("count", "asset"):
        return f"{v:,.0f}"
    if unit == "SAR":
        return f"{v:,.0f} ريال"
    return f"{v:,.2f}" if not float(v).is_integer() else f"{v:,.0f}"


def render_tokens(text, metrics):
    if not isinstance(text, str):
        return text

    def rep(m):
        res = metrics.get(m.group(1)) or {}
        path = [p for p in m.group(2).split(".") if p]
        if (not path or path == ["value"]) and growth_riyal_only(res) and is_num(res.get("value")):
            return fmt_value_text(res.get("abs_change"), "SAR") + " (تغير بالريال؛ لا نسبة)"
        if not path:
            return fmt_value_text(res.get("value"), res.get("unit"))
        cur = res
        for p in path:
            cur = cur.get(p) if isinstance(cur, dict) else None
        f = path[-1]
        if f == "coverage_num_share" and is_num(cur):
            return f"{cur * 100:,.1f}%"
        if f in COUNT_FIELDS:
            return fmt_value_text(cur, "count")
        if f in MONEY_FIELDS and res.get("unit") in ("SAR", "%"):
            return fmt_value_text(cur, "SAR")
        if isinstance(cur, list):
            return "، ".join(map(str, cur))
        return fmt_value_text(cur, res.get("unit"))
    return TOKEN_RE.sub(rep, text)


def status_cell(status, vocab, res=None):
    """Arabic metric status. An empty result names its cause (C3): no qualifying record in the file (empty_cause
    scope, the default for an unfiltered export) or none under the current filters."""
    if status == "empty":
        cause = (res or {}).get("empty_cause")
        lab = "لا توجد سجلات ضمن التصفية الحالية" if cause == "filters" else "لا توجد سجلات مؤهلة في الملف"
        fill, color = METRIC_STATUS_STYLE.get(status, (None, None))
        return C(lab, fill=fill, color=color)
    lab = ((vocab.get("metric_status") or {}).get(status) or {}).get("ar") or status or "—"
    fill, color = METRIC_STATUS_STYLE.get(status, (None, None))
    return C(lab, fill=fill, color=color)


def unknown_text(tv, vocab):
    """«غير متاح — نص في حقل رقمي: «6 محلات + 4 مكاتب»» — the status label plus the original cell text (brief §4: keep
    the original value), so a reviewer sees what the file holds without opening it."""
    s = tv.get("s")
    lab = ((vocab.get("status") or {}).get(s) or {}).get("ar") or s or "غير متاح"
    text = lab if s in ("not_available", "not_applicable") or lab == "غير متاح" else f"غير متاح — {lab}"
    raw = tv.get("raw")
    raw = clean_text(raw) if isinstance(raw, (str, int, float)) and not isinstance(raw, bool) else ""
    if raw and s not in ("blank", "filler") and raw != lab:
        text += f": «{raw if len(raw) <= 200 else raw[:200] + '…'}»"
    return text


_DEED_FIELD_RE = re.compile(r"deed|صك|صكوك", re.IGNORECASE)


def is_deed_field(field, meta=None) -> bool:
    """A deed-number field (deed_no, «رقم الصك»): its 10-digit values are deeds, not national IDs."""
    label = str(((meta or {}).get(field) or {}).get("label") or "")
    return bool(_DEED_FIELD_RE.search(str(field or "")) or _DEED_FIELD_RE.search(label))


def tv_cell(tv, vocab, fmt=None):
    """Known -> number (or text); unknown -> gray text label with the original text, never a zero."""
    if not isinstance(tv, dict):
        return C("غير موجود", fill=GRAY_FILL, color=GRAY_TXT)
    s, v = tv.get("s"), tv.get("v")
    if s in KNOWN_STATUSES:
        if is_num(v):
            return C(v, fmt=fmt or (FMT_INT if float(v).is_integer() else FMT_DEC))
        return C(v)
    if s == "value" and v is not None:
        return C(v)
    return C(unknown_text(tv, vocab), fill=GRAY_FILL, color=GRAY_TXT)


def ar(vocab, dim, value):
    if value is None:
        return "غير محدد"
    if isinstance(value, list):
        return "، ".join(ar(vocab, dim, v) for v in value) or "غير محدد"
    if isinstance(value, bool):
        return "نعم" if value else "لا"
    table = vocab.get(dim) or {}
    return table.get(value, value) if isinstance(table, dict) else value


def list_of(x):
    return x if isinstance(x, list) else ([] if x in (None, "") else [x])


def _opt(path):
    if os.path.isfile(path):
        try:
            return load_json(path)
        except Exception as exc:
            print(f"تحذير: تعذر قراءة {os.path.basename(path)}: {exc}", file=sys.stderr)
    return None


# ---------------------------------------------------------------------------
# Builder
# ---------------------------------------------------------------------------
class Export:
    def __init__(self, run_dir):
        self.run_dir = run_dir
        self.rec_doc = load_json(os.path.join(run_dir, "03_records.json"))
        self.met_doc = load_json(os.path.join(run_dir, "05_metrics.json"))
        self.wbmap = _opt(os.path.join(run_dir, "01_workbook_map.json")) or {}
        self.mapping = _opt(os.path.join(run_dir, "02_mapping.json")) or {}
        self.config = _opt(os.path.join(run_dir, "04_metrics_config.json")) or {}
        self.validation = _opt(os.path.join(run_dir, "07_validation.json"))
        self.analysis = _opt(os.path.join(run_dir, "analysis.json"))
        narr = _opt(os.path.join(run_dir, "06_narrative.json"))
        if isinstance(self.analysis, dict) and isinstance(self.analysis.get("narrative"), dict):
            narr = self.analysis["narrative"]  # already rendered by build_dashboard.py
        self.narrative = narr or {}
        self.narrative_missing = narr is None
        self.metrics = self.met_doc.get("metrics") or {}
        self.records = [r for r in self.rec_doc.get("records") or [] if isinstance(r, dict)]
        self.meta = self.rec_doc.get("field_meta") or {}
        try:
            self.vocab = load_vocab()
        except Exception:
            self.vocab = {}
        self.pii_fields = self._pii_fields()
        present = set(self.meta).union(*[set((r.get("fields") or {}).keys()) for r in self.records])
        PRIVACY.fields_dropped = sorted(self.pii_fields & present)
        self.spec_by_id = {m.get("id"): m for m in self.config.get("metrics") or [] if isinstance(m, dict)}
        self.vc_file = _opt(os.path.join(run_dir, "version_comparison.json"))
        an = self.analysis if isinstance(self.analysis, dict) else {}
        self.facts = an.get("facts") if isinstance(an.get("facts"), dict) else {}

    def _pii_fields(self) -> set:
        """Fields never written: mapped as kind pii, a personal key (PII_KEY_RE), group/kind pii in field_meta, or a
        text-like field whose label reads as personal (is_pii_header / PII_KEY_RE) unless its column says pii_ok."""
        out, pii_ok = set(), set()
        for src in self.mapping.get("sources") or []:
            for spec in ((src or {}).get("columns") or {}).values():
                if not isinstance(spec, dict) or not spec.get("field"):
                    continue
                if spec.get("kind") == "pii":
                    out.add(spec["field"])
                elif spec.get("pii_ok") is True:
                    pii_ok.add(spec["field"])
        fields = set(self.meta)
        for r in self.records:
            fields |= set((r.get("fields") or {}).keys())
        for f in fields:
            m = self.meta.get(f) or {}
            label = str(m.get("label") or "")
            if PII_KEY_RE.search(f) or m.get("group") == "pii" or m.get("kind") == "pii":
                out.add(f)
            elif f not in pii_ok and m.get("group") in (None, "text", "other", "pii") and \
                    (is_pii_header(label) or PII_KEY_RE.search(label)):
                out.add(f)
        return out

    def r(self, x):
        return self.hz(render_tokens(x, self.metrics))

    def scope_ar(self, scope):
        """A gap scope as the board reads it: «portfolio» -> المحفظة كاملة, «group:dim=value» -> the group by labels."""
        s = str(scope or "").strip()
        if not s:
            return "غير محدد"
        if s == "portfolio":
            return "المحفظة كاملة"
        m = re.match(r"^group:([A-Za-z_]+)=(.+)$", s)
        if m:
            dim, val = m.group(1), m.group(2)
            dim_ar = (self.vocab.get("dims") or {}).get(dim)
            return f"مجموعة: {dim_ar if isinstance(dim_ar, str) else dim} = {ar(self.vocab, dim, val)}"
        return self.hz(s)

    @staticmethod
    def excluded_of(res: dict) -> list:
        """Records a metric excluded, with their reason; a not_computable metric that included and excluded none lists
        its scope (engine scope_ids) with the reason, as the dashboard's calculation log does."""
        if res.get("excluded"):
            return list(res["excluded"])
        why = res.get("reason") or res.get("unavailable_reason") or "مدخل غير متوفر"
        return [{"id": i, "reason_text": why} for i in res.get("scope_ids") or []]

    def metric_label(self, mid: str) -> str:
        res = self.metrics.get(mid) or {}
        return res.get("short_label") or res.get("label") or (self.spec_by_id.get(mid) or {}).get("label") or mid

    def hz(self, x):
        """Board wording of a text: field keys -> «Arabic label», metric ids -> «metric label», engine enum words ->
        Arabic (dashboard-spec §4, BP-10); keys stay only in the «المرجع التقني» columns."""
        if not isinstance(x, str) or not re.search(r"[a-z]", x):
            return x
        for rx, to in ENUM_TEXT_AR:
            x = rx.sub(to, x)

        def rep(m):
            k = m.group(0)
            star = k.endswith("*")
            k0 = k.rstrip("*")
            if k0 in self.metrics:
                return f"«{self.metric_label(k0)}»"
            if star and k0.endswith("_"):
                fam = list(OrderedDict.fromkeys(self.metric_label(i) for i in self.metrics if i.startswith(k0)))
                if fam:
                    return "«" + "، ".join(fam) + "»"
            base = k0[:-len(ORIGINAL_SUFFIX)] if k0.endswith(ORIGINAL_SUFFIX) else k0
            if ("." in k0 or "_" in k0) and (base in self.meta or k0 in self.meta):
                return f"«{self.field_label(k0)}»"
            return k
        return _KEY_IN_TEXT_RE.sub(rep, x)

    def ev_text(self, ref) -> str:
        """An evidence reference as the board reads it: «m:id[.field]» -> metric label (and value), «f:key» -> fact
        label; cell references, record, issue and anomaly ids stay as they are."""
        ref = (ref.get("ref") if isinstance(ref, dict) else ref)
        if not isinstance(ref, str):
            return ref
        m = re.match(r"^m:([A-Za-z0-9_\-]+)(?:\.([A-Za-z0-9_.]+))?$", ref)
        if m:
            mid, fld = m.group(1), m.group(2)
            if mid in self.metrics:
                return "مؤشر: " + self.metric_label(mid) + (f" ({FIELD_PART_AR.get(fld, fld)})" if fld and fld != "value" else "")
            bd = (self.met_doc.get("breakdowns") or {}).get(mid)
            if bd:
                return "توزيع: " + str((bd or {}).get("label") or mid)
            if mid == "trend":
                return "التطور الزمني للإيراد"
            return "مؤشر غير متاح"
        fm = re.match(r"^f:(.+)$", ref)
        if fm:
            fx = (self.facts or {}).get(fm.group(1)) or {}
            return "حقيقة من الملف: " + str(fx.get("label") or "قيمة مدونة")
        return self.hz(ref)

    def allowed_field(self, f):
        return f not in self.pii_fields

    # ----------------------------------------------------------------- sheets
    def build(self, out_path):
        wb = Workbook()
        wb.remove(wb.active)
        wb.properties.title = "المرفق الداعم لدراسة محفظة الأصول"
        wb.properties.creator = "asset-portfolio-analysis"
        self.summary(wb)
        self.assets(wb)
        self.calculations(wb)
        self.exclusions(wb)
        self.reconciliation(wb)
        self.issues(wb)
        self.gaps_recs(wb)
        self.plan(wb)
        self.followup(wb)
        self.versions(wb)
        self.assumptions_sources(wb)
        wb.active = 0
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        wb.save(out_path)
        return [ws.title for ws in wb.worksheets]

    def _file_line(self):
        src = self.rec_doc.get("source") or {}
        name = os.path.basename(src.get("file") or "") or "—"
        sha = (src.get("sha256") or "")[:12]
        return f"الملف المصدر: {name}" + (f" · البصمة {sha}…" if sha else "")

    def _validation_line(self, v: dict) -> str:
        """Blocking, open and acknowledged warnings (with their check ids) of 07_validation.json, and when it ran: the
        export reads the validation that precedes it, so the final one may differ (BP-11)."""
        if not v:
            return "لم يُشغَّل بعد"
        acks = [c for c in v.get("checks") or [] if isinstance(c, dict) and (c.get("level") == "acknowledged" or c.get("acknowledged"))]
        ids = OrderedDict()
        for c in acks:
            title = c.get("title") if isinstance(c.get("title"), str) and c.get("title") != c.get("id") else "فحص آلي"
            ids[title] = ids.get(title, 0) + 1
        n_ack = v.get("acknowledged_count") if is_num(v.get("acknowledged_count")) else len(acks)
        ack_txt = (f" · مُقرّ بها {n_ack}" + (": " + "؛ ".join(f"«{k}» ×{n}" if n > 1 else f"«{k}»" for k, n in ids.items()) if ids else "")
                   if n_ack else "")
        when = str(v.get("generated_at") or v.get("checked_at") or "")[:16].replace("T", " ")
        return (f"أخطاء مانعة {v.get('blocking_count')} · تحذيرات مفتوحة {v.get('warning_count')}{ack_txt}"
                + (f" — نتيجة التحقق السابق للتصدير ({when})؛ تُعاد بعد التصدير لفحص هذا المرفق" if when else ""))

    def summary(self, wb):
        sh = Sheet(wb, "ملخص", "دراسة محفظة الأصول — المرفق الداعم", self._file_line())
        m = self.mapping
        val = self.config.get("valuation") or {}
        st = self.rec_doc.get("stats") or {}
        by_scope = st.get("by_scope") or {}
        v = self.validation or {}
        rows = [
            ("الجهة", (m.get("entity") or {}).get("name") or "غير مذكورة في الملف"),
            ("فترة البيانات", (m.get("data_period") or {}).get("label") or (self.met_doc.get("default_filters") or {}).get("period")),
            ("تاريخ إعداد الدراسة", ((m.get("study_date") or {}).get("value") or "غير محدد") +
             (f" ({(m.get('study_date') or {}).get('basis')})" if (m.get("study_date") or {}).get("basis") else "")),
            ("تاريخ التحليل", m.get("analysis_date") or _dt.date.today().isoformat()),
            ("العملة", m.get("currency") or self.config.get("currency") or "SAR"),
            ("التقييم المستخدم في المؤشرات", " — ".join(x for x in (val.get("label"), val.get("basis"), val.get("date_text")) if x) or "—"),
            ("السجلات", f"{st.get('records', len(self.records))} سجل: ضمن النطاق {by_scope.get('in_scope', 0)}، "
                        f"معزول {by_scope.get('isolated', 0)}، مستبعد {by_scope.get('excluded', 0)}"),
            ("التحقق الآلي", self._validation_line(v)),
            ("دليل الألوان", "الرمادي = غير متاح (لا يُجمع صفرًا) · الأحمر = مشكلة مثبتة فقط · الأصفر = جزئي · الأخضر = مكتمل للنطاق"),
        ]
        sh.table("tblContext", ["البند", "القيمة"], [[C(k, bold=True), x] for k, x in rows], filters=False)
        sh.heading("الخلاصة التنفيذية")
        es = self.r(self.narrative.get("executive_summary")) if not self.narrative_missing else None
        sh.note(es or "لم يُكتب السرد (06_narrative.json) بعد.", ncols=9)
        sh.row += 1
        sh.heading("المؤشرات")
        pos = {mid: k for k, mid in enumerate(self.metrics)}
        ids = sorted(self.metrics, key=lambda i: (not (self.metrics[i] or {}).get("card"),
                                                  (self.metrics[i] or {}).get("card_order") or 999, pos[i]))
        rows = []
        for mid in ids:
            res = self.metrics[mid] or {}
            val_cell, unit_lab = metric_value_cell(res, "غير قابل للاحتساب بالبيانات الحالية" if res.get("status") != "empty"
                                                   else empty_text(res))
            rows.append([C(res.get("label") or mid, bold=bool(res.get("card"))), val_cell, unit_lab,
                         growth_status_cell(res) if growth_riyal_only(res) and is_num(res.get("value"))
                         else status_cell(res.get("status"), self.vocab, res),
                         f"{res.get('n_included', 0)} من {res.get('n_scope', 0)}",
                         res.get("scope_text"), res.get("basis"), res.get("valuation_date"),
                         self.hz(result_notes(res, self.field_label)) or ("بطاقة في الداشبورد" if res.get("card") else "")])
        sh.table("tblKpis", ["المؤشر", "القيمة", "الوحدة", "الحالة", "المشمول من النطاق", "النطاق", "الأساس",
                             "تاريخ التقييم", "السبب أو الملاحظة"], rows)
        gaps = {g.get("id"): g for g in self.narrative.get("gaps") or [] if isinstance(g, dict)}
        recs = {x.get("id"): x for x in self.narrative.get("recommendations") or [] if isinstance(x, dict)}
        # two separate lists (brief §11.8): the n-th gap and the n-th action are not a pair; an action names the gaps it
        # addresses (gap_ids) explicitly
        tg = [gaps[g] for g in list(self.narrative.get("top_gaps") or []) if g in gaps][:3]
        ta = [recs[a] for a in list(self.narrative.get("top_actions") or []) if a in recs][:3]
        sh.heading("أهم ثلاث فجوات")
        sh.table("tblTopGaps", ["الفجوة", "الإثبات", "#", "المعرّف"],
                 [[self.r(g.get("title")), "مشكلة مثبتة" if g.get("is_proven_problem") is True else "محتملة — تحتاج تحققًا",
                   i, g.get("id")] for i, g in enumerate(tg, 1)], filters=False)
        sh.heading("أهم ثلاثة إجراءات أو قرارات مطلوبة")
        sh.table("tblTopActions", ["الإجراء أو القرار المطلوب", "المسؤول المقترح", "#", "يعالج الفجوة"],
                 [[self.r(a.get("first_action") or a.get("title")), a.get("owner_role") or "—", i,
                   "، ".join(str(x) for x in a.get("gap_ids") or []) or "—"] for i, a in enumerate(ta, 1)], filters=False)
        sh.ws.freeze_panes = "A4"
        sh.finish(fixed={1: 40, 2: 22, 3: 7, 4: 20, 5: 11, 6: 34, 7: 30, 8: 16, 9: 30}, print_titles=False, fit_width=1)

    def field_label(self, f: str) -> str:
        """Arabic label of a field; «<field>__original» is the stored value a documented derivation replaced."""
        lab = (self.meta.get(f) or {}).get("label")
        if lab:
            return lab
        if isinstance(f, str) and f.endswith(ORIGINAL_SUFFIX):
            base = f[:-len(ORIGINAL_SUFFIX)]
            return f"{(self.meta.get(base) or {}).get('label') or base} كما ورد قبل الاشتقاق (قيمة مخزنة معيبة — غير مستخدمة)"
        return f

    def header_label(self, f: str) -> str:
        """A column header: the field's Arabic label; a field the mapping left unlabelled reads «حقل بلا تسمية» (its key
        is in the «المرجع التقني» row above the header)."""
        lab = self.field_label(f)
        return "حقل بلا تسمية في التصنيف" if lab == f and re.fullmatch(r"[A-Za-z0-9_.]+", str(f)) else lab

    def _metric_fields(self) -> set:
        """Fields the metrics read: the valuation field, every period's revenue fields, resolved metric fields."""
        if getattr(self, "_mf", None) is None:
            used = set()
            val = (self.config.get("valuation") or {}).get("field")
            if val:
                used.add(val)
            for p in self.config.get("periods") or []:
                if isinstance(p, dict):
                    used.update(v for k, v in p.items() if isinstance(v, str) and k.endswith("field"))
            for res in self.metrics.values():
                for v in ((res or {}).get("resolved_fields") or {}).values():
                    used.update(v if isinstance(v, list) else [v])
            self._mf = {u for u in used if isinstance(u, str)}
        return self._mf

    def _field_lists(self):
        numeric, text = [], []
        groups_order = {"revenue": 0, "expense": 1, "valuation": 2}
        seen = OrderedDict()
        for f in self.meta:
            seen[f] = True
        for r in self.records:
            for f in (r.get("fields") or {}):
                seen.setdefault(f, True)
        for f in seen:
            if not self.allowed_field(f):
                continue
            vals = [((r.get("fields") or {}).get(f) or {}) for r in self.records]
            has_num = any(is_num(tv.get("v")) for tv in vals if isinstance(tv, dict))
            grp = (self.meta.get(f) or {}).get("group")
            if grp in groups_order or has_num:
                numeric.append(f)
            else:
                text.append(f)
        numeric.sort(key=lambda f: (groups_order.get((self.meta.get(f) or {}).get("group"), 3), f))
        return numeric, text

    def assets(self, wb):
        sh = Sheet(wb, "الأصول", "سجل الأصول والسجلات كما طُبِّعت (كل سجل مرة واحدة)",
                   "الخلايا الرمادية قيم غير متاحة في الملف ولا تُعامل صفرًا؛ الصفر المسجل يظهر 0. عمود «المرجع» يحدد الورقة والصف في الملف الأصلي.")
        numeric, text = self._field_lists()
        dims = [("asset_type", "نوع الأصل"), ("location", "الموقع"), ("district", "الحي"), ("use", "الاستخدام"),
                ("income_nature", "طبيعة الدخل"), ("ownership", "الملكية والحقوق"), ("waqf", "صفة الوقف"),
                ("managed_for_others", "مدار للغير"), ("rent_allocation", "تخصيص الريع"), ("restriction", "قيود الريع")]
        headers = ["المعرّف", "الاسم", "نوع السجل", "النطاق", "سبب النطاق"] + [lab for _, lab in dims]
        headers += [self.header_label(f) for f in numeric]
        headers += [self.header_label(f) for f in text]
        headers += ["المرجع", "مراجع القيم الرقمية", "مشكلات البيانات"]
        # the field keys sit in a «المرجع التقني» row above the Arabic headers, never in the headers themselves
        tech = [None] * (5 + len(dims)) + list(numeric) + list(text) + [None] * 3
        rows = []
        for rec in self.records:
            d = rec.get("dims") or {}
            fields = rec.get("fields") or {}
            scope = rec.get("scope")
            row = [rec.get("id"), rec.get("name"), ar(self.vocab, "record_type", rec.get("record_type")),
                   C(ar(self.vocab, "scope", scope), fill=GRAY_FILL if scope != "in_scope" else None,
                     color=GRAY_TXT if scope != "in_scope" else None),
                   rec.get("scope_reason")]
            for k, _ in dims:
                vv = d.get(k)
                vocab_key = {"use": "use", "income_nature": "income_nature", "ownership": "ownership", "waqf": "waqf"}.get(k)
                txt = ar(self.vocab, vocab_key, vv) if vocab_key else (ar(self.vocab, "", vv) if isinstance(vv, (bool, list)) else vv)
                row.append(C(txt if txt not in (None, "") else "غير محدد", fill=GRAY_FILL if vv in (None, "unknown", "") else None,
                             color=GRAY_TXT if vv in (None, "unknown", "") else None))
            for f in numeric:
                if f.endswith(ORIGINAL_SUFFIX) and isinstance(fields.get(f), dict):
                    # the stored value a documented derivation replaced: shown, flagged, never used (N7)
                    tv0 = fields[f]
                    row.append(C(tv0.get("v"), fmt=FMT_INT, fill=AMBER_FILL, color=AMBER_TXT) if is_num(tv0.get("v"))
                               else tv_cell(tv0, self.vocab))
                    continue
                row.append(tv_cell(fields.get(f), self.vocab))
            for f in text:
                tv = fields.get(f)
                row.append(tv_cell(tv, self.vocab) if isinstance(tv, dict) and tv.get("s") not in KNOWN_STATUSES else
                           C((tv or {}).get("v"), raw_ok=is_deed_field(f, self.meta)))
            src = rec.get("source") or {}
            # every numeric source cell, the fields the metrics use first (valuation, period revenue): no cut-off (N8)
            used = self._metric_fields()
            ordered = [f for f in numeric if f in used] + [f for f in numeric if f not in used]
            refs = [f"{self.field_label(f)}: {(fields.get(f) or {}).get('ref')}" for f in ordered if (fields.get(f) or {}).get("ref")]
            row += [src.get("ref_name") or (f"{src.get('sheet')}!{src.get('row')}" if src.get("sheet") else None),
                    "\n".join(refs), "، ".join(rec.get("issue_ids") or [])]
            rows.append(row)
        hr, first, _ = sh.table("tblAssets", headers, rows, tech=tech)
        sh.ws.freeze_panes = f"C{first}"
        sh.finish(fixed={2: 32})

    def calculations(self, wb):
        sh = Sheet(wb, "الحسابات", "سجل الحسابات — تعريف كل مؤشر ونطاقه ونتيجته",
                   "القيم من 05_metrics.json (المصدر الحسابي الوحيد للداشبورد). التفصيل أدناه يعيد الجمع من السجلات المشمولة بصيغ SUM للمراجعة.")
        rows = []
        for mid, res in self.metrics.items():
            res = res or {}
            unit = res.get("unit")
            spec = self.spec_by_id.get(mid) or {}
            rf = res.get("resolved_fields") or {}
            excl = self.excluded_of(res)
            val_cell, unit_lab = metric_value_cell(res, empty_text(res) if res.get("status") == "empty" else "غير متاح")
            share = res.get("coverage_num_share")
            rows.append([
                C(mid, code=True), res.get("label"), KIND_AR.get(res.get("kind"), res.get("kind")), val_cell, unit_lab,
                growth_status_cell(res) if growth_riyal_only(res) and is_num(res.get("value"))
                else status_cell(res.get("status"), self.vocab, res), res.get("n_scope"), res.get("n_included"), len(excl),
                "؛ ".join("، ".join(self.field_label(x) for x in (v if isinstance(v, list) else [v]) if isinstance(x, str))
                          for k, v in rf.items()) if rf else "—",
                C(res.get("num"), fmt=FMT_INT) if is_num(res.get("num")) else None,
                C(res.get("den"), fmt=FMT_INT) if is_num(res.get("den")) else None,
                C(share, fmt=FMT_SHARE) if is_num(share) else None,
                res.get("definition"), res.get("scope_text"), res.get("basis"), res.get("valuation_date"),
                res.get("period_text"), res.get("source_text"), self.hz(result_notes(res, self.field_label)) or None,
                "، ".join(res.get("included_ids") or []),
                "، ".join(f"{e.get('id')} ({e.get('reason_text') or EXCL_REASON_AR.get(e.get('reason'), e.get('reason'))})" for e in excl[:60]),
                ("نعم" if spec.get("card") else "لا"),
                C("؛ ".join(f"{k}: {'، '.join(v) if isinstance(v, list) else v}" for k, v in rf.items()) if rf else None, code=True),
            ])
        sh.table("tblMetrics", [TECH, "المؤشر", "النوع", "القيمة", "الوحدة", "الحالة", "سجلات النطاق", "المشمول",
                                "المستبعد", "الحقول المستخدمة", "البسط", "المقام", "حصة الإيراد المعروف المغطاة",
                                "التعريف", "النطاق", "الأساس", "تاريخ التقييم", "الفترة", "المصدر", "السبب",
                                "السجلات المشمولة", "السجلات المستبعدة وسببها", "بطاقة", TECH + " (الحقول)"], rows,
                 freeze="C5")
        self._calc_detail(sh)
        self._breakdowns(sh)
        sh.finish(fixed={2: 36, 14: 50, 21: 30, 22: 40}, print_titles=False)

    def _calc_detail(self, sh):
        recs = {r.get("id"): r for r in self.records}
        sh.heading("تفصيل الحساب: إعادة الجمع من السجلات المشمولة")
        for mid, res in self.metrics.items():
            res = res or {}
            kind = res.get("kind")
            if kind not in ("sum", "ratio", "diff", "growth") or res.get("status") not in ("complete", "partial"):
                continue
            rf = res.get("resolved_fields") or {}
            cols = {"sum": ["field"], "ratio": ["num", "den"], "diff": ["a", "b"], "growth": ["prev", "curr"]}[kind]
            fields = [rf.get(k) for k in cols]
            if not all(isinstance(f, str) for f in fields):
                continue
            ids = res.get("included_ids") or []
            sh.note(f"{res.get('label') or mid} — {len(ids)} سجل مشمول", ncols=6)
            sh.tech_row([mid, None] + list(fields) + [None], label_col=2)
            hr = sh.row
            heads = ["المعرّف", "الاسم"] + [self.header_label(f) for f in fields] + ["المرجع"]
            for j, h in enumerate(_unique_headers(heads), 1):
                sh._put(hr, j, h, header=True)
            r = hr
            for rid in ids:
                r += 1
                rec = recs.get(rid) or {}
                fl = rec.get("fields") or {}
                vals = [rid, rec.get("name")] + [tv_cell(fl.get(f), self.vocab) for f in fields] + \
                       [" | ".join(str((fl.get(f) or {}).get("ref") or "") for f in fields)]
                for j, v in enumerate(vals, 1):
                    sh._put(r, j, v)
            tot = r + 1
            sh._put(tot, 1, C("المجموع (SUM)", bold=True))
            sh._put(tot, 2, C("قيمة المحرك في 05 للمقارنة ←", color=GRAY_TXT))
            for k in range(len(fields)):
                col = get_column_letter(3 + k)
                sh._put(tot, 3 + k, C(f"=SUM({col}{hr + 1}:{col}{max(r, hr + 1)})", fmt=FMT_INT, bold=True, formula=True))
            engine = {"sum": [res.get("value")], "ratio": [res.get("num"), res.get("den")],
                      "diff": [res.get("a_sum"), res.get("b_sum")], "growth": [res.get("prev_sum"), res.get("curr_sum")]}[kind]
            sh._put(tot, 3 + len(fields), C("محرك: " + " / ".join(f"{x:,.2f}" if is_num(x) else "—" for x in engine),
                                            color=GRAY_TXT))
            if kind == "ratio" and is_num(res.get("value")):
                sh._put(tot + 1, 1, C("النتيجة", bold=True))
                scale = res.get("scale") if is_num(res.get("scale")) else 1
                sh._put(tot + 1, 3, C(f"=IF(D{tot}=0,\"\",C{tot}/D{tot}*{scale})", fmt=FMT_PCT_UNIT if res.get("unit") == "%" else FMT_DEC, bold=True,
                                      formula=True))
                sh._put(tot + 1, 4, C(res.get("value"), fmt=FMT_PCT_UNIT if res.get("unit") == "%" else FMT_DEC, color=GRAY_TXT))
                tot += 1
            sh.row = tot + 2

    def _breakdowns(self, sh):
        bds = self.met_doc.get("breakdowns") or {}
        if not bds:
            return
        sh.heading("التوزيعات")
        rows = []
        for bid, b in bds.items():
            b = b or {}
            for row in b.get("rows") or []:
                key = row.get("key")
                key_ar = "غير مصنف" if key in (None, "__none__") else ar(self.vocab, b.get("by"), key)
                rows.append([b.get("label") or bid, key_ar,
                             C(row.get("value"), fmt=FMT_INT) if is_num(row.get("value")) else C("غير متاح", fill=GRAY_FILL, color=GRAY_TXT),
                             C(row.get("share"), fmt=FMT_SHARE) if is_num(row.get("share")) else None,
                             row.get("n_known"), row.get("n_unknown"), b.get("note") or ("قد تتداخل الفئات" if b.get("overlap") else ""),
                             C(bid, code=True)])
        sh.table("tblBreakdowns", ["التوزيع", "الفئة", "القيمة", "الحصة", "سجلات معروفة", "سجلات غير معروفة", "ملاحظة", TECH], rows)

    def exclusions(self, wb):
        sh = Sheet(wb, "المستبعدات", "المستبعدات والمعزولات وأسبابها",
                   "السجلات المعزولة لم يتحقق نطاقها فلا تدخل في المؤشرات؛ الاستبعاد من مؤشر يعني نقص أحد مدخلاته لا انعدام قيمته.")
        rows = []
        for rec in self.records:
            if rec.get("scope") in ("isolated", "excluded"):
                rows.append([ar(self.vocab, "scope", rec.get("scope")), f"{rec.get('id')} — {rec.get('name') or ''}",
                             (rec.get("source") or {}).get("ref_name"), rec.get("scope_reason") or "—", "جميع المؤشرات"])
        for ex in self.rec_doc.get("excluded_rows") or []:
            rows.append(["صف مستبعد من المصدر", f"الصف {ex.get('row')}", f"{ex.get('sheet')}!{ex.get('row')}",
                         self.hz(ex.get("reason")), "جميع المؤشرات"])
        for ex in self.narrative.get("exclusions") or []:
            if isinstance(ex, dict):
                rows.append(["استبعاد تحليلي", self.r(ex.get("what")), ex.get("where"), self.r(ex.get("reason")), "—"])
        for mid, res in self.metrics.items():
            res = res or {}
            for e in self.excluded_of(res):
                st = e.get("reason")
                reason = e.get("reason_text") or EXCL_REASON_AR.get(st) or (((self.vocab.get("status") or {}).get(st) or {}).get("ar")) or st
                rows.append(["مستبعد من مؤشر", e.get("id"), None, reason, res.get("label") or mid])
        sh.table("tblExclusions", ["النوع", "البند", "الموضع", "السبب", "المؤشر المتأثر"], rows, freeze=True)
        sh.finish(fixed={2: 40, 4: 50, 5: 40})

    def reconciliation_rows(self):
        """Recorded totals (03 excluded_rows[].values) vs the sum of the records of the same source and field.

        Same rules as validate_analysis.py: the full in-scope sum first, then the section above the total row (a
        subtotal), then every record of the source (isolated/excluded included), then the sum without numbers that the
        file stores as text (Excel's SUM skips them). Joined fields are not part of the source's own rows."""
        entries = [e for e in self.rec_doc.get("excluded_rows") or []
                   if isinstance(e, dict) and isinstance(e.get("values"), dict) and e["values"]]
        if not entries:
            return []
        by_src = defaultdict(list)
        for r in self.records:
            by_src[(r.get("source") or {}).get("source_id")].append(r)
        total_rows = defaultdict(list)
        for e in entries:
            if isinstance(e.get("row"), int):
                total_rows[e.get("source_id")].append(e["row"])
        out = []
        for e in entries:
            sid, sheet, row = e.get("source_id"), e.get("sheet"), e.get("row")
            above = [x for x in total_rows.get(sid, []) if isinstance(row, int) and x < row]
            lo = max(above) if above else None

            def agg(field, in_scope_only, section, original=False):
                s, n_known, n_unknown, s_text, ids = 0.0, 0, 0, 0.0, []
                for r in by_src.get(sid, []):
                    if in_scope_only and r.get("scope") != "in_scope":
                        continue
                    rr = (r.get("source") or {}).get("row")
                    if section and not (isinstance(rr, int) and isinstance(row, int) and (lo is None or rr > lo) and rr < row):
                        continue
                    fl = r.get("fields") or {}
                    tv = (fl.get(field + "__original") or fl.get(field)) if original else fl.get(field)
                    tv = tv if isinstance(tv, dict) else {}
                    if tv.get("joined_from"):
                        continue
                    if tv.get("s") in KNOWN_STATUSES and is_num(tv.get("v")):
                        s += tv["v"]
                        n_known += 1
                        ids.append(r.get("id"))
                        if isinstance(tv.get("raw"), str) and str(tv.get("note") or "").startswith("نص رقمي"):
                            s_text += tv["v"]
                    elif tv:
                        n_unknown += 1
                return round(s, 2), n_known, n_unknown, round(s_text, 2), ids
            for field, tv in e["values"].items():
                if not self.allowed_field(field):
                    continue
                tv = tv if isinstance(tv, dict) else {"v": tv, "s": "value" if is_num(tv) else "text", "raw": tv}
                item = OrderedDict([("source_id", sid), ("sheet", sheet), ("row", row),
                                    ("ref", tv.get("ref") or (f"{sheet}!{row}" if sheet else f"{sid}:{row}")),
                                    ("field", field), ("reason", e.get("reason")), ("tv", tv)])
                if tv.get("s") not in KNOWN_STATUSES or not is_num(tv.get("v")):
                    item.update(recorded=None, status="recorded_unknown")
                    out.append(item)
                    continue
                recorded = float(tv["v"])
                tol = max(0.5, 1e-9 * abs(recorded))
                full, sec = agg(field, True, False), agg(field, True, True)
                all_rec, all_sec = agg(field, False, False), agg(field, False, True)
                status, matched = "mismatch", None
                for st, res in (("match", full), ("match_section", sec), ("match_all_records", all_rec),
                                ("match_section_all_records", all_sec)):
                    if res[1] and abs(recorded - res[0]) <= tol:
                        status, matched = st, res[0]
                        break
                if status == "mismatch" and full[3] and abs(recorded - (full[0] - full[3])) <= tol:
                    status, matched = "match_without_text_numbers", round(full[0] - full[3], 2)
                if status == "mismatch" and any(field + "__original" in (r.get("fields") or {}) for r in by_src.get(sid, [])):
                    for res in (agg(field, True, False, True), agg(field, True, True, True)):
                        if res[1] and abs(recorded - res[0]) <= tol:
                            status, matched = "match_original", res[0]
                            break
                # no in-scope value (isolated or supplementary source), or matched by all records: the section sum and
                # the difference come from every record of the source — never «recorded − 0» beside «matches» (N4)
                use_all = not full[1] or status in ("match_all_records", "match_section_all_records")
                base, base_sec = (all_rec, all_sec) if use_all else (full, sec)
                # «السجلات المجموعة»: the records of the sum the total was compared with — a section (sub)total lists the
                # lines of its own section, not the whole source
                summed = {"match_section": sec, "match_section_all_records": all_sec}.get(status, base)
                item.update(recorded=recorded, sum_in_scope=full[0], n_known=full[1], n_unknown=full[2],
                            ids=summed[4], difference=round(recorded - (matched if matched is not None else base[0]), 2),
                            difference_basis="all_records" if use_all else "in_scope", sum_all_records=all_rec[0],
                            section_sum=base_sec[0] if base_sec[1] and (lo is not None or base_sec[1] != base[1]) else None,
                            matched_sum=matched, status=status)
                out.append(item)
        return out

    def reconciliation(self, wb):
        items = self.reconciliation_rows()
        if not items:
            return
        sh = Sheet(wb, "المطابقة مع الإجماليات", "مطابقة الإجماليات المدونة في الملف مع مجموع السجلات المؤهلة",
                   "القيمة المدونة من صف الإجمالي المستبعد من السجلات؛ المجموع من سجلات المصدر نفسه ضمن النطاق بقيم "
                   "معروفة فقط (المجهول لا يُجمع صفرًا). الفرق = المدون − المجموع؛ الفرق غير المفسَّر يُذكر في مشكلات البيانات. "
                   "«السجلات المجموعة» سجلات المجموع الذي قورن به الإجمالي (سجلات قسمه لإجمالي القسم).")
        rows = []
        for it in items:
            st = it["status"]
            ok = st in ("match", "match_section", "match_all_records", "match_section_all_records")
            fill, color = ((GREEN_FILL, GREEN_TXT) if ok else
                           (GRAY_FILL, GRAY_TXT) if st == "recorded_unknown" else (AMBER_FILL, AMBER_TXT))
            lab = (self.meta.get(it["field"]) or {}).get("label") or it["field"]
            ids = it.get("ids") or []
            known = it.get("recorded") is not None
            rows.append([
                str(it.get("sheet") or "").strip() or it["source_id"], it["ref"], lab,
                C(it["recorded"], fmt=FMT_INT if float(it["recorded"]).is_integer() else FMT_DEC) if known
                else C(unknown_text(it["tv"], self.vocab), fill=GRAY_FILL, color=GRAY_TXT),
                C(it["sum_in_scope"], fmt=FMT_INT) if known else None,
                it.get("n_known") if known else None, it.get("n_unknown") if known else None,
                C(it["difference"], fmt=FMT_INT, fill=None if ok else AMBER_FILL, color=None if ok else AMBER_TXT)
                if known else None,
                C(it["section_sum"], fmt=FMT_INT) if known and it.get("section_sum") is not None else None,
                C(it["sum_all_records"], fmt=FMT_INT) if known else None,
                (({"all_records": "سجلات قسمه في المصدر (لا سجلات ضمن النطاق أو المطابقة بها)",
                   "in_scope": "سجلات قسمه ضمن النطاق"} if st in ("match_section", "match_section_all_records") else
                  {"all_records": "كل سجلات المصدر (لا سجلات ضمن النطاق أو المطابقة بها)", "in_scope": "السجلات ضمن النطاق"})
                 .get(it.get("difference_basis"), "—") if known else None),
                C(RECON_STATUS_AR[st], fill=fill, color=color),
                "، ".join(ids[:40]) + (f"… و{len(ids) - 40} سجلًا آخر (القائمة كاملة في 05_metrics.json)" if len(ids) > 40 else ""), self.hz(it.get("reason")),
                C(f"{it['source_id']} · {it['field']}", code=True),
            ])
        sh.table("tblReconciliation", ["الورقة", "صف الإجمالي", "الحقل", "القيمة المدونة", "مجموع السجلات ضمن النطاق",
                                       "قيم معروفة", "قيم غير معروفة", "الفرق (المدون − المجموع)", "مجموع سجلات القسم",
                                       "مجموع كل سجلات المصدر", "أساس الفرق والقسم", "النتيجة", "السجلات المجموعة",
                                       "سبب استبعاد الصف", TECH],
                 rows, freeze=True)
        n_bad = sum(1 for it in items if it["status"] == "mismatch")
        sh.note(f"بنود المطابقة: {len(items)}؛ فروق غير مفسَّرة: {n_bad}. «قيم غير معروفة» سجلات ضمن النطاق قيمتها في الحقل "
                "غير متاحة، وقد تفسر جزءًا من الفرق.", ncols=8)
        sh.finish(fixed={3: 34, 11: 30, 12: 30, 13: 40, 14: 40})

    def issues(self, wb):
        sh = Sheet(wb, "مشكلات البيانات", "سجل مشكلات البيانات (منفصل عن أداء المحفظة)",
                   "المصدر: فحص المصنف (01)، التطبيع (03)، أو ملاحظات المحلل (06). الخطورة تخص أثر المشكلة على الحساب، وليست حكمًا على الأصل.")
        sev_ar = dict(SEVERITY_AR_EXTRA, **(self.vocab.get("severity") or {}))
        rows = []

        def sev(s):
            fill, color = SEVERITY_STYLE.get(s, (None, None))
            return C(sev_ar.get(s, s or "—"), fill=fill, color=color)

        def tech(x):
            keys = [k for k in (x.get("type"), x.get("field")) if isinstance(k, str) and k and not re.search(r"[\u0600-\u06FF]", k)]
            return C(" · ".join(keys) or None, code=True)
        # the type is shown by its Arabic label; the type and field keys only in «المرجع التقني» (BP-10)
        for a in self.wbmap.get("anomalies") or []:
            if isinstance(a, dict):
                rows.append([a.get("id"), "فحص المصنف", issue_type_ar(a.get("type")), sev(a.get("severity")), a.get("where"),
                             None, self.hz(a.get("detail")), None, None, tech(a)])
        for i in self.rec_doc.get("issues") or []:
            if isinstance(i, dict):
                rows.append([i.get("id"), "التطبيع", issue_type_ar(i.get("type")), sev(i.get("severity")), i.get("where"),
                             "، ".join(i.get("record_ids") or []), self.hz(i.get("detail")), self.hz(i.get("impact")),
                             self.hz(i.get("treatment")), tech(i)])
        for i in self.narrative.get("data_issues_extra") or []:
            if isinstance(i, dict):
                rows.append([i.get("id"), "ملاحظة المحلل", issue_type_ar(i.get("type")) if i.get("type") else "—",
                             sev(i.get("severity")), i.get("where"),
                             "، ".join(i.get("record_ids") or []), self.r(i.get("detail")), self.r(i.get("impact")),
                             self.r(i.get("treatment")), tech(i)])
        sh.table("tblIssues", ["المعرّف", "المصدر", "النوع", "الخطورة", "الموضع", "السجلات", "المشكلة", "الأثر", "المعالجة",
                               TECH], rows, freeze=True)
        sh.finish(fixed={7: 60, 8: 45, 9: 45})

    def gaps_recs(self, wb):
        sh = Sheet(wb, "الفجوات والتوصيات", "الفجوات المشخّصة والتوصيات المرتبطة بها",
                   "الأحمر للمشكلة المثبتة فقط؛ غير المثبت يحتاج تحققًا قبل القرار. كل توصية مرتبطة بفجوة ودليل.")
        n = self.narrative
        top_g, top_a = set(n.get("top_gaps") or []), set(n.get("top_actions") or [])
        rows = []
        for g in n.get("gaps") or []:
            if not isinstance(g, dict):
                continue
            proven = g.get("is_proven_problem") is True
            imp = g.get("impact") or {}
            mref = imp.get("metric")
            mval = None
            if isinstance(mref, str) and mref.startswith("m:"):
                mid_, _, fld = mref[2:].partition(".")
                res = self.metrics.get(mid_) or {}
                if fld and fld != "value":
                    cur = res.get(fld)
                    mval = fmt_value_text(cur, "SAR" if fld in ("abs_change", "start_change", "prev_sum", "curr_sum", "num", "den", "total")
                                          else ("count" if fld.startswith("n_") else res.get("unit")))
                elif growth_riyal_only(res) and is_num(res.get("value")):
                    mval = (fmt_value_text(res.get("abs_change"), "SAR") + f" (تغير بالريال؛ قابلية المقارنة {growth_comp_ar(res)}، لا نسبة)")
                else:
                    mval = fmt_value_text(res.get("value"), res.get("unit"))
            ev = [self.ev_text(e) for e in g.get("evidence") or []]
            rows.append([g.get("id"), "نعم" if g.get("id") in top_g else "", self.r(g.get("title")), self.scope_ar(g.get("scope")),
                         (self.vocab.get("severity") or {}).get(g.get("severity"), g.get("severity")),
                         C("مشكلة مثبتة", fill=RED_FILL, color=RED_TXT, bold=True) if proven else C("محتملة — تحتاج تحققًا"),
                         ev, self.r(imp.get("text")), f"{self.ev_text(mref)}: {mval}" if mref else None, self.r(g.get("proven_cause")) or "—",
                         [self.r(x) for x in g.get("possible_causes") or []], [self.r(x) for x in g.get("verification_needed") or []],
                         self.r(g.get("proposed_action"))])
        sh.heading("الفجوات")
        sh.table("tblGaps", ["المعرّف", "ضمن الأهم", "الفجوة", "النطاق", "الخطورة", "الإثبات", "الأدلة (مراجع)", "الأثر",
                             "مؤشر الأثر", "السبب المثبت", "الأسباب المحتملة", "ما يلزم للتحقق", "الإجراء المقترح"],
                 rows, freeze="A6")
        rows = []
        for x in sorted([x for x in n.get("recommendations") or [] if isinstance(x, dict)],
                        key=lambda x: (x.get("priority") if isinstance(x.get("priority"), int) else 999, str(x.get("id")))):
            rk = x.get("ranking") or {}
            rows.append([x.get("id"), x.get("priority"), "نعم" if x.get("id") in top_a else "", "، ".join(x.get("gap_ids") or []),
                         self.r(x.get("target")), self.r(x.get("problem")), [self.ev_text(e) for e in list_of(x.get("evidence"))],
                         self.r(x.get("first_action")),
                         [self.r(a) for a in x.get("alternatives") or []], self.r(x.get("recommended_option")),
                         self.r(x.get("rationale")), x.get("owner_role"), x.get("duration"), self.r(x.get("cost")),
                         self.r(x.get("expected_impact")), self.r(x.get("success_kpi")), x.get("approval_required") or "لا يلزم",
                         LEVEL_AR.get(rk.get("impact"), rk.get("impact")), LEVEL_AR.get(rk.get("urgency"), rk.get("urgency")),
                         LEVEL_AR.get(rk.get("evidence_strength"), rk.get("evidence_strength")),
                         LEVEL_AR.get(rk.get("feasibility"), rk.get("feasibility")),
                         LEVEL_AR.get(rk.get("cost_level"), rk.get("cost_level"))])
        sh.heading("التوصيات (مرتبة بالأولوية)")
        sh.table("tblRecommendations", ["المعرّف", "الأولوية", "ضمن الأهم", "الفجوات", "الأصل أو النشاط", "المشكلة", "الدليل",
                                        "الإجراء الأول", "البدائل", "الخيار الموصى به", "سبب الاختيار", "المسؤول المقترح",
                                        "المدة", "التكلفة أو متطلبات تقديرها", "الأثر المتوقع", "مؤشر النجاح",
                                        "الاعتماد المطلوب", "الأثر", "الإلحاح", "قوة الدليل", "قابلية التنفيذ", "مستوى التكلفة"],
                 rows)
        sh.finish(fixed={3: 36, 6: 36, 7: 30, 8: 40, 13: 40}, print_titles=False)

    def plan(self, wb):
        sh = Sheet(wb, "الخطة التشغيلية", "خطة التنفيذ: 30 و60 و90 يومًا",
                   "للمشروعات الطويلة يُذكر المخرج المرحلي خلال المدة (دراسة، عروض، تقييم) دون الإيحاء بانتهاء المشروع. عمود «الحالة» قائمة منسدلة للمتابعة.")
        recs = {x.get("id"): x for x in self.narrative.get("recommendations") or [] if isinstance(x, dict)}
        plan = self.narrative.get("plan") or {}
        rows = []
        for key, label in (("d30", "خلال 30 يومًا"), ("d60", "خلال 60 يومًا"), ("d90", "خلال 90 يومًا")):
            for it in plan.get(key) or []:
                if not isinstance(it, dict):
                    continue
                rec = recs.get(it.get("rec_id")) or {}
                status = it.get("status") if it.get("status") in PLAN_STATUSES else "لم يبدأ"
                rows.append([label, it.get("rec_id"), rec.get("priority"), self.r(rec.get("target")), self.r(it.get("action")),
                             self.r(it.get("deliverable")), it.get("owner_role"), status, None, None])
        hr, first, last = sh.table("tblPlan", ["المرحلة", "التوصية", "الأولوية", "الأصل أو النشاط", "الإجراء", "المخرج",
                                               "المسؤول", "الحالة", "تاريخ المتابعة", "ملاحظات"], rows, freeze=True)
        dv = DataValidation(type="list", formula1='"' + ",".join(PLAN_STATUSES) + '"', allow_blank=True,
                            showErrorMessage=True, errorTitle="قيمة غير مسموحة",
                            error="اختر: " + "، ".join(PLAN_STATUSES), promptTitle="الحالة", prompt="اختر حالة البند")
        sh.ws.add_data_validation(dv)
        dv.add(f"H{first}:H{max(last, first) + 50}")
        sh.finish(fixed={4: 30, 5: 45, 6: 35, 10: 30})

    # ------------------------------------------------- brief §12: versions and recommendation follow-up
    def _cat_label(self, labels, key):
        return (labels or {}).get(key) or dict(VC_CATEGORIES).get(key) or key

    def _signal_lines(self, item, labels):
        seen, out = set(), []
        for sg in item.get("data_signals") or []:
            if not isinstance(sg, dict):
                continue
            k = (sg.get("record_id"), sg.get("category"), sg.get("field"))
            if k in seen:
                continue
            seen.add(k)
            fld = sg.get("field")
            fl = "الاسم" if fld == "name" else ((self.meta.get(fld) or {}).get("label") or fld) if fld else None
            d = sg.get("delta")
            out.append(f"{sg.get('record_id')}: {self._cat_label(labels, sg.get('category'))}" + (f" — {fl}" if fl else "") +
                       (f" ({d:+,.0f})" if is_num(d) else ""))
        return out

    def followup(self, wb):
        """«متابعة التوصيات»: one row per previous recommendation, status dropdown «لم يبدأ/قيد التنفيذ/منجز/متعثر»."""
        items = [x for x in self.narrative.get("recommendation_followup") or [] if isinstance(x, dict)]
        if not items:
            return
        sh = Sheet(wb, "متابعة التوصيات", "متابعة توصيات النسخة السابقة",
                   "تُحدَّث الحالة («لم يبدأ» أو «قيد التنفيذ» أو «منجز» أو «متعثر») بما تثبته البيانات أو إفادة المستخدم فقط؛ "
                   "تحسن رقم أو تغير في الملف لا يثبت تنفيذ التوصية. عمود «الحالة» قائمة منسدلة.")
        recs = {x.get("id"): x for x in self.narrative.get("recommendations") or [] if isinstance(x, dict)}
        labels = (self.narrative.get("version_comparison") or {}).get("category_labels") \
            if isinstance(self.narrative.get("version_comparison"), dict) else None
        rows = []
        for it in items:
            rec = recs.get(it.get("rec_id")) or {}
            status = it.get("status")
            ev = it.get("evidence")
            has_ev = bool(ev) and not (isinstance(ev, str) and not ev.strip())
            note = self.r(it.get("note")) or ""
            if status in PLAN_STATUSES:
                fill, color = FOLLOW_STYLE.get(status, (None, None))
                if status == "منجز" and not has_ev:
                    fill, color = AMBER_FILL, AMBER_TXT
                    note = (note + "\n" if note else "") + "لا دليل مسجل — لا تُعد التوصية منجزة لمجرد تحسن رقم"
                st_cell = C(status, fill=fill, color=color)
            else:
                st_cell = C(status or "—", fill=AMBER_FILL, color=AMBER_TXT)
                note = (note + "\n" if note else "") + "الحالة خارج القائمة المعتمدة: " + "، ".join(PLAN_STATUSES)
            title = it.get("title") or rec.get("title") or it.get("target") or rec.get("target")
            rows.append([it.get("rec_id"), self.r(title), self.r(it.get("first_action") or rec.get("first_action")),
                         st_cell, [self.r(x.get("ref") or x.get("text") or str(x)) if isinstance(x, dict) else self.r(x)
                                   for x in ev] if isinstance(ev, list) and ev else
                         (self.r(ev) if has_ev else C("لا يوجد دليل بعد", color=GRAY_TXT)),
                         note, self._signal_lines(it, labels) or "—", None])
        hr, first, last = sh.table("tblFollowup", ["التوصية", "العنوان أو الأصل", "الإجراء الأول", "الحالة", "الدليل",
                                                   "ملاحظة", "إشارات من الملف (لا تثبت التنفيذ)", "تاريخ المتابعة"],
                                   rows, freeze=True)
        dv = DataValidation(type="list", formula1='"' + ",".join(PLAN_STATUSES) + '"', allow_blank=True,
                            showErrorMessage=True, errorTitle="قيمة غير مسموحة",
                            error="اختر: " + "، ".join(PLAN_STATUSES), promptTitle="الحالة", prompt="اختر حالة التوصية")
        sh.ws.add_data_validation(dv)
        dv.add(f"D{first}:D{max(last, first) + 50}")
        sh.finish(fixed={2: 34, 3: 40, 5: 36, 6: 44, 7: 44})

    def _vc_value(self, x, field=None):
        if isinstance(x, dict) and ("s" in x or "v" in x):
            return tv_cell(x, self.vocab)
        if x is None:
            return "—"
        return x if is_num(x) else str(x)

    def versions(self, wb):
        """«مقارنة الإصدارات»: the five categories of compare_versions.py, matching notes and metric deltas."""
        vc = self.narrative.get("version_comparison")
        if not isinstance(vc, dict) or not vc:
            return
        lists_of = (lambda d: d.get("categories") if isinstance(d.get("categories"), dict)
                    else d.get("changes") if isinstance(d.get("changes"), dict) else None)
        full = self.vc_file if isinstance(self.vc_file, dict) else None
        detail = vc if lists_of(vc) is not None else (full or vc)
        from_file = detail is full and detail is not vc
        labels = vc.get("category_labels") or (detail or {}).get("category_labels") or {}
        sh = Sheet(wb, "مقارنة الإصدارات", "مقارنة بالنسخة السابقة (الموجّه §12)",
                   "كل فرق في فئة واحدة. رفع نسخة أحدث من بيانات الفترة نفسها ليس نموًا، واختفاء سجل لا يعني بيعه."
                   + (" التفاصيل من version_comparison.json في مجلد التشغيل." if from_file else ""))
        get = (lambda k: vc.get(k) if vc.get(k) not in (None, "", [], {}) else (detail or {}).get(k))
        pi, ci = get("prev") or {}, get("curr") or {}
        mc = (get("matching") or {}).get("counts") or {}
        ctx = []
        if pi or ci:
            ctx.append(("النسختان", f"{pi.get('source_file') or pi.get('data_period') or 'السابقة'} ← "
                                     f"{ci.get('source_file') or ci.get('data_period') or 'الحالية'}"
                        + (" (الملف نفسه بالبصمة)" if get("same_source_file") else "")))
        if get("period_relation"):
            ctx.append(("العلاقة بين الفترتين", VC_RELATION_AR.get(get("period_relation"), get("period_relation"))))
        if get("new_periods"):
            ctx.append(("فترات جديدة", "، ".join(map(str, get("new_periods")))))
        if get("growth_note"):
            ctx.append(("قراءة الفروق", get("growth_note")))
        if mc:
            ctx.append(("مطابقة السجلات", " · ".join(f"{VC_CONF_AR[k]} {mc[k]}" for k in VC_CONF_AR if is_num(mc.get(k))) +
                        f" · في السابقة فقط {mc.get('only_prev', 0)} · في الجديدة فقط {mc.get('only_curr', 0)}"))
        for k in ("summary", "summary_ar"):
            if isinstance(vc.get(k), str) and vc[k].strip():
                ctx.append(("ملخص المحلل", self.r(vc[k])))
        if ctx:
            sh.table("tblVersionContext", ["البند", "القيمة"], [[C(a, bold=True), b] for a, b in ctx], filters=False)
        lists = lists_of(detail or {}) or {}
        counts = get("counts") or {}
        examples = [x for x in vc.get("examples") or [] if isinstance(x, dict)]
        cat_rows = []
        for key, _ in VC_CATEGORIES:
            items = lists.get(key) if isinstance(lists.get(key), list) else None
            n = counts.get(key) if is_num(counts.get(key)) else (len(items) if items is not None else
                                                                   (sum(1 for x in examples if x.get("category") == key)
                                                                    if examples else None))
            cat_rows.append([self._cat_label(labels, key), n if n is not None else C("غير مذكور", color=GRAY_TXT)])
        sh.heading("الفئات الخمس")
        sh.table("tblVersionCategories", ["الفئة", "عدد الفروق"], cat_rows, filters=False)
        rows = []
        for key, _ in VC_CATEGORIES:
            items = lists.get(key) if isinstance(lists.get(key), list) else [x for x in examples if x.get("category") == key]
            for c in items:
                if not isinstance(c, dict):
                    continue
                fld = c.get("field")
                fl = ("الاسم" if fld == "name" else (self.meta.get(fld) or {}).get("label") or fld) if fld \
                    else (c.get("key_ar") or c.get("key"))
                if fld == "name":
                    pv, cv = c.get("prev_name") or "—", c.get("curr_name") or "—"
                elif "prev_value" in c or "curr_value" in c:
                    pv, cv = self._vc_value(c.get("prev_value")), self._vc_value(c.get("curr_value"))
                else:
                    pv, cv = self._vc_value(c.get("prev")), self._vc_value(c.get("curr"))
                rows.append([self._cat_label(labels, key), VC_SUBTYPE_AR.get(c.get("subtype"), c.get("subtype")),
                             c.get("record_id_prev") or "—", c.get("record_id_curr") or "—", c.get("name"), fl, pv, cv,
                             C(c["delta"], fmt=FMT_INT if float(c["delta"]).is_integer() else FMT_DEC) if is_num(c.get("delta")) else None,
                             VC_CONF_AR.get(c.get("match_confidence"), c.get("match_confidence")), self.r(c.get("note"))])
        sh.heading("الفروق")
        sh.table("tblVersionChanges", ["الفئة", "النوع", "المعرّف السابق", "المعرّف الحالي", "الاسم", "الحقل أو البند",
                                       "السابق", "الحالي", "الفرق", "المطابقة", "ملاحظة"], rows, freeze=None)
        m = get("matching") or {}
        rows = []
        for x in m.get("id_conflicts") or vc.get("id_conflicts") or []:
            if isinstance(x, dict):
                res = [f"السابق ← {x['prev_matched_to']}" if x.get("prev_matched_to") else None,
                       f"الحالي ← {x['curr_matched_from']}" if x.get("curr_matched_from") else None]
                rows.append(["تعارض معرّف", x.get("id"), "—", f"{x.get('prev_name') or '—'} ← {x.get('curr_name') or '—'}",
                             x.get("note"), "؛ ".join(r for r in res if r) or "لم يُطابق"])
        for x in m.get("ambiguous") or vc.get("ambiguous") or []:
            if isinstance(x, dict):
                rows.append(["حالة غامضة", "، ".join(x.get("prev_ids") or []), "، ".join(x.get("curr_ids") or []), None,
                             x.get("note"), "تحتاج تحققًا يدويًا"])
        if rows:
            sh.heading("ما يحتاج مراجعة يدوية في المطابقة")
            sh.table("tblVersionReview", ["النوع", "المعرّف السابق", "المعرّف الحالي", "الاسم", "ملاحظة", "النتيجة"], rows,
                     filters=False)
        rows = []
        for x in get("metric_deltas") or []:
            if not isinstance(x, dict):
                continue
            af = x.get("affected_by") or {}
            unit = x.get("unit")

            hidden = bool(x.get("values_hidden"))
            pp = x.get("difference_unit") == "pp" or unit == "%"

            def mv(v, st):
                if hidden:
                    return C("قابلية المقارنة غير مؤكدة", fill=GRAY_FILL, color=GRAY_TXT)
                return C(v, fmt=unit_fmt(unit)) if is_num(v) else status_cell(st, self.vocab)
            diff = None
            if is_num(x.get("difference")) and not hidden:
                # two percentages differ by percentage points (not a relative %): written as a number + «نقطة مئوية»
                diff = C(round(x["difference"], 2), fmt='0.00 "نقطة مئوية"') if pp else C(x["difference"], fmt=unit_fmt(unit))
            rows.append([x.get("label") or x.get("id"), x.get("prev_period"), x.get("curr_period"),
                         mv(x.get("prev_value"), x.get("prev_status")), mv(x.get("curr_value"), x.get("curr_status")),
                         diff,
                         (UNIT_AR.get(unit, unit) or "") + ("؛ الفرق بالنقاط المئوية" if pp and unit else ""),
                         "؛ ".join(f"{self._cat_label(labels, k)} ({len(v) if isinstance(v, list) else v})" for k, v in af.items()) or "—",
                         x.get("definition_causes") or "—", x.get("note")])
        if rows:
            sh.heading("فروق المؤشرات بين النسختين")
            sh.table("tblMetricDeltas", ["المؤشر", "الفترة السابقة", "الفترة الحالية", "السابق", "الحالي", "الفرق", "الوحدة",
                                         "يتأثر بـ (عدد السجلات)", "تغير التعريف", "ملاحظة"], rows)
        sh.finish(fixed={1: 24, 5: 30, 6: 30, 11: 44}, print_titles=False)

    def assumptions_sources(self, wb):
        sh = Sheet(wb, "الافتراضات والمصادر", "الافتراضات والمصادر والأسئلة وطلبات البيانات",
                   "يميّز المرفق بين المعلومة الواردة بالملف، والحساب المشتق، والاستنتاج التحليلي، وافتراض السيناريو.")
        n = self.narrative
        kinds = self.vocab.get("statement_kind") or {}
        sh.heading("الافتراضات والاستنتاجات")
        sh.table("tblAssumptions", ["المعرّف", "النص", "النوع", "الأساس"],
                 [[a.get("id"), self.r(a.get("text")), (kinds.get(a.get("kind")) or {}).get("ar", a.get("kind")), self.r(a.get("basis"))]
                  for a in n.get("assumptions") or [] if isinstance(a, dict)])
        sh.heading("المصادر")
        sh.table("tblSources", ["النوع", "المرجع", "التاريخ", "حدود الاستخدام"],
                 [[SOURCE_TYPE_AR.get(s.get("type"), s.get("type")), self.hz(s.get("ref")), s.get("date"), self.r(s.get("limits"))]
                  for s in n.get("sources") or [] if isinstance(s, dict)])
        sh.heading("الأسئلة المؤثرة (ثلاثة كحد أقصى)")
        sh.table("tblQuestions", ["#", "السؤال"], [[i, self.r(q)] for i, q in enumerate(n.get("questions") or [], 1)])
        sh.heading("البيانات الإضافية المطلوبة وما ستتيحه")
        sh.table("tblDataRequests", ["الأولوية", "البند", "ما سيتيحه"],
                 [[d.get("priority"), self.r(d.get("item")), self.r(d.get("enables"))] for d in
                  sorted([d for d in n.get("data_requests") or [] if isinstance(d, dict)],
                         key=lambda d: d.get("priority") if isinstance(d.get("priority"), int) else 999)])
        sc = n.get("scenarios") or {}
        sh.heading("السيناريوهات")
        if sc.get("status") == "prepared":
            rows = []
            for it in sc.get("items") or []:
                if not isinstance(it, dict):
                    continue
                for case, lab in (("conservative", "محافظ"), ("base", "أساسي"), ("optimistic", "محسّن")):
                    c = it.get(case) or {}
                    rows.append([it.get("id"), self.r(it.get("opportunity")), "، ".join(it.get("asset_ids") or []), lab,
                                 *[C(c.get(k), fmt=FMT_INT) if is_num(c.get(k)) else c.get(k)
                                   for k in ("extra_revenue", "extra_expense", "net_effect", "capex")],
                                 c.get("start"),
                                 *[C(c.get(k), fmt=FMT_INT) if is_num(c.get(k)) else c.get(k) for k in ("year1_effect", "run_rate_effect")],
                                 "متكرر" if it.get("recurring") else "لمرة واحدة",
                                 "\n".join(f"{a.get('text')} — المصدر: {a.get('source')}" for a in it.get("assumptions") or [] if isinstance(a, dict))])
            sh.table("tblScenarios", ["المعرّف", "الفرصة", "الأصول", "السيناريو", "الإيراد الإضافي", "المصروف الإضافي",
                                      "صافي الأثر", "التكلفة الرأسمالية", "بدء الأثر", "أثر السنة الأولى", "الأثر السنوي بعد الاكتمال",
                                      "طبيعة الأثر", "الافتراضات ومصادرها"], rows)
        else:
            sh.note("لم تُعدّ السيناريوهات: " + (self.r(sc.get("reason")) or "لا تتوافر مدخلات موثقة") +
                    ("\nالمدخلات اللازمة: " + "، ".join(map(str, sc.get("required_inputs") or [])) if sc.get("required_inputs") else ""),
                    ncols=4)
            sh.row += 1
        notes = n.get("methodology_notes") or []
        if notes:
            sh.heading("ملاحظات المنهجية")
            sh.table("tblMethodology", ["#", "الملاحظة"], [[i, self.r(x)] for i, x in enumerate(notes, 1)])
        sh.ws.freeze_panes = "A4"
        sh.finish(fixed={2: 60, 3: 30, 4: 40}, print_titles=False)


def privacy_self_check(path: str):
    """Scan every cell of the saved workbook: text cells, and numbers under a personal-data header (a phone stored as a
    number). Other numbers are amounts or deed numbers — a 9-digit amount would look like a mobile number.
    Returns (block_hits, review_hits) as lists of (sheet, cell, kinds); values are never returned."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    block, review = [], []
    try:
        for ws in wb.worksheets:
            above = {}
            for row in ws.iter_rows():
                for c in row:
                    v = getattr(c, "value", None)
                    if v is None or isinstance(v, bool):
                        continue
                    col = getattr(c, "column", None)
                    if isinstance(v, (int, float)):
                        if not is_pii_header(above.get(col, "")):
                            continue
                        t = str(int(v)) if float(v).is_integer() else str(v)
                    elif isinstance(v, str):
                        t = v
                        above[col] = v
                    else:
                        continue
                    kb = pii_value_hits(t, ("block",))
                    if kb:
                        block.append((ws.title, c.coordinate, kb))
                    kr = pii_value_hits(t, ("review",))
                    if kr:
                        review.append((ws.title, c.coordinate, kr))
    finally:
        wb.close()
    return block, review


def _kinds_ar(hits):
    return "، ".join(sorted({PII_KIND_AR.get(k, k) for _s, _c, ks in hits for k in ks}))


def _refs(hits, n=10):
    return "، ".join(f"«{s}»!{c}" for s, c, _k in hits[:n]) + ("…" if len(hits) > n else "")


def export(run_dir: str, out: str) -> int:
    """Build, save and privacy-check the workbook. 0 = written; 2 = personal data found after saving (file deleted)."""
    PRIVACY.reset()
    exp = Export(run_dir)
    sheets = exp.build(out)
    block, review = privacy_self_check(out)
    if block:
        try:
            os.remove(out)
        except OSError:
            pass
        print(f"خطأ خصوصية: بعد الحفظ وُجد ما يشبه {_kinds_ar(block)} في {len(block)} خلية ({_refs(block)})؛ حُذف الملف "
              "ولم يُسلَّم. أزل مصدر هذه البيانات (عمود شخصي أو نص حر) ثم أعد التصدير.", file=sys.stderr)
        return 2
    n = exp.narrative
    print(f"تم: {out}")
    print(f"  الأوراق ({len(sheets)}): " + "، ".join(sheets))
    print(f"  السجلات {len(exp.records)} · المؤشرات {len(exp.metrics)} · الفجوات/التوصيات "
          f"{len(n.get('gaps') or [])}/{len(n.get('recommendations') or [])}")
    if PRIVACY.fields_dropped:
        print(f"  خصوصية: أُسقطت {len(PRIVACY.fields_dropped)} من الحقول الشخصية: " + "، ".join(PRIVACY.fields_dropped))
    if PRIVACY.masked:
        print(f"  خصوصية: حُجبت بيانات شخصية ({'، '.join(sorted(PII_KIND_AR.get(k, k) for k in PRIVACY.kinds))}) "
              f"في {PRIVACY.masked} نصًا")
    if review:
        print(f"  تنبيه خصوصية: {len(review)} خلية تحوي {PII_KIND_AR['national_id']} أو صكًا ({_refs(review, 6)})؛ "
              "راجعها واحذف أي رقم هوية شخصية", file=sys.stderr)
    print("  فحص الخصوصية بعد الحفظ: لا بيانات اتصال أو حسابات بنكية في الخلايا")
    if exp.narrative_missing:
        print("  تنبيه: لا يوجد 06_narrative.json؛ أقسام السرد فارغة")
    return 0


# ---------------------------------------------------------------------------
# Self-test (fictional data only; run by tests/run_tests.sh or by hand)
# ---------------------------------------------------------------------------
def _st_tv(v, s="value", raw=None, ref=None):
    return {"v": v, "s": s, "raw": raw if raw is not None else (None if v is None else str(v)), "ref": ref}


def _st_write(d, name, data):
    import json
    with open(os.path.join(d, name), "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False)


def _st_run(d):
    contact = "للتواصل: (050) 123-4567 أو 011-4567890 أو ٠٥٠١٢٣٤٥٦٧ أو +966 11 456 7890 أو someone@example.com"
    meta = {"rev.2025.excl_vat": {"label": "الإيراد 2025", "group": "revenue", "period": "2025", "vat": "excl"},
            "units": {"label": "عدد الوحدات", "group": "other"}, "notes": {"label": "ملاحظات", "group": "text"},
            "tenant_label": {"label": "اسم المستأجر", "group": "text"}, "supervisor_mobile": {"label": "جوال", "group": "text"}}

    def rec(rid, row, rev, scope="in_scope", src="T", **f):
        fields = {"rev.2025.excl_vat": _st_tv(rev, ref=f"S!C{row}")}
        fields.update(f)
        return {"id": rid, "name": f"عمارة تجريبية {rid}", "record_type": "asset", "scope": scope, "dims": {},
                "source": {"source_id": src, "sheet": "S", "row": row}, "fields": fields}
    meta["val.2024.total"] = {"label": "تقييم 2024", "group": "valuation", "unit": "SAR"}
    many = {f"area.p{i}": _st_tv(100 + i, ref=f"S!Z{i}") for i in range(14)}   # N8: more than 12 numeric refs
    records = [
        rec("T-1", 2, 100000, units=_st_tv(None, "text", "6 محلات + 4 مكاتب"), notes=_st_tv(contact),
            tenant_label=_st_tv("مستأجر وهمي"), supervisor_mobile=_st_tv("0509876543"), **many,
            **{"val.2024.total": _st_tv(900000, "derived", ref="S!X2"),
               "val.2024.total__original": _st_tv(1200000, ref="S!X2")}),
        rec("T-2", 3, 50000), rec("T-3", 4, 25000, scope="isolated"), rec("U-1", 2, 10, src="U"),
        # N4: an isolated statement with two section totals that match (rows 4, 6) and a typed grand total (row 7)
        rec("V-1", 2, 40, scope="isolated", src="V"), rec("V-2", 3, 60, scope="isolated", src="V"),
        rec("V-3", 5, 30, scope="isolated", src="V"),
    ]
    excluded = [{"source_id": "T", "sheet": "S", "row": 5, "reason": "صف إجمالي",
                 "values": {"rev.2025.excl_vat": _st_tv(150000, ref="S!C5")}},
                {"source_id": "U", "sheet": "S2", "row": 3, "reason": "صف إجمالي",
                 "values": {"rev.2025.excl_vat": _st_tv(12, ref="S2!C3")}},
                {"source_id": "V", "sheet": "S", "row": 4, "reason": "إجمالي القسم",
                 "values": {"rev.2025.excl_vat": _st_tv(100, ref="S!C4")}},
                {"source_id": "V", "sheet": "S", "row": 6, "reason": "إجمالي القسم",
                 "values": {"rev.2025.excl_vat": _st_tv(30, ref="S!C6")}},
                {"source_id": "V", "sheet": "S", "row": 7, "reason": "الإجمالي العام",
                 "values": {"rev.2025.excl_vat": _st_tv(90, ref="S!C7")}}]
    _st_write(d, "03_records.json", {"contract_version": "1.0", "source": {"file": "t.xlsx"}, "field_meta": meta,
                                     "records": records, "excluded_rows": excluded, "issues": [],
                                     "stats": {"records": 4}})
    # N6/BP-01: a growth result whose comparability is unverified is exported in riyals, its % only in a review note
    _st_write(d, "05_metrics.json", {"metrics": {"revenue_growth": {
        "id": "revenue_growth", "kind": "growth", "label": "تغير الإيراد", "value": 37.5, "unit": "%", "status": "complete",
        "abs_change": 7500, "n_scope": 1, "n_included": 1, "included_ids": ["T-2"], "excluded": [],
        "resolved_fields": {"prev": "rev.2024.actual", "curr": "rev.2025.excl_vat"},
        "comparability": {"status": "unverified", "reason": "أساس الضريبة غير موضح"}}}})
    _st_write(d, "07_validation.json", {"blocking_count": 0, "warning_count": 0, "acknowledged_count": 2,
                                        "generated_at": "2026-01-02T10:20:30", "checks": [
                                            {"id": "totals_reconciliation", "level": "acknowledged", "passed": False,
                                             "title": "مطابقة الإجماليات المدونة"},
                                            {"id": "growth_comparability", "level": "acknowledged", "passed": False,
                                             "title": "قابلية مقارنة الفترات"}]})
    vc = {"prev": {"source_file": "old.xlsx"}, "curr": {"source_file": "t.xlsx"}, "period_relation": "same_period",
          "counts": {"scope_change": 1, "data_correction": 1},
          "category_labels": dict(VC_CATEGORIES),
          "categories": {"scope_change": [{"category": "scope_change", "subtype": "absent_in_new", "record_id_prev": "T-9",
                                           "record_id_curr": None, "name": "أرض تجريبية", "note": "غير موجود في النسخة الجديدة"}],
                         "data_correction": [{"category": "data_correction", "subtype": "value_changed", "record_id_prev": "T-1",
                                              "record_id_curr": "T-1", "name": "عمارة تجريبية", "field": "rev.2025.excl_vat",
                                              "prev": _st_tv(90000), "curr": _st_tv(100000), "delta": 10000,
                                              "match_confidence": "exact", "note": "تصحيح"}]},
          "matching": {"counts": {"exact": 3, "deed": 0, "probable": 0, "only_prev": 1, "only_curr": 0},
                       "id_conflicts": [], "ambiguous": []},
          "metric_deltas": [{"id": "property_revenue", "label": "الإيرادات العقارية", "prev_value": 140000,
                             "curr_value": 150000, "difference": 10000, "unit": "SAR",
                             "affected_by": {"data_correction": ["T-1"]}, "definition_causes": [], "note": "ليس نموًا"}]}
    _st_write(d, "06_narrative.json", {
        "recommendations": [{"id": "R1", "target": "T-1", "first_action": "تحقق من العقد"}],
        "recommendation_followup": [
            {"rec_id": "R1", "status": "لم يبدأ", "evidence": [], "note": "حالة افتراضية",
             "data_signals": [{"record_id": "T-1", "category": "data_correction", "field": "rev.2025.excl_vat", "delta": 10000},
                              {"record_id": "T-1", "category": "data_correction", "field": "rev.2025.excl_vat", "delta": 10000}]},
            {"rec_id": "R2", "status": "منجز", "evidence": None}],
        "version_comparison": vc})


def self_test() -> int:
    import tempfile
    import openpyxl
    fails = []

    def check(name, ok, detail=""):
        print(("  PASS  " if ok else "  FAIL  ") + name + ("" if ok else f" — {detail}"))
        if not ok:
            fails.append(name)
    with tempfile.TemporaryDirectory() as d:
        _st_run(d)
        out = os.path.join(d, "supporting_workbook.xlsx")
        rc = export(d, out)
        check("export exits 0", rc == 0 and os.path.isfile(out), rc)
        if rc == 0:
            wb = openpyxl.load_workbook(out)
            texts = [str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.value is not None]
            blob = "\n".join(texts)
            block, _review = privacy_self_check(out)
            check("D1 no phone/landline/e-mail in any cell (separators, Arabic-Indic digits)", not block, block)
            check("D1 personal columns dropped (key and label)",
                  "مستأجر وهمي" not in blob and "0509876543" not in blob and "supervisor_mobile" not in blob, "")
            check("D11 original text of a text-in-numeric cell shown", "«6 محلات + 4 مكاتب»" in blob, "")
            for name in ("المطابقة مع الإجماليات", "متابعة التوصيات", "مقارنة الإصدارات"):
                check(f"sheet «{name}» present", name in wb.sheetnames, wb.sheetnames)
            if "المطابقة مع الإجماليات" in wb.sheetnames:
                ws = wb["المطابقة مع الإجماليات"]
                src = lambda r: str(r[14] or "").split(" · ")[0]  # noqa: E731 — the technical column «V · field»
                got = {(src(r), r[11]): r[7] for r in ws.iter_rows(min_row=5, values_only=True) if src(r) in ("T", "U")}
                check("A14 reconciliation: T matches 150,000 in scope; U differs by 2",
                      got == {("T", RECON_STATUS_AR["match"]): 0, ("U", RECON_STATUS_AR["mismatch"]): 2}, got)
                vrows = {r[1]: r for r in ws.iter_rows(min_row=5, values_only=True) if src(r) == "V"}
                sec, sec2, grand = vrows.get("S!C4"), vrows.get("S!C6"), vrows.get("S!C7")
                check("item 6: a section total lists its own section's lines (S!C4: V-1، V-2; S!C6: V-3), the grand total all",
                      bool(sec) and sec[12] == "V-1، V-2" and bool(sec2) and sec2[12] == "V-3"
                      and bool(grand) and grand[12] == "V-1، V-2، V-3", (sec and sec[12], sec2 and sec2[12], grand and grand[12]))
                check("BP-10: the sheet name, the Arabic field label and the key only in «المرجع التقني»",
                      bool(sec) and sec[0] == "S" and sec[2] == "الإيراد 2025" and sec[14] == "V · rev.2025.excl_vat", sec)
                check("N4 isolated section totals: section sums 100 and 30 (all records), difference 0, never «recorded − 0»",
                      bool(sec) and sec[8] == 100 and sec[7] == 0 and bool(sec2) and sec2[8] == 30 and sec2[7] == 0,
                      (sec, sec2))
                check("N4 isolated grand total: difference 90 − 130 = −40 against all records",
                      bool(grand) and grand[7] == -40 and grand[8] is None, grand)
            ws = wb["ملخص"]
            row = next((r for r in ws.iter_rows(values_only=True) if r and r[0] == "تغير الإيراد"), None)
            check("N6 summary growth row: 7,500 riyals, no percentage value, review note",
                  bool(row) and row[1] == 7500 and row[2] == "ريال" and "لا نسبة" in str(row[3]) and "للمراجعة فقط" in str(row[8]), row)
            nums = [c.value for sh in ("ملخص", "الحسابات") for r in wb[sh].iter_rows() for c in r if isinstance(c.value, (int, float))]
            check("N6 the growth % (37.5) is no value cell in ملخص/الحسابات", 37.5 not in nums, "")
            vline = next((r[1] for r in ws.iter_rows(values_only=True) if r and r[0] == "التحقق الآلي"), "")
            check("BP-11 validation row lists the acknowledged warnings by their Arabic titles and the validation time",
                  "مُقرّ بها 2" in str(vline) and "«قابلية مقارنة الفترات»" in str(vline)
                  and "growth_comparability" not in str(vline) and "2026-01-02 10:20" in str(vline), vline)
            ws = wb["الأصول"]
            tech = [c.value for c in next(ws.iter_rows(min_row=4, max_row=4))]
            hdr = [c.value for c in next(ws.iter_rows(min_row=5, max_row=5))]
            check("N7 «__original» column labelled as the stored value before the derivation",
                  any("كما ورد قبل الاشتقاق" in str(h) for h in hdr), hdr)
            check("BP-10 headers are Arabic labels only; the keys sit in the «المرجع التقني» row above them",
                  tech[0] == TECH + ":" and "val.2024.total__original" in tech
                  and not any(re.search(r"[a-z]+\.[a-z0-9_]+", str(h or "")) for h in hdr), (tech, hdr))
            j = hdr.index("مراجع القيم الرقمية") if "مراجع القيم الرقمية" in hdr else None
            refs = next((r[j] for r in ws.iter_rows(min_row=6, values_only=True) if r[0] == "T-1"), "") if j is not None else ""
            check("N8 every numeric source cell listed (no 12-line cut), metric fields first, by their labels",
                  str(refs).count("\n") >= 15 and str(refs).startswith("الإيراد 2025: "), refs)
            if "متابعة التوصيات" in wb.sheetnames:
                ws = wb["متابعة التوصيات"]
                dvs = [dv.formula1 for dv in ws.data_validations.dataValidation]
                check("§12 follow-up status dropdown uses «لم يبدأ/قيد التنفيذ/منجز/متعثر»",
                      any(all(s in (f or "") for s in PLAN_STATUSES) for f in dvs), dvs)
                cells = [str(c.value) for row in ws.iter_rows() for c in row if c.value is not None]
                check("§12 «منجز» without evidence is flagged", any("لا دليل مسجل" in c for c in cells), "")
                check("D7 follow-up data signals de-duplicated", sum(c.count("T-1:") for c in cells) == 1, cells)
            wb.close()
        # C24: anything that slips past the masking is caught after saving; the file is deleted and the exit is 2
        global mask_pii
        real = mask_pii
        mask_pii = lambda v, levels=("block",): (str(v), [])  # noqa: E731
        try:
            rc = export(d, out)
        finally:
            mask_pii = real
        check("C24 self-check: unmasked phone -> exit 2 and no file", rc == 2 and not os.path.exists(out), rc)
    print("RESULT:", "FAIL" if fails else "PASS")
    return 1 if fails else 0


def main(argv=None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    if "--self-test" in argv:
        return self_test()
    ap = argparse.ArgumentParser(description="تصدير المرفق الداعم (supporting_workbook.xlsx) لمجلد تشغيل")
    ap.add_argument("--run-dir", required=True, help="مجلد التشغيل")
    ap.add_argument("--out", help="مسار الملف الناتج (الافتراضي RUN/supporting_workbook.xlsx)")
    ap.add_argument("--self-test", action="store_true", help="تشغيل حالات الاختبار المضمنة (بيانات وهمية)")
    args = ap.parse_args(argv)
    for f in ("03_records.json", "05_metrics.json"):
        if not os.path.isfile(os.path.join(args.run_dir, f)):
            print(f"خطأ: الملف {f} غير موجود في {args.run_dir}", file=sys.stderr)
            return 1
    out = args.out or os.path.join(args.run_dir, "supporting_workbook.xlsx")
    return export(args.run_dir, out)


if __name__ == "__main__":
    sys.exit(main())
