#!/usr/bin/env python3
"""Export the reviewable supporting workbook of a run (data-contract.md §0: supporting_workbook.xlsx).

Usage:
  python3 scripts/export_workbook.py --run-dir RUN [--out RUN/supporting_workbook.xlsx]

Reads 03_records.json and 05_metrics.json (required) plus 01, 02, 04, 06, 07 and analysis.json when present.
Sheets: «ملخص»، «الأصول»، «الحسابات»، «المستبعدات»، «مشكلات البيانات»، «الفجوات والتوصيات»،
«الخطة التشغيلية»، «الافتراضات والمصادر». Right-to-left, navy headers, filters on every table, gray = not
available (never a zero), red = a proven problem only. No personal data: PII fields and phone/e-mail-like
text are never written. Values only, except SUM totals in the calculation detail of «الحسابات».
Exit 0 on success, 1 on unusable input.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argparse  # noqa: E402
import datetime as _dt  # noqa: E402
import re  # noqa: E402
from collections import OrderedDict, defaultdict  # noqa: E402

from apa_common import (  # noqa: E402
    EMAIL_RE, KNOWN_STATUSES, load_json, load_vocab,
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
GRAY_FILL, GRAY_TXT = "E7E6E6", "595959"
RED_FILL, RED_TXT = "FFC7CE", "9C0006"
AMBER_FILL, AMBER_TXT = "FFF2CC", "7F6000"
GREEN_FILL, GREEN_TXT = "E2EFDA", "375623"
BORDER = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
FMT_INT = "#,##0"
FMT_DEC = "#,##0.00"
FMT_PCT_UNIT = '#,##0.00" %"'      # metric values already multiplied by 100
FMT_SHARE = "0.0%"                 # shares stored in 0..1
PLAN_STATUSES = ["لم يبدأ", "قيد التنفيذ", "منجز", "متعثر"]
MAX_WIDTH, MIN_WIDTH = 60, 9
PHONE_LOOSE_RE = re.compile(r"(?<!\d)(?:\+?966|00966|0)?[\s\-]?5(?:[\s\-]?\d){8}(?!\d)")
PII_KEY_RE = re.compile(
    r"(?i)(^|[._\-])(phone|mobile|tel|telephone|email|e_mail|supervisor|supervisor_name|contact_phone|contact_no|"
    r"contact_number|national_id|id_number|iban)($|[._\-])|جوال|هاتف|تواصل|المشرف|ايميل|إيميل|بريد|هوية")
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
}
KIND_AR = {"count": "عدّ", "sum": "مجموع", "ratio": "نسبة من مجاميع متطابقة", "diff": "فرق", "growth": "تغير بين فترتين",
           "concentration": "تركّز", "coverage": "تغطية بيانات"}
SOURCE_TYPE_AR = {"file": "ملف", "web": "مصدر ويب", "user": "إفادة المستخدم"}
LEVEL_AR = {"high": "مرتفع", "medium": "متوسط", "low": "منخفض", "strong": "قوي", "moderate": "متوسط", "weak": "ضعيف"}
DIM_AR_DEFAULT = {"income_generation": "توليد الدخل", "income_stability": "استقرار الدخل",
                  "asset_utilization": "كفاءة استخدام الأصول", "data_quality": "جودة البيانات"}


# ---------------------------------------------------------------------------
# Cell spec and sheet writer
# ---------------------------------------------------------------------------
class C:
    """A cell value with optional formatting."""
    __slots__ = ("v", "fmt", "fill", "color", "bold", "wrap", "formula")

    def __init__(self, v, fmt=None, fill=None, color=None, bold=False, wrap=True, formula=False):
        self.v, self.fmt, self.fill, self.color, self.bold, self.wrap = v, fmt, fill, color, bold, wrap
        self.formula = formula  # only deliberate formulas (SUM totals) are written unescaped


def is_num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def safe_text(x):
    """Text for a cell; personal phone / e-mail patterns are masked."""
    if x is None:
        return None
    if isinstance(x, (list, tuple)):
        x = "\n".join(str(i) for i in x if i is not None and str(i) != "")
    elif isinstance(x, dict):
        x = "؛ ".join(f"{k}: {v}" for k, v in x.items())
    s = str(x)
    if PHONE_LOOSE_RE.search(s.translate(str.maketrans("٠١٢٣٤٥٦٧٨٩", "0123456789"))) or EMAIL_RE.search(s):
        s = PHONE_LOOSE_RE.sub("[حُجب]", s)
        s = EMAIL_RE.sub("[حُجب]", s)
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
            v = safe_text(v)
        elif v is not None and not is_num(v) and not isinstance(v, (_dt.date, _dt.datetime)):
            v = safe_text(v)
        cell = self.ws.cell(row=r, column=col, value=v)
        cell.border = BORDER
        if header:
            cell.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, readingOrder=2)
        else:
            cell.font = Font(name=FONT, size=10, bold=spec.bold, color=spec.color)
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

    def table(self, name, headers, rows, freeze=None, filters=True):
        """Write a table with a navy header row; returns (header_row, first_data_row, last_row)."""
        headers = _unique_headers(headers)
        if not rows:
            rows = [[C("لا توجد بنود", color=GRAY_TXT)] + [None] * (len(headers) - 1)]
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


def status_cell(status, vocab):
    lab = ((vocab.get("metric_status") or {}).get(status) or {}).get("ar") or status or "—"
    fill, color = METRIC_STATUS_STYLE.get(status, (None, None))
    return C(lab, fill=fill, color=color)


def tv_cell(tv, vocab, fmt=None):
    """Known -> number (or text); unknown -> gray text label, never a zero."""
    if not isinstance(tv, dict):
        return C("غير موجود", fill=GRAY_FILL, color=GRAY_TXT)
    s, v = tv.get("s"), tv.get("v")
    if s in KNOWN_STATUSES:
        if is_num(v):
            return C(v, fmt=fmt or (FMT_INT if float(v).is_integer() else FMT_DEC))
        return C(v)
    if s == "value" and v is not None:
        return C(v)
    lab = ((vocab.get("status") or {}).get(s) or {}).get("ar") or s or "غير متاح"
    return C(f"غير متاح — {lab}", fill=GRAY_FILL, color=GRAY_TXT)


def ar(vocab, dim, value):
    if value is None:
        return "غير محدد"
    if isinstance(value, list):
        return "، ".join(ar(vocab, dim, v) for v in value) or "غير محدد"
    if isinstance(value, bool):
        return "نعم" if value else "لا"
    table = vocab.get(dim) or {}
    return table.get(value, value) if isinstance(table, dict) else value


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
        self.pii_fields = set()
        for src in self.mapping.get("sources") or []:
            for col, spec in ((src or {}).get("columns") or {}).items():
                if isinstance(spec, dict) and spec.get("kind") == "pii" and spec.get("field"):
                    self.pii_fields.add(spec["field"])
        self.spec_by_id = {m.get("id"): m for m in self.config.get("metrics") or [] if isinstance(m, dict)}

    def r(self, x):
        return render_tokens(x, self.metrics)

    def allowed_field(self, f):
        return f not in self.pii_fields and not PII_KEY_RE.search(f)

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
        self.issues(wb)
        self.gaps_recs(wb)
        self.plan(wb)
        self.assumptions_sources(wb)
        wb.active = 0
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        wb.save(out_path)

    def _file_line(self):
        src = self.rec_doc.get("source") or {}
        name = os.path.basename(src.get("file") or "") or "—"
        sha = (src.get("sha256") or "")[:12]
        return f"الملف المصدر: {name}" + (f" · البصمة {sha}…" if sha else "")

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
            ("التقييم المعتمد", " — ".join(x for x in (val.get("label"), val.get("basis"), val.get("date_text")) if x) or "—"),
            ("السجلات", f"{st.get('records', len(self.records))} سجل: ضمن النطاق {by_scope.get('in_scope', 0)}، "
                        f"معزول {by_scope.get('isolated', 0)}، مستبعد {by_scope.get('excluded', 0)}"),
            ("التحقق الآلي", (f"أخطاء مانعة {v.get('blocking_count')} · تحذيرات {v.get('warning_count')}" if v else "لم يُشغَّل بعد")),
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
            unit = res.get("unit")
            val_cell = (C(res.get("value"), fmt=unit_fmt(unit)) if is_num(res.get("value"))
                        else C("غير قابل للاحتساب بالبيانات الحالية" if res.get("status") != "empty" else "لا توجد سجلات",
                               fill=GRAY_FILL, color=GRAY_TXT))
            rows.append([C(res.get("label") or mid, bold=bool(res.get("card"))), val_cell, UNIT_AR.get(unit, unit or "—"),
                         status_cell(res.get("status"), self.vocab),
                         f"{res.get('n_included', 0)} من {res.get('n_scope', 0)}",
                         res.get("scope_text"), res.get("basis"), res.get("valuation_date"),
                         res.get("reason") or ("بطاقة في الداشبورد" if res.get("card") else "")])
        sh.table("tblKpis", ["المؤشر", "القيمة", "الوحدة", "الحالة", "المشمول من النطاق", "النطاق", "الأساس",
                             "تاريخ التقييم", "السبب أو الملاحظة"], rows)
        gaps = {g.get("id"): g for g in self.narrative.get("gaps") or [] if isinstance(g, dict)}
        recs = {x.get("id"): x for x in self.narrative.get("recommendations") or [] if isinstance(x, dict)}
        sh.heading("أهم ثلاث فجوات وأهم ثلاثة إجراءات")
        rows = []
        tg = list(self.narrative.get("top_gaps") or [])[:3]
        ta = list(self.narrative.get("top_actions") or [])[:3]
        for i in range(max(len(tg), len(ta))):
            g = gaps.get(tg[i]) if i < len(tg) else None
            a = recs.get(ta[i]) if i < len(ta) else None
            if g or a:
                rows.append([i + 1, self.r(g.get("title")) if g else "—", self.r(a.get("first_action")) if a else "—",
                             (a or {}).get("owner_role") or "—"])
        sh.table("tblTop", ["الفجوة", "الإجراء أو القرار المطلوب", "#", "المسؤول المقترح"],
                 [[r[1], r[2], r[0], r[3]] for r in rows], filters=False)
        sh.ws.freeze_panes = "A4"
        sh.finish(fixed={1: 40, 2: 22, 3: 7, 4: 20, 5: 11, 6: 34, 7: 30, 8: 16, 9: 30}, print_titles=False, fit_width=1)

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
        headers += [f"{(self.meta.get(f) or {}).get('label') or f} ({f})" for f in numeric]
        headers += [f"{(self.meta.get(f) or {}).get('label') or f} ({f})" for f in text]
        headers += ["المرجع", "مراجع القيم الرقمية", "مشكلات البيانات"]
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
                row.append(tv_cell(fields.get(f), self.vocab))
            for f in text:
                tv = fields.get(f)
                row.append(tv_cell(tv, self.vocab) if isinstance(tv, dict) and tv.get("s") not in KNOWN_STATUSES else
                           (tv or {}).get("v"))
            src = rec.get("source") or {}
            refs = [f"{f}: {(fields.get(f) or {}).get('ref')}" for f in numeric if (fields.get(f) or {}).get("ref")]
            row += [src.get("ref_name") or (f"{src.get('sheet')}!{src.get('row')}" if src.get("sheet") else None),
                    "\n".join(refs[:12]), "، ".join(rec.get("issue_ids") or [])]
            rows.append(row)
        sh.table("tblAssets", headers, rows, freeze="C5")
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
            excl = res.get("excluded") or []
            val_cell = (C(res.get("value"), fmt=unit_fmt(unit)) if is_num(res.get("value"))
                        else C("غير متاح", fill=GRAY_FILL, color=GRAY_TXT))
            share = res.get("coverage_num_share")
            rows.append([
                mid, res.get("label"), KIND_AR.get(res.get("kind"), res.get("kind")), val_cell, UNIT_AR.get(unit, unit),
                status_cell(res.get("status"), self.vocab), res.get("n_scope"), res.get("n_included"), len(excl),
                "؛ ".join(f"{k}: {v}" for k, v in rf.items()) if rf else "—",
                C(res.get("num"), fmt=FMT_INT) if is_num(res.get("num")) else None,
                C(res.get("den"), fmt=FMT_INT) if is_num(res.get("den")) else None,
                C(share, fmt=FMT_SHARE) if is_num(share) else None,
                res.get("definition"), res.get("scope_text"), res.get("basis"), res.get("valuation_date"),
                res.get("period_text"), res.get("source_text"), res.get("reason"),
                "، ".join(res.get("included_ids") or []),
                "، ".join(f"{e.get('id')} ({EXCL_REASON_AR.get(e.get('reason'), e.get('reason'))})" for e in excl[:60]),
                ("نعم" if spec.get("card") else "لا"),
            ])
        sh.table("tblMetrics", ["المعرّف", "المؤشر", "النوع", "القيمة", "الوحدة", "الحالة", "سجلات النطاق", "المشمول",
                                "المستبعد", "الحقول المستخدمة", "البسط", "المقام", "حصة الإيراد المعروف المغطاة",
                                "التعريف", "النطاق", "الأساس", "تاريخ التقييم", "الفترة", "المصدر", "السبب",
                                "السجلات المشمولة", "السجلات المستبعدة وسببها", "بطاقة"], rows, freeze="C5")
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
            sh.note(f"{res.get('label') or mid} ({mid}) — {len(ids)} سجل مشمول", ncols=6)
            hr = sh.row
            heads = ["المعرّف", "الاسم"] + [f"{(self.meta.get(f) or {}).get('label') or f} ({f})" for f in fields] + ["المرجع"]
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
                             row.get("n_known"), row.get("n_unknown"), b.get("note") or ("قد تتداخل الفئات" if b.get("overlap") else ""), bid])
        sh.table("tblBreakdowns", ["التوزيع", "الفئة", "القيمة", "الحصة", "سجلات معروفة", "سجلات غير معروفة", "ملاحظة", "المعرّف"], rows)

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
                         ex.get("reason"), "جميع المؤشرات"])
        for ex in self.narrative.get("exclusions") or []:
            if isinstance(ex, dict):
                rows.append(["استبعاد تحليلي", self.r(ex.get("what")), ex.get("where"), self.r(ex.get("reason")), "—"])
        for mid, res in self.metrics.items():
            res = res or {}
            for e in res.get("excluded") or []:
                st = e.get("reason")
                reason = EXCL_REASON_AR.get(st) or (((self.vocab.get("status") or {}).get(st) or {}).get("ar")) or st
                rows.append(["مستبعد من مؤشر", e.get("id"), None, reason, res.get("label") or mid])
        sh.table("tblExclusions", ["النوع", "البند", "الموضع", "السبب", "المؤشر المتأثر"], rows, freeze=True)
        sh.finish(fixed={2: 40, 4: 50, 5: 40})

    def issues(self, wb):
        sh = Sheet(wb, "مشكلات البيانات", "سجل مشكلات البيانات (منفصل عن أداء المحفظة)",
                   "المصدر: فحص المصنف (01)، التطبيع (03)، أو ملاحظات المحلل (06). الخطورة تخص أثر المشكلة على الحساب، وليست حكمًا على الأصل.")
        sev_ar = self.vocab.get("severity") or {}
        rows = []

        def sev(s):
            fill, color = SEVERITY_STYLE.get(s, (None, None))
            return C(sev_ar.get(s, s or "—"), fill=fill, color=color)
        for a in self.wbmap.get("anomalies") or []:
            if isinstance(a, dict):
                rows.append([a.get("id"), "فحص المصنف", a.get("type"), sev(a.get("severity")), a.get("where"), None,
                             a.get("detail"), None, None])
        for i in self.rec_doc.get("issues") or []:
            if isinstance(i, dict):
                rows.append([i.get("id"), "التطبيع", i.get("type"), sev(i.get("severity")), i.get("where"),
                             "، ".join(i.get("record_ids") or []), i.get("detail"), i.get("impact"), i.get("treatment")])
        for i in self.narrative.get("data_issues_extra") or []:
            if isinstance(i, dict):
                rows.append([i.get("id"), "ملاحظة المحلل", i.get("type") or "—", sev(i.get("severity")), i.get("where"),
                             "، ".join(i.get("record_ids") or []), self.r(i.get("detail")), self.r(i.get("impact")),
                             self.r(i.get("treatment"))])
        sh.table("tblIssues", ["المعرّف", "المصدر", "النوع", "الخطورة", "الموضع", "السجلات", "المشكلة", "الأثر", "المعالجة"],
                 rows, freeze=True)
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
                res = self.metrics.get(mref[2:].split(".")[0]) or {}
                mval = fmt_value_text(res.get("value"), res.get("unit"))
            ev = [(e.get("ref") if isinstance(e, dict) else e) for e in g.get("evidence") or []]
            rows.append([g.get("id"), "نعم" if g.get("id") in top_g else "", self.r(g.get("title")), g.get("scope"),
                         (self.vocab.get("severity") or {}).get(g.get("severity"), g.get("severity")),
                         C("مشكلة مثبتة", fill=RED_FILL, color=RED_TXT, bold=True) if proven else C("محتملة — تحتاج تحققًا"),
                         ev, self.r(imp.get("text")), f"{mref}: {mval}" if mref else None, self.r(g.get("proven_cause")) or "—",
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
                         self.r(x.get("target")), self.r(x.get("problem")), x.get("evidence"), self.r(x.get("first_action")),
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
                 [[SOURCE_TYPE_AR.get(s.get("type"), s.get("type")), s.get("ref"), s.get("date"), s.get("limits")]
                  for s in n.get("sources") or [] if isinstance(s, dict)])
        sh.heading("الأسئلة المؤثرة (ثلاثة كحد أقصى)")
        sh.table("tblQuestions", ["#", "السؤال"], [[i, self.r(q)] for i, q in enumerate(n.get("questions") or [], 1)])
        sh.heading("البيانات الإضافية المطلوبة وما ستتيحه")
        sh.table("tblDataRequests", ["الأولوية", "البند", "ما سيتيحه"],
                 [[d.get("priority"), self.r(d.get("item")), d.get("enables")] for d in
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


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="تصدير المرفق الداعم (supporting_workbook.xlsx) لمجلد تشغيل")
    ap.add_argument("--run-dir", required=True, help="مجلد التشغيل")
    ap.add_argument("--out", help="مسار الملف الناتج (الافتراضي RUN/supporting_workbook.xlsx)")
    args = ap.parse_args(argv)
    for f in ("03_records.json", "05_metrics.json"):
        if not os.path.isfile(os.path.join(args.run_dir, f)):
            print(f"خطأ: الملف {f} غير موجود في {args.run_dir}", file=sys.stderr)
            return 1
    out = args.out or os.path.join(args.run_dir, "supporting_workbook.xlsx")
    exp = Export(args.run_dir)
    exp.build(out)
    n = exp.narrative
    print(f"تم: {out}")
    print(f"  الأوراق: ملخص، الأصول ({len(exp.records)} سجل)، الحسابات ({len(exp.metrics)} مؤشر)، المستبعدات، "
          f"مشكلات البيانات، الفجوات والتوصيات ({len(n.get('gaps') or [])}/{len(n.get('recommendations') or [])})، "
          "الخطة التشغيلية، الافتراضات والمصادر")
    if exp.narrative_missing:
        print("  تنبيه: لا يوجد 06_narrative.json؛ أقسام السرد فارغة")
    return 0


if __name__ == "__main__":
    sys.exit(main())
