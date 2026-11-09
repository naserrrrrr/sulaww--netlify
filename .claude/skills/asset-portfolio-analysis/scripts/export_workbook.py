#!/usr/bin/env python3
"""Export the reviewable supporting workbook of a run (data-contract.md §0: supporting_workbook.xlsx).

Usage:
  python3 "$SKILL_DIR/scripts/export_workbook.py" --run-dir "$RUN" [--out RUN/supporting_workbook.xlsx]

Reads 03_records.json and 05_metrics.json (required) plus 01, 02, 04, 06, 07, analysis.json and
version_comparison.json when present.
Sheets: «ملخص»، «الأصول»، «الحسابات»، «المستبعدات»، «المطابقة مع الإجماليات» (when excluded total rows kept their
values)، «مشكلات البيانات»، «الفجوات والتوصيات»، «الخطة التشغيلية»، «متابعة التوصيات» and «مقارنة الإصدارات» (when
06 has recommendation_followup / version_comparison)، «الافتراضات والمصادر»، «قوائم الاختيار» (the plan's dropdown lists).
Right-to-left, navy headers, filters on every table, gray = not available (never a zero; the original text of a
non-numeric cell is shown next to its status), red = a proven problem only; riyals #,##0 and percentages 0.0"%" as the
dashboard shows them.
Formulas (finance traceability): each sum/ratio/difference/change metric of «الحسابات» is recomputed from its included
records (SUM per field — a literal 0 when no record is included), beside the value the dashboard shows (a number) and a
check cell «مطابق/فرق»; «ملخص» links its KPI value to that recomputation; «المطابقة مع الإجماليات» writes «المدون −
مجموع النطاق» and «المدون − مجموع المقارنة» (the sum the total matched, with its basis) as live differences. Each formula
stores the value it evaluates to (a preview that does not recalculate shows it); after saving, a formula whose range
holds its own cell (a circular reference) deletes the file and exits 2.
Board wording: engine keys (fields, metrics, issue types, check ids, gap scopes, evidence tokens) are shown by their
Arabic labels; the keys themselves only in code-styled «المرجع التقني» columns/rows; no working-file names or step
numbers; record codes in sentences read as the record's name; a growth whose comparability is not confirmed is shown
in riyals only (no percentage anywhere); a metric flagged for verification (05 «review», C1) reads «يحتاج تحقق — …»;
the notes of a record in «الأصول» are its 03 issues plus the analyst's notes (C3); «الأولوية n» is the recommendation's
priority everywhere (C4).

Personal data: fields whose key, label or mapping kind is personal (assets/pii_patterns.json via apa_common) are
dropped; block-level patterns (mobile, landline, e-mail, IBAN; Arabic-Indic digits included) are masked in every text
cell; review-level hits (a 10-digit number that may be a national ID or a deed) are kept and reported. After saving,
every cell is scanned again: a block-level hit deletes the file and exits 2.
Exit 0 on success, 1 on unusable input, 2 when the privacy or the circular-formula self-check fails. `--self-test`
runs built-in cases (fictional data); tests/test_workbook_export.py also recalculates the workbook in LibreOffice.
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
    pii_value_hits, study_date_meta, study_date_text,
)

try:
    from openpyxl import Workbook
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table
except ImportError:  # pragma: no cover
    print("خطأ: مكتبة openpyxl غير مثبتة (python3 -m pip install openpyxl)", file=sys.stderr)
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
FMT_INT = "#,##0"                  # riyals and counts: no decimals, as the dashboard shows them
FMT_DEC = "#,##0.00"               # non-money measures that are fractional (an area, a rate)
FMT_PCT_UNIT = '0.0"%"'            # metric values already multiplied by 100: «6.6%», as the dashboard
FMT_SHARE = "0.0%"                 # shares stored in 0..1
FMT_DATE = "yyyy-mm-dd"
PLAN_STATUSES = ["لم يبدأ", "قيد التنفيذ", "منجز", "متعثر"]  # plan items and recommendation follow-up
PLAN_PHASES = [("d30", "خلال 30 يومًا", 30), ("d60", "خلال 60 يومًا", 60), ("d90", "خلال 90 يومًا", 90)]
# generic owner roles offered in the «المسؤول» dropdown after the roles the narrative names (any other text is accepted)
GENERIC_OWNERS = ["مجلس الإدارة", "لجنة الاستثمار", "لجنة المراجعة", "المدير التنفيذي", "المدير المالي",
                  "مدير إدارة الأوقاف والاستثمار", "مسؤول إدارة الأملاك", "المستشار القانوني"]
LISTS_SHEET = "قوائم الاختيار"
REVIEW_STATUS = "needs_verification"   # C1: a metric spec flagged for verification (04 → 05 «review»)
REVIEW_AR = "يحتاج تحقق"
DASH_VALUE_AR = "القيمة المعروضة في الداشبورد"
MATCH_AR, DIFF_AR = "مطابق", "فرق"
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
    brief §6), exactly as the dashboard's «تغير بالريال فقط». The percentage is written nowhere in the workbook."""
    comp = res.get("comparability") if isinstance(res.get("comparability"), dict) else {}
    return res.get("kind") == "growth" and comp.get("status") != "confirmed"


def growth_comp_ar(res: dict) -> str:
    comp = res.get("comparability") if isinstance(res.get("comparability"), dict) else {}
    return "غير صالحة" if comp.get("status") == "not_comparable" else "غير متحققة"


def growth_status_cell(res: dict):
    return C(f"قابلية المقارنة {growth_comp_ar(res)} — لا نسبة", fill=AMBER_FILL, color=AMBER_TXT)


def growth_pct_note(res: dict) -> str:
    """The dashboard's wording for an unconfirmed growth: the change in riyals only, no growth percentage."""
    if not is_num(res.get("value")):
        return ""
    comp = res.get("comparability") if isinstance(res.get("comparability"), dict) else {}
    why = comp.get("reason") if isinstance(comp.get("reason"), str) and comp.get("reason").strip() else ""
    return (f"تغير بالريال فقط — قابلية المقارنة {growth_comp_ar(res)}، فلا تُعرض نسبة نمو"
            + (f" ({why.strip()})" if why else ""))


def metric_review(res: dict, spec=None):
    """C1: the «review» flag a metric spec carries (copied by compute_metrics into the 05 result): its reason when the
    status is «needs_verification», else None."""
    for src in (res, spec or {}):
        rv = src.get("review") if isinstance(src, dict) else None
        if isinstance(rv, dict) and rv.get("status") == REVIEW_STATUS and isinstance(rv.get("reason"), str) \
                and rv["reason"].strip():
            return rv["reason"].strip()
    return None


def review_cell(reason: str):
    return C(f"{REVIEW_AR} — {reason}", fill=AMBER_FILL, color=AMBER_TXT)


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
# a record code in a sentence (prefix-number, e.g. «T-1»); only codes of records in 03 are replaced by their names
_REC_ID_RE = re.compile(r"(?<![A-Za-z0-9_.\-])[A-Za-z][A-Za-z0-9_]*-[A-Za-z0-9_]+(?:-[A-Za-z0-9_]+)*(?![A-Za-z0-9_\-])")
FIELD_PART_AR = {"abs_change": "التغير بالريال", "num": "البسط", "den": "المقام", "n_included": "المشمول",
                 "n_scope": "سجلات النطاق", "coverage_num_share": "حصة الإيراد المعروف", "prev_sum": "مجموع الفترة السابقة",
                 "curr_sum": "مجموع الفترة الحالية", "start_change": "تغير بدء الإيراد", "total": "المجموع",
                 "top_sum": "مجموع الأكبر", "a_sum": "مجموع الإيراد", "b_sum": "مجموع المصروف", "hhi": "مؤشر هيرفندال"}
KIND_AR = {"count": "عدّ", "sum": "مجموع", "ratio": "نسبة من مجاميع متطابقة", "diff": "فرق", "growth": "تغير بين فترتين",
           "concentration": "تركّز", "coverage": "تغطية بيانات"}
SOURCE_TYPE_AR = {"file": "ملف", "web": "مصدر ويب", "user": "إفادة المستخدم"}
# Levels agree with their heading, as in the dashboard: masculine for «الأثر» and «الإلحاح», feminine for «الأهمية»
# (gap and issue severity), «قوة الدليل», «قابلية التنفيذ» and «درجة التكلفة».
LEVEL_AR = {"high": "مرتفع", "medium": "متوسط", "low": "منخفض", "strong": "قوي", "moderate": "متوسط", "weak": "ضعيف",
            "info": "معلوماتي"}
LEVEL_F = {"high": "مرتفعة", "medium": "متوسطة", "low": "منخفضة", "strong": "قوية", "moderate": "متوسطة", "weak": "ضعيفة",
           "info": "معلوماتية"}
RANKING_FEM = {"evidence_strength", "feasibility", "cost_level"}
# Reconciliation of recorded totals (brief §4/§5.8) — the statuses of validate_analysis.py (07.reconciliation). The
# sheet's «الفرق عن مجموع النطاق» is always the recorded value minus the in-scope sum; each text says what that
# difference is when the total matched another sum («مجموع المقارنة»).
RECON_STATUS_AR = {
    "match": "مطابق لمجموع السجلات ضمن النطاق",
    "match_section": "مطابق لمجموع سجلات قسمه ضمن النطاق (إجمالي فرعي)؛ الفرق عن مجموع النطاق سببه الأقسام الأخرى",
    "match_all_records": "يطابق مجموع كل سجلات المصدر؛ الفرق عن مجموع النطاق يعود إلى سجلات معزولة أو مستبعدة",
    "match_section_all_records": "يطابق مجموع سجلات قسمه كلها؛ الفرق عن مجموع النطاق يعود إلى سجلات معزولة أو مستبعدة",
    "match_without_text_numbers": "يطابق مجموع السجلات ضمن النطاق دون أرقام مخزنة كنص في الملف (لا تجمعها صيغ Excel)؛ "
                                  "الإجمالي المدون ناقص بمقدار الفرق عن مجموع النطاق",
    "match_original": "يطابق القيم الأصلية قبل الاشتقاق الموثق في التصنيف؛ الفرق عن مجموع النطاق سببه الاشتقاق لا نقص في "
                      "السجلات",
    "mismatch": "فرق غير مفسَّر: لا يطابق الإجمالي المدون أي مجموع محسوب (انظر «الفرق عن مجموع المقارنة»)",
    "recorded_unknown": "القيمة المدونة في صف الإجمالي غير معروفة (فارغة أو نص أو خطأ)",
}
# when the source has no in-scope value in the field (an isolated statement), there is no in-scope sum to differ from
RECON_STATUS_NO_SCOPE_AR = {
    "match_all_records": "يطابق مجموع كل سجلات المصدر (لا قيم ضمن النطاق في هذا المصدر؛ سجلاته معزولة أو مستبعدة)",
    "match_section_all_records": "يطابق مجموع سجلات قسمه كلها (لا قيم ضمن النطاق في هذا المصدر؛ سجلاته معزولة أو مستبعدة)",
}
RECON_BASIS_AR = {
    "match": "السجلات ضمن النطاق",
    "match_section": "سجلات قسمه ضمن النطاق",
    "match_all_records": "كل سجلات المصدر (المعزولة والمستبعدة ضمنها)",
    "match_section_all_records": "سجلات قسمه كلها (المعزولة والمستبعدة ضمنها)",
    "match_without_text_numbers": "السجلات ضمن النطاق دون الأرقام المخزنة كنص",
    "match_original": "القيم الأصلية قبل الاشتقاق (السجلات ضمن النطاق)",
    "match_original_section": "القيم الأصلية قبل الاشتقاق (سجلات قسمه ضمن النطاق)",
    "mismatch_in_scope": "السجلات ضمن النطاق — لا مطابقة",
    "mismatch_all_records": "كل سجلات المصدر (لا قيم ضمن النطاق) — لا مطابقة",
}
# pipeline words and working-file names never reach a board reader (plain Arabic instead)
PLAIN_AR = [
    (re.compile(r"\b05_metrics\.json\b"), "نتائج المؤشرات"), (re.compile(r"\b03_records\.json\b"), "السجلات المنظّفة"),
    (re.compile(r"\b06_narrative\.json\b"), "السرد التحليلي"), (re.compile(r"\b02_mapping\.json\b"), "تصنيف الملف"),
    (re.compile(r"\b01_workbook_map\.(?:json|md)\b"), "فحص بنية الملف"),
    (re.compile(r"\b04_metrics_config\.json\b"), "إعداد المؤشرات"),
    (re.compile(r"\b07_validation\.(?:json|md)\b"), "التحقق الآلي"),
    (re.compile(r"\bversion_comparison\.json\b"), "مقارنة الإصدارات"),
    (re.compile(r"(?:خريطة|فحص) المصنف\s*\(0?1\)"), "فحص بنية الملف"),
    (re.compile(r"التطبيع\s*\(0?3\)"), "تنظيف البيانات"),
    (re.compile(r"(ملاحظات? المحلل|السرد)\s*\(0?6\)"), r"\1"),
    (re.compile(r"(?:قيمة )?المحرك في 05"), DASH_VALUE_AR),
    (re.compile(r"بعد التطبيع"), "بعد توحيد الكتابة"),
]


def plain_ar(x):
    """A text without pipeline file names or step numbers (finance and board readers see plain Arabic)."""
    if not isinstance(x, str):
        return x
    for rx, to in PLAIN_AR:
        x = rx.sub(to, x)
    return x


def n_ar(n, one, two, few, many, zero=None):
    """Arabic counted noun: «لا …» for 0 (when given), 1 and 2 by their own forms, 3–10 plural, 11–99 singular
    accusative, a round hundred singular genitive («100 سجل»)."""
    if n == 0 and zero:
        return zero
    if n == 1:
        return one
    if n == 2:
        return two
    r = n % 100
    if 3 <= r <= 10:
        return f"{n} {few}"
    if r >= 11:
        return f"{n} {many}"
    return f"{n} {many.rstrip('اًٌ').rstrip('ً')}" if n >= 100 else f"{n} {many}"


SHARED_REASON_MIN_LEN, SHARED_REASON_MIN_ROWS, SHORT_REASON_LEN = 80, 3, 70


def _short_reason(text: str) -> str:
    """The opening of a long reason for a table row: parentheticals dropped, cut at a word boundary."""
    t = re.sub(r"\s*\([^()]*\)", "", str(text or ""))
    t = re.sub(r"\s+", " ", t).strip()
    if len(t) <= SHORT_REASON_LEN:
        return t
    cut = t[:SHORT_REASON_LEN].rsplit(" ", 1)[0].rstrip("،؛,:— ")
    return cut + "…"


def n_records(n):
    return n_ar(n, "سجل واحد", "سجلان", "سجلات", "سجلًا", zero="لا سجلات")


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
    __slots__ = ("v", "fmt", "fill", "color", "bold", "wrap", "formula", "raw_ok", "code", "cache")

    def __init__(self, v, fmt=None, fill=None, color=None, bold=False, wrap=True, formula=False, raw_ok=False, code=False,
                 cache=None):
        self.v, self.fmt, self.fill, self.color, self.bold, self.wrap = v, fmt, fill, color, bold, wrap
        self.code = code
        self.formula = formula  # only deliberate formulas (totals, checks, links) are written unescaped
        self.cache = cache      # a formula's value computed here: stored in the file so a viewer that does not
        #                         recalculate (a preview) still shows the number; Excel and LibreOffice recalculate
        self.raw_ok = raw_ok    # a deed number: review-level digits are kept (block-level patterns are still masked)


def F(formula, cache, **kw):
    """A deliberate formula with the value it evaluates to (see C.cache)."""
    return C(formula, formula=True, cache=cache, **kw)


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
            c = self.ws.cell(row=2, column=1, value=safe_text(plain_ar(subtitle)))
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
            if not spec.code:
                v = plain_ar(v)
        elif v is not None and not is_num(v) and not isinstance(v, (_dt.date, _dt.datetime)):
            v = safe_text(v, review=not spec.raw_ok)
        cell = self.ws.cell(row=r, column=col, value=v)
        cell.border = BORDER
        if spec.formula:
            _formula_cache(self.wb)[(self.ws.title, cell.coordinate)] = spec.cache
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
        text = plain_ar(text)
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


def _formula_cache(wb) -> dict:
    """{(sheet, cell): value} of the deliberate formulas written into this workbook (stored after saving)."""
    if not hasattr(wb, "_apa_formula_cache"):
        wb._apa_formula_cache = {}
    return wb._apa_formula_cache


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
        self._rec_by_id = {r.get("id"): r for r in self.records if isinstance(r.get("id"), str)}
        self._notes_by_rec = self._index_notes()
        self.calc_links = {}   # metric id -> (cell in «الحسابات», cached value, number format) for «ملخص»

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
        """A narrative text as the board reads it: tokens rendered, keys by their labels, record codes by names."""
        return self.names_in_prose(self.hz(self._facts_rendered(render_tokens(x, self.metrics))))

    def prose(self, x):
        """A generated text (issue detail, impact): keys by their labels and record codes by names."""
        return self.names_in_prose(self.hz(x))

    def _facts_rendered(self, x):
        """{{f:key}} left in a text (06 read directly, without analysis.json): rendered from the facts of the run."""
        if not isinstance(x, str) or "{{" not in x:
            return x
        try:
            import facts as facts_mod
        except Exception:  # pragma: no cover
            return x
        if not self.facts:
            try:
                self.facts = facts_mod.compute_facts(self.run_dir)
            except Exception:
                self.facts = {}

        def rep(m):
            fx = (self.facts or {}).get(m.group(1)) or {}
            v, unit = fx.get("value"), fx.get("unit")
            if not is_num(v):
                return m.group(0)
            return fmt_value_text(v, {"pct": "%", "SAR": "SAR", "count": "count"}.get(unit, unit))
        return facts_mod.FACT_TOKEN_RE.sub(rep, x)

    def rec_name(self, rid):
        rec = self._rec_by_id.get(rid) if isinstance(rid, str) else None
        name = clean_text((rec or {}).get("name")) if rec else ""
        return name if name and name != rid else None

    def names_in_prose(self, x):
        """Record codes in a prose cell -> the record's name «…» (finding: codes stay in id, list and technical columns,
        never bare in a sentence). A code with no name is kept."""
        if not isinstance(x, str) or "-" not in x:
            return x

        def rep(m):
            nm = self.rec_name(m.group(0))
            return f"«{nm}»" if nm else m.group(0)
        return _REC_ID_RE.sub(rep, x)

    def target_text(self, x):
        """«الأصل أو النشاط» of a recommendation: a bare record code reads as the record's name."""
        if isinstance(x, str) and self.rec_name(x.strip()):
            return self.rec_name(x.strip())
        return self.r(x)

    def field_fmt(self, f):
        """#,##0 for riyals (revenue, expense, valuation fields), as the dashboard shows them; None lets a fractional
        non-money measure keep its decimals."""
        base = f[:-len(ORIGINAL_SUFFIX)] if isinstance(f, str) and f.endswith(ORIGINAL_SUFFIX) else f
        m, mb = self.meta.get(f) or {}, self.meta.get(base) or {}
        if (m.get("unit") or mb.get("unit")) == "SAR" or (m.get("group") or mb.get("group")) in ("revenue", "expense",
                                                                                                  "valuation"):
            return FMT_INT
        if isinstance(base, str) and re.match(r"^(?:rev|val|exp|noi)\.", base):
            return FMT_INT
        return None

    def num_cell(self, v, f=None, **kw):
        fmt = self.field_fmt(f) if f else None
        return C(v, fmt=fmt or (FMT_INT if float(v).is_integer() else FMT_DEC), **kw)

    def data_notes(self, rid):
        """C3: the data notes of one record — 03 issues and 06 data_issues_extra (and 01 anomalies) whose record ids
        include it; the same count and list as the dashboard's asset chip and detail panel."""
        return self._notes_by_rec.get(rid, [])

    def _index_notes(self):
        by = defaultdict(list)

        def add(rid, nid):
            if isinstance(rid, str) and isinstance(nid, str) and nid not in by[rid]:
                by[rid].append(nid)
        for i in self.rec_doc.get("issues") or []:
            if isinstance(i, dict):
                for rid in i.get("record_ids") or []:
                    add(rid, i.get("id"))
        for i in self.narrative.get("data_issues_extra") or []:
            if isinstance(i, dict):
                for rid in i.get("record_ids") or []:
                    add(rid, i.get("id"))
        for a in self.wbmap.get("anomalies") or []:
            if isinstance(a, dict):
                for rid in a.get("record_ids") or []:
                    add(rid, a.get("id"))
        return by

    def level(self, key, x):
        """A level in agreement with its heading (dashboard rankLevel)."""
        return (LEVEL_F if key in RANKING_FEM else LEVEL_AR).get(x, x)

    def severity(self, x):
        """«الأهمية» of a gap or a data issue: feminine, as the dashboard («مرتفعة/متوسطة/منخفضة/معلوماتية»)."""
        return LEVEL_F.get(x, x or "—")

    def owner_roles(self):
        """Owner roles the narrative names (recommendations and plan items), then generic roles, without repeats."""
        seen = OrderedDict()
        for x in self.narrative.get("recommendations") or []:
            if isinstance(x, dict) and isinstance(x.get("owner_role"), str) and x["owner_role"].strip():
                seen[x["owner_role"].strip()] = True
        for key, _lab, _d in PLAN_PHASES:
            for it in (self.narrative.get("plan") or {}).get(key) or []:
                if isinstance(it, dict) and isinstance(it.get("owner_role"), str) and it["owner_role"].strip():
                    seen[it["owner_role"].strip()] = True
        for g in GENERIC_OWNERS:
            seen.setdefault(g, True)
        return list(seen)

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
        parts = [p.strip() for p in re.split(r"[،,]", s) if p.strip()]
        if parts and all(self.rec_name(p) for p in parts):
            return "، ".join(self.rec_name(p) for p in parts)   # a record (or records): by name, as the dashboard
        return self.names_in_prose(self.hz(s))

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
        self.calc_links = {}
        self.calculations(wb)   # first: «ملخص» links its KPI values to the calculation cells
        self.summary(wb)
        self.assets(wb)
        self.exclusions(wb)
        self.reconciliation(wb)
        self.issues(wb)
        self.gaps_recs(wb)
        self.plan(wb)
        self.followup(wb)
        self.versions(wb)
        self.assumptions_sources(wb)
        self.lists_sheet(wb)
        for k, title in enumerate(("ملخص", "الأصول", "الحسابات")):
            if title in wb.sheetnames:
                wb.move_sheet(title, offset=k - wb.sheetnames.index(title))
        wb.active = 0
        os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
        wb.save(out_path)
        self.formula_cache = dict(_formula_cache(wb))
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
        ack_txt = (f" · تنبيهات معلَّلة {n_ack}" + (": " + "؛ ".join(f"«{k}» ×{n}" if n > 1 else f"«{k}»" for k, n in ids.items()) if ids else "")
                   if n_ack else "")
        when = str(v.get("generated_at") or v.get("checked_at") or "")[:16].replace("T", " ")
        return (f"الأخطاء المانعة {v.get('blocking_count')} · التنبيهات {v.get('warning_count')}{ack_txt}"
                + (f" — نتيجة التحقق السابق للتصدير ({when})؛ تُعاد بعد التصدير لفحص هذا المرفق" if when else ""))

    def _study_date_meta(self) -> dict:
        """The dashboard's study-date meta (analysis.json), else the same rule applied here: a date the file states is
        the study date; the workbook's last-modified timestamp never is (it stays as context)."""
        meta = (self.analysis or {}).get("meta") if isinstance(self.analysis, dict) else None
        if isinstance(meta, dict) and "study_date_stated" in meta:
            return meta
        return study_date_meta(self.mapping, self.wbmap)

    def summary(self, wb):
        sh = Sheet(wb, "ملخص", "دراسة محفظة الأصول — المرفق الداعم", self._file_line())
        m = self.mapping
        val = self.config.get("valuation") or {}
        st = self.rec_doc.get("stats") or {}
        by_scope = st.get("by_scope") or {}
        v = self.validation or {}
        n_all = st.get("records", len(self.records))
        rows = [
            ("الجهة", (m.get("entity") or {}).get("name") or "غير مذكورة في الملف"),
            ("فترة البيانات", (m.get("data_period") or {}).get("label") or (self.met_doc.get("default_filters") or {}).get("period")),
            ("تاريخ إعداد الدراسة", study_date_text(self._study_date_meta())),
            ("تاريخ التحليل", m.get("analysis_date") or _dt.date.today().isoformat()),
            ("العملة", m.get("currency") or self.config.get("currency") or "SAR"),
            ("التقييم المستخدم في المؤشرات", " — ".join(x for x in (val.get("label"), val.get("basis"), val.get("date_text")) if x) or "—"),
            ("السجلات", (n_records(n_all) if is_num(n_all) else str(n_all)) +
             f": ضمن النطاق {by_scope.get('in_scope', 0)}، معزول {by_scope.get('isolated', 0)}، "
             f"مستبعد {by_scope.get('excluded', 0)}"),
            ("التحقق الآلي", self._validation_line(v)),
            ("دليل الألوان", "الرمادي = غير متاح (لا يُجمع صفرًا) · الأحمر = مشكلة مثبتة فقط · الأصفر = جزئي أو يحتاج تحقق · "
                             "الأخضر = مكتمل للنطاق"),
        ]
        sh.table("tblContext", ["البند", "القيمة"], [[C(k, bold=True), x] for k, x in rows], filters=False)
        sh.heading("الخلاصة التنفيذية")
        es = self.r(self.narrative.get("executive_summary")) if not self.narrative_missing else None
        sh.note(es or "لم يُكتب السرد التحليلي بعد.", ncols=9)
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
            note = self.hz(result_notes(res, self.field_label))
            link = self.calc_links.get(mid)
            if link and is_num(val_cell.v) and is_num(link[1]) and \
                    abs(link[1] - val_cell.v) < (0.05 if link[2] == FMT_PCT_UNIT else 0.005 if link[2] == FMT_DEC else 0.5):
                # the value is the recomputation of «الحسابات» (a live link); its check cell compares it with the dashboard
                coord, cached, fmt = link
                val_cell = F(f"='الحسابات'!{coord}", cached, fmt=fmt or val_cell.fmt)
                note = "؛ ".join(x for x in (note, f"القيمة مرتبطة بإعادة الجمع في «الحسابات»!{coord}") if x)
            review = metric_review(res, self.spec_by_id.get(mid))
            st_cell = (review_cell(review) if review else
                       growth_status_cell(res) if growth_riyal_only(res) and is_num(res.get("value"))
                       else status_cell(res.get("status"), self.vocab, res))
            rows.append([C(res.get("label") or mid, bold=bool(res.get("card"))), val_cell, unit_lab, st_cell,
                         f"{res.get('n_included', 0)} من {res.get('n_scope', 0)}",
                         res.get("scope_text"), res.get("basis"), res.get("valuation_date"),
                         note or ("بطاقة في الداشبورد" if res.get("card") else "")])
        sh.table("tblKpis", ["المؤشر", "القيمة", "الوحدة", "الحالة", "المشمول من النطاق", "النطاق", "الأساس",
                             "تاريخ التقييم", "السبب أو الملاحظة"], rows)
        gaps = {g.get("id"): g for g in self.narrative.get("gaps") or [] if isinstance(g, dict)}
        recs = {x.get("id"): x for x in self.narrative.get("recommendations") or [] if isinstance(x, dict)}
        # two separate lists (brief §11.8): the n-th gap and the n-th action are not a pair; an action names the gaps it
        # addresses (gap_ids) explicitly
        tg = [gaps[g] for g in list(self.narrative.get("top_gaps") or []) if g in gaps][:3]
        ta = [recs[a] for a in list(self.narrative.get("top_actions") or []) if a in recs][:3]
        # C4: an action carries its recommendation's «الأولوية n», the same number as every other list, in that order
        ta = [a for _k, a in sorted(enumerate(ta), key=lambda ka: (
            ka[1].get("priority") if is_num(ka[1].get("priority")) else 999, ka[0]))]
        sh.heading("أهم ثلاث فجوات")
        sh.table("tblTopGaps", ["الفجوة", "الإثبات", "#", "المعرّف"],
                 [[self.r(g.get("title")), "مشكلة مثبتة" if g.get("is_proven_problem") is True else "محتملة — تحتاج تحققًا",
                   i, g.get("id")] for i, g in enumerate(tg, 1)], filters=False)
        sh.heading("أهم ثلاثة إجراءات أو قرارات مطلوبة")
        sh.table("tblTopActions", ["الإجراء أو القرار المطلوب", "المسؤول المقترح", "الأولوية", "يعالج الفجوة"],
                 [[self.r(a.get("first_action") or a.get("title")), a.get("owner_role") or "—",
                   a.get("priority") if is_num(a.get("priority")) else "—",
                   "، ".join(str(x) for x in a.get("gap_ids") or []) or "—"] for a in ta], filters=False)
        sh.ws.freeze_panes = "A4"
        sh.finish(fixed={1: 40, 2: 22, 3: 9, 4: 22, 5: 11, 6: 34, 7: 30, 8: 16, 9: 30}, print_titles=False, fit_width=1)

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

    def shared_reasons(self, texts):
        """Long reasons repeated on many rows: {full text: (k, n, short text)}. Each row keeps a short, filterable reason
        and the full text is written once above the table."""
        cnt = OrderedDict()
        for t in texts:
            if isinstance(t, str) and len(t.strip()) >= SHARED_REASON_MIN_LEN:
                cnt[t] = cnt.get(t, 0) + 1
        out = OrderedDict()
        for t, n in cnt.items():
            if n >= SHARED_REASON_MIN_ROWS:
                k = len(out) + 1
                out[t] = (k, n, f"السبب المشترك ({k}): {_short_reason(self.hz(t))} — نصه كاملًا أعلى الجدول")
        return out

    def shared_note(self, shared) -> str:
        return ("أسباب مشتركة بين صفوف كثيرة (نصها الكامل هنا مرة واحدة، وفي كل صف اختصارها؛ التصفية بعمود السبب تجمعها):\n"
                + "\n".join(f"({k}) {self.hz(t)} — {n_records(n)}." for t, (k, n, _s) in shared.items()))

    def assets(self, wb):
        sh = Sheet(wb, "الأصول", "سجل الأصول والسجلات كما طُبِّعت (كل سجل مرة واحدة)",
                   "الخلايا الرمادية قيم غير متاحة في الملف ولا تُعامل صفرًا؛ الصفر المسجل يظهر 0. عمود «المرجع» يحدد الورقة "
                   "والصف في الملف الأصلي. «الموقع» كما ورد في الملف، و«الحي» مستقل عنه.")
        numeric, text = self._field_lists()
        dims = [("asset_type", "نوع الأصل"), ("location", "الموقع"), ("district", "الحي"), ("use", "الاستخدام"),
                ("income_nature", "طبيعة الدخل"), ("ownership", "الملكية والحقوق"), ("waqf", "صفة الوقف"),
                ("managed_for_others", "مدار للغير"), ("rent_allocation", "تخصيص الريع"), ("restriction", "قيود الريع")]
        headers = ["المعرّف", "الاسم", "نوع السجل", "النطاق", "سبب النطاق"] + [lab for _, lab in dims]
        headers += [self.header_label(f) for f in numeric]
        headers += [self.header_label(f) for f in text]
        headers += ["المرجع", "مراجع القيم الرقمية", "عدد ملاحظات البيانات", "ملاحظات البيانات"]
        # the field keys sit in a «المرجع التقني» row above the Arabic headers, never in the headers themselves
        tech = [None] * (5 + len(dims)) + list(numeric) + list(text) + [None] * 4
        shared = self.shared_reasons(rec.get("scope_reason") for rec in self.records)
        if shared:
            sh.note(self.shared_note(shared), ncols=8)
        used = self._metric_fields()
        ordered = [f for f in numeric if f in used] + [f for f in numeric if f not in used]
        rows = []
        for rec in self.records:
            d = rec.get("dims") or {}
            fields = rec.get("fields") or {}
            scope = rec.get("scope")
            reason = rec.get("scope_reason")
            row = [rec.get("id"), rec.get("name"), ar(self.vocab, "record_type", rec.get("record_type")),
                   C(ar(self.vocab, "scope", scope), fill=GRAY_FILL if scope != "in_scope" else None,
                     color=GRAY_TXT if scope != "in_scope" else None),
                   shared[reason][2] if reason in shared else reason]
            district = d.get("district")
            for k, _ in dims:
                vv = d.get(k)
                vocab_key = {"use": "use", "income_nature": "income_nature", "ownership": "ownership", "waqf": "waqf"}.get(k)
                txt = ar(self.vocab, vocab_key, vv) if vocab_key else (ar(self.vocab, "", vv) if isinstance(vv, (bool, list)) else vv)
                if k == "location" and vv in (None, "", "unknown") and district not in (None, "", "unknown", []):
                    # the file names no location but a district: never «غير محدد» beside a known district
                    row.append(C(f"غير مذكور (الحي: {ar(self.vocab, '', district)})", color=GRAY_TXT))
                    continue
                row.append(C(txt if txt not in (None, "") else "غير محدد", fill=GRAY_FILL if vv in (None, "unknown", "") else None,
                             color=GRAY_TXT if vv in (None, "unknown", "") else None))
            for f in numeric:
                if f.endswith(ORIGINAL_SUFFIX) and isinstance(fields.get(f), dict):
                    # the stored value a documented derivation replaced: shown, flagged, never used (N7)
                    tv0 = fields[f]
                    row.append(self.num_cell(tv0.get("v"), f, fill=AMBER_FILL, color=AMBER_TXT) if is_num(tv0.get("v"))
                               else tv_cell(tv0, self.vocab))
                    continue
                row.append(tv_cell(fields.get(f), self.vocab, fmt=self.field_fmt(f)))
            for f in text:
                tv = fields.get(f)
                row.append(tv_cell(tv, self.vocab) if isinstance(tv, dict) and tv.get("s") not in KNOWN_STATUSES else
                           C((tv or {}).get("v"), raw_ok=is_deed_field(f, self.meta)))
            src = rec.get("source") or {}
            # every numeric source cell, the fields the metrics use first (valuation, period revenue): no cut-off (N8)
            refs = [f"{self.field_label(f)}: {(fields.get(f) or {}).get('ref')}" for f in ordered if (fields.get(f) or {}).get("ref")]
            notes = self.data_notes(rec.get("id"))
            row += [src.get("ref_name") or (f"{src.get('sheet')}!{src.get('row')}" if src.get("sheet") else None),
                    "\n".join(refs), C(len(notes), fmt=FMT_INT), "، ".join(notes)]
            rows.append(row)
        hr, first, _ = sh.table("tblAssets", headers, rows, tech=tech)
        sh.ws.freeze_panes = f"C{first}"
        sh.finish(fixed={2: 32, 5: 40})

    def calculations(self, wb):
        sh = Sheet(wb, "الحسابات", "سجل الحسابات — تعريف كل مؤشر ونطاقه ونتيجته",
                   f"«القيمة» هي {DASH_VALUE_AR}. «تفصيل الحساب» أدناه يعيد جمع كل مؤشر من سجلاته المشمولة بصيغ حية، "
                   f"ويقارن الناتج بالقيمة المعروضة في الداشبورد في صف «التحقق» («{MATCH_AR}» أو «{DIFF_AR}»).")
        rows = []
        for mid, res in self.metrics.items():
            res = res or {}
            spec = self.spec_by_id.get(mid) or {}
            rf = res.get("resolved_fields") or {}
            excl = self.excluded_of(res)
            val_cell, unit_lab = metric_value_cell(res, empty_text(res) if res.get("status") == "empty" else "غير متاح")
            share = res.get("coverage_num_share")
            review = metric_review(res, spec)
            rows.append([
                C(mid, code=True), res.get("label"), KIND_AR.get(res.get("kind"), res.get("kind")), val_cell, unit_lab,
                review_cell(review) if review else
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
        sh.finish(fixed={1: 30, 2: 36, 14: 50, 21: 30, 22: 40}, print_titles=False)

    @staticmethod
    def _check(sh, row, col, a_row, b_row, ndigits, a, b):
        """«مطابق»/«فرق»: the recomputation (row a_row) against the dashboard value (row b_row) in one column, rounded to
        ndigits as displayed (riyals to 0, percentages to 1 decimal)."""
        c = get_column_letter(col)
        ok = is_num(a) and is_num(b) and abs(a - b) < 0.5 * 10 ** (-ndigits)
        sh._put(row, col, F(f'=IF(ISNUMBER({c}{a_row}),IF(ROUND({c}{a_row}-{c}{b_row},{ndigits})=0,"{MATCH_AR}","{DIFF_AR}"),"—")',
                            (MATCH_AR if ok else DIFF_AR) if is_num(a) else "—", bold=True))
        # live colour: green «مطابق», amber «فرق» — follows the formula if a reviewer edits a value above
        for text, (fill, color) in ((MATCH_AR, (GREEN_FILL, GREEN_TXT)), (DIFF_AR, (AMBER_FILL, AMBER_TXT))):
            sh.ws.conditional_formatting.add(f"{c}{row}", CellIsRule(
                operator="equal", formula=[f'"{text}"'], fill=PatternFill("solid", fgColor=fill), font=Font(color=color, bold=True)))

    def _result_rows(self, sh, row, label, formula, cached, engine, fmt, ndigits):
        """Three rows under a block: a result formula, the dashboard value, and their check. Returns (next row, cell)."""
        sh._put(row, 1, C(label, bold=True))
        sh._put(row, 3, F(formula, cached, fmt=fmt, bold=True))
        sh._put(row + 1, 1, C(DASH_VALUE_AR, color=GRAY_TXT))
        sh._put(row + 1, 3, C(engine, fmt=fmt) if is_num(engine) else C("—", color=GRAY_TXT))
        sh._put(row + 2, 1, C("التحقق", bold=True))
        if is_num(engine):
            self._check(sh, row + 2, 3, row, row + 1, ndigits, cached, engine)
        else:
            sh._put(row + 2, 3, C("—", color=GRAY_TXT))
        return row + 3, f"C{row}"

    def _calc_detail(self, sh):
        """Each block recomputes a metric from its included records with live formulas: SUM per field, the dashboard
        value per field as a number, a check per field, then the result (ratio, difference, change) with its own check.
        A block with no included record writes a literal 0 (an empty range would end at the total cell itself).
        «ملخص» links its KPI value to the block's result cell (self.calc_links)."""
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
            ids = [i for i in res.get("included_ids") or [] if isinstance(i, str)]
            n_inc = n_ar(len(ids), "سجل واحد مشمول", "سجلان مشمولان", "سجلات مشمولة", "سجلًا مشمولًا", zero="لا سجلات مشمولة")
            sh.note(f"{res.get('label') or mid} — {n_inc}", ncols=6)
            sh.tech_row([mid, None] + list(fields) + [None], label_col=2)
            hr = sh.row
            heads = ["المعرّف", "الاسم"] + [self.header_label(f) for f in fields] + ["المرجع"]
            for j, h in enumerate(_unique_headers(heads), 1):
                sh._put(hr, j, h, header=True)
            r = hr
            sums = [0.0] * len(fields)
            for rid in ids:
                r += 1
                rec = recs.get(rid) or {}
                fl = rec.get("fields") or {}
                vals = [rid, rec.get("name")] + [tv_cell(fl.get(f), self.vocab, fmt=self.field_fmt(f)) for f in fields] + \
                       [" | ".join(str((fl.get(f) or {}).get("ref") or "") for f in fields)]
                for k, f in enumerate(fields):
                    tv = fl.get(f) or {}
                    if tv.get("s") in KNOWN_STATUSES and is_num(tv.get("v")):
                        sums[k] += tv["v"]
                for j, v in enumerate(vals, 1):
                    sh._put(r, j, v)
            if not ids:
                r += 1
                sh._put(r, 1, C("لا سجلات مشمولة في هذا المؤشر؛ المجموع صفر", color=GRAY_TXT))
            sums = [round(x, 6) for x in sums]
            tot = r + 1
            sh._put(tot, 1, C("المجموع (SUM)", bold=True))
            sh._put(tot, 2, C("إعادة الجمع من السجلات أعلاه", color=GRAY_TXT))
            for k, f in enumerate(fields):
                col = get_column_letter(3 + k)
                fmt = self.field_fmt(f) or FMT_INT
                sh._put(tot, 3 + k, F(f"=SUM({col}{hr + 1}:{col}{r})", sums[k], fmt=fmt, bold=True) if ids
                        else C(0, fmt=fmt, bold=True))
            engine = {"sum": [res.get("value")], "ratio": [res.get("num"), res.get("den")],
                      "diff": [res.get("a_sum"), res.get("b_sum")], "growth": [res.get("prev_sum"), res.get("curr_sum")]}[kind]
            sh._put(tot + 1, 1, C(DASH_VALUE_AR, color=GRAY_TXT))
            sh._put(tot + 2, 1, C("التحقق", bold=True))
            for k, f in enumerate(fields):
                x = engine[k]
                sh._put(tot + 1, 3 + k, C(x, fmt=self.field_fmt(f) or FMT_INT) if is_num(x) else C("—", color=GRAY_TXT))
                if is_num(x):
                    self._check(sh, tot + 2, 3 + k, tot, tot + 1, 0, sums[k], x)
                else:
                    sh._put(tot + 2, 3 + k, C("—", color=GRAY_TXT))
            nxt, link, link_cache, link_fmt = tot + 3, f"C{tot}", sums[0], FMT_INT
            unit = res.get("unit")
            if kind == "ratio":
                scale = res.get("scale") if is_num(res.get("scale")) else (100 if unit == "%" else 1)
                sc = f"{scale:g}"
                fmt = FMT_PCT_UNIT if unit == "%" else FMT_DEC
                cached = (sums[0] / sums[1] * scale) if sums[1] else ""
                nxt, link = self._result_rows(sh, nxt, f"النتيجة (البسط ÷ المقام{' × ' + sc if scale != 1 else ''})",
                                              f'=IF(D{tot}=0,"",C{tot}/D{tot}*{sc})', cached, res.get("value"), fmt,
                                              1 if unit == "%" else 2)
                link_cache, link_fmt = cached, fmt
            elif kind == "diff":
                cached = round(sums[0] - sums[1], 6)
                nxt, link = self._result_rows(sh, nxt, "الفرق (العمود الأول − الثاني)", f"=C{tot}-D{tot}", cached,
                                              res.get("value"), FMT_INT, 0)
                link_cache = cached
            elif kind == "growth":
                cached = round(sums[1] - sums[0], 6)
                nxt, link = self._result_rows(sh, nxt, "التغير بالريال (الفترة الحالية − السابقة)", f"=D{tot}-C{tot}",
                                              cached, res.get("abs_change"), FMT_INT, 0)
                link_cache = cached
                if not growth_riyal_only(res) and is_num(res.get("value")):
                    # a confirmed comparability: the growth % is shown (and checked); otherwise riyals only
                    pct = (sums[1] - sums[0]) / sums[0] * 100 if sums[0] > 0 else ""
                    nxt, link = self._result_rows(sh, nxt, "نسبة التغير (قابلية المقارنة مؤكدة)",
                                                  f'=IF(C{tot}>0,(D{tot}-C{tot})/C{tot}*100,"")', pct, res.get("value"),
                                                  FMT_PCT_UNIT, 1)
                    link_cache, link_fmt = pct, FMT_PCT_UNIT
            if is_num(link_cache):
                self.calc_links[mid] = (link, link_cache, link_fmt)
            sh.row = nxt + 1

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
        iso = [rec for rec in self.records if rec.get("scope") in ("isolated", "excluded")]
        shared = self.shared_reasons(rec.get("scope_reason") for rec in iso)
        if shared:
            sh.note(self.shared_note(shared), ncols=5)
        for rec in iso:
            reason = rec.get("scope_reason")
            rows.append([ar(self.vocab, "scope", rec.get("scope")), f"{rec.get('id')} — {rec.get('name') or ''}",
                         (rec.get("source") or {}).get("ref_name"),
                         shared[reason][2] if reason in shared else (self.hz(reason) or "—"), "جميع المؤشرات"])
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
                nm = self.rec_name(e.get("id"))
                rows.append(["مستبعد من مؤشر", f"{e.get('id')} — {nm}" if nm else e.get("id"), None, reason,
                             res.get("label") or mid])
        sh.table("tblExclusions", ["النوع", "البند", "الموضع", "السبب", "المؤشر المتأثر"], rows, freeze=True)
        sh.finish(fixed={2: 40, 4: 50, 5: 40})

    def reconciliation_rows(self):
        """Recorded totals (03 excluded_rows[].values) vs the sum of the records of the same source and field.

        Same rules as validate_analysis.py: the full in-scope sum first, then the section above the total row (a
        subtotal), then every record of the source (isolated/excluded included), then the sum without numbers that the
        file stores as text (Excel's SUM skips them), then the stored values a documented derivation replaced. Joined
        fields are not part of the source's own rows.

        Each item: «difference» = recorded − in-scope sum (None when the source has no in-scope value in the field);
        «cmp_sum»/«cmp_basis»/«ids» = the sum the total matched (or, on a mismatch, the in-scope sum — every record when
        none is in scope) and «cmp_difference» = recorded − cmp_sum. The sheet writes both differences as live formulas."""
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
                status, basis, mres, matched = "mismatch", None, None, None
                for st, res in (("match", full), ("match_section", sec), ("match_all_records", all_rec),
                                ("match_section_all_records", all_sec)):
                    if res[1] and abs(recorded - res[0]) <= tol:
                        status, basis, mres, matched = st, st, res, res[0]
                        break
                if status == "mismatch" and full[3] and abs(recorded - (full[0] - full[3])) <= tol:
                    status = basis = "match_without_text_numbers"
                    mres, matched = full, round(full[0] - full[3], 2)
                if status == "mismatch" and any(field + "__original" in (r.get("fields") or {}) for r in by_src.get(sid, [])):
                    for b, res in (("match_original", agg(field, True, False, True)),
                                   ("match_original_section", agg(field, True, True, True))):
                        if res[1] and abs(recorded - res[0]) <= tol:
                            status, basis, mres, matched = "match_original", b, res, res[0]
                            break
                if status == "mismatch":
                    # no match: compared with the in-scope sum, or every record when no in-scope value exists
                    mres = full if full[1] else all_rec
                    basis = "mismatch_in_scope" if full[1] else "mismatch_all_records"
                    matched = mres[0]
                no_scope = not full[1]
                use_all = no_scope or status in ("match_all_records", "match_section_all_records")
                base, base_sec = (all_rec, all_sec) if use_all else (full, sec)
                item.update(recorded=recorded, sum_in_scope=None if no_scope else full[0], n_known=full[1],
                            n_unknown=full[2], ids=mres[4],
                            difference=None if no_scope else round(recorded - full[0], 2),
                            cmp_sum=matched, cmp_basis=basis, cmp_difference=round(recorded - matched, 2),
                            sum_all_records=all_rec[0],
                            section_sum=base_sec[0] if base_sec[1] and (lo is not None or base_sec[1] != base[1]) else None,
                            matched_sum=matched if status != "mismatch" else None, status=status)
                out.append(item)
        return out

    def reconciliation(self, wb):
        items = self.reconciliation_rows()
        if not items:
            return
        sh = Sheet(wb, "المطابقة مع الإجماليات", "مطابقة الإجماليات المدونة في الملف مع مجموع السجلات المؤهلة",
                   "القيمة المدونة من صف الإجمالي المستبعد من السجلات؛ المجاميع من سجلات المصدر نفسه بقيم معروفة فقط (المجهول "
                   "لا يُجمع صفرًا).")
        sh.note("«الفرق عن مجموع النطاق» = القيمة المدونة − مجموع السجلات ضمن النطاق دائمًا (صيغة حية). «مجموع المقارنة» هو "
                "المجموع الذي طابقه الإجمالي وأساسه في العمود المجاور، و«الفرق عن مجموع المقارنة» صيغة حية كذلك؛ عند عدم "
                "المطابقة يُقارن بمجموع النطاق (أو بكل سجلات المصدر إن لم تكن فيه قيم ضمن النطاق) ويُذكر الفرق في مشكلات "
                "البيانات. «السجلات المجموعة» سجلات مجموع المقارنة.", ncols=8)
        hr = sh.row
        rows = []
        for i, it in enumerate(items, 1):
            r = hr + i
            st = it["status"]
            ok = st in ("match", "match_section", "match_all_records", "match_section_all_records")
            fill, color = ((GREEN_FILL, GREEN_TXT) if ok else
                           (GRAY_FILL, GRAY_TXT) if st == "recorded_unknown" else (AMBER_FILL, AMBER_TXT))
            lab = (self.meta.get(it["field"]) or {}).get("label") or it["field"]
            ids = it.get("ids") or []
            known = it.get("recorded") is not None
            f = it["field"]
            in_scope = known and it.get("sum_in_scope") is not None
            status_txt = (RECON_STATUS_NO_SCOPE_AR.get(st) if known and not in_scope else None) or RECON_STATUS_AR[st]
            more = len(ids) - 40
            rows.append([
                str(it.get("sheet") or "").strip() or it["source_id"], it["ref"], lab,
                self.num_cell(it["recorded"], f) if known else C(unknown_text(it["tv"], self.vocab), fill=GRAY_FILL, color=GRAY_TXT),
                (self.num_cell(it["sum_in_scope"], f) if in_scope else
                 C("لا قيم معروفة ضمن النطاق", fill=GRAY_FILL, color=GRAY_TXT) if known else None),
                it.get("n_known") if known else None, it.get("n_unknown") if known else None,
                F(f"=D{r}-E{r}", it["difference"], fmt=self.field_fmt(f) or FMT_INT) if in_scope else None,
                self.num_cell(it["cmp_sum"], f) if known else None,
                RECON_BASIS_AR.get(it.get("cmp_basis"), "—") if known else None,
                F(f"=D{r}-I{r}", it["cmp_difference"], fmt=self.field_fmt(f) or FMT_INT,
                  fill=AMBER_FILL if st == "mismatch" else None, color=AMBER_TXT if st == "mismatch" else None)
                if known else None,
                C(status_txt, fill=fill, color=color),
                "، ".join(ids[:40]) + (f"… و{n_ar(more, 'سجل آخر', 'سجلان آخران', 'سجلات أخرى', 'سجلًا آخر')}" if more > 0 else ""),
                self.num_cell(it["section_sum"], f) if known and it.get("section_sum") is not None else None,
                self.num_cell(it["sum_all_records"], f) if known else None,
                self.hz(it.get("reason")),
                C(f"{it['source_id']} · {it['field']}", code=True),
            ])
        sh.table("tblReconciliation", ["الورقة", "صف الإجمالي", "الحقل", "القيمة المدونة", "مجموع السجلات ضمن النطاق",
                                       "قيم معروفة ضمن النطاق", "قيم غير معروفة ضمن النطاق",
                                       "الفرق عن مجموع النطاق (المدون − مجموع النطاق)", "مجموع المقارنة",
                                       "أساس مجموع المقارنة", "الفرق عن مجموع المقارنة (المدون − مجموع المقارنة)", "النتيجة",
                                       "السجلات المجموعة", "مجموع سجلات القسم", "مجموع كل سجلات المصدر", "سبب استبعاد الصف",
                                       TECH],
                 rows, freeze=True)
        n_bad = sum(1 for it in items if it["status"] == "mismatch")
        sh.note(f"بنود المطابقة: {len(items)}؛ فروق غير مفسَّرة: {n_bad}. «قيم غير معروفة ضمن النطاق» سجلات ضمن النطاق قيمتها "
                "في الحقل غير متاحة، وقد تفسر جزءًا من الفرق.", ncols=8)
        sh.finish(fixed={3: 34, 10: 30, 12: 40, 13: 40, 16: 40})

    def issues(self, wb):
        sh = Sheet(wb, "مشكلات البيانات", "سجل مشكلات البيانات (منفصل عن أداء المحفظة)",
                   "المصدر: فحص بنية الملف، أو تنظيف السجلات، أو ملاحظة المحلل. الأهمية تخص أثر المشكلة على الحساب، وليست "
                   "حكمًا على الأصل.")
        rows = []

        def sev(s):
            fill, color = SEVERITY_STYLE.get(s, (None, None))
            return C(self.severity(s), fill=fill, color=color)

        def tech(x):
            keys = [k for k in (x.get("type"), x.get("field")) if isinstance(k, str) and k and not re.search(r"[؀-ۿ]", k)]
            return C(" · ".join(keys) or None, code=True)
        # the type is shown by its Arabic label; the type and field keys only in «المرجع التقني» (BP-10)
        for a in self.wbmap.get("anomalies") or []:
            if isinstance(a, dict):
                rows.append([a.get("id"), "فحص بنية الملف", issue_type_ar(a.get("type")), sev(a.get("severity")), a.get("where"),
                             "، ".join(a.get("record_ids") or []) or None, self.prose(a.get("detail")), None, None, tech(a)])
        for i in self.rec_doc.get("issues") or []:
            if isinstance(i, dict):
                rows.append([i.get("id"), "تنظيف السجلات", issue_type_ar(i.get("type")), sev(i.get("severity")), i.get("where"),
                             "، ".join(i.get("record_ids") or []), self.prose(i.get("detail")), self.prose(i.get("impact")),
                             self.prose(i.get("treatment")), tech(i)])
        for i in self.narrative.get("data_issues_extra") or []:
            if isinstance(i, dict):
                rows.append([i.get("id"), "ملاحظة المحلل", issue_type_ar(i.get("type")) if i.get("type") else "—",
                             sev(i.get("severity")), i.get("where"),
                             "، ".join(i.get("record_ids") or []), self.r(i.get("detail")), self.r(i.get("impact")),
                             self.r(i.get("treatment")), tech(i)])
        sh.table("tblIssues", ["المعرّف", "المصدر", "النوع", "الأهمية", "الموضع", "السجلات", "المشكلة", "الأثر", "المعالجة",
                               TECH], rows, freeze=True)
        sh.finish(fixed={7: 60, 8: 45, 9: 45})

    def gaps_recs(self, wb):
        sh = Sheet(wb, "الفجوات والتوصيات", "الفجوات المشخّصة والتوصيات المرتبطة بها",
                   "الأحمر للمشكلة المثبتة فقط؛ غير المثبت يحتاج تحققًا قبل القرار. كل توصية مرتبطة بفجوة ودليل. عمود «الحالة» في "
                   "جدول التوصيات قائمة منسدلة للمتابعة.")
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
                         self.severity(g.get("severity")),
                         C("مشكلة مثبتة", fill=RED_FILL, color=RED_TXT, bold=True) if proven else C("محتملة — تحتاج تحققًا"),
                         ev, self.r(imp.get("text")), f"{self.ev_text(mref)}: {mval}" if mref else None, self.r(g.get("proven_cause")) or "—",
                         [self.r(x) for x in g.get("possible_causes") or []], [self.r(x) for x in g.get("verification_needed") or []],
                         self.r(g.get("proposed_action"))])
        sh.heading("الفجوات")
        sh.table("tblGaps", ["المعرّف", "ضمن الأهم", "الفجوة", "النطاق", "الأهمية", "الإثبات", "الأدلة (مراجع)", "الأثر",
                             "مؤشر الأثر", "السبب المثبت", "الأسباب المحتملة", "ما يلزم للتحقق", "الإجراء المقترح"],
                 rows, freeze="A6")
        rows = []
        for x in sorted([x for x in n.get("recommendations") or [] if isinstance(x, dict)],
                        key=lambda x: (x.get("priority") if isinstance(x.get("priority"), int) else 999, str(x.get("id")))):
            rk = x.get("ranking") or {}
            # a new recommendation starts «لم يبدأ»; the dropdown lets the board track it
            status = x.get("status") if x.get("status") in PLAN_STATUSES else "لم يبدأ"
            rows.append([x.get("id"), x.get("priority"), "نعم" if x.get("id") in top_a else "", status,
                         "، ".join(x.get("gap_ids") or []),
                         self.target_text(x.get("target")), self.r(x.get("problem")), [self.ev_text(e) for e in list_of(x.get("evidence"))],
                         self.r(x.get("first_action")),
                         [self.r(a) for a in x.get("alternatives") or []], self.r(x.get("recommended_option")),
                         self.r(x.get("rationale")), x.get("owner_role"), x.get("duration"), self.r(x.get("cost")),
                         self.r(x.get("expected_impact")), self.r(x.get("success_kpi")), self.r(x.get("approval_required")) or "لا يلزم",
                         *[self.level(k, rk.get(k)) for k in ("impact", "urgency", "evidence_strength", "feasibility", "cost_level")]])
        sh.heading("التوصيات (مرتبة بالأولوية)")
        hr, first, last = sh.table("tblRecommendations", [
            "المعرّف", "الأولوية", "ضمن الأهم", "الحالة", "الفجوات", "الأصل أو النشاط", "المشكلة", "الدليل", "الإجراء الأول",
            "البدائل", "الخيار الموصى به", "سبب الاختيار", "المسؤول المقترح", "المدة", "التكلفة أو متطلبات تقديرها",
            "الأثر المتوقع", "مؤشر النجاح", "الاعتماد المطلوب", "الأثر", "الإلحاح", "قوة الدليل", "قابلية التنفيذ",
            "درجة التكلفة"], rows)
        if rows:
            self._status_dv(sh, f"D{first}:D{max(last, first) + 20}", "اختر حالة التوصية")
        sh.finish(fixed={3: 36, 6: 36, 7: 30, 8: 40, 13: 40}, print_titles=False)

    # ------------------------------------------------------------- dropdowns (data validation)
    def _status_dv(self, sh, ref, prompt):
        dv = DataValidation(type="list", formula1='"' + ",".join(PLAN_STATUSES) + '"', allow_blank=True,
                            showErrorMessage=True, errorTitle="قيمة غير مسموحة",
                            error="اختر: " + "، ".join(PLAN_STATUSES), promptTitle="الحالة", prompt=prompt)
        sh.ws.add_data_validation(dv)
        dv.add(ref)

    def _date_dv(self, sh, col, first, last):
        """A date on or after 2000-01-01 (cells formatted yyyy-mm-dd); blank allowed."""
        dv = DataValidation(type="date", operator="greaterThanOrEqual", formula1="36526", allow_blank=True,
                            showErrorMessage=True, errorTitle="تاريخ غير صحيح", error="أدخل تاريخًا بصيغة yyyy-mm-dd",
                            promptTitle="تاريخ المتابعة", prompt="تاريخ بصيغة yyyy-mm-dd")
        sh.ws.add_data_validation(dv)
        dv.add(f"{col}{first}:{col}{last}")
        for r in range(first, last + 1):
            sh.ws[f"{col}{r}"].number_format = FMT_DATE

    def _list_dv(self, sh, ref, list_ref, title, prompt, strict=True):
        """A dropdown fed by a column of «قوائم الاختيار» (no 255-character limit). strict=False: any other text is
        accepted (an owner the list does not name)."""
        dv = DataValidation(type="list", formula1=list_ref, allow_blank=True, showErrorMessage=strict,
                            errorTitle="قيمة غير مسموحة", error="اختر من القائمة", promptTitle=title, prompt=prompt)
        sh.ws.add_data_validation(dv)
        dv.add(ref)

    def _lists(self):
        """Columns of «قوائم الاختيار»: (header, values) — referenced by the plan's dropdowns."""
        return [("المرحلة", [lab for _k, lab, _d in PLAN_PHASES]), ("المسؤول", self.owner_roles()),
                ("الحالة", list(PLAN_STATUSES))]

    def _list_ref(self, header):
        for j, (h, vals) in enumerate(self._lists(), 1):
            if h == header:
                col = get_column_letter(j)
                return f"'{LISTS_SHEET}'!${col}$5:${col}${4 + max(1, len(vals))}"
        return None

    def lists_sheet(self, wb):
        # always written with «الخطة التشغيلية», whose dropdowns read it (rows start at 5: title, subtitle, header at 4)
        sh = Sheet(wb, LISTS_SHEET, "قوائم الاختيار في «الخطة التشغيلية»",
                   "تغذي هذه القوائم القوائم المنسدلة. عمود «المسؤول» يقبل أي مسمى آخر يُكتب يدويًا.")
        lists = self._lists()
        hr = sh.row
        for j, (h, vals) in enumerate(lists, 1):
            sh._put(hr, j, h, header=True)
            for i, v in enumerate(vals, 1):
                sh._put(hr + i, j, v)
        sh.row = hr + max(len(v) for _h, v in lists) + 2
        sh.finish(fixed={1: 18, 2: 44, 3: 16}, print_titles=False)

    def plan(self, wb):
        sh = Sheet(wb, "الخطة التشغيلية", "خطة التنفيذ: 30 و60 و90 يومًا",
                   "للمشروعات الطويلة يُذكر المخرج المرحلي خلال المدة (دراسة، عروض، تقييم) دون الإيحاء بانتهاء المشروع. أعمدة "
                   "«المرحلة» و«المسؤول» و«الحالة» قوائم منسدلة (يقبل «المسؤول» أي مسمى آخر)، و«تاريخ المتابعة» تاريخ.")
        recs = {x.get("id"): x for x in self.narrative.get("recommendations") or [] if isinstance(x, dict)}
        plan = self.narrative.get("plan") or {}
        rows = []
        for key, label, days in PLAN_PHASES:
            for it in plan.get(key) or []:
                if not isinstance(it, dict):
                    continue
                rec = recs.get(it.get("rec_id")) or {}
                status = it.get("status") if it.get("status") in PLAN_STATUSES else "لم يبدأ"
                duration = it.get("duration") or (f"مدة التوصية: {rec['duration']}" if rec.get("duration") else f"{days} يومًا")
                rows.append([label, it.get("rec_id"), rec.get("priority"), self.target_text(rec.get("target")),
                             self.r(it.get("action")), self.r(it.get("deliverable")), self.r(duration),
                             self.r(it.get("success_kpi") or rec.get("success_kpi")) or "—", it.get("owner_role"),
                             self.r(rec.get("approval_required")) or "لا يلزم", status, None, None])
        hr, first, last = sh.table("tblPlan", ["المرحلة", "التوصية", "الأولوية", "الأصل أو النشاط", "الإجراء", "المخرج",
                                               "المدة", "مؤشر النجاح", "المسؤول", "الاعتماد المطلوب", "الحالة",
                                               "تاريخ المتابعة", "ملاحظات"], rows, freeze=True)
        end = max(last, first) + 50
        self._list_dv(sh, f"A{first}:A{end}", self._list_ref("المرحلة"), "المرحلة", "اختر المرحلة")
        self._list_dv(sh, f"I{first}:I{end}", self._list_ref("المسؤول"), "المسؤول",
                      "اختر من القائمة أو اكتب مسمى آخر", strict=False)
        self._status_dv(sh, f"K{first}:K{end}", "اختر حالة البند")
        self._date_dv(sh, "L", first, end)
        sh.finish(fixed={4: 30, 5: 45, 6: 35, 7: 18, 8: 40, 9: 26, 10: 28, 12: 14, 13: 30})

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
        self._status_dv(sh, f"D{first}:D{max(last, first) + 50}", "اختر حالة التوصية")
        self._date_dv(sh, "H", first, max(last, first) + 50)
        sh.finish(fixed={2: 34, 3: 40, 5: 36, 6: 44, 7: 44, 8: 14})

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
                   + (" التفاصيل من ملف مقارنة الإصدارات في مجلد التشغيل." if from_file else ""))
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
                             self.num_cell(c["delta"], fld) if is_num(c.get("delta")) else None,
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
                diff = C(round(x["difference"], 2), fmt='0.0 "نقطة مئوية"') if pp else C(x["difference"], fmt=unit_fmt(unit))
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


# a cell or range reference in a formula, with an optional sheet prefix ('quoted name'! or name!)
_FORMULA_REF_RE = re.compile(r"(?:(?:'((?:[^']|'')+)'|([A-Za-z0-9_.؀-ۿ]+))!)?(?<![A-Za-z0-9_])\$?([A-Z]{1,3})\$?([0-9]+)"
                             r"(?::\$?([A-Z]{1,3})\$?([0-9]+))?(?![A-Za-z0-9_(])")
_FORMULA_STR_RE = re.compile(r'"(?:[^"]|"")*"')


def formula_self_refs(path: str):
    """Formulas whose reference or range contains their own cell (a circular reference: Excel warns on opening and
    LibreOffice shows Err:522). Returns [(sheet, cell, formula)]."""
    import openpyxl
    from openpyxl.utils import column_index_from_string as col_ix
    wb = openpyxl.load_workbook(path, read_only=True)
    hits = []
    try:
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                for c in row:
                    v = getattr(c, "value", None)
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue
                    f = _FORMULA_STR_RE.sub('""', v)
                    for m in _FORMULA_REF_RE.finditer(f):
                        sheet = (m.group(1) or "").replace("''", "'") or m.group(2)
                        if sheet and sheet != ws.title:
                            continue
                        c1, r1 = col_ix(m.group(3)), int(m.group(4))
                        c2, r2 = (col_ix(m.group(5)), int(m.group(6))) if m.group(5) else (c1, r1)
                        if min(c1, c2) <= c.column <= max(c1, c2) and min(r1, r2) <= c.row <= max(r1, r2):
                            hits.append((ws.title, c.coordinate, v))
                            break
    finally:
        wb.close()
    return hits


def store_formula_values(path: str, cache: dict) -> int:
    """Write the value each deliberate formula evaluates to (computed by the export) into the saved file, so a viewer
    that does not recalculate (a file preview) shows the numbers; Excel and LibreOffice still recalculate on opening
    (fullCalcOnLoad). Returns the number of cells written."""
    import html
    import zipfile
    by_sheet = defaultdict(dict)
    for (sheet, coord), val in (cache or {}).items():
        if val is not None:
            by_sheet[sheet][coord] = val
    if not by_sheet:
        return 0
    with zipfile.ZipFile(path) as z:
        infos = z.infolist()
        data = {i.filename: z.read(i.filename) for i in infos}
    wbx = data["xl/workbook.xml"].decode("utf-8")
    rels = data["xl/_rels/workbook.xml.rels"].decode("utf-8")
    def attr(a, key):
        mm = re.search(r'(?:^|\s)' + key + r'="([^"]*)"', a)
        return html.unescape(mm.group(1)) if mm else None
    targets = {attr(m.group(1), "Id"): attr(m.group(1), "Target") for m in re.finditer(r"<Relationship\b([^>]*)>", rels)}
    n = 0
    for m in re.finditer(r"<sheet\b([^>]*)>", wbx):
        name, rid = attr(m.group(1), "name"), attr(m.group(1), "r:id")
        vals = by_sheet.get(name)
        t = targets.get(rid)
        if not vals or not t:
            continue
        part = t.lstrip("/") if t.startswith("/") else "xl/" + t
        if part not in data:
            continue
        xml = data[part].decode("utf-8")

        def rep(mm):
            nonlocal n
            coord, attrs, f = mm.group(1), mm.group(2), mm.group(3)
            if coord not in vals:
                return mm.group(0)
            v = vals[coord]
            attrs = re.sub(r'\s+t="[^"]*"', "", attrs)
            n += 1
            if is_num(v):
                txt = str(int(v)) if float(v).is_integer() and abs(v) < 1e15 else repr(float(v))
                return f'<c r="{coord}"{attrs}><f>{f}</f><v>{txt}</v></c>'
            return f'<c r="{coord}"{attrs} t="str"><f>{f}</f><v>{html.escape(str(v), quote=False)}</v></c>'
        data[part] = re.sub(r'<c r="([A-Z]+[0-9]+)"([^>]*)><f>([^<]*)</f><v\s*(?:/>|>\s*</v>)</c>', rep, xml).encode("utf-8")
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w") as zo:
        for i in infos:
            zo.writestr(i, data[i.filename])
    os.replace(tmp, path)
    return n


def _kinds_ar(hits):
    return "، ".join(sorted({PII_KIND_AR.get(k, k) for _s, _c, ks in hits for k in ks}))


def _refs(hits, n=10):
    return "، ".join(f"«{s}»!{c}" for s, c, _k in hits[:n]) + ("…" if len(hits) > n else "")


def export(run_dir: str, out: str) -> int:
    """Build, save and privacy-check the workbook. 0 = written; 2 = personal data found after saving (file deleted)."""
    PRIVACY.reset()
    exp = Export(run_dir)
    sheets = exp.build(out)
    circular = formula_self_refs(out)
    if circular:
        try:
            os.remove(out)
        except OSError:
            pass
        print(f"خطأ في الصيغ: {len(circular)} صيغة تشير إلى خليتها نفسها (مرجع دائري يظهر خطأً عند فتح الملف) في "
              + "، ".join(f"«{s}»!{c}" for s, c, _f in circular[:10]) + ("…" if len(circular) > 10 else "")
              + "؛ حُذف الملف ولم يُسلَّم. هذا خلل في سكربت التصدير لا في البيانات: لا تسلّم المرفق قبل إصلاحه.",
              file=sys.stderr)
        return 2
    store_formula_values(out, getattr(exp, "formula_cache", {}))
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
            "rev.2024.actual": {"label": "الإيراد 2024", "group": "revenue", "period": "2024"},
            "units": {"label": "عدد الوحدات", "group": "other"}, "notes": {"label": "ملاحظات", "group": "text"},
            "tenant_label": {"label": "اسم المستأجر", "group": "text"}, "supervisor_mobile": {"label": "جوال", "group": "text"}}
    long_reason = ("ورقة تجريبية مخفية بعنوان عام وأقسام متعددة، لكن بنودها بلا معرّف أو صك مشترك مع السجل الرئيسي، "
                   "فلم يتحقق انتماؤها للمحفظة ولا تدخل أي مؤشر")

    def rec(rid, row, rev, scope="in_scope", src="T", name=None, dims=None, reason=None, **f):
        fields = {"rev.2025.excl_vat": _st_tv(rev, ref=f"S!C{row}")}
        fields.update(f)
        return {"id": rid, "name": name or f"عمارة تجريبية {rid}", "record_type": "asset", "scope": scope,
                "scope_reason": reason, "dims": dims or {}, "issue_ids": [],
                "source": {"source_id": src, "sheet": "S", "row": row}, "fields": fields}
    meta["val.2024.total"] = {"label": "تقييم 2024", "group": "valuation", "unit": "SAR"}
    many = {f"area.p{i}": _st_tv(100 + i, ref=f"S!Z{i}") for i in range(14)}   # N8: more than 12 numeric refs
    records = [
        rec("T-1", 2, 100000, name="عمارة النخيل التجريبية", dims={"district": "حي تجريبي"},
            units=_st_tv(None, "text", "6 محلات + 4 مكاتب"), notes=_st_tv(contact),
            tenant_label=_st_tv("مستأجر وهمي"), supervisor_mobile=_st_tv("0509876543"), **many,
            **{"val.2024.total": _st_tv(900000, "derived", ref="S!X2"),
               "val.2024.total__original": _st_tv(1200000, ref="S!X2"),
               "rev.2024.actual": _st_tv(80000.4, ref="S!T2")}),
        rec("T-2", 3, 50000, **{"rev.2024.actual": _st_tv(62499.6, ref="S!T3")}),
        rec("T-3", 4, 25000, scope="isolated", reason="سبب قصير"), rec("U-1", 2, 10, src="U"),
        # N4: an isolated statement with two section totals that match (rows 4, 6) and a typed grand total (row 7)
        rec("V-1", 2, 40, scope="isolated", src="V", reason=long_reason),
        rec("V-2", 3, 60, scope="isolated", src="V", reason=long_reason),
        rec("V-3", 5, 30, scope="isolated", src="V", reason=long_reason),
    ]
    excluded = [{"source_id": "T", "sheet": "S", "row": 5, "reason": "صف إجمالي",
                 "values": {"rev.2025.excl_vat": _st_tv(150000, ref="S!C5"),
                            # matches the stored values before the derivation, not the in-scope sum (finding: H = D − E)
                            "val.2024.total": _st_tv(1200000, ref="S!X5")}},
                {"source_id": "U", "sheet": "S2", "row": 3, "reason": "صف إجمالي",
                 "values": {"rev.2025.excl_vat": _st_tv(12, ref="S2!C3")}},
                {"source_id": "V", "sheet": "S", "row": 4, "reason": "إجمالي القسم",
                 "values": {"rev.2025.excl_vat": _st_tv(100, ref="S!C4")}},
                {"source_id": "V", "sheet": "S", "row": 6, "reason": "إجمالي القسم",
                 "values": {"rev.2025.excl_vat": _st_tv(30, ref="S!C6")}},
                {"source_id": "V", "sheet": "S", "row": 7, "reason": "الإجمالي العام — خريطة المصنف (01) تعده إجماليًا",
                 "values": {"rev.2025.excl_vat": _st_tv(90, ref="S!C7")}}]
    issues = [{"id": "DQ-001", "type": "zero_as_missing", "severity": "high", "where": "S!X2", "record_ids": ["T-1"],
               "field": "val.2024.total", "detail": "قيمة تجريبية في السجل T-1", "impact": "أثر", "treatment": "معالجة"}]
    _st_write(d, "03_records.json", {"contract_version": "1.0", "source": {"file": "t.xlsx"}, "field_meta": meta,
                                     "records": records, "excluded_rows": excluded, "issues": issues,
                                     "stats": {"records": 7}})
    # N6/BP-01: a growth result whose comparability is unverified is exported in riyals, without any percentage
    metrics = {
        "revenue_growth": {
            "id": "revenue_growth", "kind": "growth", "label": "تغير الإيراد", "value": 37.5, "unit": "%",
            "status": "complete", "abs_change": 7500, "prev_sum": 142500, "curr_sum": 150000, "n_scope": 2,
            "n_included": 2, "included_ids": ["T-1", "T-2"], "excluded": [],
            "resolved_fields": {"prev": "rev.2024.actual", "curr": "rev.2025.excl_vat"},
            "comparability": {"status": "unverified", "reason": "أساس الضريبة غير موضح"}},
        "property_revenue": {
            "id": "property_revenue", "kind": "sum", "label": "الإيرادات العقارية", "value": 150000, "unit": "SAR",
            "status": "complete", "n_scope": 2, "n_included": 2, "included_ids": ["T-1", "T-2"], "excluded": [],
            "resolved_fields": {"field": "rev.2025.excl_vat"}, "card": True, "card_order": 1,
            # C1: a metric spec flagged for verification
            "review": {"status": REVIEW_STATUS, "reason": "رقم بديل للإيراد في بيان آخر لم يُحسم بعد"}},
        "gross_yield": {
            "id": "gross_yield", "kind": "ratio", "label": "نسبة الإيراد إلى التقييم", "value": 100000 / 900000 * 100,
            "unit": "%", "scale": 100, "status": "complete", "num": 100000, "den": 900000, "n_scope": 1, "n_included": 1,
            "included_ids": ["T-1"], "excluded": [], "resolved_fields": {"num": "rev.2025.excl_vat", "den": "val.2024.total"},
            "card": True, "card_order": 2},
        # finding 1: a sum with no included record (its check must not refer to its own cell)
        "zero_rev_value": {
            "id": "zero_rev_value", "kind": "sum", "label": "قيمة الأصول بإيراد صفري", "value": 0, "unit": "SAR",
            "status": "complete", "n_scope": 0, "n_included": 0, "included_ids": [], "excluded": [],
            "resolved_fields": {"field": "val.2024.total"}},
    }
    _st_write(d, "05_metrics.json", {"metrics": metrics})
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
    ranking = {"impact": "high", "urgency": "high", "evidence_strength": "strong", "feasibility": "high", "cost_level": "low"}
    _st_write(d, "06_narrative.json", {
        "executive_summary": "خلاصة تجريبية تذكر الأصل T-1 والقيم من 05_metrics.json.",
        "gaps": [{"id": "G1", "title": "فجوة تجريبية", "scope": "T-1", "severity": "high", "is_proven_problem": False,
                  "evidence": ["m:property_revenue"], "impact": {"text": "أثر تجريبي على T-1"}}],
        "recommendations": [
            {"id": "R1", "priority": 1, "gap_ids": ["G1"], "target": "T-1", "first_action": "تحقق من العقد",
             "owner_role": "مدير تجريبي", "duration": "60 يومًا", "success_kpi": "عقد موثق خلال 60 يومًا",
             "approval_required": "مجلس الإدارة", "ranking": ranking},
            {"id": "R2", "priority": 2, "gap_ids": ["G1"], "target": "المحفظة", "first_action": "إجراء ثان تجريبي",
             "owner_role": "مدير مالي تجريبي", "ranking": ranking}],
        "top_gaps": ["G1"], "top_actions": ["R2", "R1"],
        "plan": {"d30": [{"rec_id": "R1", "action": "طلب العقد", "deliverable": "نسخة العقد", "owner_role": "مدير تجريبي"}],
                 "d60": [{"rec_id": "R2", "action": "مراجعة", "deliverable": "مذكرة", "owner_role": "مدير مالي تجريبي"}],
                 "d90": [{"rec_id": "R1", "action": "اعتماد", "deliverable": "قرار", "owner_role": "مدير تجريبي"}]},
        "data_issues_extra": [{"id": "DQ-X1", "severity": "medium", "record_ids": ["T-1"], "detail": "ملاحظة محلل تجريبية"}],
        "recommendation_followup": [
            {"rec_id": "R1", "status": "لم يبدأ", "evidence": [], "note": "حالة افتراضية",
             "data_signals": [{"record_id": "T-1", "category": "data_correction", "field": "rev.2025.excl_vat", "delta": 10000},
                              {"record_id": "T-1", "category": "data_correction", "field": "rev.2025.excl_vat", "delta": 10000}]},
            {"rec_id": "R2", "status": "منجز", "evidence": None}],
        "version_comparison": vc})
    # board#17: a study date that is only the workbook's last-modified property is not the study date
    _st_write(d, "01_workbook_map.json", {"file": {"properties": {"modified": "2031-03-14T09:00:00"}}, "sheets": []})
    _st_write(d, "02_mapping.json", {"study_date": {"value": "2031-03-14", "basis": "آخر تعديل للملف (خصائص المصنف)",
                                                    "evidence": "file.properties.modified"}})


def self_test() -> int:
    import tempfile
    import openpyxl
    fails = []

    def check(name, ok, detail=""):
        print(("  PASS  " if ok else "  FAIL  ") + name + ("" if ok else f" — {detail}"))
        if not ok:
            fails.append(name)

    def rows_of(ws, lo=1):
        return [list(r) for r in ws.iter_rows(min_row=lo, values_only=True)]
    with tempfile.TemporaryDirectory() as d:
        _st_run(d)
        out = os.path.join(d, "supporting_workbook.xlsx")
        rc = export(d, out)
        check("export exits 0", rc == 0 and os.path.isfile(out), rc)
        if rc == 0:
            wb = openpyxl.load_workbook(out)
            wv = openpyxl.load_workbook(out, data_only=True)   # the values stored for formulas
            texts = [str(c.value) for ws in wb.worksheets for row in ws.iter_rows() for c in row if c.value is not None]
            blob = "\n".join(texts)
            block, _review = privacy_self_check(out)
            check("D1 no phone/landline/e-mail in any cell (separators, Arabic-Indic digits)", not block, block)
            check("D1 personal columns dropped (key and label)",
                  "مستأجر وهمي" not in blob and "0509876543" not in blob and "supervisor_mobile" not in blob, "")
            check("D11 original text of a text-in-numeric cell shown", "«6 محلات + 4 مكاتب»" in blob, "")
            for name in ("المطابقة مع الإجماليات", "متابعة التوصيات", "مقارنة الإصدارات", LISTS_SHEET):
                check(f"sheet «{name}» present", name in wb.sheetnames, wb.sheetnames)
            check("sheet order: «ملخص»، «الأصول»، «الحسابات» first", wb.sheetnames[:3] == ["ملخص", "الأصول", "الحسابات"],
                  wb.sheetnames)
            # finding 1: no formula refers to its own cell; every formula has a stored value
            check("no circular formula (a range containing its own cell)", formula_self_refs(out) == [],
                  formula_self_refs(out))
            no_cache = [f"{ws.title}!{c.coordinate}" for ws in wb.worksheets for row in ws.iter_rows() for c in row
                        if isinstance(c.value, str) and c.value.startswith("=")
                        and wv[ws.title][c.coordinate].value is None]
            check("every formula carries its computed value (a preview shows numbers)", not no_cache, no_cache[:6])
            ws, wsv = wb["الحسابات"], wv["الحسابات"]
            calc = rows_of(ws)
            calcv = rows_of(wsv)
            zi = next((i for i, r in enumerate(calc) if str(r[0] or "").startswith("قيمة الأصول بإيراد صفري —")), None)
            zt = next((i for i in range(zi or 0, len(calc)) if calc[i][0] == "المجموع (SUM)"), None) if zi is not None else None
            check("zero-included block: «لا سجلات مشمولة», a literal 0 total, its check «مطابق»",
                  zi is not None and "لا سجلات مشمولة" in calc[zi][0] and zt is not None and calc[zt][2] == 0
                  and calc[zt + 1][0] == DASH_VALUE_AR and calcv[zt + 2][2] == MATCH_AR,
                  (zi, zt, zt is not None and calc[zt][2]))
            check("block headings count records in Arabic («سجلان مشمولان»)",
                  any(str(r[0] or "") == "الإيرادات العقارية — سجلان مشمولان" for r in calc), "")
            check("engine values are numbers beside each SUM, with a visible check («مطابق»)",
                  sum(1 for r in calc if r[0] == DASH_VALUE_AR and any(is_num(x) for x in r[2:4])) >= 4
                  and sum(1 for r in calcv if r[0] == "التحقق" and MATCH_AR in r[2:4]) >= 6
                  and not any(DIFF_AR in [x for x in r[2:4]] for r in calcv if r[0] == "التحقق"), "")
            check("finding 5: no pipeline jargon (file names, «محرك», step numbers) outside technical cells",
                  not any(t for ws_ in wb.worksheets for row in ws_.iter_rows() for c in row
                          if isinstance(c.value, str) and (c.font is None or c.font.name != CODE_FONT)
                          for t in ("05_metrics", "06_narrative", "محرك", "(01)", "(03)", "(06)") if t in c.value), "")
            # «ملخص»: KPI values linked to the calculation cells, C1 review status, C4 priorities
            ws, wsv = wb["ملخص"], wv["ملخص"]
            srow = {r[0]: r for r in rows_of(ws)}
            svrow = {r[0]: r for r in rows_of(wsv)}
            pr = srow.get("الإيرادات العقارية")
            check("«ملخص» KPI value is a live link to «الحسابات», showing 150,000",
                  bool(pr) and str(pr[1]).startswith("='الحسابات'!C") and svrow["الإيرادات العقارية"][1] == 150000, pr)
            yv = svrow.get("نسبة الإيراد إلى التقييم")
            check("ratio linked too, percent format 0.0\"%\"",
                  bool(yv) and abs(yv[1] - 100000 / 900000 * 100) < 1e-9
                  and next(c for c in ws[[i for i, r in enumerate(rows_of(ws), 1) if r[0] == "نسبة الإيراد إلى التقييم"][0]]
                           if c.column == 2).number_format == FMT_PCT_UNIT, yv)
            check("C1: a reviewed metric shows «يحتاج تحقق — <reason>» in «ملخص» and «الحسابات»",
                  bool(pr) and str(pr[3]).startswith(REVIEW_AR + " — رقم بديل")
                  and any(r[0] == "property_revenue" and str(r[5]).startswith(REVIEW_AR + " — ") for r in calc), pr)
            row = srow.get("تغير الإيراد")
            check("N6 summary growth row: 7,500 riyals, «تغير بالريال فقط», no percentage anywhere",
                  bool(row) and svrow["تغير الإيراد"][1] == 7500 and row[2] == "ريال" and "لا نسبة" in str(row[3])
                  and "تغير بالريال فقط" in str(row[8]) and "37.5" not in blob and "37.50" not in blob, row)
            nums = [c.value for sh in ("ملخص", "الحسابات") for r in wv[sh].iter_rows() for c in r if is_num(c.value)]
            check("N6 the growth % (37.5) is no value cell in ملخص/الحسابات", not any(abs(x - 37.5) < 1e-9 for x in nums), "")
            acts = [r for r in rows_of(ws) if r[0] in ("تحقق من العقد", "إجراء ثان تجريبي")]
            check("C4: top actions in priority order with «الأولوية n»",
                  [(r[0], r[2]) for r in acts] == [("تحقق من العقد", 1), ("إجراء ثان تجريبي", 2)], acts)
            check("record codes in prose read as names (executive summary)",
                  "عمارة النخيل التجريبية" in str(next((r[0] for r in rows_of(ws) if str(r[0] or "").startswith("خلاصة")), ""))
                  and "نتائج المؤشرات" in blob, "")
            vline = srow.get("التحقق الآلي", [None, ""])[1]
            check("BP-11 validation row lists the acknowledged warnings by their Arabic titles and the validation time",
                  "تنبيهات معلَّلة 2" in str(vline) and "الأخطاء المانعة" in str(vline) and "«قابلية مقارنة الفترات»" in str(vline)
                  and "growth_comparability" not in str(vline) and "2026-01-02 10:20" in str(vline), vline)
            sd = str(srow.get("تاريخ إعداد الدراسة", [None, ""])[1])
            check("board#17: a last-modified date reads «غير مذكور في الملف — آخر تعديل للملف المصدر: …»",
                  sd == "غير مذكور في الملف — آخر تعديل للملف المصدر: 2031-03-14", sd)
            check("records counted in Arabic («7 سجلات»)", str(srow.get("السجلات", [None, ""])[1]).startswith("7 سجلات:"),
                  srow.get("السجلات"))
            # finding 2: reconciliation — H = D − E always (live), the matched sum and its basis beside it
            if "المطابقة مع الإجماليات" in wb.sheetnames:
                ws, wsv = wb["المطابقة مع الإجماليات"], wv["المطابقة مع الإجماليات"]
                rr = rows_of(ws)
                rv = rows_of(wsv)
                src = lambda r: str(r[16] or "").split(" · ")[0]  # noqa: E731 — the technical column «V · field»
                by = {(src(r), r[1]): (r, v, n) for n, (r, v) in enumerate(zip(rr, rv), 1)
                      if len(r) > 16 and src(r) and r[16] != TECH}
                t_rev, t_val = by.get(("T", "S!C5")), by.get(("T", "S!X5"))
                check("A14 T matches 150,000 in scope: H = 0, K = 0", bool(t_rev) and t_rev[1][7] == 0
                      and t_rev[1][10] == 0 and t_rev[1][11] == RECON_STATUS_AR["match"], t_rev and t_rev[1])
                check("finding 2: a total matched through original values — H = D − E (300,000) live, I = 1,200,000, "
                      "K = 0, L explains the difference",
                      bool(t_val) and t_val[0][7] == f"=D{t_val[2]}-E{t_val[2]}" and t_val[1][7] == 300000 and t_val[1][8] == 1200000
                      and t_val[1][10] == 0 and "الفرق عن مجموع النطاق سببه الاشتقاق" in str(t_val[1][11])
                      and t_val[1][9] == RECON_BASIS_AR["match_original"], t_val and t_val[1])
                u = by.get(("U", "S2!C3"))
                check("A14 U differs by 2 (unexplained)", bool(u) and u[1][7] == 2 and u[1][10] == 2
                      and u[1][11] == RECON_STATUS_AR["mismatch"], u and u[1])
                incoherent = [r[1] for _r, r, _n in by.values() if is_num(r[3]) and (
                    (is_num(r[4]) and abs((r[7] or 0) - (r[3] - r[4])) > 1e-6) or
                    (is_num(r[8]) and abs((r[10] or 0) - (r[3] - r[8])) > 1e-6) or
                    (r[11] != RECON_STATUS_AR["mismatch"] and is_num(r[10]) and abs(r[10]) > 0.5))]
                check("every row coherent: H = D − E, K = D − I, a matched row has K ≈ 0", not incoherent, incoherent)
                sec, sec2, grand = by.get(("V", "S!C4")), by.get(("V", "S!C6")), by.get(("V", "S!C7"))
                check("item 6: a section total lists its own section's lines (S!C4: V-1، V-2; S!C6: V-3), the grand total all",
                      bool(sec) and sec[1][12] == "V-1، V-2" and bool(sec2) and sec2[1][12] == "V-3"
                      and bool(grand) and grand[1][12] == "V-1، V-2، V-3",
                      (sec and sec[1][12], sec2 and sec2[1][12], grand and grand[1][12]))
                check("BP-10: the sheet name, the Arabic field label and the key only in «المرجع التقني»",
                      bool(sec) and sec[1][0] == "S" and sec[1][2] == "الإيراد 2025" and sec[1][16] == "V · rev.2025.excl_vat",
                      sec and sec[1])
                check("N4 isolated section totals: no in-scope sum (no «recorded − 0»), matched section sums 100 and 30",
                      bool(sec) and sec[1][7] is None and sec[1][8] == 100 and sec[1][10] == 0 and sec[1][13] == 100
                      and bool(sec2) and sec2[1][8] == 30 and sec2[1][10] == 0, (sec and sec[1], sec2 and sec2[1]))
                check("N4 isolated grand total: 90 − 130 = −40 against all records", bool(grand) and grand[1][10] == -40
                      and grand[1][8] == 130 and grand[1][13] is None, grand and grand[1])
                check("jargon in a reason replaced («فحص بنية الملف»)", bool(grand) and "فحص بنية الملف" in str(grand[1][15]),
                      grand and grand[1][15])
            # «الأصول»: C3 notes, location vs district, shared reasons, SAR without decimals
            ws = wb["الأصول"]
            ar_rows = rows_of(ws)
            hi = next(i for i, r in enumerate(ar_rows) if r[0] == "المعرّف")
            hdr, tech = ar_rows[hi], ar_rows[hi - 1]
            t1 = next((r for r in ar_rows if r[0] == "T-1"), None)
            check("N7 «__original» column labelled as the stored value before the derivation",
                  any("كما ورد قبل الاشتقاق" in str(h) for h in hdr), hdr)
            check("BP-10 headers are Arabic labels only; the keys sit in the «المرجع التقني» row above them",
                  tech[0] == TECH + ":" and "val.2024.total__original" in tech
                  and not any(re.search(r"[a-z]+\.[a-z0-9_]+", str(h or "")) for h in hdr), (tech, hdr))
            j = hdr.index("مراجع القيم الرقمية")
            check("N8 every numeric source cell listed (no 12-line cut), metric fields first, by their labels",
                  bool(t1) and str(t1[j]).count("\n") >= 15
                  and {x.split(": ")[0] for x in str(t1[j]).split("\n")[:3]} == {"الإيراد 2024", "الإيراد 2025", "تقييم 2024"},
                  t1 and t1[j])
            check("C3: notes count and list = 03 issues + 06 data_issues_extra («2», «DQ-001، DQ-X1»)",
                  bool(t1) and t1[hdr.index("عدد ملاحظات البيانات")] == 2
                  and t1[hdr.index("ملاحظات البيانات")] == "DQ-001، DQ-X1", t1 and t1[-2:])
            check("location not stated but district known: «غير مذكور (الحي: …)»",
                  bool(t1) and t1[hdr.index("الموقع")] == "غير مذكور (الحي: حي تجريبي)", t1 and t1[hdr.index("الموقع")])
            v1 = next((r for r in ar_rows if r[0] == "V-1"), None)
            check("a long reason repeated on many rows: a short reason per row and the full text once above the table",
                  bool(v1) and str(v1[4]).startswith("السبب المشترك (1): ") and len(str(v1[4])) < 140
                  and any("(1) ورقة تجريبية مخفية" in str(r[0] or "") for r in ar_rows[:hi]), v1 and v1[4])
            rcol = hdr.index("الإيراد 2024") + 1
            t1r = hi + 1 + next(i for i, r in enumerate(ar_rows[hi + 1:], 1) if r[0] == "T-1")
            check("SAR shown without decimals (#,##0) even when fractional",
                  ws.cell(row=t1r, column=rcol).number_format == FMT_INT and ws.cell(row=t1r, column=rcol).value == 80000.4,
                  ws.cell(row=t1r, column=rcol).number_format)
            # «مشكلات البيانات» and «الفجوات والتوصيات»: feminine levels, record names, status dropdown
            ws = wb["مشكلات البيانات"]
            irows = {r[0]: r for r in rows_of(ws)}
            check("severity «مرتفعة/متوسطة» under «الأهمية»; issue source in plain Arabic",
                  irows.get("المعرّف", [None] * 4)[3] == "الأهمية" and irows.get("DQ-001", [None] * 4)[3] == "مرتفعة"
                  and irows.get("DQ-X1", [None] * 4)[3] == "متوسطة" and irows["DQ-001"][1] == "تنظيف السجلات", irows.get("DQ-001"))
            ws = wb["الفجوات والتوصيات"]
            grows = rows_of(ws)
            g1 = next((r for r in grows if r[0] == "G1"), None)
            r1 = next((r for r in grows if r[0] == "R1"), None)
            rh = next((r for r in grows if r[0] == "المعرّف" and "الأولوية" in r), None)
            check("gap scope and recommendation target by the record's name; gap severity «مرتفعة»",
                  bool(g1) and g1[3] == "عمارة النخيل التجريبية" and g1[4] == "مرتفعة" and "«عمارة النخيل التجريبية»" in str(g1[7])
                  and bool(r1) and r1[rh.index("الأصل أو النشاط")] == "عمارة النخيل التجريبية", (g1, r1))
            check("ranking levels agree with their headings (الأثر مرتفع · قوة الدليل قوية · قابلية التنفيذ مرتفعة · "
                  "درجة التكلفة منخفضة)",
                  bool(r1) and [r1[rh.index(h)] for h in ("الأثر", "قوة الدليل", "قابلية التنفيذ", "درجة التكلفة")]
                  == ["مرتفع", "قوية", "مرتفعة", "منخفضة"], r1)
            dvs = {str(dv.sqref): dv for dv in ws.data_validations.dataValidation}
            col = get_column_letter(rh.index("الحالة") + 1)
            check("recommendations table: «الحالة» dropdown (لم يبدأ/قيد التنفيذ/منجز/متعثر)",
                  any(k.startswith(col) and all(s in (dv.formula1 or "") for s in PLAN_STATUSES) for k, dv in dvs.items())
                  and r1[rh.index("الحالة")] == "لم يبدأ", list(dvs))
            # finding 4: «الخطة التشغيلية» columns and validations
            ws = wb["الخطة التشغيلية"]
            prow = rows_of(ws)
            ph = next(r for r in prow if r[0] == "المرحلة")
            p1 = next((r for r in prow if r[1] == "R1"), None)
            check("plan columns «المدة»، «مؤشر النجاح»، «الاعتماد المطلوب» filled from the recommendation",
                  bool(p1) and p1[ph.index("المدة")] == "مدة التوصية: 60 يومًا"
                  and p1[ph.index("مؤشر النجاح")] == "عقد موثق خلال 60 يومًا"
                  and p1[ph.index("الاعتماد المطلوب")] == "مجلس الإدارة"
                  and p1[ph.index("الأصل أو النشاط")] == "عمارة النخيل التجريبية", p1)
            kinds = {}
            for dv in ws.data_validations.dataValidation:
                for rng in str(dv.sqref).split():
                    kinds[rng[0]] = (dv.type, dv.formula1, dv.showErrorMessage)
            lc = lambda h: get_column_letter(ph.index(h) + 1)  # noqa: E731
            check("plan dropdowns: «المرحلة» and «المسؤول» from «قوائم الاختيار» (owner free text allowed), status list, "
                  "date on «تاريخ المتابعة»",
                  kinds.get(lc("المرحلة"), ("",))[0] == "list" and LISTS_SHEET in str(kinds[lc("المرحلة")][1])
                  and kinds.get(lc("المسؤول"), ("",))[0] == "list" and LISTS_SHEET in str(kinds[lc("المسؤول")][1])
                  and not kinds[lc("المسؤول")][2]
                  and kinds.get(lc("الحالة"), ("",))[0] == "list" and kinds.get(lc("تاريخ المتابعة"), ("",))[0] == "date",
                  kinds)
            owners = [r[1] for r in rows_of(wb[LISTS_SHEET], 5) if r[1]]
            check("owner list: the narrative's roles first, then generic roles",
                  owners[:2] == ["مدير تجريبي", "مدير مالي تجريبي"] and "مجلس الإدارة" in owners, owners)
            if "متابعة التوصيات" in wb.sheetnames:
                ws = wb["متابعة التوصيات"]
                dvs = [dv.formula1 for dv in ws.data_validations.dataValidation]
                check("§12 follow-up status dropdown uses «لم يبدأ/قيد التنفيذ/منجز/متعثر»",
                      any(all(s in (f or "") for s in PLAN_STATUSES) for f in dvs), dvs)
                cells = [str(c.value) for row in ws.iter_rows() for c in row if c.value is not None]
                check("§12 «منجز» without evidence is flagged", any("لا دليل مسجل" in c for c in cells), "")
                check("D7 follow-up data signals de-duplicated", sum(c.count("T-1:") for c in cells) == 1, cells)
            wb.close()
            wv.close()
        # finding 1: the post-save guard deletes a workbook with a circular formula and exits 2
        orig = Export.summary

        def bad_summary(self, wb):
            orig(self, wb)
            wb["ملخص"]["Z9"] = "=SUM(Z5:Z9)"
        Export.summary = bad_summary
        try:
            rc = export(d, out)
        finally:
            Export.summary = orig
        check("finding 1 guard: a formula whose range holds its own cell -> exit 2 and no file",
              rc == 2 and not os.path.exists(out), rc)
        probe = os.path.join(d, "probe.xlsx")
        pw = openpyxl.Workbook()
        pw.active.title = "ورقة"
        other = pw.create_sheet("أخرى")
        pw["ورقة"]["C5"] = "=SUM(C1:C4)"
        pw["ورقة"]["D5"] = '=IF(C5=0,"D5","")'
        pw["ورقة"]["E5"] = "='أخرى'!E5"
        pw["ورقة"]["F7"] = "=SUM($F$2:$F$9)"
        other["A1"] = "=A1"
        pw.save(probe)
        hits = sorted((s, c) for s, c, _f in formula_self_refs(probe))
        check("guard: own cell inside a range or a bare reference is caught; another sheet and text are not",
              hits == [("أخرى", "A1"), ("ورقة", "F7")], hits)
        # C24: anything that slips past the masking is caught after saving; the file is deleted and the exit is 2
        global mask_pii
        real = mask_pii
        mask_pii = lambda v, levels=("block",): (str(v), [])  # noqa: E731
        try:
            rc = export(d, out)
        finally:
            mask_pii = real
        check("C24 self-check: unmasked phone -> exit 2 and no file", rc == 2 and not os.path.exists(out), rc)
    check("Arabic counts: 0/1/2/3/11/100 records",
          [n_records(k) for k in (0, 1, 2, 3, 11, 100)] == ["لا سجلات", "سجل واحد", "سجلان", "3 سجلات", "11 سجلًا", "100 سجل"],
          [n_records(k) for k in (0, 1, 2, 3, 11, 100)])
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
