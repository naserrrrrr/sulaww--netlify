#!/usr/bin/env python3
"""Validate a run folder before (and after) the dashboard is built.

Usage:
  python3 scripts/validate_analysis.py --run-dir RUN [--xlsx SOURCE.xlsx] [--html RUN/dashboard.html]
                                       [--xlsx-out RUN/supporting_workbook.xlsx]

Reads RUN/01..06 (data-contract.md §0-§6), plus RUN/analysis.json when it exists, RUN/09_message.md and
RUN/supporting_workbook.xlsx (or --xlsx-out) when they exist.
Writes RUN/07_validation.json ({blocking_count, warning_count, acknowledged_count, checks:[...], reconciliation})
and RUN/07_validation.md (Arabic).
Exit codes: 0 = no blocking error, 2 = at least one blocking error, 1 = the run folder itself is unusable.
Warnings acknowledged in 06 `acknowledged_warnings` ({check_id, reason ≥ 8 words}) get level "acknowledged" and are
not counted in warning_count; blocking checks can never be acknowledged.

Privacy: values of personal-data columns are read from --xlsx into memory only, to look for them in the
outputs. They are never printed, logged or written; findings name the source cell and the output file only.
"""
from __future__ import annotations

import os
import sys

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argparse  # noqa: E402
import datetime as _dt  # noqa: E402
import html as _html  # noqa: E402
import json  # noqa: E402
import math  # noqa: E402
import re  # noqa: E402
from collections import OrderedDict, defaultdict  # noqa: E402

from apa_common import (  # noqa: E402
    ALL_STATUSES, CONTRACT_VERSION, EMAIL_RE, KNOWN_STATUSES, NA_TOKENS, NOT_APPLICABLE_TOKENS,
    PII_HEADER_RE, SKILL_DIR, clean_text, load_json, normalize_arabic, save_json, sha256_file, word_count_ar,
)

RUN_FILES = OrderedDict([
    ("wb", "01_workbook_map.json"),
    ("mapping", "02_mapping.json"),
    ("records", "03_records.json"),
    ("config", "04_metrics_config.json"),
    ("metrics", "05_metrics.json"),
    ("narrative", "06_narrative.json"),
])
SCHEMA_PATH = os.path.join(SKILL_DIR, "schemas", "narrative.schema.json")
METRIC_STATUSES = {"complete", "partial", "not_computable", "not_comparable", "empty"}
PAIRWISE_KINDS = {"ratio": ("num", "den"), "diff": ("a", "b"), "growth": ("prev", "curr")}
NUM_KEYS_IN_RESULT = ("value", "num", "den", "num_known_total", "prev_sum", "curr_sum", "abs_change", "total",
                      "top_sum", "a_sum", "b_sum")
REC_REQUIRED = ("id", "gap_ids", "target", "problem", "evidence", "first_action", "alternatives",
                "recommended_option", "rationale", "owner_role", "duration", "cost", "expected_impact",
                "success_kpi", "ranking", "priority")
RANKING_KEYS = ("impact", "urgency", "evidence_strength", "feasibility", "cost_level")
REL_TOL = 1e-9

_AR_INDIC = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
TOKEN_RE = re.compile(r"\{\{\s*m:([A-Za-z0-9_\-]+)((?:\.[A-Za-z0-9_]+)*)\s*\}\}")
LOOSE_TOKEN_RE = re.compile(r"\{\{|\}\}")
# Saudi mobile, tolerant to one space/dash between digits (visible text and strings only — never raw numbers)
PHONE_LOOSE_RE = re.compile(r"(?<!\d)(?:\+?966|00966|0)?[\s\-]?5(?:[\s\-]?\d){8}(?!\d)")
HEX_RE = re.compile(r"^[0-9a-fA-F]{16,}$")
PII_KEY_RE = re.compile(
    r"(?i)(^|[._\-])(phone|mobile|tel|telephone|email|e_mail|supervisor|supervisor_name|contact_phone|contact_no|"
    r"contact_number|national_id|id_number|iban)($|[._\-])|جوال|هاتف|تواصل|المشرف|ايميل|إيميل|بريد|هوية")
# Generic, non-personal words that may sit in a "supervisor" column (normalised); never treated as PII values.
GENERIC_NON_PII = {normalize_arabic(x) for x in (
    "الجمعية", "جمعية", "الإدارة", "إدارة", "الادارة", "المكتب", "مكتب", "شركة", "مؤسسة", "المالك", "مالك",
    "الوقف", "وقف", "لا يوجد", "غير محدد", "غير معروف", "بدون", "لايوجد",
)} | {normalize_arabic(x) for x in NA_TOKENS | NOT_APPLICABLE_TOKENS}

NEGATIONS = {"ليس", "ليست", "لا", "غير", "دون", "بدون", "ولا", "وليس", "وليست", "لم", "ليسا", "ولم", "بلا"}
NET_RE = re.compile(r"(?<![\w])(?:ال)?صافي(?:ا|ه)?(?![\w])")
TOTAL_RETURN_RE = re.compile(r"(?<![\w])(?:ال)?عايدا?\s+(?:(?:ال)?استثماريا?\s+)?(?:ال)?كليا?(?![\w])")
NET_YIELD_RE = re.compile(r"(?<![\w])(?:ال)?عايدا?\s+(?:ال)?صافيا?(?![\w])|(?<![\w])صافي\s+(?:ال)?عايد(?![\w])")
BEFORE_EXP_RE = re.compile(r"قبل\s+(?:ال)?مصروفات")
SCORE_RE = re.compile(r"من\s*100|/\s*100(?!\d)|(?<!\d)\d{1,2}\s*/\s*10(?!\d)|درجه\s*\d|\d+\s*نقاط?(?![\w])|(?<![\w])علامه\s*\d")
PCT_RE = re.compile(r"\d+(?:[.,]\d+)?\s*[%٪]|بالمايه|في المايه")
JUDGE_RE = re.compile(r"ممتاز|جيد|ضعيف|سيي|مقبول|متميز|رايع|مرتفع|منخفض|متدن|عالي|قوي|ممتازه|ضعيفه")
SUPERLATIVE_RE = re.compile(r"ممتاز|متميز|رايع|مثالي|ضعيفه?\s+جدا|سييه?\s+جدا|سيي\s+للغايه|كارثي")
_AR_LETTER = "\u0621-\u064a"
DRASTIC_RE = re.compile(rf"(?<![{_AR_LETTER}])(?:و|ف)?(?:ب|ل)?(?:ال)?(?:بيع|هدم)(?:ه|ها|هم|هما)?(?![{_AR_LETTER}])|تغيير\s+(?:ال)?استخدام")
STUDY_RE = re.compile(r"دراس|تحقق|التحقق|تقييم|مراجع|جدوي|استشار|فحص|مقارن|عروض|نظامي|موافق|اعتماد|حصر|قيود|متطلبات|تثبت")
IMPLEMENTED_RE = re.compile(r"منفذ|نفذت|تم التنفيذ|مكتمل|منجز")
ELEMENTS = OrderedDict([
    ("state", ("وصف حالة المحفظة", re.compile(r"المحفظه|وضع|حاله"))),
    ("strengths", ("نقاط القوة", re.compile(r"قو[هي]|ايجابي|ميزه"))),
    ("gaps", ("أبرز الفجوات", re.compile(r"فجو"))),
    ("priorities", ("الأولويات", re.compile(r"اولوي"))),
    ("completeness", ("اكتمال البيانات", re.compile(r"اكتمال|مكتمل|تغطيه|نقص البيانات"))),
])
NUMBER_RE = re.compile(r"(?<![\w.,\-/])(\d{1,3}(?:[,٬]\d{3})+(?:[.٫]\d+)?|\d{4,}(?:[.٫]\d+)?)(?![\w%٪])(?![.,٫٬]\d)")
DATE_RE = re.compile(r"\d{1,4}[/\-]\d{1,2}[/\-]\d{1,4}")
NUMBER_SKIP_KEYS = {"id", "ref", "evidence", "gap_ids", "rec_id", "top_gaps", "top_actions", "sources", "scenarios",
                    "asset_ids", "version_comparison", "recommendation_followup", "metric", "where", "date",
                    "contract_version", "record_ids"}

CHECK_TITLES = OrderedDict([
    ("files_exist", "وجود ملفات التشغيل المطلوبة (01–06) وصحة صيغتها"),
    ("narrative_schema", "مطابقة 06_narrative.json لمخطط السرد"),
    ("exec_summary_length", "طول الخلاصة التنفيذية بين 120 و180 كلمة"),
    ("metric_tokens", "صحة رموز المؤشرات {{m:…}}"),
    ("duplicate_record_ids", "عدم تكرار معرّفات السجلات"),
    ("total_rows_in_scope", "عدم إدخال صفوف الإجمالي في النطاق أو في المؤشرات"),
    ("tv_status_consistency", "اتساق القيم مع حالاتها (المجهول لا يتحول صفرًا)"),
    ("metric_status_value", "اتساق حالة كل مؤشر مع قيمته وسببه"),
    ("pairwise_sets", "تطابق مجموعة السجلات في البسط والمقام"),
    ("metrics_reproducible", "إعادة احتساب 05 من 03 و04 بالنتيجة نفسها"),
    ("label_net", "عدم تسمية العائد «صافيًا» دون مصروفات"),
    ("label_total_return", "عدم تسمية «العائد الاستثماري الكلي» دون مدخلاته"),
    ("label_before_expenses", "ذكر «قبل المصروفات» في تسمية العائد عند غياب المصروفات"),
    ("questions_max", "ثلاثة أسئلة كحد أقصى"),
    ("narrative_ids", "تفرد معرّفات الفجوات والتوصيات"),
    ("gap_evidence", "لكل فجوة دليل يُحال إلى مرجع موجود"),
    ("recommendation_links", "ارتباط كل توصية بفجوة موجودة"),
    ("recommendation_fields", "اكتمال حقول كل توصية"),
    ("plan_links", "ارتباط بنود خطة 30/60/90 بتوصيات موجودة"),
    ("top_lists", "أهم ثلاث فجوات وثلاثة إجراءات موجودة"),
    ("scenarios", "سلامة السيناريوهات (مصادر الافتراضات وعدم الازدواج)"),
    ("pii_text_patterns", "خلو السرد والسجلات من أرقام جوال أو بريد"),
    ("pii_keys", "خلو analysis.json من حقول بيانات شخصية"),
    ("pii_html_patterns", "خلو الداشبورد من أرقام جوال أو بريد"),
    ("pii_values", "خلو المخرجات من قيم أعمدة البيانات الشخصية في الملف"),
    ("assessment_scores", "التقييم دون درجات أو أحكام مطلقة بلا معيار"),
    ("superlatives", "عدم استخدام أحكام مطلقة في الخلاصة ونقاط القوة"),
    ("hand_typed_numbers", "الأرقام المكتوبة يدويًا في السرد"),
    ("numbers_as_tokens", "أرقام يدوية تطابق مؤشرًا (يُفضّل الرمز)"),
    ("valuation_years_mixed", "عدم خلط تقييمات سنوات مختلفة في مجموع واحد"),
    ("growth_comparability", "النمو المحسوب مع قابلية مقارنة غير مؤكدة"),
    ("drastic_recommendations", "توصيات البيع أو الهدم أو تغيير الاستخدام مسبوقة بدراسة وتحقق"),
    ("proven_problem_cause", "المشكلة المثبتة لها سبب مثبت"),
    ("data_coverage_low", "تغطية البيانات 50% فأكثر"),
    ("unknown_classification", "تصنيف الاستخدام وطبيعة الدخل للسجلات ضمن النطاق"),
    ("card_count", "عدد البطاقات بين 6 و8"),
    ("exec_summary_elements", "اشتمال الخلاصة على عناصرها الخمسة"),
    ("token_null_value", "رموز تشير إلى مؤشر بلا قيمة"),
    ("evidence_refs_other", "مراجع أدلة التقييم ونقاط القوة والتوصيات"),
    ("isolated_in_metrics", "عدم دخول سجلات معزولة أو مستبعدة في المؤشرات"),
    ("narrative_claims", "عدم ادعاء عائد صافٍ أو كلي في السرد دون مدخلاته"),
    ("followup_status", "متابعة التوصيات السابقة دون افتراض التنفيذ"),
    ("source_hash", "تطابق ملف Excel مع الملف المفحوص في 01"),
    ("gaps_without_recommendation", "فجوات بلا توصية مرتبطة"),
])

LEVEL_AR = {"blocking": "مانع", "warning": "تحذير", "info": "معلومة"}


class Skip(Exception):
    """A check cannot run (missing input); reported as an info entry, not a pass."""


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def nonempty(x) -> bool:
    if x is None:
        return False
    if isinstance(x, str):
        return bool(clean_text(x))
    if isinstance(x, (list, tuple, dict, set)):
        return len(x) > 0
    return True


def is_num(x) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def close(a, b, tol=REL_TOL) -> bool:
    if a is None or b is None:
        return a is None and b is None
    if not (is_num(a) and is_num(b)):
        return a == b
    return abs(a - b) <= max(tol * max(abs(a), abs(b)), 1e-9)


def as_list(x) -> list:
    return x if isinstance(x, list) else []


def as_dict(x) -> dict:
    return x if isinstance(x, dict) else {}


def short_list(items, n=12) -> str:
    items = [str(i) for i in items]
    return "، ".join(items[:n]) + (f" … (+{len(items) - n})" if len(items) > n else "")


def fmt_num(x) -> str:
    if not is_num(x):
        return str(x)
    if float(x).is_integer():
        return f"{int(x):,}"
    return f"{x:,.2f}"


def walk_strings(obj, path="", skip_keys=frozenset()):
    """Yield (path, string) for every string in a JSON-like structure."""
    if isinstance(obj, str):
        yield path, obj
    elif isinstance(obj, dict):
        for k, v in obj.items():
            if k in skip_keys or (isinstance(k, str) and k.startswith("x_") and "x_" in skip_keys):
                continue
            yield from walk_strings(v, f"{path}.{k}" if path else str(k), skip_keys)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from walk_strings(v, f"{path}[{i}]", skip_keys)


def walk_keys(obj, path=""):
    if isinstance(obj, dict):
        for k, v in obj.items():
            p = f"{path}.{k}" if path else str(k)
            yield p, k
            yield from walk_keys(v, p)
    elif isinstance(obj, list):
        for i, v in enumerate(obj):
            yield from walk_keys(v, f"{path}[{i}]")


def negated(norm_text: str, start: int, window: int = 3) -> bool:
    before = norm_text[:start].split()[-window:]
    return any(w in NEGATIONS for w in before)


def affirmative_hits(regex, text: str) -> list:
    """Matches of regex in the normalised text that are not preceded by a negation word."""
    n = normalize_arabic(text)
    return [m.group(0) for m in regex.finditer(n) if not negated(n, m.start())]


def col_to_idx(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def year_of_field(field: str, meta: dict):
    y = as_dict(meta.get(field)).get("year")
    if y is not None:
        return str(y)
    m = re.match(r"^val\.(\d{4})\.", field or "")
    return m.group(1) if m else None


def field_group(field, meta: dict):
    g = as_dict(meta.get(field)).get("group")
    if g:
        return g
    if not isinstance(field, str):
        return None
    for pre, grp in (("rev.", "revenue"), ("val.", "valuation"), ("exp.", "expense")):
        if field.startswith(pre):
            return grp
    return None


def gid(item, i, key="id"):
    return f"{i}({item.get(key)})" if isinstance(item, dict) and item.get(key) else str(i)


# ---------------------------------------------------------------------------
# Context: loads the run folder once
# ---------------------------------------------------------------------------
class Ctx:
    def __init__(self, run_dir: str, xlsx: str = None, html_path: str = None):
        self.run_dir = run_dir
        self.xlsx = xlsx
        self.html_path = html_path
        self.data = {}
        self.load_errors = OrderedDict()
        for key, name in RUN_FILES.items():
            p = os.path.join(run_dir, name)
            if not os.path.isfile(p):
                self.load_errors[key] = "missing"
                continue
            try:
                self.data[key] = load_json(p)
            except Exception as exc:  # invalid JSON
                self.load_errors[key] = f"invalid: {exc}"
        self.analysis = None
        self.analysis_error = None
        ap = os.path.join(run_dir, "analysis.json")
        if os.path.isfile(ap):
            try:
                self.analysis = load_json(ap)
            except Exception as exc:
                self.analysis_error = str(exc)
        self.html_raw = None
        if html_path and os.path.isfile(html_path):
            with open(html_path, "r", encoding="utf-8", errors="replace") as fh:
                self.html_raw = fh.read()
        self._resolver = None

    # typed accessors -------------------------------------------------------
    def need(self, *keys):
        for k in keys:
            if k not in self.data:
                raise Skip(f"الملف {RUN_FILES[k]} غير متاح أو غير صالح")
            if not isinstance(self.data[k], dict):
                raise Skip(f"الملف {RUN_FILES[k]} ليس كائن JSON")

    @property
    def wb(self):
        return as_dict(self.data.get("wb"))

    @property
    def mapping(self):
        return as_dict(self.data.get("mapping"))

    @property
    def rec_doc(self):
        return as_dict(self.data.get("records"))

    @property
    def records(self):
        return [r for r in as_list(self.rec_doc.get("records")) if isinstance(r, dict)]

    @property
    def rec_by_id(self):
        out = {}
        for r in self.records:
            out.setdefault(r.get("id"), r)
        return out

    @property
    def field_meta(self):
        return as_dict(self.rec_doc.get("field_meta")) or as_dict(self.mapping.get("field_meta"))

    @property
    def config(self):
        return as_dict(self.data.get("config"))

    @property
    def met_doc(self):
        return as_dict(self.data.get("metrics"))

    @property
    def metrics(self):
        return as_dict(self.met_doc.get("metrics"))

    @property
    def narrative(self):
        return as_dict(self.data.get("narrative"))

    def resolver(self):
        if self._resolver is None:
            self._resolver = Resolver(self)
        return self._resolver


# ---------------------------------------------------------------------------
# Evidence reference resolution
# ---------------------------------------------------------------------------
CELL_COORD_RE = re.compile(
    r"^(?:\$?(?P<c1>[A-Z]{1,3})\$?(?P<r1>\d+)(?::\$?(?P<c2>[A-Z]{1,3})\$?(?P<r2>\d+))?"
    r"|\$?(?P<cc1>[A-Z]{1,3}):\$?(?P<cc2>[A-Z]{1,3})|(?P<rr1>\d+):(?P<rr2>\d+))(?:\s*[(（].*[)）])?\s*$")


class Resolver:
    def __init__(self, ctx: Ctx):
        self.metrics = ctx.metrics
        self.breakdowns = as_dict(ctx.met_doc.get("breakdowns"))
        self.has_trend = isinstance(ctx.met_doc.get("trend"), dict)
        self.dq_ids = {i.get("id") for i in as_list(ctx.rec_doc.get("issues")) if isinstance(i, dict)}
        self.dq_ids |= {i.get("id") for i in as_list(ctx.narrative.get("data_issues_extra")) if isinstance(i, dict)}
        self.wb_ids = {a.get("id") for a in as_list(ctx.wb.get("anomalies")) if isinstance(a, dict)}
        self.record_ids = {r.get("id") for r in ctx.records}
        self.sheets = {s.get("name"): s for s in as_list(ctx.wb.get("sheets")) if isinstance(s, dict)}
        self.sheets_clean = {clean_text(n): n for n in self.sheets if isinstance(n, str)}

    def resolve(self, ref):
        """Return (ok, note). note explains a failure, or a non-blocking remark on success."""
        if not isinstance(ref, str) or not clean_text(ref):
            return False, "مرجع فارغ"
        r = ref.strip()
        ok, note = self._single(r)
        if ok:
            return ok, note
        parts = [p for p in re.split(r"\s*[,،;]\s*|\s+\+\s+", r) if p.strip()]
        if len(parts) > 1:
            notes = []
            for p in parts:
                pok, pnote = self._single(p.strip())
                if not pok:
                    return False, f"الجزء «{p.strip()}»: {pnote}"
                if pnote:
                    notes.append(pnote)
            return True, "؛ ".join(notes) or None
        return False, note

    def _single(self, r):
        if r.startswith("m:"):
            body = r[2:]
            mid, _, path = body.partition(".")
            if mid in self.metrics:
                if path:
                    ok = resolve_path(self.metrics[mid], path.split("."))[0]
                    return (True, None) if ok else (False, f"الحقل «{path}» غير موجود في نتيجة المؤشر {mid}")
                return True, None
            if mid in self.breakdowns or (mid == "trend" and self.has_trend):
                return True, None
            return False, f"المؤشر «{mid}» غير موجود في 05_metrics.json"
        if r.startswith("DQ-"):
            return (True, None) if r in self.dq_ids else (False, "معرّف مشكلة بيانات غير موجود في 03 أو data_issues_extra")
        if r.startswith("WB-"):
            return (True, None) if r in self.wb_ids else (False, "معرّف ملاحظة مصنف غير موجود في 01")
        if r in self.record_ids:
            return True, None
        if "!" in r:
            return self._cell(r)
        return False, "مرجع غير معروف (المسموح: m:<مؤشر>، DQ-…، WB-…، معرّف سجل، أو Sheet!A1)"

    def _cell(self, r):
        sheet, _, coord = r.rpartition("!")
        sheet_q = sheet.strip()
        if len(sheet_q) > 1 and sheet_q[0] == "'" and sheet_q[-1] == "'":
            sheet_q = sheet_q[1:-1].replace("''", "'")
        m = CELL_COORD_RE.match(coord.strip())
        if not m:
            return False, f"إحداثي الخلية «{coord.strip()}» غير صالح"
        note = None
        if sheet in self.sheets:
            s = self.sheets[sheet]
        elif sheet_q in self.sheets:
            s = self.sheets[sheet_q]
        elif clean_text(sheet_q) in self.sheets_clean:
            exact = self.sheets_clean[clean_text(sheet_q)]
            s = self.sheets[exact]
            note = f"اسم الورقة يختلف في المسافات عن الاسم الدقيق «{exact}»"
        else:
            return False, f"الورقة «{sheet_q}» غير موجودة في خريطة المصنف 01"
        max_row, max_col = s.get("max_row"), s.get("max_col")
        rows = [int(x) for x in (m.group("r1"), m.group("r2"), m.group("rr1"), m.group("rr2")) if x]
        cols = [col_to_idx(x) for x in (m.group("c1"), m.group("c2"), m.group("cc1"), m.group("cc2")) if x]
        if isinstance(max_row, int) and rows and max(rows) > max_row:
            return False, f"الصف {max(rows)} خارج نطاق الورقة (آخر صف {max_row})"
        if isinstance(max_col, int) and cols and max(cols) > max_col:
            return False, f"العمود خارج نطاق الورقة (آخر عمود رقم {max_col})"
        return True, note


def resolve_path(obj, parts):
    cur = obj
    for p in parts:
        if isinstance(cur, dict) and p in cur:
            cur = cur[p]
        elif isinstance(cur, list) and p.isdigit() and int(p) < len(cur):
            cur = cur[int(p)]
        else:
            return False, None
    return True, cur


# ---------------------------------------------------------------------------
# BLOCKING checks — each returns a list of (detail, where)
# ---------------------------------------------------------------------------
def c_files_exist(ctx: Ctx):
    out = []
    for key, err in ctx.load_errors.items():
        name = RUN_FILES[key]
        if err == "missing":
            out.append((f"الملف {name} غير موجود في مجلد التشغيل", name))
        else:
            out.append((f"الملف {name} ليس JSON صالحًا: {err[9:200]}", name))
    for key, val in ctx.data.items():
        if not isinstance(val, dict):
            out.append((f"الملف {RUN_FILES[key]} يجب أن يكون كائن JSON", RUN_FILES[key]))
        elif val.get("contract_version") not in (None, CONTRACT_VERSION):
            out.append((f"إصدار العقد في {RUN_FILES[key]} = {val.get('contract_version')!r} والمتوقع {CONTRACT_VERSION}",
                        RUN_FILES[key]))
    if ctx.analysis_error:
        out.append((f"analysis.json ليس JSON صالحًا: {ctx.analysis_error[:200]}", "analysis.json"))
    if ctx.html_path and ctx.html_raw is None:
        out.append(("ملف الداشبورد المحدد في --html غير موجود", ctx.html_path))
    return out


def _schema_msg(e) -> str:
    v = e.validator
    inst = e.instance
    if v == "required":
        m = re.search(r"'([^']+)' is a required property", e.message)
        return f"الحقل المطلوب «{m.group(1) if m else e.message}» مفقود"
    if v == "enum":
        return f"القيمة {inst!r} غير مسموحة؛ المسموح: {', '.join(map(str, e.validator_value))}"
    if v == "const":
        return f"القيمة يجب أن تكون {e.validator_value!r}"
    if v == "maxItems":
        return f"عدد العناصر {len(inst)} يتجاوز الحد الأقصى {e.validator_value}"
    if v == "minItems":
        return f"عدد العناصر {len(inst) if isinstance(inst, list) else '?'} أقل من الحد الأدنى {e.validator_value}"
    if v == "type":
        return f"النوع يجب أن يكون {e.validator_value} (الموجود: {type(inst).__name__})"
    if v == "additionalProperties":
        known = set((e.schema or {}).get("properties") or {})
        pats = list(((e.schema or {}).get("patternProperties") or {}).keys())
        extra = [k for k in (inst if isinstance(inst, dict) else {})
                 if k not in known and not any(re.search(p, k) for p in pats)]
        return f"مفتاح غير معروف في هذا الموضع: {', '.join(extra) or e.message} (للامتدادات استخدم بادئة x_)"
    if v == "contains":
        dim = (((e.validator_value or {}).get("properties") or {}).get("dimension") or {}).get("const")
        return f"البعد «{dim}» مفقود من التقييم" if dim else "عنصر مطلوب مفقود من القائمة"
    if v in ("pattern",):
        return f"القيمة {inst!r} لا تطابق الصيغة المطلوبة {e.validator_value}"
    if v == "minLength":
        return "نص فارغ"
    if v == "uniqueItems":
        return "عناصر مكررة في القائمة"
    if v == "minimum":
        return f"القيمة {inst!r} أقل من {e.validator_value}"
    if v == "oneOf":
        try:
            from jsonschema.exceptions import best_match
            sub = best_match(e.context)
            if sub is not None:
                return ("لا يطابق أيًا من الصيغتين المسموحتين (not_prepared مع reason وrequired_inputs، أو prepared مع items)"
                        f"؛ أقرب خطأ: {_schema_msg(sub)} عند {'/'.join(map(str, sub.relative_path)) or '(الجذر)'}")
        except Exception:
            pass
        return "لا يطابق أيًا من الصيغ المسموحة"
    return e.message


def c_narrative_schema(ctx: Ctx):
    ctx.need("narrative")
    try:
        import jsonschema
    except ImportError:
        raise Skip("مكتبة jsonschema غير مثبتة")
    schema = load_json(SCHEMA_PATH)
    validator = jsonschema.Draft202012Validator(schema)
    out = []
    for e in sorted(validator.iter_errors(ctx.data["narrative"]), key=lambda e: list(map(str, e.absolute_path))):
        path = "/".join(map(str, e.absolute_path)) or "(الجذر)"
        out.append((_schema_msg(e), f"06_narrative.json:{path}"))
    return out


def summary_word_count(text: str) -> int:
    return word_count_ar(TOKEN_RE.sub(" 0 ", text))


def c_exec_summary_length(ctx: Ctx):
    ctx.need("narrative")
    s = ctx.narrative.get("executive_summary")
    if not isinstance(s, str):
        raise Skip("الخلاصة التنفيذية ليست نصًا")
    n = summary_word_count(s)
    if 120 <= n <= 180:
        return []
    return [(f"الخلاصة التنفيذية {n} كلمة (المطلوب 120–180، كل رمز {{{{m:…}}}} يُحتسب كلمة واحدة)",
             "06_narrative.json:executive_summary")]


def c_metric_tokens(ctx: Ctx):
    ctx.need("narrative", "metrics")
    out = []
    for path, s in walk_strings(ctx.narrative):
        for m in TOKEN_RE.finditer(s):
            mid, fpath = m.group(1), m.group(2).lstrip(".")
            if mid not in ctx.metrics:
                out.append((f"الرمز {m.group(0)} يشير إلى مؤشر غير موجود في 05_metrics.json", f"06_narrative.json:{path}"))
            elif fpath and not resolve_path(ctx.metrics[mid], fpath.split("."))[0]:
                out.append((f"الرمز {m.group(0)} يشير إلى حقل «{fpath}» غير موجود في نتيجة المؤشر",
                            f"06_narrative.json:{path}"))
        rest = FACT_TOKEN_RE.sub("", TOKEN_RE.sub("", s))  # {{f:…}} keys are resolved by fact_tokens
        if LOOSE_TOKEN_RE.search(rest):
            frag = rest[max(0, rest.find("{{") - 10): rest.find("{{") + 30] if "{{" in rest else rest[:40]
            out.append((f"رمز مشوّه سيظهر نصًا خامًا في الداشبورد: «{frag.strip()}» (الصيغة: {{{{m:metric_id}}}} أو {{{{m:metric_id.field}}}})",
                        f"06_narrative.json:{path}"))
    return out


def c_duplicate_record_ids(ctx: Ctx):
    ctx.need("records")
    seen = defaultdict(list)
    out = []
    for i, r in enumerate(ctx.records):
        rid = r.get("id")
        if not isinstance(rid, str) or not rid:
            out.append((f"سجل بلا معرّف صالح (الموضع {i})", f"03_records.json:records[{i}]"))
            continue
        seen[rid].append(r.get("source", {}).get("ref_name") if isinstance(r.get("source"), dict) else i)
    for rid, locs in seen.items():
        if len(locs) > 1:
            out.append((f"المعرّف {rid} مكرر {len(locs)} مرات", f"03_records.json:{rid} ({short_list(locs, 4)})"))
    return out


def c_total_rows_in_scope(ctx: Ctx):
    ctx.need("records")
    out = []
    totals = {r.get("id") for r in ctx.records if r.get("record_type") == "total"}
    for r in ctx.records:
        if r.get("record_type") == "total" and r.get("scope") == "in_scope":
            out.append((f"صف إجمالي {r.get('id')} مصنف ضمن النطاق؛ يجب استبعاده أو عزله حتى لا يُجمع مع تفاصيله",
                        f"03_records.json:{r.get('id')}"))
    if totals and "metrics" in ctx.data:
        for mid, res in ctx.metrics.items():
            used = totals & set(as_list(as_dict(res).get("included_ids")))
            if used:
                out.append((f"المؤشر {mid} يشمل صفوف إجمالي: {short_list(sorted(used))}", f"05_metrics.json:{mid}"))
    return out


def c_tv_status_consistency(ctx: Ctx):
    ctx.need("records")
    out = []
    for r in ctx.records:
        for fname, tv in as_dict(r.get("fields")).items():
            if not isinstance(tv, dict):
                out.append((f"الحقل {fname} ليس قيمة مصنفة (TV)", f"03_records.json:{r.get('id')}.{fname}"))
                continue
            s, v = tv.get("s"), tv.get("v")
            where = f"03_records.json:{r.get('id')}.{fname} ({tv.get('ref') or '-'})"
            if s not in ALL_STATUSES:
                out.append((f"حالة غير معروفة {s!r}", where))
            elif s not in KNOWN_STATUSES and v is not None:
                tag = "صفر" if v == 0 else "قيمة"
                out.append((f"الحالة {s} تعني قيمة مجهولة لكن v = {v!r} ({tag})؛ المجهول لا يُحوَّل إلى {tag}", where))
            elif s in ("zero_recorded", "zero_confirmed") and v != 0:
                out.append((f"الحالة {s} تتطلب v = 0 لكن v = {v!r}", where))
            elif s in ("value", "derived") and v is None:
                out.append((f"الحالة {s} تعني قيمة معروفة لكن v فارغة", where))
            elif s in ("value", "derived") and isinstance(v, float) and not math.isfinite(v):
                out.append((f"قيمة غير منتهية {v!r}", where))
    return out


def c_metric_status_value(ctx: Ctx):
    ctx.need("metrics")
    out = []
    for mid, res in ctx.metrics.items():
        if not isinstance(res, dict):
            out.append(("نتيجة المؤشر ليست كائنًا", f"05_metrics.json:{mid}"))
            continue
        st, val = res.get("status"), res.get("value")
        where = f"05_metrics.json:{mid}"
        if st not in METRIC_STATUSES:
            out.append((f"حالة مؤشر غير معروفة {st!r}", where))
            continue
        if is_num(val) and isinstance(val, float) and not math.isfinite(val):
            out.append((f"قيمة غير منتهية {val!r}", where))
        if st in ("complete", "partial") and val is None:
            if res.get("kind") == "growth" and is_num(res.get("prev_sum")) and res.get("prev_sum") <= 0:
                continue  # contract §5: pct is null when prev_sum is 0 ("بدء إيراد")
            out.append((f"الحالة {st} لكن القيمة فارغة؛ إما أن تُحسب القيمة أو تكون الحالة not_computable مع سبب", where))
        if st in ("not_computable", "not_comparable") and not nonempty(res.get("reason")):
            out.append((f"الحالة {st} دون سبب (reason)؛ البطاقة ستظهر مبهمة", where))
    if "config" in ctx.data:
        for spec in as_list(ctx.config.get("metrics")):
            if isinstance(spec, dict) and spec.get("id") not in ctx.metrics:
                out.append((f"المؤشر {spec.get('id')} معرّف في 04 لكنه غير موجود في 05؛ أعد تشغيل compute_metrics.py",
                            "05_metrics.json"))
    return out


def _fv(rec, field):
    tv = as_dict(as_dict(rec.get("fields")).get(field))
    return tv.get("s"), tv.get("v")


def _known(rec, field):
    s, v = _fv(rec, field)
    return s in KNOWN_STATUSES and is_num(v)


def c_pairwise_sets(ctx: Ctx):
    ctx.need("metrics", "records")
    out = []
    recs = ctx.rec_by_id
    for mid, res in ctx.metrics.items():
        if not isinstance(res, dict) or res.get("kind") not in PAIRWISE_KINDS:
            continue
        where = f"05_metrics.json:{mid}"
        st = res.get("status")
        inc = as_list(res.get("included_ids"))
        exc = [e.get("id") for e in as_list(res.get("excluded")) if isinstance(e, dict)]
        n_inc, n_scope = res.get("n_included"), res.get("n_scope")
        if n_inc != len(inc):
            out.append((f"n_included = {n_inc} لا يساوي عدد included_ids = {len(inc)}", where))
        if len(set(inc)) != len(inc):
            out.append(("سجلات مكررة داخل included_ids", where))
        both = set(inc) & set(exc)
        if both:
            out.append((f"سجلات مشمولة ومستبعدة معًا: {short_list(sorted(both))}", where))
        evaluated = st in ("complete", "partial") or (st == "not_computable" and (inc or exc))
        if evaluated and is_num(n_scope) and len(inc) + len(exc) != n_scope:
            out.append((f"المشمول ({len(inc)}) + المستبعد ({len(exc)}) ≠ n_scope ({n_scope}); جزء من النطاق غير مفسَّر", where))
        if not evaluated:
            continue
        k1, k2 = PAIRWISE_KINDS[res["kind"]]
        rf = as_dict(res.get("resolved_fields"))
        f1, f2 = rf.get(k1), rf.get(k2)
        if not (isinstance(f1, str) and isinstance(f2, str)):
            out.append((f"resolved_fields لا يحدد حقلي {k1}/{k2}", where))
            continue
        bad_inc, bad_exc, missing = [], [], []
        s1 = s2 = sd = 0.0
        for rid in inc:
            r = recs.get(rid)
            if r is None:
                missing.append(rid)
                continue
            ok = _known(r, f1) and _known(r, f2)
            if ok and res["kind"] == "ratio":
                ok = _fv(r, f2)[1] > 0
            if not ok:
                bad_inc.append(rid)
                continue
            v1, v2 = _fv(r, f1)[1], _fv(r, f2)[1]
            s1 += v1
            s2 += v2
            sd += v1 - v2
        for rid in exc:
            r = recs.get(rid)
            if r is None:
                missing.append(rid)
                continue
            ok = _known(r, f1) and _known(r, f2)
            if ok and res["kind"] == "ratio":
                ok = _fv(r, f2)[1] > 0
            if ok:
                bad_exc.append(rid)
        if missing:
            out.append((f"معرّفات غير موجودة في 03: {short_list(missing)}", where))
        if bad_inc:
            out.append((f"سجلات مشمولة لا يتوافر لها {f1} و{f2} معًا (البسط والمقام لا يغطيان المجموعة نفسها): "
                        f"{short_list(bad_inc)}", where))
        if bad_exc:
            out.append((f"سجلات مستبعدة يتوافر لها الحقلان {f1} و{f2}؛ استبعادها يخل بتطابق المجاميع: {short_list(bad_exc)}",
                        where))
        if bad_inc or missing or not inc:
            continue
        if res["kind"] == "ratio":
            if not close(res.get("num"), s1, 1e-7) or not close(res.get("den"), s2, 1e-7):
                out.append((f"البسط/المقام المسجلان ({fmt_num(res.get('num'))} / {fmt_num(res.get('den'))}) لا يساويان "
                            f"مجموع السجلات المشمولة ({fmt_num(s1)} / {fmt_num(s2)})", where))
            elif s2 > 0 and is_num(res.get("value")):
                scale = res.get("scale") if is_num(res.get("scale")) else 1
                if not close(res.get("value"), s1 / s2 * scale, 1e-7):
                    out.append((f"القيمة {res.get('value')} لا تساوي البسط ÷ المقام × {scale}", where))
        elif res["kind"] == "diff":
            if is_num(res.get("value")) and not close(res.get("value"), sd, 1e-7):
                out.append((f"القيمة {fmt_num(res.get('value'))} لا تساوي مجموع الفروق للسجلات المشمولة ({fmt_num(sd)})", where))
        elif res["kind"] == "growth":
            if not close(res.get("prev_sum"), s1, 1e-7) or not close(res.get("curr_sum"), s2, 1e-7):
                out.append((f"مجموعا الفترتين ({fmt_num(res.get('prev_sum'))} / {fmt_num(res.get('curr_sum'))}) لا يطابقان "
                            f"السجلات المشمولة ({fmt_num(s1)} / {fmt_num(s2)})", where))
    return out


def c_metrics_reproducible(ctx: Ctx):
    ctx.need("metrics", "records", "config")
    try:
        import metric_engine as me  # sibling module (Python twin of the dashboard engine)
        eng = me.Engine(ctx.config, ctx.records)
    except Exception as exc:
        raise Skip(f"تعذر تحميل محرك المؤشرات لإعادة الاحتساب: {exc}")
    filters = as_dict(ctx.met_doc.get("default_filters")) or None
    out = []
    for spec in as_list(ctx.config.get("metrics")):
        if not isinstance(spec, dict):
            continue
        mid = spec.get("id")
        old = as_dict(ctx.metrics.get(mid))
        if not old:
            continue  # reported by metric_status_value
        try:
            new = eng.compute(spec, filters)
        except Exception as exc:
            out.append((f"تعذر إعادة احتساب المؤشر: {exc}", f"04_metrics_config.json:{mid}"))
            continue
        diffs = []
        for k in ("value", "status", "n_included"):
            if not close(old.get(k), new.get(k), 1e-9):
                diffs.append(f"{k}: {old.get(k)!r} ← الآن {new.get(k)!r}")
        if diffs:
            out.append(("05 لا يطابق إعادة الاحتساب من 03 و04 الحاليين (" + "؛ ".join(diffs) +
                        ")؛ أعد تشغيل compute_metrics.py", f"05_metrics.json:{mid}"))
    return out


def _periods(ctx):
    ps = [p for p in as_list(ctx.config.get("periods")) if isinstance(p, dict)]
    return ps or [{}]


def _resolve_ph(name, period: dict, config: dict):
    if not isinstance(name, str):
        return None
    if name == "{REV}":
        return period.get("rev_field")
    if name == "{REV_PREV}":
        return period.get("prev_rev_field")
    if name == "{EXP}":
        return period.get("exp_field")
    if name == "{VAL}":
        return as_dict(config.get("valuation")).get("field")
    return name


def _spec_texts(spec: dict, period_id):
    keys = ("label", "short_label", "definition", "basis")
    out = {k: spec.get(k) for k in keys}
    tbp = as_dict(as_dict(spec.get("text_by_period")).get(period_id)) if period_id is not None else {}
    for k in keys:
        if tbp.get(k) is not None:
            out[k] = tbp[k]
    return {k: v for k, v in out.items() if isinstance(v, str)}


def _ratio_profile(spec, period, ctx):
    """(is_revenue_to_valuation_ratio, expense_resolved, n_valuation_years)."""
    meta = ctx.field_meta
    names = [spec.get(k) for k in ("num", "den", "field", "a", "b", "prev", "curr")] + as_list(spec.get("fields"))
    resolved = [_resolve_ph(n, period, ctx.config) for n in names if isinstance(n, str)]
    exp_resolved = any(isinstance(f, str) and (field_group(f, meta) == "expense" or
                                               re.match(r"^(noi|net)[._]", f)) for f in resolved)
    val_years = {year_of_field(f, meta) for f in resolved if isinstance(f, str) and field_group(f, meta) == "valuation"}
    is_rv = False
    if spec.get("kind") == "ratio":
        num = _resolve_ph(spec.get("num"), period, ctx.config)
        den = _resolve_ph(spec.get("den"), period, ctx.config)
        num_rev = spec.get("num") in ("{REV}", "{REV_PREV}") or field_group(num, meta) == "revenue"
        den_val = spec.get("den") == "{VAL}" or field_group(den, meta) == "valuation"
        is_rv = bool(num_rev and den_val)
    return is_rv, exp_resolved, len({y for y in val_years if y})


def _label_findings(ctx, test):
    ctx.need("config")
    out = []
    for spec in as_list(ctx.config.get("metrics")):
        if not isinstance(spec, dict):
            continue
        bad_periods = defaultdict(list)
        for period in _periods(ctx):
            texts = _spec_texts(spec, period.get("id"))
            prof = _ratio_profile(spec, period, ctx)
            msg = test(spec, texts, prof)
            if msg:
                bad_periods[msg].append(period.get("id") or "-")
        for msg, pids in bad_periods.items():
            out.append((f"{msg} (الفترات: {short_list(pids)})", f"04_metrics_config.json:{spec.get('id')}"))
    return out


def c_label_net(ctx: Ctx):
    def test(spec, texts, prof):
        is_rv, exp_res, _ = prof
        if not is_rv or exp_res:
            return None
        hits = [k for k in ("label", "short_label", "definition") if affirmative_hits(NET_RE, texts.get(k, ""))]
        if hits:
            return (f"نسبة إيراد إلى تقييم موصوفة بـ«صافي» في {', '.join(hits)} دون أي حقل مصروفات؛ "
                    "سمّها «الإيراد المسجل … إلى تقييم … للأصول نفسها، قبل المصروفات»")
        return None
    return _label_findings(ctx, test)


def c_label_total_return(ctx: Ctx):
    def test(spec, texts, prof):
        _, exp_res, n_val_years = prof
        hits = [k for k in ("label", "short_label", "definition") if affirmative_hits(TOTAL_RETURN_RE, texts.get(k, ""))]
        if hits and not (exp_res and n_val_years >= 2):
            return ("تسمية «العائد الاستثماري الكلي» دون مدخلاته (صافي الدخل وقيمتا بداية الفترة ونهايتها والتدفقات) في "
                    f"{', '.join(hits)}")
        return None
    return _label_findings(ctx, test)


def c_label_before_expenses(ctx: Ctx):
    def test(spec, texts, prof):
        is_rv, exp_res, _ = prof
        if not is_rv or exp_res:
            return None
        visible = " ".join(texts.get(k, "") for k in ("label", "short_label", "basis"))
        if not BEFORE_EXP_RE.search(normalize_arabic(visible)):
            return "عائد إيراد إلى تقييم دون مصروفات لا يذكر «قبل المصروفات» في التسمية أو الأساس الظاهر تحت البطاقة"
        return None
    return _label_findings(ctx, test)


def c_questions_max(ctx: Ctx):
    ctx.need("narrative")
    q = ctx.narrative.get("questions")
    if isinstance(q, list) and len(q) > 3:
        return [(f"{len(q)} أسئلة؛ الحد الأقصى ثلاثة أسئلة مؤثرة في كل دفعة", "06_narrative.json:questions")]
    return []


def c_narrative_ids(ctx: Ctx):
    ctx.need("narrative")
    out = []
    for key in ("gaps", "recommendations", "assumptions", "data_issues_extra"):
        ids = [x.get("id") for x in as_list(ctx.narrative.get(key)) if isinstance(x, dict)]
        dups = sorted({i for i in ids if ids.count(i) > 1 and i is not None})
        if dups:
            out.append((f"معرّفات مكررة في {key}: {short_list(dups)}", f"06_narrative.json:{key}"))
    if "records" in ctx.data:
        clash = {x.get("id") for x in as_list(ctx.narrative.get("data_issues_extra")) if isinstance(x, dict)} & \
                {i.get("id") for i in as_list(ctx.rec_doc.get("issues")) if isinstance(i, dict)}
        if clash:
            out.append((f"معرّفات data_issues_extra تتعارض مع معرّفات 03: {short_list(sorted(clash))}",
                        "06_narrative.json:data_issues_extra"))
    return out


def c_gap_evidence(ctx: Ctx):
    ctx.need("narrative", "records", "metrics", "wb")
    res = ctx.resolver()
    out = []
    for i, g in enumerate(as_list(ctx.narrative.get("gaps"))):
        if not isinstance(g, dict):
            continue
        base = f"06_narrative.json:gaps[{gid(g, i)}]"
        ev = as_list(g.get("evidence"))
        if not ev:
            out.append(("الفجوة بلا أي دليل (evidence)", base))
        for j, e in enumerate(ev):
            ref = e.get("ref") if isinstance(e, dict) else e
            ok, note = res.resolve(ref)
            if not ok:
                out.append((f"المرجع «{ref}» لا يُحال إلى شيء موجود: {note}", f"{base}.evidence[{j}]"))
        imp = as_dict(g.get("impact")).get("metric")
        if isinstance(imp, str) and imp.strip():
            ok, note = res.resolve(imp)
            if not ok:
                out.append((f"مؤشر الأثر «{imp}» غير موجود: {note}", f"{base}.impact.metric"))
    return out


def c_recommendation_links(ctx: Ctx):
    ctx.need("narrative")
    gap_ids = {g.get("id") for g in as_list(ctx.narrative.get("gaps")) if isinstance(g, dict)}
    out = []
    for i, r in enumerate(as_list(ctx.narrative.get("recommendations"))):
        if not isinstance(r, dict):
            continue
        where = f"06_narrative.json:recommendations[{gid(r, i)}].gap_ids"
        gids = as_list(r.get("gap_ids"))
        if not gids:
            out.append(("التوصية لا تستند إلى أي فجوة (gap_ids فارغ)", where))
        missing = [g for g in gids if g not in gap_ids]
        if missing:
            out.append((f"فجوات غير موجودة: {short_list(missing)}", where))
    return out


def c_recommendation_fields(ctx: Ctx):
    ctx.need("narrative")
    out = []
    for i, r in enumerate(as_list(ctx.narrative.get("recommendations"))):
        if not isinstance(r, dict):
            continue
        where = f"06_narrative.json:recommendations[{gid(r, i)}]"
        empty = [k for k in REC_REQUIRED if not nonempty(r.get(k))]
        if isinstance(r.get("alternatives"), list) and r["alternatives"] and not any(nonempty(a) for a in r["alternatives"]):
            empty.append("alternatives")
        if "approval_required" not in r:
            empty.append("approval_required (المفتاح مفقود؛ يُسمح بقيمة null)")
        elif r["approval_required"] is not None and not nonempty(r["approval_required"]):
            empty.append("approval_required (نص فارغ؛ استخدم null إن لم يلزم اعتماد)")
        rk = as_dict(r.get("ranking"))
        if rk:
            empty += [f"ranking.{k}" for k in RANKING_KEYS if not nonempty(rk.get(k))]
        if empty:
            out.append((f"حقول ناقصة أو فارغة: {short_list(sorted(set(empty)), 20)}", where))
    return out


def c_plan_links(ctx: Ctx):
    ctx.need("narrative")
    rec_ids = {r.get("id") for r in as_list(ctx.narrative.get("recommendations")) if isinstance(r, dict)}
    plan = as_dict(ctx.narrative.get("plan"))
    out = []
    for key in ("d30", "d60", "d90"):
        for i, it in enumerate(as_list(plan.get(key))):
            if isinstance(it, dict) and it.get("rec_id") not in rec_ids:
                out.append((f"بند الخطة يشير إلى توصية غير موجودة «{it.get('rec_id')}»", f"06_narrative.json:plan.{key}[{i}]"))
    return out


def c_top_lists(ctx: Ctx):
    ctx.need("narrative")
    n = ctx.narrative
    out = []
    for key, pool_key, label in (("top_gaps", "gaps", "الفجوات"), ("top_actions", "recommendations", "التوصيات")):
        lst = n.get(key)
        if not isinstance(lst, list):
            continue
        pool = {x.get("id") for x in as_list(n.get(pool_key)) if isinstance(x, dict)}
        if len(lst) > 3:
            out.append((f"{key} يضم {len(lst)} عناصر؛ الحد الأقصى ثلاثة", f"06_narrative.json:{key}"))
        missing = [x for x in lst if x not in pool]
        if missing:
            out.append((f"عناصر غير موجودة في {label}: {short_list(missing)}", f"06_narrative.json:{key}"))
        if len(set(map(str, lst))) != len(lst):
            out.append(("عناصر مكررة", f"06_narrative.json:{key}"))
    return out


def _scenario_type(item):
    return normalize_arabic(item.get("opportunity_type") or item.get("opportunity") or "")


def scenario_findings(ctx: Ctx):
    """Return (blocking, warnings)."""
    sc = as_dict(ctx.narrative.get("scenarios"))
    if sc.get("status") != "prepared":
        return [], []
    blocking, warnings = [], []
    items = [x for x in as_list(sc.get("items")) if isinstance(x, dict)]
    if not items:
        blocking.append(("السيناريوهات معلنة «prepared» دون بنود (items)", "06_narrative.json:scenarios"))
    rec_ids = {r.get("id") for r in ctx.records} if "records" in ctx.data else None
    for i, it in enumerate(items):
        base = f"06_narrative.json:scenarios.items[{gid(it, i)}]"
        for j, a in enumerate(as_list(it.get("assumptions"))):
            if not isinstance(a, dict) or not nonempty(a.get("source")):
                blocking.append(("افتراض بلا مصدر أو تفسير (source)؛ لا تُعد مبالغ أو نسب تحسن دون سند",
                                 f"{base}.assumptions[{j}]"))
        if not as_list(it.get("assumptions")):
            blocking.append(("بند سيناريو بلا افتراضات موثقة", base))
        if rec_ids is not None:
            missing = [a for a in as_list(it.get("asset_ids")) if a not in rec_ids]
            if missing:
                blocking.append((f"أصول غير موجودة في 03: {short_list(missing)}", f"{base}.asset_ids"))
    for i in range(len(items)):
        for j in range(i + 1, len(items)):
            a, b = items[i], items[j]
            common = set(map(str, as_list(a.get("asset_ids")))) & set(map(str, as_list(b.get("asset_ids"))))
            if not common:
                continue
            declared = b.get("id") in as_list(a.get("overlaps_with")) or a.get("id") in as_list(b.get("overlaps_with"))
            if declared:
                continue
            where = f"06_narrative.json:scenarios.items[{a.get('id')}↔{b.get('id')}]"
            if _scenario_type(a) == _scenario_type(b):
                blocking.append((f"فرصتان من النوع نفسه على الأصول نفسها ({short_list(sorted(common))}) دون overlaps_with؛ "
                                 "يُحتسب الأثر مرتين", where))
            else:
                warnings.append((f"فرصتان مختلفتان على الأصول نفسها ({short_list(sorted(common))}) دون overlaps_with؛ "
                                 "تحقق من عدم ازدواج الأثر (مثل معالجة الشغور ورفع الإيجار)", where))
    return blocking, warnings


def c_scenarios(ctx: Ctx):
    ctx.need("narrative")
    return scenario_findings(ctx)[0]


def _scan_pattern_strings(obj, label, levels=("block",)):
    """Personal-data patterns (assets/pii_patterns.json, the given levels) in the strings of a JSON-like object.
    Raw text of numeric TVs, hashes and file paths are skipped. Returns [(detail, where)]; values are never printed."""
    out = []

    def walk(o, path):
        if isinstance(o, dict):
            numeric_tv = "s" in o and "v" in o and is_num(o.get("v"))
            for k, v in o.items():
                if numeric_tv and k == "raw":
                    continue
                if k in ("sha256", "sha256_short", "path"):
                    continue
                walk(v, f"{path}.{k}" if path else str(k))
        elif isinstance(o, list):
            for i, v in enumerate(o):
                walk(v, f"{path}[{i}]")
        elif isinstance(o, str):
            if HEX_RE.match(o.strip()):
                return
            kinds = pii_value_hits(o, levels)
            if kinds:
                out.append((f"نص يشبه {pii_kinds_ar(kinds)} (القيمة لا تُطبع)", f"{label}:{path}"))
    walk(obj, "")
    return out


def c_pii_text_patterns(ctx: Ctx):
    ctx.need("narrative")
    out = _scan_pattern_strings(ctx.narrative, "06_narrative.json")
    if "records" in ctx.data:
        out += _scan_pattern_strings({"records": ctx.records}, "03_records.json")
    if "mapping" in ctx.data:  # mapping prose (notes, reasons, evidence) is shown in the dashboard's workbook section
        out += _scan_pattern_strings(ctx.mapping, "02_mapping.json")
    msg = g4b_message_text(ctx)
    if msg is not None:
        out += g4b_scan_lines(msg, MESSAGE_NAME, ("block",))
    return out


def pii_field_names(ctx: Ctx) -> set:
    names = set()
    for src in as_list(ctx.mapping.get("sources")):
        for col, spec in as_dict(as_dict(src).get("columns")).items():
            if isinstance(spec, dict) and spec.get("kind") == "pii" and isinstance(spec.get("field"), str):
                names.add(spec["field"])
    return names


def embedded_data(html_raw: str):
    if not html_raw:
        return None
    m = re.search(r"const\s+DATA\s*=\s*", html_raw)
    if not m:
        return None
    try:
        obj, _ = json.JSONDecoder().raw_decode(html_raw, m.end())
        return obj
    except Exception:
        return None


def c_pii_keys(ctx: Ctx):
    # Post-build check only: a stale analysis.json must not block the rebuild that replaces it.
    if not ctx.html_path:
        raise Skip("لم يُمرَّر --html (يُفحص analysis.json والبيانات المضمنة بعد البناء)")
    targets = []
    if ctx.analysis is not None:
        targets.append(("analysis.json", ctx.analysis))
    data = embedded_data(ctx.html_raw)
    if data is not None:
        targets.append((f"{os.path.basename(ctx.html_path)} (DATA)", data))
    if not targets:
        raise Skip("لا يوجد analysis.json ولا بيانات مضمنة في الداشبورد بعد")
    names = pii_field_names(ctx)
    out = []
    for label, obj in targets:
        hits = [p for p, k in walk_keys(obj) if isinstance(k, str)
                and (k in names or PII_KEY_RE.search(k) or SHARED_PII_KEY_RE.search(k))]
        for p in hits[:20]:
            out.append(("مفتاح يدل على حقل بيانات شخصية؛ يجب حذفه من الحزمة المضمنة", f"{label}:{p}"))
        if len(hits) > 20:
            out.append((f"و{len(hits) - 20} مفتاحًا آخر", label))
    return out


def visible_text(html_raw: str) -> str:
    t = re.sub(r"(?is)<(script|style|template)\b.*?</\1\s*>", " ", html_raw)
    t = re.sub(r"(?s)<!--.*?-->", " ", t)
    t = re.sub(r"(?s)<[^>]+>", " ", t)
    return _html.unescape(t)


def c_pii_html_patterns(ctx: Ctx):
    if not ctx.html_path:
        raise Skip("لم يُمرَّر --html")
    if ctx.html_raw is None:
        raise Skip("ملف الداشبورد غير موجود")
    name = os.path.basename(ctx.html_path)
    out = g4b_scan_lines(visible_text(ctx.html_raw), f"{name}: نص ظاهر", ("block",))
    data = embedded_data(ctx.html_raw)
    if data is not None:
        out += _scan_pattern_strings(data, f"{name} (DATA)")
    elif "const DATA" not in ctx.html_raw:
        out.append(("لم يُعثر على البيانات المضمنة (const DATA = …) لفحصها", name))
    if ctx.analysis is not None:  # post-build only (--html), like pii_keys: a stale bundle must not block the rebuild
        out += _scan_pattern_strings(ctx.analysis, "analysis.json")
    return out


# --- PII values from the workbook (memory only) ------------------------------
def _pii_columns(ctx: Ctx) -> list:
    cols = OrderedDict()
    hdr_rows = {}
    for s in as_list(ctx.wb.get("sheets")):
        if isinstance(s, dict):
            hr = s.get("primary_header_row")
            if not isinstance(hr, int):
                cands = as_list(s.get("header_candidates"))
                hr = cands[0].get("row") if cands and isinstance(cands[0], dict) else None
            hdr_rows[s.get("name")] = hr if isinstance(hr, int) else None
    for p in as_list(ctx.wb.get("pii_columns")):
        if isinstance(p, dict) and p.get("sheet") and p.get("col"):
            cols[(p["sheet"], p["col"])] = {"header": p.get("header"), "header_row": hdr_rows.get(p["sheet"])}
    for src in as_list(ctx.mapping.get("sources")):
        src = as_dict(src)
        for col, spec in as_dict(src.get("columns")).items():
            if isinstance(spec, dict) and spec.get("kind") == "pii" and src.get("sheet"):
                key = (src["sheet"], col)
                info = cols.setdefault(key, {"header": None, "header_row": None})
                if isinstance(src.get("header_row"), int):
                    info["header_row"] = src["header_row"]
    # A column the analyst explicitly cleared with "pii_ok": true (e.g. a notes column whose garbled header
    # mentions «المالك») is not a personal-data column; its values are still pattern-scanned elsewhere.
    for src in as_list(ctx.mapping.get("sources")):
        src = as_dict(src)
        for col, spec in as_dict(src.get("columns")).items():
            if isinstance(spec, dict) and spec.get("pii_ok") is True and spec.get("kind") != "pii":
                cols.pop((src.get("sheet"), col), None)
    return [(sh, col, info) for (sh, col), info in cols.items()]


def _value_patterns(v, header):
    """Compiled patterns for one PII cell value. The value itself is never returned or stored elsewhere."""
    if v is None or isinstance(v, bool):
        return []
    if isinstance(v, float) and v.is_integer():
        s = str(int(v))
    elif isinstance(v, (int, float)):
        s = str(v)
    elif isinstance(v, (_dt.date, _dt.datetime)):
        return []
    else:
        s = clean_text(v)
    s = s.translate(_AR_INDIC)
    if not s:
        return []
    ns = normalize_arabic(s)
    if ns in GENERIC_NON_PII or (header and ns == normalize_arabic(header)) or PII_HEADER_RE.fullmatch(s.strip()):
        return []
    pats = []
    # Digit runs may be written with the separators of assets/pii_patterns.json («(050) 123-4567», «050.123.4567»).
    sep = PII_PATTERNS.get("digit_separators") or r"[\s\-]"

    def digit_cores(rx):
        out = []
        for m in re.finditer(rx, s):
            d = re.sub(r"\D", "", m.group(0))
            out.append(re.sub(r"^(00966|966|0)", "", d) if len(d) >= 9 else d)
        return out
    # a run with any separators is one phone/ID unless it is longer than one (two numbers joined by « / »)
    wide = [c for c in digit_cores(r"\+?\d(?:" + sep + r"{0,3}\d)*") if len(c) <= 12]
    # parts of such a run («123» of «(050) 123-4567») would match unrelated numbers in the outputs
    cores = list(dict.fromkeys(wide + [c for c in digit_cores(r"\+?\d[\d\s\-]*\d")
                                       if not any(c in w for w in wide if w != c)]))
    for core in cores:
        if len(core) >= 7:  # phone / ID-like: any separators between digits, optional +966 / 0 prefix
            body = (sep + "{0,3}").join(core)
            pats.append(("digits", re.compile(r"(?<!\d)(?:(?:\+|00)\s?966" + sep + r"{0,3}|0)?" + sep + "{0,3}"
                                              + body + r"(?!\d)")))
        elif len(core) >= 4:
            body = r"[\s\-]?".join(core)
            pats.append(("digits", re.compile(r"(?<!\d)(?:\+?966|00966|0)?[\s\-]?" + body + r"(?!\d)")))
    letters = re.sub(r"\+?\d(?:" + sep + r"{0,3}\d)*", " ", s)
    for part in re.split(r"[/,،;|\n]+|\s-\s", letters):
        n = normalize_arabic(part)
        n = " ".join(re.sub(r"[^\w\s]", " ", n).split())
        if len(n.replace(" ", "")) < 4 or n in GENERIC_NON_PII or not re.search(r"[^\W\d_]", n):
            continue
        pats.append(("text", re.compile(r"(?<!\w)" + r"\s+".join(re.escape(t) for t in n.split()) + r"(?!\w)")))
    return pats


def _decode_u(s: str) -> str:
    return re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), s)


def c_pii_values(ctx: Ctx):
    if not ctx.xlsx:
        raise Skip("لم يُمرَّر --xlsx؛ فحص قيم الأعمدة الشخصية لم يُنفذ")
    if not os.path.isfile(ctx.xlsx):
        return [("ملف Excel المحدد في --xlsx غير موجود", ctx.xlsx)]
    cols = _pii_columns(ctx)
    if not cols:
        raise Skip("لا توجد أعمدة بيانات شخصية مسجلة في 01 أو 02")
    targets = []
    if ctx.html_raw is not None:
        targets.append((os.path.basename(ctx.html_path), ctx.html_raw))
    names = ("analysis.json", "06_narrative.json", "03_records.json") if ctx.html_path else \
        ("06_narrative.json", "03_records.json")  # analysis.json only after a build (see c_pii_keys)
    for name in names + ("02_mapping.json", MESSAGE_NAME):
        p = os.path.join(ctx.run_dir, name)
        if os.path.isfile(p):
            with open(p, "r", encoding="utf-8", errors="replace") as fh:
                targets.append((name, fh.read()))
    exp = g4b_export_cells(ctx)
    if exp and exp.get("cells"):
        targets.append((exp["label"], "\n".join(c[2] for c in exp["cells"])))
    if not targets:
        raise Skip("لا توجد مخرجات لفحصها")
    prepared = []
    for label, raw in targets:
        t = _html.unescape(raw)
        dec = _decode_u(t)
        if dec != t:
            t = t + "\n" + dec
        prepared.append((label, t.translate(_AR_INDIC), normalize_arabic(t)))
    try:
        import warnings
        import openpyxl
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")  # e.g. "Data Validation extension is not supported"
            wb = openpyxl.load_workbook(ctx.xlsx, read_only=True, data_only=True)
    except Exception as exc:
        return [(f"تعذر فتح ملف Excel للقراءة: {exc}", ctx.xlsx)]
    out = []
    n_values = 0
    try:
        names = {clean_text(n): n for n in wb.sheetnames}
        for sheet, col, info in cols:
            real = sheet if sheet in wb.sheetnames else names.get(clean_text(sheet))
            if real is None or not re.fullmatch(r"[A-Z]{1,3}", str(col)):
                continue
            ws = wb[real]
            ci = col_to_idx(col)
            hr = info.get("header_row") or 0
            for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row or 1, min_col=ci, max_col=ci,
                                                     values_only=True), start=1):
                if r_idx <= hr or not row:
                    continue
                pats = _value_patterns(row[0], info.get("header"))
                if pats:
                    n_values += 1
                for kind, pat in pats:
                    for label, digits_txt, norm_txt in prepared:
                        if pat.search(digits_txt if kind == "digits" else norm_txt):
                            out.append((f"قيمة من عمود بيانات شخصية ظهرت في المخرجات ({'رقم' if kind == 'digits' else 'نص'}؛ "
                                        "القيمة لا تُطبع)", f"{label} ← مصدرها {sheet}!{col}{r_idx}"))
    finally:
        wb.close()
    if not n_values:
        raise Skip("أعمدة البيانات الشخصية لا تحوي قيمًا قابلة للفحص")
    return out


# ---------------------------------------------------------------------------
# --- data integrity & privacy checks (G4b) ---
# Scope and double counting (units, totals, duplicate ids, overlapping rows, sheet decisions), controlled
# vocabularies, study context in 02, reconciliation of recorded totals, {{f:…}} fact tokens, personal-data
# patterns over every delivered output (09 message, supporting workbook), and acknowledged warnings.
# Registered through G4B_CHECKS (one line in CHECKS) and G4B_TITLES.
# ---------------------------------------------------------------------------
from apa_common import (  # noqa: E402
    PII_KEY_RE as SHARED_PII_KEY_RE, PII_PATTERNS, RECORD_TYPES, SCOPES, is_total_label, load_vocab, mask_pii,
    name_key, pii_value_hits,
)
from facts import FACT_TOKEN_RE, compute_facts  # noqa: E402

MESSAGE_NAME = "09_message.md"
EXPORT_NAME = "supporting_workbook.xlsx"
ASSET_LIKE = ("asset", "land_parcel")
CHILD_TYPES = ("unit", "contract")
VOCAB_DIMS = ("use", "income_nature", "waqf")          # single code each
ACK_MIN_WORDS = 8
PII_KIND_AR = {"mobile": "رقم جوال", "landline": "رقم هاتف ثابت", "email": "بريد إلكتروني",
               "iban": "رقم حساب بنكي (IBAN)", "national_id": "رقم من 10 خانات يبدأ بـ1 أو 2 (قد يكون هوية وطنية)"}
RECON_STATUS_AR = {
    "match": "مطابق لمجموع السجلات ضمن النطاق",
    "match_section": "مطابق لمجموع سجلات قسمه ضمن النطاق (إجمالي فرعي)",
    "match_all_records": "يطابق مجموع كل سجلات المصدر؛ الفرق يعود إلى سجلات معزولة أو مستبعدة",
    "match_section_all_records": "يطابق مجموع سجلات قسمه كلها؛ الفرق يعود إلى سجلات معزولة أو مستبعدة",
    "match_without_text_numbers": "يطابق المجموع دون أرقام مخزنة كنص في الملف (لا تجمعها صيغ Excel)؛ الإجمالي المدون ناقص",
    "mismatch": "فرق غير مفسَّر",
    "recorded_unknown": "القيمة المدونة في صف الإجمالي غير معروفة (فارغة أو نص أو خطأ)",
}
LEVEL_AR["acknowledged"] = "مُقرّ به"


def pii_kinds_ar(kinds) -> str:
    return " أو ".join(PII_KIND_AR.get(k, k) for k in kinds)


def _words_ar(n: int) -> str:
    return {0: "دون كلمات", 1: "كلمة واحدة", 2: "كلمتان"}.get(n, f"{n} كلمات" if 3 <= n <= 10 else f"{n} كلمة")


RECORDS_AR = ("سجل واحد", "سجلان", "سجلات", "سجلًا")
CELLS_AR = ("خلية واحدة", "خليتان", "خلايا", "خلية")
PLACES_AR = ("موضع واحد", "موضعان", "مواضع", "موضعًا")


def _n_ar(n: int, forms) -> str:
    """Arabic count phrase: «سجل واحد» / «سجلان» / «3 سجلات» / «12 سجلًا»."""
    one, two, few, many = forms
    return one if n == 1 else two if n == 2 else f"{n} {few if 3 <= n <= 10 else many}"


def _n_values_ar(n: int, known: bool) -> str:
    """«قيمة واحدة معروفة» / «قيمتان غير معروفتين» / «3 قيم معروفة» / «12 قيمة غير معروفة»."""
    adj1, adj2 = ("معروفة", "معروفتان") if known else ("غير معروفة", "غير معروفتين")
    if n == 1:
        return f"قيمة واحدة {adj1}"
    if n == 2:
        return f"قيمتان {adj2}"
    return f"{n} {'قيم' if 3 <= n <= 10 else 'قيمة'} {adj1}"


def _in_scope(r) -> bool:
    return isinstance(r, dict) and r.get("scope") == "in_scope"


def _src(r) -> dict:
    return as_dict(r.get("source"))


def _lookup_sheet(d: dict, sheet):
    """Value for a sheet name: exact, else equal after trimming spaces (names may carry trailing spaces)."""
    if sheet in d:
        return d[sheet]
    c = clean_text(sheet)
    for k, v in d.items():
        if clean_text(k) == c:
            return v
    return None


def _issues_by_type(ctx: Ctx, *types) -> list:
    """[(issue, record_ids)] for 03 issues of the given types; record ids from the issue and from records' issue_ids."""
    linked = defaultdict(list)
    for r in ctx.records:
        for iid in as_list(r.get("issue_ids")):
            linked[iid].append(r.get("id"))
    out = []
    for i in as_list(ctx.rec_doc.get("issues")):
        if isinstance(i, dict) and i.get("type") in types:
            rids = [x for x in as_list(i.get("record_ids")) if isinstance(x, str)]
            rids += [x for x in linked.get(i.get("id"), []) if isinstance(x, str) and x not in rids]
            out.append((i, rids))
    return out


def _issue_where(i: dict) -> str:
    return f"03_records.json:issues[{i.get('id')}] ({i.get('where') or '-'})"


def _evidenced_override(ctx: Ctx, rid) -> str:
    """The reason when 02 records/<rid> deliberately sets record_type or scope AND explains it (scope_reason or notes)."""
    o = as_dict(as_dict(ctx.mapping.get("records")).get(rid))
    if not ("record_type" in o or "scope" in o):
        return ""
    for k in ("scope_reason", "notes"):
        if nonempty(o.get(k)):
            return clean_text(o[k])
    return ""


def _amounts(r: dict, meta: dict) -> dict:
    """Known, non-zero revenue / valuation amounts of a record: {field: v}."""
    out = {}
    for f, tv in as_dict(r.get("fields")).items():
        if field_group(f, meta) in ("revenue", "valuation") and _known(r, f) and as_dict(tv).get("v") != 0:
            out[f] = as_dict(tv).get("v")
    return out


def _deed_keys(r: dict) -> set:
    keys = set()
    for f, tv in as_dict(r.get("fields")).items():
        if not (f == "deed_no" or f.startswith("deed_no.")):
            continue
        tv = as_dict(tv)
        v = tv.get("v")
        if tv.get("s") not in KNOWN_STATUSES or v in (None, ""):
            continue
        s = str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)
        k = re.sub(r"[\s()\[\]{}]", "", normalize_arabic(s))
        if k and re.search(r"\d", k) and not re.fullmatch(r"0+", k):
            keys.add(k)
    return keys


# --- A3: total rows kept in scope --------------------------------------------------------------------------------
def _row_id(r: dict) -> bool:
    """True when the record id was generated from its row number (no serial in the file), e.g. REG-r15."""
    row = _src(r).get("row")
    return isinstance(row, int) and bool(re.search(rf"(?:^|[-_.])r{row}(?:-\d+)?$", str(r.get("id") or "")))


def _sum_of_others(ctx: Ctx) -> dict:
    """{record id: {"match": [fields], "level": "strong"|"weak"}} — in-scope records whose known amount equals the sum of
    ≥ 2 other non-zero in-scope amounts of the same source and field (the arithmetic signature of a total row that no
    label or formula revealed). Round valuations make one such equality a plausible coincidence for a named asset, so:
    strong = equal in ≥ 2 fields, or in one field for a row without a name or without the serial its siblings carry;
    weak = equal in every testable field of the last row of its source. Other single-field equalities are not reported."""
    meta = ctx.field_meta
    groups = defaultdict(list)
    last_row, n_src, n_row_ids = {}, defaultdict(int), defaultdict(int)
    for r in ctx.records:
        if not _in_scope(r):
            continue
        sid, row = _src(r).get("source_id"), _src(r).get("row")
        n_src[sid] += 1
        n_row_ids[sid] += _row_id(r)
        if isinstance(row, int):
            last_row[sid] = max(last_row.get(sid, row), row)
        for f, tv in as_dict(r.get("fields")).items():
            if field_group(f, meta) in ("revenue", "valuation", "expense") and _known(r, f) \
                    and not as_dict(tv).get("joined_from"):
                groups[(sid, f)].append((r.get("id"), float(as_dict(tv)["v"])))
    tested, matched = defaultdict(list), defaultdict(list)
    for (_sid, f), items in groups.items():
        nonzero = [x for x in items if x[1] != 0]
        if len(nonzero) < 3:  # the record itself + at least two non-zero others
            continue
        tot = sum(v for _, v in nonzero)
        for rid, v in nonzero:
            if v > 0:
                tested[rid].append(f)
                if abs(v - (tot - v)) <= max(0.5, 1e-9 * v):
                    matched[rid].append(f)
    out = {}
    for rid, fs in matched.items():
        r = as_dict(ctx.rec_by_id.get(rid))
        sid = _src(r).get("source_id")
        no_serial = _row_id(r) and n_row_ids[sid] * 2 < n_src[sid]  # its siblings mostly carry a serial
        if len(fs) >= 2 or not nonempty(r.get("name")) or no_serial:
            out[rid] = {"match": fs, "level": "strong"}
        elif len(fs) == len(tested[rid]) and _src(r).get("row") == last_row.get(sid):
            out[rid] = {"match": fs, "level": "weak"}
    return out


def g4b_total_like(ctx: Ctx):
    """(blocking, warnings): in-scope records flagged total_like_row, or equal to the sum of their siblings."""
    ctx.need("records")
    recs = ctx.rec_by_id
    sums = _sum_of_others(ctx)
    blocking, warnings, seen = [], [], set()
    for iss, rids in _issues_by_type(ctx, "total_like_row"):
        for rid in rids:
            r = recs.get(rid)
            if not _in_scope(r) or r.get("record_type") == "total" or rid in seen:
                continue  # record_type total in scope is reported by total_rows_in_scope
            seen.add(rid)
            where = f"03_records.json:{rid} ← {iss.get('id')} ({iss.get('where') or '-'})"
            reason = _evidenced_override(ctx, rid)
            if reason and rid not in sums:
                warnings.append((f"السجل {rid} يشبه صف إجمالي ({iss.get('id')}) وأُبقي ضمن النطاق بتجاوز موثق في 02 "
                                 f"(«{reason[:160]}»)؛ تأكد أنه ليس مجموعًا لسجلات أخرى", where))
            elif reason:
                blocking.append((f"السجل {rid} يشبه صف إجمالي ({iss.get('id')}) وقيمته في "
                                 f"{short_list(sums[rid]['match'], 3)} تساوي مجموع بقية سجلات مصدره ضمن النطاق، فالتجاوز "
                                 "الموثق في 02 لا يكفي؛ استبعده (exclude_rows) أو اجعل نوعه total", where))
            else:
                blocking.append((f"السجل {rid} ضمن النطاق رغم أن {iss.get('id')} يصفه صف إجمالي؛ سيُجمع مع تفاصيله. "
                                 "أضف الصف إلى exclude_rows أو اجعل نوعه total، أو أثبت أنه ليس إجماليًا بـ records/<id> "
                                 "في 02 (record_type مع scope_reason أو notes)", where))
    for r in ctx.records:  # a total label («الإجمالي»، «جملة»، «Total») in any text cell, whatever 03's issues say
        rid = r.get("id")
        if not _in_scope(r) or r.get("record_type") == "total" or rid in seen:
            continue
        cells = [("name", _src(r).get("ref_name"))] if is_total_label(r.get("name")) else []
        cells += [(f, as_dict(tv).get("ref")) for f, tv in as_dict(r.get("fields")).items()
                  if is_total_label(as_dict(tv).get("v"))]
        cells += [(f"dims.{k}", None) for k, v in as_dict(r.get("dims")).items() if is_total_label(v)]
        if not cells:
            continue
        seen.add(rid)
        at = short_list([f"{f} ({ref})" if ref else f for f, ref in cells], 3)
        reason = _evidenced_override(ctx, rid)
        if reason:
            warnings.append((f"السجل {rid} يحمل عبارة إجمالي في {at} وأُبقي ضمن النطاق بتجاوز موثق في 02 "
                             f"(«{reason[:160]}»)؛ تأكد أنه ليس مجموعًا لسجلات أخرى", f"03_records.json:{rid}"))
        else:
            blocking.append((f"السجل {rid} ضمن النطاق ويحمل عبارة إجمالي (مثل «الإجمالي» أو «جملة» أو Total) في {at}؛ "
                             "سيُجمع مع تفاصيله. أضف الصف إلى exclude_rows، أو أثبت أنه ليس إجماليًا بـ records/<id> في 02 "
                             "(record_type مع scope_reason أو notes)", f"03_records.json:{rid}"))
    for rid, hit in sums.items():
        if rid in seen:
            continue
        reason = _evidenced_override(ctx, rid)
        msg = (f"قيمة السجل {rid} في {short_list(hit['match'], 3)} تساوي مجموع بقية سجلات مصدره ضمن النطاق، وهذه سمة صف "
               "إجمالي لم يرصده اسم أو صيغة")
        if reason:
            warnings.append((f"{msg}؛ أُبقي ضمن النطاق بتجاوز موثق في 02 («{reason[:160]}»)", f"03_records.json:{rid}"))
        elif hit["level"] == "strong":
            blocking.append((f"{msg}؛ سيُجمع مع تفاصيله. استبعده (exclude_rows) إن كان إجماليًا، أو أثبت أنه أصل مستقل بـ "
                             "records/<id> في 02 (record_type مع scope_reason أو notes)", f"03_records.json:{rid}"))
        else:
            warnings.append((f"{msg}، وهو آخر صف في مصدره؛ قد يكون التساوي مصادفة لأن للسجل اسمًا، فتحقق منه واستبعده "
                             "(exclude_rows) إن كان إجماليًا", f"03_records.json:{rid}"))
    return blocking, warnings


# --- A2: an asset and its units / contracts counted together -----------------------------------------------------
def _parent_of(child: dict, parents: list):
    """(parent record, basis) for a unit/contract record: id branch (REG-1 → REG-1.2), same deed, or a name that starts
    with the parent's name (≥ 2 words). (None, None) when nothing links it."""
    cid = str(child.get("id") or "")
    best = None
    for p in parents:
        pid = str(p.get("id") or "")
        if pid and len(cid) > len(pid) + 1 and cid.startswith(pid) and cid[len(pid)] in "./_-":
            if best is None or len(pid) > len(str(best.get("id"))):
                best = p
    if best is not None:
        return best, f"معرّفه متفرع من {best.get('id')}"
    dk = _deed_keys(child)
    if dk:
        for p in parents:
            if dk & _deed_keys(p):
                return p, f"رقم الصك نفسه في {p.get('id')}"
    ck = name_key(child.get("name") or "")
    cands = []
    for p in parents:
        pk = name_key(p.get("name") or "")
        if ck and len(pk.split()) >= 2 and (ck == pk or ck.startswith(pk + " ")):
            cands.append((len(pk), p))
    if cands:
        p = max(cands, key=lambda x: x[0])[1]
        return p, f"اسمه يبدأ باسم {p.get('id')}"
    return None, None


def g4b_unit_levels(ctx: Ctx):
    """(blocking, warnings) for in-scope unit/contract records that carry revenue or valuation amounts."""
    ctx.need("records")
    meta = ctx.field_meta
    ins = [r for r in ctx.records if _in_scope(r)]
    parents = [r for r in ins if r.get("record_type") in ASSET_LIKE]
    by_parent, unlinked = OrderedDict(), []
    for c in ins:
        if c.get("record_type") not in CHILD_TYPES:
            continue
        amounts = _amounts(c, meta)
        if not amounts:
            continue
        p, basis = _parent_of(c, parents)
        if p is None:
            unlinked.append(c.get("id"))
        else:
            by_parent.setdefault(p.get("id"), (p, []))[1].append((c, basis, amounts))
    blocking, warnings = [], []
    fix = ("اختر مستوى واحدًا للجمع: الأصل بمبلغه الكامل والوحدات excluded بسبب «تفصيل لأصل مسجل»، أو الوحدات وحدها "
           "بعد استبعاد مبلغ الأصل")
    for pid, (p, kids) in by_parent.items():
        p_amounts = _amounts(p, meta)
        hard, soft = [], []
        for c, basis, amounts in kids:
            desc = f"{c.get('id')} ({c.get('record_type')}؛ {basis})"
            (soft if _evidenced_override(ctx, c.get("id")) else hard).append((desc, set(amounts)))
        pname = f" «{p.get('name')}»" if p.get("name") else ""
        if hard:
            child_fields = set().union(*(fs for _, fs in hard))
            shared = sorted(child_fields & set(p_amounts))
            effect = (f"الحقول نفسها ({short_list(shared, 4)}) في المستويين، فيتضاعف المبلغ في أي مجموع يجمعهما" if shared
                      else f"مبالغ ({short_list(sorted(child_fields), 4)}) لا يحملها الأصل، ومؤشرات الإيراد والقيمة تُحتسب من "
                           "مستوى الأصل، فيسقط مبلغ الوحدات منها")
            blocking.append((f"الأصل {pid}{pname} ووحدات أو عقود تابعة له ضمن النطاق تحمل مبالغ إيراد أو تقييم: "
                             f"{short_list([d for d, _ in hard], 6)}؛ {effect}. {fix}، أو أثبت في 02 أنها ليست جزءًا "
                             "منه (records/<id> مع scope_reason أو notes)", f"03_records.json:{pid}"))
        if soft:
            warnings.append((f"الأصل {pid}{pname} ومعه ضمن النطاق وحدات أو عقود بمبالغ إيراد أو تقييم أُبقيت بتجاوز موثق "
                             f"في 02: {short_list([d for d, _ in soft], 6)}؛ تأكد أنها ليست جزءًا منه فلا يُحتسب المبلغ "
                             "مرتين", f"03_records.json:{pid}"))
    if unlinked:
        warnings.append((f"وحدات أو عقود ضمن النطاق ({len(unlinked)}) بمبالغ إيراد أو تقييم دون أصل أم محدد (لا معرّف "
                         f"متفرع ولا صك ولا اسم مشترك): {short_list(unlinked, 8)}؛ مؤشرات الإيراد والقيمة لا تقرأ مستوى "
                         "الوحدة أو العقد. إن كانت تتبع أصلًا مسجلًا فاجعلها excluded بسبب «تفصيل لأصل مسجل»، وإن كانت "
                         "أصلًا مستقلًا فاجعل نوعها asset",
                         "03_records.json:records[record_type∈unit|contract]"))
    return blocking, warnings


# --- A11: duplicate ids kept in scope ------------------------------------------------------------------------------
def g4b_duplicate_ids(ctx: Ctx):
    """(blocking, warnings) for duplicate_id issues whose records are both still in scope."""
    ctx.need("records")
    recs = ctx.rec_by_id
    blocking, warnings = [], []
    for iss, rids in _issues_by_type(ctx, "duplicate_id"):
        ins = [x for x in rids if _in_scope(recs.get(x))]
        if len(ins) < 2:
            continue
        reasons = [_evidenced_override(ctx, x) for x in ins[1:]]
        if all(reasons):
            warnings.append((f"{iss.get('id')}: المعرّف مكرر في {short_list(ins)} وأُبقيت السجلات ضمن النطاق بتجاوز موثق "
                             f"في 02 («{reasons[0][:160]}»)؛ تأكد من اختلاف الصك والموقع", _issue_where(iss)))
        else:
            blocking.append((f"{iss.get('id')}: المعرّف مكرر والسجلات {short_list(ins)} كلها ضمن النطاق؛ قد يُحتسب الأصل "
                             "مرتين في العدد والقيمة والإيراد. استبعد المكرر في 02 (records/<id>: scope = excluded مع "
                             "السبب)، أو أثبت أنهما أصلان مختلفان (records/<id>: scope = in_scope مع scope_reason يذكر الصك "
                             "أو الموقع)", _issue_where(iss)))
    return blocking, warnings


# --- A10: the same sheet row read into several records ----------------------------------------------------------
def g4b_overlaps(ctx: Ctx):
    """(blocking, warnings). Blocking when ≥ 2 in-scope records read the same cell, or are two assets from one row."""
    ctx.need("records")
    issue_of = {}
    for iss, rids in _issues_by_type(ctx, "overlapping_source_rows"):
        for x in rids:
            issue_of.setdefault(x, iss.get("id"))
    groups = defaultdict(list)
    for r in ctx.records:
        s = _src(r)
        if isinstance(s.get("sheet"), str) and isinstance(s.get("row"), int):
            groups[(clean_text(s["sheet"]), s["row"])].append(r)
    hard, soft = defaultdict(list), defaultdict(list)
    for (sheet, row), rs in sorted(groups.items(), key=lambda kv: (kv[0][0], kv[0][1])):
        if len(rs) < 2:
            continue
        ins = [r for r in rs if _in_scope(r)]
        cell_owners = defaultdict(set)
        for r in ins:
            for f, tv in as_dict(r.get("fields")).items():
                ref = as_dict(tv).get("ref")
                if isinstance(ref, str) and _known(r, f) and not as_dict(tv).get("joined_from"):
                    cell_owners[ref].add(r.get("id"))
        same_cells = sorted(ref for ref, ids in cell_owners.items() if len(ids) > 1)
        n_assets = sum(1 for r in ins if r.get("record_type") in ASSET_LIKE)
        item = (row, [r.get("id") for r in rs], same_cells)
        exact = _src(rs[0]).get("sheet")  # the exact name (may carry trailing spaces) for the reference
        (hard if len(ins) >= 2 and (same_cells or n_assets >= 2) else soft)[exact].append(item)
    blocking, warnings = [], []
    for sheet, items in hard.items():
        rows = short_list([r for r, _, _ in items], 8)
        ids = [i for _, rids, _ in items for i in rids]
        cells = [c for _, _, cs in items for c in cs]
        iss = sorted({issue_of[i] for i in ids if i in issue_of})
        blocking.append((f"الصفوف {rows} من الورقة «{sheet}» مقروءة في أكثر من سجل ضمن النطاق ({short_list(ids, 8)})"
                         + (f"، والخلايا نفسها مكررة ({short_list(cells, 4)})" if cells else "")
                         + "؛ يتضاعف العدد أو القيمة أو الإيراد. اجعل نطاقات data_rows للمصادر غير متداخلة أو استبعد "
                         "المكرر" + (f" ({short_list(iss, 4)})" if iss else ""), f"{sheet}!{rows}"))
    for sheet, items in soft.items():
        rows = short_list([r for r, _, _ in items], 8)
        ids = [i for _, rids, _ in items for i in rids]
        warnings.append((f"الصفوف {rows} من الورقة «{sheet}» مقروءة في أكثر من سجل ({short_list(ids, 8)}) دون خلية "
                         "معروفة مكررة ضمن النطاق؛ تحقق أن ذلك مقصود (مثل فصل نشاط عن أصل في الصف نفسه)",
                         f"{sheet}!{rows}"))
    return blocking, warnings


# --- A6: controlled vocabularies ---------------------------------------------------------------------------------
def c_dims_vocabulary(ctx: Ctx):
    ctx.need("records")
    voc = load_vocab()
    allowed = {d: list(as_dict(voc.get(d)).keys()) for d in VOCAB_DIMS + ("ownership",)}
    recs = ctx.rec_by_id
    out = []
    for iss, rids in _issues_by_type(ctx, "dim_out_of_vocabulary"):
        ins = [x for x in rids if _in_scope(recs.get(x))]
        if ins:
            out.append((f"{iss.get('id')}: {clean_text(iss.get('detail'))[:240]} — منها {_n_ar(len(ins), RECORDS_AR)} "
                        f"ضمن النطاق: {short_list(ins, 8)}؛ صنّفها بالرموز المعتمدة عبر dim_rules أو records بدليل",
                        _issue_where(iss)))
    bad = OrderedDict()

    def add(key, rid):
        bad.setdefault(key, []).append(rid)
    for r in ctx.records:
        rid = r.get("id")
        if r.get("record_type") not in RECORD_TYPES:
            add(("record_type", repr(r.get("record_type"))), rid)
        if r.get("scope") not in SCOPES:
            add(("scope", repr(r.get("scope"))), rid)
        if not _in_scope(r):
            continue
        d = as_dict(r.get("dims"))
        for dim in VOCAB_DIMS:
            v = d.get(dim)
            if v not in (None, "") and v not in allowed[dim]:
                add((dim, f"«{v}»" if isinstance(v, str) else repr(v)), rid)
        own = d.get("ownership")
        if own not in (None, ""):
            for v in (own if isinstance(own, list) else [own]):
                if v not in allowed["ownership"]:
                    add(("ownership", f"«{v}»" if isinstance(v, str) else repr(v)), rid)
            if not isinstance(own, list):
                add(("ownership", None), rid)
        m = d.get("managed_for_others")
        if m is not None and not isinstance(m, bool):
            add(("managed_for_others", repr(m)), rid)
    allowed_txt = {"record_type": RECORD_TYPES, "scope": SCOPES, "managed_for_others": ("true", "false", "null")}
    effect = {"record_type": "لا يُعرف هل يُعد أصلًا أو وحدة أو إجماليًا",
              "scope": "لا يدخل المؤشرات ولا قسم المعزول والمستبعد"}
    for (dim, shown), ids in bad.items():
        where = f"03_records.json:records[].{'' if dim in effect else 'dims.'}{dim}"
        if shown is None:
            out.append((f"ownership ليست قائمة في {_n_ar(len(ids), RECORDS_AR)} ضمن النطاق: {short_list(ids, 8)}؛ "
                        "الملكية قائمة غير حصرية (مثل [\"owned\"])", where))
            continue
        ok = allowed_txt.get(dim) or allowed.get(dim) or ()
        out.append((f"القيمة {shown} في {dim} خارج المفردات المعتمدة (assets/vocab.json) في {_n_ar(len(ids), RECORDS_AR)}"
                    f"{'' if dim in effect else ' ضمن النطاق'}: {short_list(ids, 8)}؛ المسموح: {'، '.join(map(str, ok))}. "
                    f"{effect.get(dim, 'المؤشرات المقيدة بهذا البعد لا تراها')}", where))
    return out


# --- A9: sheet decisions, managed-for-others sheets, hidden sheets ---------------------------------------------------
def g4b_sheet_scope(ctx: Ctx):
    """(blocking, warnings)."""
    ctx.need("records")
    blocking, warnings, covered = [], [], set()
    for iss, rids in _issues_by_type(ctx, "sheet_scope_mismatch", "managed_unflagged"):
        blocking.append((f"{iss.get('id')} ({iss.get('type')}): {clean_text(iss.get('detail'))[:320]}", _issue_where(iss)))
        covered.update(rids)
    decisions = {s.get("sheet"): s for s in as_list(ctx.mapping.get("sheets"))
                 if isinstance(s, dict) and isinstance(s.get("sheet"), str)}
    overrides = as_dict(ctx.mapping.get("records"))
    wrong, unflagged = OrderedDict(), OrderedDict()
    for r in ctx.records:
        rid = r.get("id")
        if not _in_scope(r) or rid in covered:
            continue
        sheet = _src(r).get("sheet")
        s = as_dict(_lookup_sheet(decisions, sheet)) if decisions and isinstance(sheet, str) else {}
        if s.get("decision") in ("isolated", "excluded"):
            o = as_dict(overrides.get(rid))
            if not (o.get("scope") == "in_scope" and nonempty(o.get("scope_reason"))):
                wrong.setdefault((sheet, s.get("decision")), []).append(rid)
        if s.get("role") == "managed_properties" and as_dict(r.get("dims")).get("managed_for_others") is not True:
            unflagged.setdefault(sheet, []).append(rid)
    for (sheet, dec), ids in wrong.items():
        blocking.append((f"{_n_ar(len(ids), RECORDS_AR)} ضمن النطاق من الورقة «{sheet}» وقرارها في 02 «{dec}»: "
                         f"{short_list(ids, 8)}؛ بيانات لم يتحقق انتماؤها للمحفظة تدخل المؤشرات. اجعل default_scope للمصدر "
                         "موافقًا لقرار الورقة، أو أدرج سجلًا بعينه بـ records/<id>: scope = in_scope مع scope_reason يذكر "
                         "الرابط المؤكد", f"02_mapping.json:sheets[{sheet}]"))
    for sheet, ids in unflagged.items():
        blocking.append((f"{_n_ar(len(ids), RECORDS_AR)} ضمن النطاق من ورقة الأملاك المدارة للغير «{sheet}» دون "
                         f"managed_for_others = true: {short_list(ids, 8)}؛ تدخل قيمة عقارات الغير وكامل إيراداتها في "
                         "أرقام الجمعية. صنّفها managed_for_others: true أو اعزل المصدر",
                         f"02_mapping.json:sheets[{sheet}]"))
    states = {s.get("name"): s.get("state") for s in as_list(ctx.wb.get("sheets")) if isinstance(s, dict)}
    hidden = OrderedDict()
    for r in ctx.records:
        sheet = _src(r).get("sheet")
        if _in_scope(r) and isinstance(sheet, str) and _lookup_sheet(states, sheet) in ("hidden", "veryHidden"):
            hidden.setdefault((sheet, _lookup_sheet(states, sheet)), []).append(r.get("id"))
    for (sheet, st), ids in hidden.items():
        reason = clean_text(as_dict(_lookup_sheet(decisions, sheet)).get("reason")) if decisions else ""
        kind = "المخفية بشدة (veryHidden)" if st == "veryHidden" else "المخفية"
        warnings.append((f"{_n_ar(len(ids), RECORDS_AR)} ضمن النطاق من الورقة {kind} «{sheet}»: {short_list(ids, 8)}؛ "
                         "تحقق من انتماء بياناتها للمحفظة ومطابقتها للسجل الرئيسي"
                         + (f" (سبب القرار في 02: «{reason[:160]}»)" if reason else "، واذكر السبب في sheets[].reason"),
                         f"01_workbook_map.json:sheets[{sheet}]"))
    return blocking, warnings


# --- A13: study context and the sheet map in 02 --------------------------------------------------------------------
def c_mapping_context(ctx: Ctx):
    ctx.need("mapping")
    m = ctx.mapping
    out = []
    w = "02_mapping.json"
    labels = {"entity": "الجهة", "data_period": "فترة البيانات", "study_date": "تاريخ إعداد الدراسة"}
    for key, vkey in (("entity", "name"), ("data_period", "label"), ("study_date", "value")):
        o = m.get(key)
        if key not in m:
            out.append((f"«{key}» ({labels[key]}) مفقود؛ اذكر القيمة أو null مع الدليل أو سبب الغياب في evidence", f"{w}:{key}"))
        elif not isinstance(o, dict):
            out.append((f"«{key}» يجب أن يكون كائنًا {{{vkey}, evidence}} (القيمة قد تكون null)", f"{w}:{key}"))
        elif not nonempty(o.get("evidence")):
            out.append((f"«{key}» ({labels[key]}) بلا دليل (evidence): مرجع خلية أو خاصية ملف، أو سبب عدم توافر القيمة",
                        f"{w}:{key}.evidence"))
    if "currency" not in m:
        out.append(("«currency» (العملة) مفقود؛ اذكر العملة الواردة في الملف أو null", f"{w}:currency"))
    vd = m.get("valuation_dates")
    if "valuation_dates" not in m:
        out.append(("«valuation_dates» (تواريخ التقييم) مفقود؛ اذكر سنة كل تقييم وأساسه ودليله، أو قائمة فارغة إن لم يرد تقييم",
                    f"{w}:valuation_dates"))
    elif not isinstance(vd, list):
        out.append(("«valuation_dates» يجب أن يكون قائمة", f"{w}:valuation_dates"))
    else:
        for i, it in enumerate(vd):
            if not isinstance(it, dict) or not nonempty(it.get("evidence")):
                out.append(("تاريخ تقييم بلا دليل (evidence)", f"{w}:valuation_dates[{i}]"))
    sheets = m.get("sheets")
    if not isinstance(sheets, list) or not sheets:
        out.append(("«sheets» (خريطة الأوراق) مفقودة أو فارغة؛ لكل ورقة role وdecision وreason، بما فيها المخفية",
                    f"{w}:sheets"))
        return out
    listed = []
    for i, s in enumerate(sheets):
        if not isinstance(s, dict) or not isinstance(s.get("sheet"), str):
            out.append(("عنصر في sheets بلا اسم ورقة (sheet)", f"{w}:sheets[{i}]"))
            continue
        listed.append(s["sheet"])
        miss = [k for k in ("role", "decision") if not nonempty(s.get(k))]
        if miss:
            out.append((f"الورقة «{s['sheet']}» بلا {' و'.join(miss)}", f"{w}:sheets[{i}]"))
    wb_sheets = [s for s in as_list(ctx.wb.get("sheets")) if isinstance(s, dict) and isinstance(s.get("name"), str)]
    if wb_sheets:
        listed_c = {clean_text(n) for n in listed}
        state_ar = {"hidden": "مخفية", "veryHidden": "مخفية بشدة"}
        missing = [f"«{s['name']}»" + (f" ({state_ar[s.get('state')]})" if s.get("state") in state_ar else "")
                   for s in wb_sheets if clean_text(s["name"]) not in listed_c]
        if missing:
            out.append((f"أوراق في المصنف (01) غير مذكورة في sheets[]: {short_list(missing, 10)}؛ أضف لكل منها role وdecision "
                        "وreason، فالفحص يشمل كل الأوراق بما فيها المخفية", f"{w}:sheets"))
        wb_c = {clean_text(s["name"]) for s in wb_sheets}
        extra = [f"«{n}»" for n in listed if clean_text(n) not in wb_c]
        if extra:
            out.append((f"أوراق في sheets[] غير موجودة في المصنف (01): {short_list(extra, 10)}؛ صحّح الاسم كما هو في الملف",
                        f"{w}:sheets"))
    return out


# --- A14: recorded totals vs the sum of qualifying records -----------------------------------------------------------
def g4b_reconciliation(ctx: Ctx):
    """(rows for 07.reconciliation, warnings for unexplained differences). Cached on ctx."""
    if getattr(ctx, "_g4b_recon", None) is not None:
        return ctx._g4b_recon
    ctx.need("records")
    entries = [e for e in as_list(ctx.rec_doc.get("excluded_rows"))
               if isinstance(e, dict) and isinstance(e.get("values"), dict) and e["values"]]
    if not entries:
        raise Skip("لا توجد صفوف إجمالي مستبعدة بقيم محفوظة (excluded_rows[].values) للمطابقة")
    by_src = defaultdict(list)
    for r in ctx.records:
        by_src[_src(r).get("source_id")].append(r)
    total_rows = defaultdict(list)
    for e in entries:
        if isinstance(e.get("row"), int):
            total_rows[e.get("source_id")].append(e["row"])
    rows, warns = [], []
    for e in entries:
        sid, sheet, row = e.get("source_id"), e.get("sheet"), e.get("row")
        prev = [x for x in total_rows.get(sid, []) if isinstance(row, int) and x < row]
        lo = max(prev) if prev else None

        def agg(field, in_scope_only, section):
            """(sum of known values, n known, n unknown, sum of numbers that were stored as text in the file)."""
            s, n_known, n_unknown, s_text = 0.0, 0, 0, 0.0
            for r in by_src.get(sid, []):
                if in_scope_only and not _in_scope(r):
                    continue
                rr = _src(r).get("row")
                if section and not (isinstance(rr, int) and isinstance(row, int) and (lo is None or rr > lo) and rr < row):
                    continue
                tv = as_dict(as_dict(r.get("fields")).get(field))
                if tv.get("joined_from"):
                    continue
                if tv.get("s") in KNOWN_STATUSES and is_num(tv.get("v")):
                    s += tv["v"]
                    n_known += 1
                    if isinstance(tv.get("raw"), str) and str(tv.get("note") or "").startswith("نص رقمي"):
                        s_text += tv["v"]  # SUM/SUBTOTAL in Excel skip numbers stored as text
                else:
                    n_unknown += 1
            return round(s, 2), n_known, n_unknown, round(s_text, 2)
        for field, tv in e["values"].items():
            tv = tv if isinstance(tv, dict) else {"v": tv, "s": "value" if is_num(tv) else "text"}
            ref = tv.get("ref") or (f"{sheet}!{row}" if sheet else f"{sid}:{row}")
            item = OrderedDict([("source_id", sid), ("sheet", sheet), ("row", row), ("ref", ref), ("field", field)])
            if tv.get("s") not in KNOWN_STATUSES or not is_num(tv.get("v")):
                item.update(recorded=None, status="recorded_unknown", status_ar=RECON_STATUS_AR["recorded_unknown"])
                rows.append(item)
                continue
            recorded = float(tv["v"])
            tol = max(0.5, 1e-9 * abs(recorded))
            full, sec = agg(field, True, False), agg(field, True, True)
            all_rec, all_sec = agg(field, False, False), agg(field, False, True)
            status, matched = "mismatch", None
            for st, (s, n_known, _u, _t) in (("match", full), ("match_section", sec), ("match_all_records", all_rec),
                                             ("match_section_all_records", all_sec)):
                if n_known and abs(recorded - s) <= tol:
                    status, matched = st, s
                    break
            if status == "mismatch" and full[3] and abs(recorded - (full[0] - full[3])) <= tol:
                status, matched = "match_without_text_numbers", round(full[0] - full[3], 2)
            item.update(recorded=recorded, sum_in_scope=full[0], n_in_scope_known=full[1], n_in_scope_unknown=full[2],
                        difference=round(recorded - full[0], 2), sum_all_records=all_rec[0], matched_sum=matched,
                        status=status, status_ar=RECON_STATUS_AR[status])
            rows.append(item)
            if status == "mismatch":
                extra = f"، و{_n_values_ar(full[2], False)} قد تفسر جزءًا منه" if full[2] else ""
                other = (f"؛ مجموع كل سجلات المصدر بما فيها المعزولة والمستبعدة {fmt_num(all_rec[0])}"
                         if abs(all_rec[0] - full[0]) > tol else "")
                warns.append((f"الإجمالي المدون {fmt_num(recorded)} للحقل {field} لا يطابق مجموع السجلات ضمن النطاق من المصدر "
                              f"{sid} ({fmt_num(full[0])} من {_n_values_ar(full[1], True)}{extra})؛ الفرق "
                              f"{fmt_num(round(recorded - full[0], 2))}{other}. تحقق من صيغة الإجمالي والصفوف المستبعدة، "
                              "واذكر الفرق وسببه في مشكلات البيانات", f"{ref} ← 03_records.json:excluded_rows[{sid}:{row}]"))
    ctx._g4b_recon = (rows, warns)
    return ctx._g4b_recon


def g4b_md_reconciliation(report: dict) -> list:
    rows = report.get("reconciliation") or []
    if not rows:
        return []
    lines = [f"## مطابقة الإجماليات المدونة مع مجموع السجلات ({len(rows)})", "",
             "| الموضع | الحقل | المدون | مجموع السجلات ضمن النطاق | الفرق | مجموع كل سجلات المصدر | المجموع المطابق للمدون "
             "| النتيجة |",
             "|---|---|---|---|---|---|---|---|"]

    def num(x):
        return _md_cell(fmt_num(x) if x is not None else None)
    for r in rows:
        lines.append(f"| {_md_cell(r.get('ref'))} | {_md_cell(r.get('field'))} | {num(r.get('recorded'))} | "
                     f"{num(r.get('sum_in_scope'))} | {num(r.get('difference'))} | {num(r.get('sum_all_records'))} | "
                     f"{num(r.get('matched_sum'))} | {_md_cell(r.get('status_ar'))} |")
    lines.append("")
    return lines


# --- C14: {{f:key}} fact tokens -------------------------------------------------------------------------------------
def c_fact_tokens(ctx: Ctx):
    ctx.need("narrative")
    found = [(p, m) for p, s in walk_strings(ctx.narrative) for m in FACT_TOKEN_RE.finditer(s)]
    if not found:
        return []
    facts = compute_facts(ctx.run_dir)
    out = []
    for p, m in found:
        key = m.group(1)
        if key in facts:
            continue
        near = [k for k in facts if k.split(".")[0] == key.split(".")[0]]
        hint = (f"المفاتيح المتاحة بالبادئة نفسها: {short_list(near, 6)}" if near
                else "اعرض المفاتيح المتاحة بـ scripts/facts.py --run-dir")
        out.append((f"الرمز {m.group(0)} يشير إلى حقيقة غير موجودة في مجلد التشغيل؛ {hint}", f"06_narrative.json:{p}"))
    return out


# --- C24 / D1: delivered outputs (supporting workbook, 09 message) --------------------------------------------------
def g4b_message_text(ctx: Ctx):
    if not hasattr(ctx, "_g4b_msg"):
        p = os.path.join(ctx.run_dir, MESSAGE_NAME)
        ctx._g4b_msg = None
        if os.path.isfile(p):
            with open(p, "r", encoding="utf-8", errors="replace") as fh:
                ctx._g4b_msg = fh.read()
    return ctx._g4b_msg


def g4b_scan_lines(text: str, label: str, levels) -> list:
    out = []
    for n, line in enumerate(text.splitlines(), 1):
        kinds = pii_value_hits(line, levels)
        if kinds:
            out.append((f"نص يشبه {pii_kinds_ar(kinds)} (القيمة لا تُطبع)", f"{label}، السطر {n}"))
    return out


def g4b_export_cells(ctx: Ctx):
    """None when there is no supporting workbook to check; else {path, label, cells: [(sheet, coord, text, is_number,
    header)], error, stale}; header = the nearest text cell above in the same column. Read into memory only."""
    if hasattr(ctx, "_g4b_export"):
        return ctx._g4b_export
    explicit = getattr(ctx, "xlsx_out", None)
    path = explicit or os.path.join(ctx.run_dir, EXPORT_NAME)
    res = None
    if explicit or os.path.isfile(path):
        res = {"path": path, "label": os.path.basename(path), "cells": [], "error": None, "stale": False}
        if not os.path.isfile(path):
            res["error"] = "الملف المحدد في --xlsx-out غير موجود"
        else:
            try:
                import warnings
                import openpyxl
                with warnings.catch_warnings():
                    warnings.simplefilter("ignore")
                    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)
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
                                    t = str(int(v)) if float(v).is_integer() else str(v)
                                    res["cells"].append((ws.title, c.coordinate, t, True, above.get(col, "")))
                                elif isinstance(v, str):
                                    res["cells"].append((ws.title, c.coordinate, v, False, above.get(col, "")))
                                    above[col] = v
                finally:
                    wb.close()
            except Exception as exc:
                res["error"] = f"تعذر فتح المرفق الداعم لفحصه: {exc}"
            newer = [n for n in ("03_records.json", "06_narrative.json")
                     if os.path.isfile(os.path.join(ctx.run_dir, n))
                     and os.path.getmtime(os.path.join(ctx.run_dir, n)) > os.path.getmtime(path) + 1]
            res["stale"] = bool(newer)
    ctx._g4b_export = res
    return res


def _export_hits(exp: dict, levels, advice: str) -> list:
    """Text cells, and numbers under a personal-data header (a phone stored as a number). Other numbers are amounts or
    deed numbers: a 9-digit amount would look like a mobile number."""
    by_sheet = OrderedDict()
    for sheet, coord, text, is_number, header in exp["cells"]:
        if is_number and not PII_HEADER_RE.search(header or ""):
            continue
        kinds = pii_value_hits(text, levels)
        if kinds:
            d = by_sheet.setdefault(sheet, {"kinds": set(), "refs": []})
            d["kinds"].update(kinds)
            d["refs"].append(coord)
    stale = "؛ الملف أقدم من 03 أو 06 وقد يكون من تشغيل سابق" if exp.get("stale") else ""
    return [(f"{_n_ar(len(d['refs']), CELLS_AR)} في الورقة «{sheet}» تحوي ما يشبه {pii_kinds_ar(sorted(d['kinds']))} (القيم لا تُطبع)"
             f"{stale}؛ {advice}", f"{exp['label']} ← «{sheet}»!{short_list(d['refs'], 5)}")
            for sheet, d in by_sheet.items()]


def c_pii_export_patterns(ctx: Ctx):
    exp = g4b_export_cells(ctx)
    if exp is None:
        raise Skip(f"لا يوجد {EXPORT_NAME} في مجلد التشغيل ولم يُمرَّر --xlsx-out")
    if exp["error"]:
        return [(exp["error"], exp["path"])]
    return _export_hits(exp, ("block",), "احذف الملف أو أعد تصديره (export_workbook.py) بعد إزالة مصدر هذه البيانات")


def c_pii_review_patterns(ctx: Ctx):
    """Review-level patterns (a 10-digit number starting with 1 or 2 may be a national id or a deed number) in every
    delivered output; one warning per output."""
    found = []
    if isinstance(ctx.data.get("narrative"), dict):
        found += _scan_pattern_strings(ctx.narrative, "06_narrative.json", ("review",))
    msg = g4b_message_text(ctx)
    if msg is not None:
        found += g4b_scan_lines(msg, MESSAGE_NAME, ("review",))
    if ctx.html_path and ctx.html_raw is not None:
        name = os.path.basename(ctx.html_path)
        found += g4b_scan_lines(visible_text(ctx.html_raw), f"{name}: نص ظاهر", ("review",))
        data = embedded_data(ctx.html_raw)
        if data is not None:
            found += _scan_pattern_strings(data, f"{name} (DATA)", ("review",))
        if ctx.analysis is not None:
            found += _scan_pattern_strings(ctx.analysis, "analysis.json", ("review",))
    by_out = OrderedDict()
    for _detail, where in found:
        by_out.setdefault(re.split(r"[:،]", where, 1)[0], []).append(where)
    advice = ("راجعها: احذف أي رقم هوية شخصية، وإن كانت أرقام صكوك أو سجلات فاذكر ذلك في acknowledged_warnings "
              "بسبب واضح")
    out = [(f"{label}: {_n_ar(len(ws), PLACES_AR)} بها {PII_KIND_AR['national_id']} (القيم لا تُطبع)؛ {advice}",
            short_list(ws, 4)) for label, ws in by_out.items()]
    exp = g4b_export_cells(ctx)
    if exp and not exp["error"]:
        out += _export_hits(exp, ("review",), advice)
    return out


# --- 09 message: hand-typed numbers ------------------------------------------------------------------------------
MSG_NUM_RE = re.compile(r"(?<![\w.,/\-−])[-−]?(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d+))?(?![\d,.]\d)"
                        r"(\s*(?:%|٪|مليار(?:ات)?|ملايين|مليون|آلاف|الاف|ألف|الف)(?![\w]))?")
MSG_SCALE = {"ألف": 1e3, "الف": 1e3, "آلاف": 1e3, "الاف": 1e3, "مليون": 1e6, "ملايين": 1e6, "مليار": 1e9,
             "مليارات": 1e9}


def _all_numbers(obj, acc: set):
    if is_num(obj) and math.isfinite(obj):
        acc.add(float(obj))
    elif isinstance(obj, dict):
        for v in obj.values():
            _all_numbers(v, acc)
    elif isinstance(obj, list):
        for v in obj:
            _all_numbers(v, acc)
    return acc


def _message_candidates(ctx: Ctx) -> set:
    acc = set()
    if isinstance(ctx.data.get("metrics"), dict):
        _all_numbers(ctx.met_doc, acc)
    try:
        for f in compute_facts(ctx.run_dir).values():
            _all_numbers(f.get("value"), acc)
    except Exception:
        pass
    if isinstance(ctx.data.get("records"), dict):
        acc |= _file_numbers(ctx)
    n = ctx.narrative if isinstance(ctx.data.get("narrative"), dict) else {}
    for key in ("gaps", "recommendations", "strengths", "assumptions", "questions", "data_requests", "top_gaps",
                "top_actions", "exclusions", "data_issues_extra"):
        acc.add(float(len(as_list(n.get(key)))))
    plan = as_dict(n.get("plan"))
    acc |= {float(len(as_list(plan.get(k)))) for k in ("d30", "d60", "d90")}
    for name in ("07_validation.json", "08_dashboard_check.json"):  # «نتيجة الفحوص» quotes the last runs
        p = os.path.join(ctx.run_dir, name)
        if os.path.isfile(p):
            try:
                _all_numbers(load_json(p), acc)
            except Exception:
                pass
    acc |= {c * 100 for c in list(acc) if abs(c) <= 1}  # shares written as percentages
    return {abs(c) for c in acc}


def c_message_numbers(ctx: Ctx):
    text = g4b_message_text(ctx)
    if text is None:
        raise Skip(f"لا يوجد {MESSAGE_NAME} بعد")
    cands = sorted(_message_candidates(ctx))
    out = []
    for line_no, line in enumerate(text.splitlines(), 1):
        t = mask_pii(line, ("block", "review"))[0]          # personal data is reported by the privacy checks
        t = t.translate(_AR_INDIC).replace("٬", ",").replace("٫", ".")
        t = re.sub(r"`[^`]*`", " ", t)                       # code spans (paths, commands)
        t = re.sub(r"\S*[/\\]\S*", " ", t)                   # paths, dates with slashes, «30/60/90»
        t = DATE_RE.sub(" ", t)
        t = re.sub(r"\d{1,2}:\d{2}(?::\d{2})?", " ", t)     # times
        t = re.sub(r"\b[0-9a-fA-F]{12,}\b", " ", t)           # hashes
        t = re.sub(r"^\s*(?:[-*+]\s+)?\d{1,2}[.)]\s", " ", t)  # list markers «1.» «2)»
        for m in MSG_NUM_RE.finditer(t):
            int_part, dec_part, unit = m.group(1), m.group(2), (m.group(3) or "").strip()
            x = float(int_part.replace(",", "") + ("." + dec_part if dec_part else ""))
            dec = len(dec_part) if dec_part else 0
            after = t[m.end():m.end() + 8]
            before = t[max(0, m.start() - 1):m.start()]
            if not unit and not dec:
                if re.fullmatch(r"\d{4}", int_part) and (1300 <= x <= 1500 or 1900 <= x <= 2100):
                    continue  # a year
                if x <= 20 and (after.startswith(")") or before in ("(", "≤", "≥", "<", ">")):
                    continue  # «1) …» ordinal or «(≤3)»
                if x in (30, 60, 90) and re.match(r"\s*(?:يوم|أيام|ايام)", after):
                    continue  # the 30/60/90-day plan
            scale = MSG_SCALE.get(unit, 1.0)
            val = x * scale
            tol = 0.5 * (10 ** -dec) * scale + 1e-9
            if any(abs(val - c) <= tol or abs(val - c) <= 1e-9 * max(c, 1) for c in cands):
                continue
            out.append((f"الرقم «{m.group(0).strip()}» مكتوب يدويًا ولا يطابق أي قيمة في 05_metrics.json ولا حقيقة {{{{f:…}}}} "
                        "ولا قيمة في الملف؛ انسخه من 05 بالتقريب المعروض أو صحّحه أو اذكر مصدره",
                        f"{MESSAGE_NAME}، السطر {line_no}"))
    return out


# --- C13: acknowledged warnings -----------------------------------------------------------------------------------
def g4b_finalize(ctx: Ctx, entries: list) -> "OrderedDict":
    """Apply 06 acknowledged_warnings (C13) to the check entries in place and return extra report keys
    (acknowledged_count, acknowledged_ids, reconciliation)."""
    nar = ctx.data.get("narrative")
    acks = as_list(nar.get("acknowledged_warnings")) if isinstance(nar, dict) else []
    if acks:
        levels = {cid: lvl for cid, lvl, _fn in CHECKS}
        levels.setdefault("scenario_overlap_cross_type", "warning")
        failing = {e["id"] for e in entries if e["level"] == "warning" and not e["passed"]}
        accepted, problems = OrderedDict(), []
        for i, a in enumerate(acks):
            a = as_dict(a)
            cid = a.get("check_id")
            reason = a.get("reason") if isinstance(a.get("reason"), str) else ""
            n = word_count_ar(reason)
            if cid not in levels:
                problems.append(("warning", f"معرّف الفحص «{cid}» غير معروف؛ انسخه من الحقل id في 07_validation.json", i))
            elif levels[cid] != "warning":
                problems.append(("warning", f"الفحص «{cid}» {'مانع' if levels[cid] == 'blocking' else 'ليس تحذيرًا'}؛ لا يُقرّ "
                                 "إلا بالتحذيرات، أما الأخطاء المانعة فتُصلح ولا يُقرّ بها", i))
            elif n < ACK_MIN_WORDS:
                problems.append(("warning", f"سبب الإقرار بالتحذير «{cid}» {_words_ar(n)}؛ المطلوب {ACK_MIN_WORDS} كلمات فأكثر "
                                 "تشرح لماذا يبقى التحذير دون إصلاح", i))
            elif cid not in failing:
                problems.append(("info", f"لا يوجد تحذير «{cid}» في هذا التحقق؛ احذف الإقرار إن لم يعد لازمًا", i))
            else:
                accepted.setdefault(cid, clean_text(reason))
        for e in entries:
            if e["level"] == "warning" and not e["passed"] and e["id"] in accepted:
                e["level"] = "acknowledged"
                e["acknowledged_reason"] = accepted[e["id"]]
        for level, detail, i in problems:
            entries.append({"id": "acknowledged_warnings", "level": level, "passed": False, "detail": detail,
                            "where": f"06_narrative.json:acknowledged_warnings[{i}]",
                            "title": CHECK_TITLES.get("acknowledged_warnings")})
    ack = [e for e in entries if e["level"] == "acknowledged"]
    extra = OrderedDict([("acknowledged_count", len(ack)), ("acknowledged_ids", sorted({e["id"] for e in ack}))])
    recon = getattr(ctx, "_g4b_recon", None)
    if recon and recon[0]:
        extra["reconciliation"] = recon[0]
    return extra


G4B_TITLES = OrderedDict([
    ("total_like_in_scope", "عدم إبقاء صفوف تشبه الإجمالي ضمن النطاق"),
    ("unit_double_count", "عدم إبقاء الأصل ووحداته أو عقوده بمبالغها معًا ضمن النطاق (مستوى واحد للجمع)"),
    ("duplicate_ids_in_scope", "عدم إبقاء سجلين بمعرّف مكرر ضمن النطاق"),
    ("overlapping_rows", "عدم قراءة خلايا الصف نفسه في أكثر من سجل ضمن النطاق"),
    ("dims_vocabulary", "تصنيفات السجلات ضمن المفردات المعتمدة (assets/vocab.json)"),
    ("sheet_scope", "اتساق نطاق السجلات مع قرار كل ورقة ودورها في 02"),
    ("mapping_context", "سياق الدراسة وخريطة الأوراق مكتملان في 02 بأدلتها"),
    ("fact_tokens", "صحة رموز الحقائق {{f:…}}"),
    ("pii_export_patterns", "خلو المرفق الداعم من بيانات اتصال أو حسابات بنكية"),
    ("total_like_records", "سجلات ضمن النطاق تشبه صفوف الإجمالي (تجاوز موثق، أو آخر صف يساوي مجموع غيره)"),
    ("unit_levels_mixed", "وحدات أو عقود ضمن النطاق بمبالغ إيراد أو تقييم (دون أصل محدد أو بتجاوز موثق)"),
    ("duplicate_ids_kept", "معرّفات مكررة أُبقيت ضمن النطاق بتجاوز موثق"),
    ("overlapping_rows_other", "صفوف مقروءة في أكثر من سجل دون تكرار خلية معروفة"),
    ("hidden_sheet_in_scope", "سجلات ضمن النطاق من أوراق مخفية"),
    ("totals_reconciliation", "مطابقة الإجماليات المدونة مع مجموع السجلات المؤهلة"),
    ("pii_review_patterns", "أرقام تشبه الهوية الوطنية في المخرجات (للمراجعة)"),
    ("message_numbers", "أرقام رسالة التسليم (09_message.md) مطابقة لـ05 أو للحقائق أو للملف"),
    ("acknowledged_warnings", "صحة الإقرار بالتحذيرات في 06 (acknowledged_warnings)"),
    ("pii_text_patterns", "خلو السرد والسجلات والربط ورسالة التسليم من بيانات اتصال أو حسابات بنكية"),
    ("pii_html_patterns", "خلو الداشبورد وanalysis.json من بيانات اتصال أو حسابات بنكية"),
])
CHECK_TITLES.update(G4B_TITLES)

G4B_CHECKS = [
    ("total_like_in_scope", "blocking", lambda ctx: g4b_total_like(ctx)[0]),
    ("unit_double_count", "blocking", lambda ctx: g4b_unit_levels(ctx)[0]),
    ("duplicate_ids_in_scope", "blocking", lambda ctx: g4b_duplicate_ids(ctx)[0]),
    ("overlapping_rows", "blocking", lambda ctx: g4b_overlaps(ctx)[0]),
    ("dims_vocabulary", "blocking", c_dims_vocabulary),
    ("sheet_scope", "blocking", lambda ctx: g4b_sheet_scope(ctx)[0]),
    ("mapping_context", "blocking", c_mapping_context),
    ("fact_tokens", "blocking", c_fact_tokens),
    ("pii_export_patterns", "blocking", c_pii_export_patterns),
    ("total_like_records", "warning", lambda ctx: g4b_total_like(ctx)[1]),
    ("unit_levels_mixed", "warning", lambda ctx: g4b_unit_levels(ctx)[1]),
    ("duplicate_ids_kept", "warning", lambda ctx: g4b_duplicate_ids(ctx)[1]),
    ("overlapping_rows_other", "warning", lambda ctx: g4b_overlaps(ctx)[1]),
    ("hidden_sheet_in_scope", "warning", lambda ctx: g4b_sheet_scope(ctx)[1]),
    ("totals_reconciliation", "warning", lambda ctx: g4b_reconciliation(ctx)[1]),
    ("pii_review_patterns", "warning", c_pii_review_patterns),
    ("message_numbers", "warning", c_message_numbers),
]


# ---------------------------------------------------------------------------
# WARNING checks
# ---------------------------------------------------------------------------
def c_assessment_scores(ctx: Ctx):
    ctx.need("narrative")
    out = []
    for i, a in enumerate(as_list(ctx.narrative.get("assessment"))):
        if not isinstance(a, dict) or nonempty(a.get("criterion")):
            continue
        txt = a.get("judgment") if isinstance(a.get("judgment"), str) else ""
        n = normalize_arabic(TOKEN_RE.sub(" ", txt)).translate(_AR_INDIC)
        reasons = []
        if SCORE_RE.search(n):
            reasons.append("درجة رقمية")
        for m in PCT_RE.finditer(n):
            window = n[max(0, m.start() - 40): m.end() + 40]
            if JUDGE_RE.search(window):
                reasons.append("نسبة مئوية مقرونة بحكم")
                break
        if SUPERLATIVE_RE.search(n):
            reasons.append("حكم مطلق")
        if reasons:
            out.append((f"التقييم يستخدم {('، '.join(reasons))} دون معيار مفسّر (criterion = null)؛ اكتفِ بحكم وصفي تدعمه الأرقام "
                        "أو اذكر المعيار المعتمد", f"06_narrative.json:assessment[{gid(a, i, 'dimension')}].judgment"))
    return out


def c_superlatives(ctx: Ctx):
    ctx.need("narrative")
    out = []
    parts = [("executive_summary", ctx.narrative.get("executive_summary"))]
    parts += [(f"strengths[{i}].text", s.get("text")) for i, s in enumerate(as_list(ctx.narrative.get("strengths")))
              if isinstance(s, dict)]
    for path, txt in parts:
        if isinstance(txt, str) and SUPERLATIVE_RE.search(normalize_arabic(txt)):
            out.append(("حكم مطلق (مثل «ممتاز» أو «ضعيف جدًا») دون معيار مفسّر", f"06_narrative.json:{path}"))
    return out


def _metric_numbers(ctx):
    """value -> list of metric ids (numeric facts computed by the scripts)."""
    out = defaultdict(set)

    def add(v, mid):
        if is_num(v) and math.isfinite(v):
            out[float(v)].add(mid)
    for mid, res in ctx.metrics.items():
        res = as_dict(res)
        for k in NUM_KEYS_IN_RESULT:
            add(res.get(k), f"{mid}" if k == "value" else f"{mid}.{k}")
        for t in as_list(res.get("top_shares")):
            add(as_dict(t).get("value"), f"{mid}.top_shares")
    for bid, b in as_dict(ctx.met_doc.get("breakdowns")).items():
        for row in as_list(as_dict(b).get("rows")):
            add(as_dict(row).get("value"), f"breakdown:{bid}")
        add(as_dict(b).get("total"), f"breakdown:{bid}.total")
    for p in as_list(as_dict(ctx.met_doc.get("trend")).get("points")):
        add(as_dict(p).get("value"), "trend")
    return out


def _parse_num(s: str):
    t = s.translate(_AR_INDIC).replace("٬", ",").replace("٫", ".").replace(",", "")
    try:
        return float(t), (len(t.split(".")[1]) if "." in t else 0)
    except ValueError:
        return None, 0


def _file_numbers(ctx):
    nums = set()
    texts = []
    for r in ctx.records:
        for tv in as_dict(r.get("fields")).values():
            tv = as_dict(tv)
            if is_num(tv.get("v")):
                nums.add(float(tv["v"]))
            elif isinstance(tv.get("v"), str):
                texts.append(tv["v"])  # e.g. deed numbers or areas typed as text in the file
            for inp in as_list(tv.get("inputs")):
                if is_num(as_dict(inp).get("v")):
                    nums.add(float(inp["v"]))
    texts += [i.get("detail") for i in as_list(ctx.rec_doc.get("issues")) if isinstance(i, dict)]
    for a in as_list(ctx.wb.get("anomalies")):
        if not isinstance(a, dict):
            continue
        texts.append(a.get("detail"))
        for _, v in walk_strings(as_dict(a.get("evidence"))):
            texts.append(v)

        def walk_nums(o):
            if is_num(o):
                nums.add(float(o))
            elif isinstance(o, dict):
                for v in o.values():
                    walk_nums(v)
            elif isinstance(o, list):
                for v in o:
                    walk_nums(v)
        walk_nums(a.get("evidence"))
    for t in texts:
        if isinstance(t, str):
            for m in NUMBER_RE.finditer(DATE_RE.sub(" ", t.translate(_AR_INDIC))):
                v, _ = _parse_num(m.group(1))
                if v is not None:
                    nums.add(v)
            for m in re.finditer(r"(?<![\w.])\d+(?:\.\d+)?(?![\w])", t.translate(_AR_INDIC).replace(",", "")):
                try:
                    nums.add(float(m.group(0)))
                except ValueError:
                    pass
    return nums


def _matches(x, dec, candidates):
    tol = 0.5 * (10 ** -dec) + 1e-9
    for c in candidates:
        if abs(x - c) <= tol or abs(x - c) <= 1e-9 * max(abs(c), 1):
            return c
    return None


def number_findings(ctx: Ctx):
    """Return (warnings, infos) for hand-typed numbers >= 1,000 in the narrative."""
    ctx.need("narrative")
    metric_nums = _metric_numbers(ctx) if "metrics" in ctx.data else {}
    file_nums = _file_numbers(ctx)
    warns, infos = [], []
    for path, s in walk_strings(ctx.narrative, skip_keys=frozenset(NUMBER_SKIP_KEYS | {"x_"})):
        t = TOKEN_RE.sub(" ", s).translate(_AR_INDIC)
        t = DATE_RE.sub(" ", t)
        for m in NUMBER_RE.finditer(t):
            raw = m.group(1)
            x, dec = _parse_num(raw)
            if x is None or x < 1000:
                continue
            if re.fullmatch(r"\d{4}", raw) and (1300 <= x <= 1500 or 1900 <= x <= 2100):
                continue  # a year
            c = _matches(x, dec, metric_nums.keys())
            if c is not None:
                # metric values first, then metric fields; breakdown/trend matches cannot be tokens
                ids = sorted(metric_nums[c], key=lambda s: (s.startswith(("breakdown:", "trend")), "." in s, s))
                tokenable = [i for i in ids if not i.startswith(("breakdown:", "trend"))]
                hint = (f"استخدم الرمز {{{{m:{tokenable[0]}}}}} ليبقى متسقًا مع الحساب" if tokenable
                        else "يطابق توزيعًا أو تطورًا زمنيًا؛ اذكره دون رقم يدوي أو أحل إلى الرسم")
                infos.append((f"الرقم {raw} مكتوب يدويًا ويطابق {short_list(ids, 3)}؛ {hint}", f"06_narrative.json:{path}"))
            elif _matches(x, dec, file_nums) is None:
                warns.append((f"الرقم {raw} مكتوب يدويًا ولا يطابق أي مؤشر في 05 ولا قيمة في الملف؛ استخدم رمز {{{{m:…}}}} "
                              "أو صحّحه أو اذكر مصدره", f"06_narrative.json:{path}"))
    return warns, infos


def _effective_years(tv, field, fy, col_years):
    tv = as_dict(tv)
    ins = [as_dict(i).get("field") for i in as_list(tv.get("inputs"))]
    ys = {m.group(1) for f in ins if isinstance(f, str) for m in [re.match(r"^val\.(\d{4})\.", f)] if m}
    if ys:
        return ys
    if tv.get("year") is not None:
        return {str(tv["year"])}
    ref = tv.get("ref")
    if isinstance(ref, str) and "!" in ref:
        sheet, _, coord = ref.rpartition("!")
        m = re.fullmatch(r"\$?([A-Z]{1,3})\$?\d+", coord.strip())
        if m:
            hy = {str(y) for y in col_years.get((sheet, m.group(1)), []) if 1900 <= int(y) <= 2100}
            if hy and fy not in hy:
                return hy
    return {fy} if fy else set()


def c_valuation_years_mixed(ctx: Ctx):
    ctx.need("metrics", "records")
    col_years = {}
    for s in as_list(ctx.wb.get("sheets")):
        for c in as_list(as_dict(s).get("columns")):
            c = as_dict(c)
            if c.get("col"):
                col_years[(s.get("name"), c["col"])] = [y for y in as_list(c.get("years_in_header")) if is_num(y)]
    meta = ctx.field_meta
    recs = ctx.rec_by_id
    out = []
    for mid, res in ctx.metrics.items():
        res = as_dict(res)
        if res.get("kind") not in ("sum", "ratio", "diff", "concentration"):
            continue  # only kinds that add valuations together
        rf = as_dict(res.get("resolved_fields"))
        fields = [f for v in rf.values() for f in (v if isinstance(v, list) else [v]) if isinstance(f, str)]
        for vf in [f for f in fields if field_group(f, meta) == "valuation"]:
            fy = year_of_field(vf, meta)
            by_year = defaultdict(list)
            for rid in as_list(res.get("included_ids")):
                r = recs.get(rid)
                if r is None:
                    continue
                for y in _effective_years(as_dict(r.get("fields")).get(vf), vf, fy, col_years):
                    by_year[y].append(rid)
            if len(by_year) > 1:
                desc = "؛ ".join(f"{y}: {short_list(ids, 5)}" for y, ids in sorted(by_year.items()))
                out.append((f"المؤشر يجمع تقييمات من سنوات مختلفة عبر الحقل {vf} ({desc})؛ لا يُعرض المجموع تقييمًا موحدًا "
                            "في تاريخ واحد", f"05_metrics.json:{mid}"))
    return out


def c_growth_comparability(ctx: Ctx):
    ctx.need("metrics")
    out = []
    for mid, res in ctx.metrics.items():
        res = as_dict(res)
        if res.get("kind") != "growth" or res.get("status") not in ("complete", "partial"):
            continue
        comp = as_dict(res.get("comparability"))
        if comp.get("status") != "confirmed":
            card = " ويظهر كبطاقة" if res.get("card") else ""
            out.append((f"نمو الإيراد محسوب{card} بينما قابلية المقارنة «{comp.get('status') or 'غير محددة'}»: "
                        f"{comp.get('reason') or 'لم يثبت تماثل الأساس الضريبي والمحاسبي واكتمال الفترة'}؛ "
                        "لا يُعرض حكمًا بالنمو قبل التحقق", f"05_metrics.json:{mid}"))
    return out


def c_drastic_recommendations(ctx: Ctx):
    ctx.need("narrative")
    out = []
    for i, r in enumerate(as_list(ctx.narrative.get("recommendations"))):
        if not isinstance(r, dict):
            continue
        all_txt = " ".join(s for _, s in walk_strings(r, skip_keys=frozenset({"id", "gap_ids", "evidence"})))
        if not DRASTIC_RE.search(normalize_arabic(all_txt)):
            continue
        steps = " ".join([r.get("first_action") or ""] + [a for a in as_list(r.get("alternatives")) if isinstance(a, str)])
        notes = []
        if not STUDY_RE.search(normalize_arabic(steps)):
            notes.append("الإجراء الأول والبدائل لا تتضمن دراسة أو تحققًا (جدوى، تقييم معتمد، قيود نظامية ووقفية) قبل المفاضلة")
        if r.get("approval_required") in (None, ""):
            notes.append("لا يُذكر الاعتماد المطلوب (مجلس الإدارة/الجهات المختصة)")
        if notes:
            out.append(("توصية تذكر البيع أو الهدم أو تغيير الاستخدام: " + "؛ ".join(notes),
                        f"06_narrative.json:recommendations[{gid(r, i)}]"))
    return out


def c_proven_problem_cause(ctx: Ctx):
    ctx.need("narrative")
    return [("الفجوة موسومة «مشكلة مثبتة» دون سبب مثبت (proven_cause)؛ لا يُحوَّل الاحتمال إلى حقيقة",
             f"06_narrative.json:gaps[{gid(g, i)}]")
            for i, g in enumerate(as_list(ctx.narrative.get("gaps")))
            if isinstance(g, dict) and g.get("is_proven_problem") is True and not nonempty(g.get("proven_cause"))]


def c_data_coverage_low(ctx: Ctx):
    ctx.need("metrics")
    out = []
    for mid, res in ctx.metrics.items():
        res = as_dict(res)
        if res.get("kind") == "coverage" and is_num(res.get("value")) and res["value"] < 50:
            out.append((f"تغطية البيانات {res['value']:.1f}% أقل من 50%؛ النتائج تمثل جزءًا محدودًا من المحفظة ويجب إبراز ذلك",
                        f"05_metrics.json:{mid}"))
    return out


def c_unknown_classification(ctx: Ctx):
    ctx.need("records")
    out = []
    for dim, label in (("use", "الاستخدام"), ("income_nature", "طبيعة الدخل")):
        ids = [r.get("id") for r in ctx.records if r.get("scope") == "in_scope"
               and as_dict(r.get("dims")).get(dim) in (None, "", "unknown")]
        if ids:
            out.append((f"{len(ids)} سجل ضمن النطاق {label} فيه «غير محدد»: {short_list(ids)}؛ صنّفه بدليل أو وضّح أثر النقص",
                        f"03_records.json:dims.{dim}"))
    return out


def c_card_count(ctx: Ctx):
    ctx.need("config")
    groups = []
    for spec in as_list(ctx.config.get("metrics")):
        if isinstance(spec, dict) and spec.get("card") is True:
            g = spec.get("card_group") or spec.get("id")
            if g not in groups:
                groups.append(g)
    if 6 <= len(groups) <= 8:
        return []
    return [(f"عدد البطاقات المرئية {len(groups)} (المطلوب 6–8): {short_list(groups)}", "04_metrics_config.json:metrics[].card")]


def c_exec_summary_elements(ctx: Ctx):
    ctx.need("narrative")
    s = ctx.narrative.get("executive_summary")
    if not isinstance(s, str):
        raise Skip("الخلاصة التنفيذية ليست نصًا")
    n = normalize_arabic(s)
    missing = [label for label, rx in ELEMENTS.values() if not rx.search(n)]
    if missing:
        return [(f"لم يُعثر في الخلاصة على: {', '.join(missing)} (فحص تقريبي بالكلمات المفتاحية)",
                 "06_narrative.json:executive_summary")]
    return []


def c_token_null_value(ctx: Ctx):
    ctx.need("narrative", "metrics")
    out = []
    for path, s in walk_strings(ctx.narrative):
        for m in TOKEN_RE.finditer(s):
            mid, fpath = m.group(1), m.group(2).lstrip(".")
            res = ctx.metrics.get(mid)
            if not isinstance(res, dict):
                continue
            ok, val = resolve_path(res, fpath.split(".")) if fpath else (True, res.get("value"))
            if ok and val is None:
                out.append((f"الرمز {m.group(0)} قيمته فارغة (حالة المؤشر: {res.get('status')})؛ سيظهر «غير متاح» في النص",
                            f"06_narrative.json:{path}"))
    return out


def c_evidence_refs_other(ctx: Ctx):
    ctx.need("narrative", "records", "metrics", "wb")
    res = ctx.resolver()
    out = []
    groups = [("assessment", "evidence"), ("strengths", "evidence"), ("recommendations", "evidence")]
    for key, ek in groups:
        for i, item in enumerate(as_list(ctx.narrative.get(key))):
            if not isinstance(item, dict):
                continue
            for j, ref in enumerate(as_list(item.get(ek))):
                ref = ref.get("ref") if isinstance(ref, dict) else ref
                ok, note = res.resolve(ref)
                if not ok:
                    out.append((f"المرجع «{ref}» لا يُحال إلى شيء موجود: {note}",
                                f"06_narrative.json:{key}[{gid(item, i, 'id' if key != 'assessment' else 'dimension')}].{ek}[{j}]"))
    # spacing remarks on gap refs (resolved but not exact)
    for i, g in enumerate(as_list(ctx.narrative.get("gaps"))):
        for j, e in enumerate(as_list(as_dict(g).get("evidence"))):
            ref = e.get("ref") if isinstance(e, dict) else e
            ok, note = res.resolve(ref)
            if ok and note:
                out.append((f"المرجع «{ref}»: {note}", f"06_narrative.json:gaps[{gid(g, i)}].evidence[{j}]"))
    return out


def c_isolated_in_metrics(ctx: Ctx):
    ctx.need("metrics", "records")
    scope = {r.get("id"): r.get("scope") for r in ctx.records}
    out = []
    for mid, res in ctx.metrics.items():
        bad = [i for i in as_list(as_dict(res).get("included_ids")) if scope.get(i) in ("isolated", "excluded")]
        if bad:
            out.append((f"المؤشر يشمل سجلات معزولة أو مستبعدة: {short_list(bad)}؛ تحقق من base_where", f"05_metrics.json:{mid}"))
    return out


def c_narrative_claims(ctx: Ctx):
    ctx.need("narrative")
    has_exp = any(field_group(f, ctx.field_meta) == "expense" for f in ctx.field_meta)
    has_exp = has_exp or any(as_dict(p).get("exp_field") for p in as_list(ctx.config.get("periods")))
    out = []
    for path, s in walk_strings(ctx.narrative, skip_keys=frozenset(NUMBER_SKIP_KEYS - {"evidence"})):
        if affirmative_hits(TOTAL_RETURN_RE, s):
            out.append(("السرد يذكر «العائد الاستثماري الكلي» إثباتًا؛ لا يُحسب إلا بقيم بداية ونهاية الفترة والدخل والتدفقات",
                        f"06_narrative.json:{path}"))
        if not has_exp and affirmative_hits(NET_YIELD_RE, s):
            out.append(("السرد يصف العائد بأنه «صافٍ» ولا توجد بيانات مصروفات في الملف", f"06_narrative.json:{path}"))
    return out


def c_followup_status(ctx: Ctx):
    ctx.need("narrative")
    out = []
    for i, f in enumerate(as_list(ctx.narrative.get("recommendation_followup"))):
        if not isinstance(f, dict):
            continue
        st = f.get("status") if isinstance(f.get("status"), str) else ""
        if IMPLEMENTED_RE.search(normalize_arabic(st)) and not nonempty(f.get("evidence")):
            out.append(("متابعة توصية سابقة موسومة منفذة دون دليل أو إفادة مستخدم؛ لا تُعد التوصية منفذة لمجرد تحسن رقم",
                        f"06_narrative.json:recommendation_followup[{gid(f, i, 'rec_id')}]"))
    return out


def c_source_hash(ctx: Ctx):
    if not ctx.xlsx:
        raise Skip("لم يُمرَّر --xlsx")
    if not os.path.isfile(ctx.xlsx):
        raise Skip("ملف Excel غير موجود")
    expected = as_dict(ctx.wb.get("file")).get("sha256")
    if not expected:
        raise Skip("01 لا يتضمن بصمة الملف")
    if sha256_file(ctx.xlsx) != expected:
        return [("بصمة ملف Excel المحدد لا تطابق الملف المفحوص في 01؛ قد يكون ملفًا آخر أو نسخة معدلة", ctx.xlsx)]
    return []


def c_gaps_without_recommendation(ctx: Ctx):
    ctx.need("narrative")
    used = {g for r in as_list(ctx.narrative.get("recommendations")) if isinstance(r, dict)
            for g in as_list(r.get("gap_ids"))}
    lone = [g.get("id") for g in as_list(ctx.narrative.get("gaps")) if isinstance(g, dict) and g.get("id") not in used]
    return [(f"فجوات لا ترتبط بها أي توصية: {short_list(lone)}", "06_narrative.json:gaps")] if lone else []


# ---------------------------------------------------------------------------
# Runner
# ---------------------------------------------------------------------------
CHECKS = [
    ("files_exist", "blocking", c_files_exist),
    ("narrative_schema", "blocking", c_narrative_schema),
    ("exec_summary_length", "blocking", c_exec_summary_length),
    ("metric_tokens", "blocking", c_metric_tokens),
    ("duplicate_record_ids", "blocking", c_duplicate_record_ids),
    ("total_rows_in_scope", "blocking", c_total_rows_in_scope),
    *G4B_CHECKS,  # data integrity & privacy checks (G4b): blocking and warning entries
    ("tv_status_consistency", "blocking", c_tv_status_consistency),
    ("metric_status_value", "blocking", c_metric_status_value),
    ("pairwise_sets", "blocking", c_pairwise_sets),
    ("metrics_reproducible", "blocking", c_metrics_reproducible),
    ("label_net", "blocking", c_label_net),
    ("label_total_return", "blocking", c_label_total_return),
    ("label_before_expenses", "blocking", c_label_before_expenses),
    ("questions_max", "blocking", c_questions_max),
    ("narrative_ids", "blocking", c_narrative_ids),
    ("gap_evidence", "blocking", c_gap_evidence),
    ("recommendation_links", "blocking", c_recommendation_links),
    ("recommendation_fields", "blocking", c_recommendation_fields),
    ("plan_links", "blocking", c_plan_links),
    ("top_lists", "blocking", c_top_lists),
    ("scenarios", "blocking", c_scenarios),
    ("pii_text_patterns", "blocking", c_pii_text_patterns),
    ("pii_keys", "blocking", c_pii_keys),
    ("pii_html_patterns", "blocking", c_pii_html_patterns),
    ("pii_values", "blocking", c_pii_values),
    ("assessment_scores", "warning", c_assessment_scores),
    ("superlatives", "warning", c_superlatives),
    ("hand_typed_numbers", "warning", lambda ctx: number_findings(ctx)[0]),
    ("valuation_years_mixed", "warning", c_valuation_years_mixed),
    ("growth_comparability", "warning", c_growth_comparability),
    ("drastic_recommendations", "warning", c_drastic_recommendations),
    ("proven_problem_cause", "warning", c_proven_problem_cause),
    ("data_coverage_low", "warning", c_data_coverage_low),
    ("unknown_classification", "warning", c_unknown_classification),
    ("card_count", "warning", c_card_count),
    ("exec_summary_elements", "warning", c_exec_summary_elements),
    ("token_null_value", "warning", c_token_null_value),
    ("evidence_refs_other", "warning", c_evidence_refs_other),
    ("isolated_in_metrics", "warning", c_isolated_in_metrics),
    ("narrative_claims", "warning", c_narrative_claims),
    ("followup_status", "warning", c_followup_status),
    ("source_hash", "warning", c_source_hash),
    ("numbers_as_tokens", "info", lambda ctx: number_findings(ctx)[1]),
    ("gaps_without_recommendation", "info", c_gaps_without_recommendation),
]


def run_checks(ctx: Ctx) -> list:
    entries = []
    for cid, level, fn in CHECKS:
        try:
            findings = fn(ctx)
        except Skip as s:
            entries.append({"id": cid, "level": "info", "passed": True, "skipped": True,
                            "detail": f"لم يُنفذ: {s}", "where": None, "title": CHECK_TITLES.get(cid, cid)})
            continue
        except Exception as exc:  # a crash must never look like a pass
            entries.append({"id": cid, "level": level, "passed": False,
                            "detail": f"تعذر تنفيذ الفحص بسبب بنية غير متوقعة في المدخلات: {type(exc).__name__}: {exc}",
                            "where": None, "title": CHECK_TITLES.get(cid, cid)})
            continue
        if cid == "scenarios" and "narrative" in ctx.data and isinstance(ctx.data["narrative"], dict):
            try:
                for d, w in scenario_findings(ctx)[1]:
                    entries.append({"id": "scenario_overlap_cross_type", "level": "warning", "passed": False,
                                    "detail": d, "where": w, "title": "تداخل فرص مختلفة على الأصول نفسها"})
            except Exception:
                pass
        if findings:
            for detail, where in findings:
                entries.append({"id": cid, "level": level, "passed": False, "detail": detail, "where": where,
                                "title": CHECK_TITLES.get(cid, cid)})
        else:
            entries.append({"id": cid, "level": level, "passed": True, "detail": "ناجح", "where": None,
                            "title": CHECK_TITLES.get(cid, cid)})
    return entries


def _md_cell(s) -> str:
    return str(s if s is not None else "—").replace("|", "\\|").replace("\n", " ")


def write_md(path: str, report: dict) -> None:
    checks = report["checks"]
    lines = ["# نتيجة التحقق الآلي (07_validation)", ""]
    lines.append(f"- مجلد التشغيل: `{report['run_dir']}`")
    lines.append(f"- وقت التحقق: {report['generated_at']}")
    lines.append(f"- ملف Excel للفحص: {'مُمرَّر' if report['inputs'].get('xlsx') else 'لم يُمرَّر'} · "
                 f"الداشبورد: {'مُمرَّر' if report['inputs'].get('html') else 'لم يُمرَّر'}")
    lines.append(f"- **الأخطاء المانعة: {report['blocking_count']}** · التحذيرات: {report['warning_count']} · "
                 f"مُقرّ بها: {report.get('acknowledged_count', 0)} · ملاحظات: {report['info_count']}")
    verdict = ("توجد أخطاء مانعة — لا يُبنى الداشبورد النهائي قبل إصلاحها (يمكن البناء التجريبي بـ --draft)."
               if report["blocking_count"] else "لا توجد أخطاء مانعة.")
    lines += [f"- **النتيجة:** {verdict}", ""]
    sections = [("الأخطاء المانعة", lambda c: c["level"] == "blocking" and not c["passed"]),
                ("التحذيرات (تُعالج، أو يُقرّ بها بسبب في acknowledged_warnings)",
                 lambda c: c["level"] == "warning" and not c["passed"]),
                ("تحذيرات مُقرّ بها في 06 (لا تُحتسب في عدد التحذيرات)", lambda c: c["level"] == "acknowledged"),
                ("ملاحظات وفحوص لم تُنفذ", lambda c: c["level"] == "info" and (not c["passed"] or c.get("skipped")))]
    for title, pred in sections:
        rows = [c for c in checks if pred(c)]
        lines.append(f"## {title} ({len(rows)})")
        lines.append("")
        if not rows:
            lines += ["لا يوجد.", ""]
            continue
        lines += ["| # | الفحص | التفاصيل | الموضع |", "|---|---|---|---|"]
        for i, c in enumerate(rows, 1):
            detail = c["detail"] + (f" — سبب الإقرار: {c['acknowledged_reason']}" if c.get("acknowledged_reason") else "")
            lines.append(f"| {i} | {_md_cell(c['title'])} | {_md_cell(detail)} | {_md_cell(c['where'])} |")
        lines.append("")
    lines += g4b_md_reconciliation(report)
    passed = [c for c in checks if c["passed"] and not c.get("skipped")]
    lines.append(f"## الفحوص الناجحة ({len(passed)})")
    lines.append("")
    for c in passed:
        lines.append(f"- {c['title']} ({LEVEL_AR.get(c['level'], c['level'])})")
    lines.append("")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines))


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="التحقق الآلي من مجلد تشغيل تحليل المحفظة (يكتب 07_validation.json و.md)")
    ap.add_argument("--run-dir", required=True, help="مجلد التشغيل")
    ap.add_argument("--xlsx", help="ملف Excel الأصلي (لفحص تسرب قيم الأعمدة الشخصية؛ يُقرأ في الذاكرة فقط)")
    ap.add_argument("--html", help="ملف الداشبورد المبني (لفحص الخصوصية)")
    ap.add_argument("--xlsx-out", help=f"المرفق الداعم المصدَّر لفحص خلوه من البيانات الشخصية "
                                       f"(الافتراضي: RUN/{EXPORT_NAME} إن وُجد)")
    args = ap.parse_args(argv)
    import warnings
    warnings.filterwarnings("ignore", category=UserWarning, module=r"openpyxl\..*")
    run_dir = args.run_dir
    if not os.path.isdir(run_dir):
        print(f"خطأ: مجلد التشغيل غير موجود: {run_dir}", file=sys.stderr)
        return 1
    ctx = Ctx(run_dir, xlsx=args.xlsx, html_path=args.html)
    ctx.xlsx_out = args.xlsx_out
    checks = run_checks(ctx)
    extra = g4b_finalize(ctx, checks)  # acknowledged warnings (level "acknowledged") + reconciliation table
    blocking = sum(1 for c in checks if c["level"] == "blocking" and not c["passed"])
    warnings = sum(1 for c in checks if c["level"] == "warning" and not c["passed"])
    infos = sum(1 for c in checks if c["level"] == "info" and (not c["passed"] or c.get("skipped")))
    inputs = OrderedDict()
    for key, name in RUN_FILES.items():
        p = os.path.join(run_dir, name)
        inputs[name] = sha256_file(p)[:16] if os.path.isfile(p) else None
    inputs["xlsx"] = os.path.basename(args.xlsx) if args.xlsx else None
    inputs["html"] = os.path.basename(args.html) if args.html else None
    exp = getattr(ctx, "_g4b_export", None)
    inputs["xlsx_out"] = os.path.basename(exp["path"]) if exp else None
    inputs["message"] = MESSAGE_NAME if g4b_message_text(ctx) is not None else None
    report = OrderedDict([
        ("contract_version", CONTRACT_VERSION),
        ("generator", "validate_analysis.py"),
        ("generated_at", _dt.datetime.now().isoformat(timespec="seconds")),
        ("run_dir", run_dir),
        ("inputs", inputs),
        ("blocking_count", blocking),
        ("warning_count", warnings),
        ("acknowledged_count", extra["acknowledged_count"]),
        ("info_count", infos),
        ("passed", blocking == 0),
        ("blocking_ids", sorted({c["id"] for c in checks if c["level"] == "blocking" and not c["passed"]})),
        ("warning_ids", sorted({c["id"] for c in checks if c["level"] == "warning" and not c["passed"]})),
        ("acknowledged_ids", extra["acknowledged_ids"]),
        ("checks", checks),
    ])
    if extra.get("reconciliation"):
        report["reconciliation"] = extra["reconciliation"]
    save_json(os.path.join(run_dir, "07_validation.json"), report)
    write_md(os.path.join(run_dir, "07_validation.md"), report)
    print(f"التحقق الآلي: أخطاء مانعة {blocking} · تحذيرات {warnings} · ملاحظات {infos}"
          + (f" · تحذيرات مُقرّ بها {extra['acknowledged_count']}" if extra["acknowledged_count"] else ""))
    for c in checks:
        if c["level"] == "blocking" and not c["passed"]:
            print(f"  [مانع] {c['title']}: {c['detail']} — {c['where'] or ''}")
    shown = 0
    for c in checks:
        if c["level"] == "warning" and not c["passed"]:
            if shown < 15:
                print(f"  [تحذير] {c['title']}: {c['detail']} — {c['where'] or ''}")
            shown += 1
    if shown > 15:
        print(f"  … و{shown - 15} تحذيرًا آخر في 07_validation.md")
    print(f"كُتب: {os.path.join(run_dir, '07_validation.json')} و07_validation.md")
    return 2 if blocking else 0


if __name__ == "__main__":
    sys.exit(main())
