#!/usr/bin/env python3
"""Compare two run folders of the same portfolio (brief §12, data-contract.md §6 `version_comparison`).

Usage:
  python3 scripts/compare_versions.py --prev RUN_OLD --curr RUN --out RUN/version_comparison.json

Reads 03_records.json of both runs (required), plus 04_metrics_config.json, 05_metrics.json, 06_narrative.json
and 02_mapping.json when present. Records are matched by id, then by deed number (fields.deed_no), then by
name_key + district; every difference is put in exactly one category:

  new_period_performance  fields of a period that the old run did not have
  scope_change            records present in one run only (never read as a sale or a purchase)
  valuation_change        val.* fields that changed
  data_correction         same-period fields (or names) that changed — a re-upload of the same year is NOT growth
  definition_change       metric definitions / field meta / record classification that changed

Old recommendations are listed for follow-up with status «لم يُتحقق»; nothing is ever marked implemented
automatically. An Arabic summary is printed to stdout. Exit 0 on success, 1 on unusable input.
"""
from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import argparse  # noqa: E402
import datetime as _dt  # noqa: E402
import difflib  # noqa: E402
import json  # noqa: E402
import re  # noqa: E402
from collections import Counter, OrderedDict, defaultdict  # noqa: E402

from apa_common import (  # noqa: E402
    CONTRACT_VERSION, KNOWN_STATUSES, clean_text, load_json, name_key, normalize_arabic, save_json,
)

CATEGORIES = OrderedDict([
    ("new_period_performance", "أداء فترة جديدة"),
    ("scope_change", "تغير نطاق المحفظة"),
    ("valuation_change", "تغير التقييم"),
    ("data_correction", "تصحيح بيانات للفترة نفسها"),
    ("definition_change", "تغير تعريف أو تصنيف"),
])
CONFIDENCE_AR = {"exact": "مطابقة بالمعرّف والاسم أو الصك", "deed": "مطابقة برقم الصك", "probable": "مطابقة محتملة (تحتاج تأكيدًا)"}
NEW_RECORD_TEXT = "أصل جديد في الملف أو كان ناقصًا سابقًا"
GONE_RECORD_TEXT = "غير موجود في النسخة الجديدة — لا يُفترض بيعه"
NOT_GROWTH = "تصحيح لبيانات الفترة نفسها — ليس نموًا"
FOLLOWUP_STATUS = "لم يُتحقق"
FOLLOWUP_NOTE = ("تُحدَّث الحالة بحسب ما تثبته البيانات أو إفادة المستخدم فقط؛ لا تُعد التوصية منفذة لمجرد تحسن رقم "
                 "أو تغير في الملف")
SPEC_KEYS = ("kind", "where", "field", "num", "den", "a", "b", "fields", "prev", "curr", "scale", "top", "unit",
             "label", "short_label", "definition", "basis", "scope_text", "card", "card_group")
META_KEYS = ("label", "group", "period", "year", "vat", "basis", "cash_basis", "completeness", "valuation_basis",
             "valuation_date", "unit")
NAME_SIMILARITY = 0.8
REL_TOL = 1e-9
_AR_INDIC = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------
def _opt(path):
    if os.path.isfile(path):
        try:
            return load_json(path)
        except Exception as exc:
            print(f"تحذير: تعذر قراءة {path}: {exc}", file=sys.stderr)
    return None


class Run:
    def __init__(self, d: str):
        self.dir = d
        p = os.path.join(d, "03_records.json")
        if not os.path.isfile(p):
            raise SystemExit(f"خطأ: لا يوجد 03_records.json في {d}")
        self.rec_doc = load_json(p)
        self.records = [r for r in self.rec_doc.get("records") or [] if isinstance(r, dict) and r.get("id")]
        self.by_id = OrderedDict((r["id"], r) for r in self.records)
        self.meta = self.rec_doc.get("field_meta") or {}
        self.config = _opt(os.path.join(d, "04_metrics_config.json")) or {}
        self.metrics_doc = _opt(os.path.join(d, "05_metrics.json")) or {}
        self.narrative = _opt(os.path.join(d, "06_narrative.json")) or {}
        self.mapping = _opt(os.path.join(d, "02_mapping.json")) or {}
        self.periods = self._periods()
        self.val_years = self._val_years()

    def _periods(self) -> set:
        ps = set()
        for f in self.all_fields():
            p = period_of(f, self.meta)
            if p:
                ps.add(p)
        for p in self.config.get("periods") or []:
            if isinstance(p, dict) and p.get("id"):
                ps.add(str(p["id"]))
        return ps

    def _val_years(self) -> set:
        return {y for f in self.all_fields() for y in [valuation_year(f, self.meta)] if y}

    def all_fields(self) -> set:
        out = set(self.meta)
        for r in self.records:
            out |= set((r.get("fields") or {}).keys())
        return out

    def default_period(self):
        dp = (self.metrics_doc.get("default_filters") or {}).get("period")
        if dp:
            return str(dp)
        ps = self.config.get("periods") or []
        if ps and isinstance(ps[0], dict) and ps[0].get("id"):
            return str(ps[0]["id"])
        return (self.mapping.get("data_period") or {}).get("label")

    def info(self) -> dict:
        src = self.rec_doc.get("source") or {}
        return OrderedDict([
            ("run_dir", self.dir),
            ("source_file", os.path.basename(src.get("file") or "") or None),
            ("sha256", src.get("sha256")),
            ("data_period", (self.mapping.get("data_period") or {}).get("label")),
            ("default_period", self.default_period()),
            ("periods", sorted(self.periods)),
            ("valuation_years", sorted(self.val_years)),
            ("n_records", len(self.records)),
            ("by_scope", dict(Counter(r.get("scope") for r in self.records))),
        ])


# ---------------------------------------------------------------------------
# Field helpers
# ---------------------------------------------------------------------------
def valuation_year(field: str, meta: dict):
    if not isinstance(field, str):
        return None
    m = re.match(r"^val\.(\d{4})(?:\.|$)", field)
    if (meta.get(field) or {}).get("group") == "valuation" or m:
        y = (meta.get(field) or {}).get("year")
        return str(y) if y is not None else (m.group(1) if m else None)
    return None


def period_of(field: str, meta: dict):
    """Period of a non-valuation field (revenue, expense, occupancy ...), or None."""
    if not isinstance(field, str) or valuation_year(field, meta):
        return None
    p = (meta.get(field) or {}).get("period")
    if p is not None:
        return str(p)
    m = re.match(r"^[a-z_]+\.(\d{4})(?:\.|$)", field)
    return m.group(1) if m else None


def is_valuation(field: str, meta_p: dict, meta_c: dict) -> bool:
    return bool(valuation_year(field, meta_p) or valuation_year(field, meta_c))


def tv_brief(tv):
    if not isinstance(tv, dict):
        return None
    return OrderedDict([("v", tv.get("v")), ("s", tv.get("s")), ("ref", tv.get("ref"))])


def _num(x):
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def same_tv(a, b) -> bool:
    if a is None or b is None:
        return a is None and b is None
    if a.get("s") != b.get("s"):
        return False
    va, vb = a.get("v"), b.get("v")
    if _num(va) and _num(vb):
        return abs(va - vb) <= max(REL_TOL * max(abs(va), abs(vb)), 1e-9)
    if isinstance(va, str) and isinstance(vb, str):
        return normalize_arabic(va) == normalize_arabic(vb)
    return va == vb


def delta(a, b):
    va, vb = (a or {}).get("v"), (b or {}).get("v")
    if _num(va) and _num(vb) and (a or {}).get("s") in KNOWN_STATUSES and (b or {}).get("s") in KNOWN_STATUSES:
        return vb - va
    return None


def deed_of(rec):
    tv = (rec.get("fields") or {}).get("deed_no")
    if isinstance(tv, dict) and tv.get("s") in ("value", "derived") and tv.get("v") not in (None, ""):
        v = tv["v"]
        s = str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)
        return re.sub(r"\s+", "", clean_text(s).translate(_AR_INDIC))
    return None


def district_of(rec):
    d = (rec.get("dims") or {}).get("district")
    if not d:
        tv = (rec.get("fields") or {}).get("district")
        d = tv.get("v") if isinstance(tv, dict) else None
    return normalize_arabic(d) if isinstance(d, str) and clean_text(d) else None


def nk(rec):
    return name_key(rec.get("name") or "")


# ---------------------------------------------------------------------------
# Matching
# ---------------------------------------------------------------------------
def match_records(prev: Run, curr: Run):
    pairs, id_conflicts, ambiguous = [], [], []
    up = OrderedDict(prev.by_id)
    uc = OrderedDict(curr.by_id)

    # pass 1 — same id, confirmed by name or deed
    for rid in [i for i in prev.by_id if i in curr.by_id]:
        p, c = prev.by_id[rid], curr.by_id[rid]
        dp, dc = deed_of(p), deed_of(c)
        if nk(p) == nk(c) or (dp and dp == dc):
            basis = "المعرّف والاسم" if nk(p) == nk(c) else "المعرّف ورقم الصك"
            pairs.append((p, c, "exact", basis))
        elif dp and dc and dp != dc:
            id_conflicts.append({"id": rid, "prev_name": p.get("name"), "curr_name": c.get("name"),
                                 "note": "المعرّف نفسه برقمي صك مختلفين — لا يُطابق تلقائيًا"})
            continue
        elif difflib.SequenceMatcher(None, nk(p), nk(c)).ratio() >= NAME_SIMILARITY:
            pairs.append((p, c, "probable", "المعرّف نفسه واسم متقارب"))
        else:
            id_conflicts.append({"id": rid, "prev_name": p.get("name"), "curr_name": c.get("name"),
                                 "note": "المعرّف نفسه باسم مختلف — قد يكون إعادة ترقيم؛ لا يُطابق تلقائيًا"})
            continue
        up.pop(rid, None)
        uc.pop(rid, None)

    # pass 2 — deed number, only when unique on both sides
    def unique_index(recs, keyfn):
        idx = defaultdict(list)
        for r in recs:
            k = keyfn(r)
            if k:
                idx[k].append(r)
        return idx
    ip, ic = unique_index(up.values(), deed_of), unique_index(uc.values(), deed_of)
    for k, ps in ip.items():
        cs = ic.get(k) or []
        if len(ps) == 1 and len(cs) == 1:
            pairs.append((ps[0], cs[0], "deed", "رقم الصك"))
            up.pop(ps[0]["id"], None)
            uc.pop(cs[0]["id"], None)
        elif cs:
            ambiguous.append({"key": "deed", "prev_ids": [r["id"] for r in ps], "curr_ids": [r["id"] for r in cs],
                              "note": "رقم صك مشترك بين أكثر من سجل — يلزم تحقق يدوي"})

    # pass 3 — name_key + district (both known), unique on both sides
    def nd(r):
        d = district_of(r)
        return (nk(r), d) if nk(r) and d else None
    ip, ic = unique_index(up.values(), nd), unique_index(uc.values(), nd)
    for k, ps in ip.items():
        cs = ic.get(k) or []
        if len(ps) == 1 and len(cs) == 1:
            pairs.append((ps[0], cs[0], "probable", "تطابق الاسم والحي (دون معرّف أو صك مشترك)"))
            up.pop(ps[0]["id"], None)
            uc.pop(cs[0]["id"], None)
        elif cs:
            ambiguous.append({"key": "name_district", "prev_ids": [r["id"] for r in ps],
                              "curr_ids": [r["id"] for r in cs], "note": "أسماء متشابهة في الحي نفسه — يلزم تحقق يدوي"})

    # name only (no district): never matched, listed for review
    names_c = defaultdict(list)
    for r in uc.values():
        if nk(r):
            names_c[nk(r)].append(r["id"])
    for r in up.values():
        if nk(r) and names_c.get(nk(r)):
            ambiguous.append({"key": "name_only", "prev_ids": [r["id"]], "curr_ids": names_c[nk(r)],
                              "note": "تطابق الاسم فقط دون حي أو صك — تشابه الاسم وحده لا يكفي للمطابقة"})
    return pairs, list(up.values()), list(uc.values()), id_conflicts, ambiguous


# ---------------------------------------------------------------------------
# Differences
# ---------------------------------------------------------------------------
def _change(cat, subtype, p, c, conf, field=None, prev_tv=None, curr_tv=None, note=None, **extra):
    item = OrderedDict([
        ("category", cat), ("subtype", subtype),
        ("record_id_prev", p.get("id") if p else None), ("record_id_curr", c.get("id") if c else None),
        ("name", (c or p or {}).get("name")), ("match_confidence", conf),
        ("field", field), ("prev", tv_brief(prev_tv)), ("curr", tv_brief(curr_tv)),
        ("delta", delta(prev_tv, curr_tv) if (prev_tv or curr_tv) else None), ("note", note),
    ])
    item.update(extra)
    return item


def record_changes(prev: Run, curr: Run, p, c, conf):
    out = []
    fp, fc = p.get("fields") or {}, c.get("fields") or {}
    for f in sorted(set(fp) | set(fc)):
        a, b = fp.get(f), fc.get(f)
        a = a if isinstance(a, dict) else None
        b = b if isinstance(b, dict) else None
        if same_tv(a, b):
            continue
        if is_valuation(f, prev.meta, curr.meta):
            y = valuation_year(f, curr.meta) or valuation_year(f, prev.meta)
            if a is None:
                sub, note = ("new_valuation_year" if y not in prev.val_years else "valuation_added"), \
                    "تقييم جديد؛ تغير القيمة ليس دخلًا نقديًا"
            elif b is None:
                sub, note = "valuation_absent_in_new", "التقييم غير موجود في النسخة الجديدة"
            else:
                sub, note = "valuation_value_changed", ("تقييم السنة نفسها تغير بين النسختين — قد يكون تصحيحًا أو إعادة "
                                                        "تقييم؛ تحقق من التقرير")
            out.append(_change("valuation_change", sub, p, c, conf, f, a, b, note, valuation_year=y))
            continue
        per = period_of(f, curr.meta) or period_of(f, prev.meta)
        if per and per not in prev.periods:
            out.append(_change("new_period_performance", "new_period_value", p, c, conf, f, a, b,
                               f"فترة جديدة ({per}) لم تكن في النسخة السابقة", period=per))
        elif per and per not in curr.periods:
            out.append(_change("data_correction", "period_absent_in_new", p, c, conf, f, a, b,
                               f"الفترة {per} غير موجودة في النسخة الجديدة — ليس انخفاضًا في الأداء", period=per))
        else:
            if a is None:
                sub = "field_added"
            elif b is None:
                sub = "field_removed"
            elif a.get("s") != b.get("s"):
                sub = "status_changed"
            else:
                sub = "value_changed"
            out.append(_change("data_correction", sub, p, c, conf, f, a, b,
                               NOT_GROWTH if per else "تصحيح بيانات وصفية أو كمية", period=per))
    if nk(p) != nk(c):
        out.append(_change("data_correction", "name_changed", p, c, conf, "name", None, None,
                           "تغير اسم السجل بين النسختين", prev_name=p.get("name"), curr_name=c.get("name")))
    for attr in ("record_type", "scope"):
        if p.get(attr) != c.get(attr):
            out.append(_change("definition_change", "classification", p, c, conf, attr, None, None,
                               "تغير تصنيف السجل بين النسختين (قرار تحليلي أو تغير في الملف)",
                               prev_value=p.get(attr), curr_value=c.get(attr)))
    dp, dc = p.get("dims") or {}, c.get("dims") or {}
    for k in sorted(set(dp) | set(dc)):
        if json.dumps(dp.get(k), ensure_ascii=False, sort_keys=True) != json.dumps(dc.get(k), ensure_ascii=False, sort_keys=True):
            out.append(_change("definition_change", "classification", p, c, conf, f"dims.{k}", None, None,
                               "تغير تصنيف السجل بين النسختين", prev_value=dp.get(k), curr_value=dc.get(k)))
    return out


def _short(x, n=300):
    s = json.dumps(x, ensure_ascii=False, sort_keys=True) if not isinstance(x, str) else x
    return s if len(s) <= n else s[:n] + "…"


def definition_changes(prev: Run, curr: Run):
    out = []

    def add(sub, key, a, b, note):
        out.append(OrderedDict([("category", "definition_change"), ("subtype", sub), ("key", key),
                                ("prev_value", _short(a)), ("curr_value", _short(b)), ("note", note)]))
    if prev.config and curr.config:
        mp = {m.get("id"): m for m in prev.config.get("metrics") or [] if isinstance(m, dict)}
        mc = {m.get("id"): m for m in curr.config.get("metrics") or [] if isinstance(m, dict)}
        for mid in sorted(set(mp) - set(mc)):
            add("metric_removed", mid, mp[mid].get("label"), None, "مؤشر لم يعد معرّفًا في النسخة الجديدة")
        for mid in sorted(set(mc) - set(mp)):
            add("metric_added", mid, None, mc[mid].get("label"), "مؤشر جديد في النسخة الجديدة")
        for mid in sorted(set(mp) & set(mc)):
            for k in SPEC_KEYS:
                if json.dumps(mp[mid].get(k), sort_keys=True, ensure_ascii=False) != \
                        json.dumps(mc[mid].get(k), sort_keys=True, ensure_ascii=False):
                    sub = "metric_label" if k in ("label", "short_label", "definition", "basis", "scope_text") else "metric_definition"
                    add(sub, f"{mid}.{k}", mp[mid].get(k), mc[mid].get(k),
                        "تغير تعريف المؤشر أو نطاقه — الفرق في قيمته ليس تغيرًا في الأداء" if sub == "metric_definition"
                        else "تغيرت تسمية المؤشر أو وصفه")
        for k in ("base_where", "valuation", "filters"):
            if json.dumps(prev.config.get(k), sort_keys=True, ensure_ascii=False) != \
                    json.dumps(curr.config.get(k), sort_keys=True, ensure_ascii=False):
                add("config", k, prev.config.get(k), curr.config.get(k), "تغير إعداد عام يؤثر في المؤشرات")
        pp = {str(p.get("id")): p for p in prev.config.get("periods") or [] if isinstance(p, dict)}
        pc = {str(p.get("id")): p for p in curr.config.get("periods") or [] if isinstance(p, dict)}
        for pid in sorted(set(pp) & set(pc)):
            for k in ("rev_field", "exp_field", "prev_rev_field"):
                if pp[pid].get(k) != pc[pid].get(k):
                    add("period_field", f"periods[{pid}].{k}", pp[pid].get(k), pc[pid].get(k),
                        "تغير الحقل المعتمد للفترة نفسها")
            sa = (pp[pid].get("comparable_with_prev") or {}).get("status")
            sb = (pc[pid].get("comparable_with_prev") or {}).get("status")
            if sa != sb:
                add("period_comparability", f"periods[{pid}].comparable_with_prev", sa, sb, "تغير حكم قابلية المقارنة")
    for f in sorted(set(prev.meta) & set(curr.meta)):
        a, b = prev.meta.get(f) or {}, curr.meta.get(f) or {}
        for k in META_KEYS:
            if a.get(k) != b.get(k):
                add("field_meta", f"{f}.{k}", a.get(k), b.get(k),
                    "تغير وصف الحقل (الضريبة/الأساس/الاكتمال/السنة) — يؤثر في قابلية المقارنة")
    return out


def _resolved(res: dict) -> set:
    out = set()
    for v in (res.get("resolved_fields") or {}).values():
        for f in (v if isinstance(v, list) else [v]):
            if isinstance(f, str):
                out.add(f)
    return out


def metric_deltas(prev: Run, curr: Run, changes: list):
    mp = prev.metrics_doc.get("metrics") or {}
    mc = curr.metrics_doc.get("metrics") or {}
    out = []
    for mid in [m for m in mc if m in mp]:
        a, b = mp[mid] or {}, mc[mid] or {}
        fields = _resolved(a) | _resolved(b)
        inc_ids = set(a.get("included_ids") or []) | set(b.get("included_ids") or [])
        changed_ids = OrderedDict()
        for ch in changes:
            rids = {x for x in (ch.get("record_id_prev"), ch.get("record_id_curr")) if x}
            if not rids & inc_ids:
                continue
            fld = ch.get("field")
            record_level = ch["category"] == "scope_change" or (ch["category"] == "definition_change") or fld == "name"
            if record_level or fld in fields:
                changed_ids.setdefault(ch["category"], set()).update(rids & inc_ids)
        same_period = a.get("period") == b.get("period")
        va, vb = a.get("value"), b.get("value")
        d = vb - va if _num(va) and _num(vb) else None
        if same_period:
            note = ("الفترة نفسها: الفرق ناتج عن تصحيح بيانات أو تغير نطاق أو تقييم أو تعريف — ليس نموًا"
                    if d not in (None, 0) else "لا فرق في القيمة")
        else:
            note = ("فترتان مختلفتان: لا يُقرأ الفرق نموًا؛ استخدم مؤشر revenue_growth بشروط قابلية المقارنة")
        affected = OrderedDict((cat, sorted(changed_ids[cat])) for cat in CATEGORIES if changed_ids.get(cat))
        out.append(OrderedDict([
            ("id", mid), ("label", b.get("label") or a.get("label")),
            ("prev_period", a.get("period")), ("curr_period", b.get("period")), ("same_period", same_period),
            ("prev_value", va), ("curr_value", vb), ("prev_status", a.get("status")), ("curr_status", b.get("status")),
            ("difference", d), ("unit", b.get("unit")),
            ("prev_n_included", a.get("n_included")), ("curr_n_included", b.get("n_included")),
            ("affected_by", affected), ("note", note),
        ]))
    return out


def followup_template(prev: Run, changes: list, pairs: list, only_prev: list):
    recs = [r for r in (prev.narrative.get("recommendations") or []) if isinstance(r, dict)]
    ids = set(prev.by_id)
    where_now = {p["id"]: (c["id"], conf) for p, c, conf, _ in pairs}
    gone = {r["id"] for r in only_prev}
    by_rec = defaultdict(list)
    for ch in changes:
        for k in ("record_id_prev", "record_id_curr"):
            if ch.get(k):
                by_rec[ch[k]].append(ch)
    out = []
    for r in recs:
        text = " ".join(str(x) for x in [r.get("target"), r.get("problem")] + list(r.get("evidence") or []))
        mentioned = sorted({i for i in ids if re.search(rf"(?<![\w-]){re.escape(i)}(?![\w-])", text)})
        signals = []
        for i in mentioned:
            for ch in by_rec.get(i, [])[:10]:
                signals.append(OrderedDict([("record_id", i), ("category", ch["category"]),
                                            ("field", ch.get("field")), ("delta", ch.get("delta"))]))
        matches = []
        for i in mentioned:
            if i in where_now:
                matches.append(OrderedDict([("prev_id", i), ("curr_id", where_now[i][0]), ("confidence", where_now[i][1])]))
            elif i in gone:
                matches.append(OrderedDict([("prev_id", i), ("curr_id", None), ("note", GONE_RECORD_TEXT)]))
        out.append(OrderedDict([
            ("rec_id", r.get("id")), ("target", r.get("target")), ("first_action", r.get("first_action")),
            ("success_kpi", r.get("success_kpi")), ("status", FOLLOWUP_STATUS), ("evidence", []),
            ("records", matches), ("data_signals", signals), ("note", FOLLOWUP_NOTE),
        ]))
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def compare(prev: Run, curr: Run) -> dict:
    pairs, only_prev, only_curr, id_conflicts, ambiguous = match_records(prev, curr)
    changes = []
    for p, c, conf, basis in pairs:
        changes += record_changes(prev, curr, p, c, conf)
    for r in only_prev:
        changes.append(_change("scope_change", "absent_in_new", r, None, None, note=GONE_RECORD_TEXT,
                               scope=r.get("scope"), record_type=r.get("record_type"),
                               source_ref=(r.get("source") or {}).get("ref_name")))
    for r in only_curr:
        changes.append(_change("scope_change", "new_in_file", None, r, None, note=NEW_RECORD_TEXT,
                               scope=r.get("scope"), record_type=r.get("record_type"),
                               source_ref=(r.get("source") or {}).get("ref_name")))
    changes += definition_changes(prev, curr)

    grouped = OrderedDict((k, [c for c in changes if c["category"] == k]) for k in CATEGORIES)

    new_periods = sorted(curr.periods - prev.periods)
    same_default = prev.default_period() is not None and prev.default_period() == curr.default_period()
    if same_default and not new_periods:
        relation = "same_period"
        growth_note = ("رفع نسخة أحدث من بيانات الفترة نفسها ليس نموًا زمنيًا؛ الفروق مصنفة تصحيحات أو تغير نطاق أو "
                       "تقييم أو تعريف")
    elif new_periods:
        relation = "new_period"
        growth_note = (f"توجد فترات جديدة ({'، '.join(new_periods)})؛ تُقرأ مقارنة الأداء عبر مؤشر نمو الإيراد بشروط "
                       "قابلية المقارنة، لا من فروق النسختين، مع فصل أثر الأصول المضافة أو الناقصة")
    else:
        relation = "different"
        growth_note = "الفترة الافتراضية تختلف بين النسختين؛ لا يُقرأ أي فرق نموًا قبل التحقق من تماثل الأساس"

    p_info, c_info = prev.info(), curr.info()
    match_counts = OrderedDict([("exact", 0), ("deed", 0), ("probable", 0)])
    for _, _, conf, _ in pairs:
        match_counts[conf] += 1
    match_counts["only_prev"] = len(only_prev)
    match_counts["only_curr"] = len(only_curr)
    counts = OrderedDict((k, len(v)) for k, v in grouped.items())
    followup = followup_template(prev, changes, pairs, only_prev)

    summary = [
        f"مقارنة نسختين: {p_info['source_file'] or prev.dir} ← {c_info['source_file'] or curr.dir}",
        f"- العلاقة بين الفترتين: {growth_note}",
        f"- مطابقة السجلات: {match_counts['exact']} تامة، {match_counts['deed']} بالصك، {match_counts['probable']} محتملة؛ "
        f"{match_counts['only_prev']} في السابقة فقط، {match_counts['only_curr']} في الجديدة فقط",
    ]
    if id_conflicts:
        summary.append(f"- تعارض معرّفات (المعرّف نفسه لاسم أو صك مختلف): {len(id_conflicts)} — لم تُطابق تلقائيًا")
    if ambiguous:
        summary.append(f"- حالات غامضة تحتاج مراجعة يدوية: {len(ambiguous)}")
    for k, label in CATEGORIES.items():
        extra = ""
        if k == "scope_change" and counts[k]:
            extra = (f" ({match_counts['only_curr']} «{NEW_RECORD_TEXT}»، {match_counts['only_prev']} «{GONE_RECORD_TEXT}»)")
        summary.append(f"- {label}: {counts[k]}{extra}")
    summary.append(f"- متابعة التوصيات السابقة: {len(followup)} توصية بحالة «{FOLLOWUP_STATUS}» "
                   "(تُحدَّث بالدليل أو إفادة المستخدم فقط)")
    if not prev.narrative:
        summary.append("- لا يوجد 06_narrative.json في النسخة السابقة؛ لا توصيات سابقة للمتابعة")

    return OrderedDict([
        ("contract_version", CONTRACT_VERSION),
        ("generator", "compare_versions.py"),
        ("generated_at", _dt.datetime.now().isoformat(timespec="seconds")),
        ("prev", p_info), ("curr", c_info),
        ("same_source_file", bool(p_info["sha256"]) and p_info["sha256"] == c_info["sha256"]),
        ("period_relation", relation), ("new_periods", new_periods), ("growth_note", growth_note),
        ("matching", OrderedDict([
            ("counts", match_counts),
            ("pairs", [OrderedDict([("prev_id", p["id"]), ("curr_id", c["id"]), ("confidence", conf),
                                    ("confidence_ar", CONFIDENCE_AR[conf]), ("basis", basis), ("name", c.get("name"))])
                       for p, c, conf, basis in pairs]),
            ("id_conflicts", id_conflicts), ("ambiguous", ambiguous),
        ])),
        ("counts", counts),
        ("category_labels", CATEGORIES),
        ("changes", grouped),
        ("metric_deltas", metric_deltas(prev, curr, changes)),
        ("recommendation_followup", followup),
        ("summary_ar", summary),
    ])


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="مقارنة مجلدي تشغيل لنسختين من ملف المحفظة وتصنيف الفروق")
    ap.add_argument("--prev", required=True, help="مجلد التشغيل السابق")
    ap.add_argument("--curr", required=True, help="مجلد التشغيل الحالي")
    ap.add_argument("--out", required=True, help="ملف JSON الناتج (مثل RUN/version_comparison.json)")
    args = ap.parse_args(argv)
    for d in (args.prev, args.curr):
        if not os.path.isdir(d):
            print(f"خطأ: المجلد غير موجود: {d}", file=sys.stderr)
            return 1
    try:
        prev, curr = Run(args.prev), Run(args.curr)
    except SystemExit as exc:
        print(str(exc), file=sys.stderr)
        return 1
    result = compare(prev, curr)
    save_json(args.out, result)
    print("\n".join(result["summary_ar"]))
    print(f"كُتب: {args.out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
