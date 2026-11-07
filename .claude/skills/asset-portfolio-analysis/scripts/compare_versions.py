#!/usr/bin/env python3
"""Compare two run folders of the same portfolio (brief §12, data-contract.md §6 `version_comparison`).

Usage:
  python3 scripts/compare_versions.py --prev RUN_OLD --curr RUN --out RUN/version_comparison.json

Reads 03_records.json of both runs (required), plus 04_metrics_config.json, 05_metrics.json, 06_narrative.json
and 02_mapping.json when present. Records are matched by id first, then by deed number (fields.deed_no), then by
name_key + district:
  - same id and same deed (or same name)                      → exact
  - same id but two different known deeds                     → id conflict, matched later by deed if possible
  - same id, different name, no contradicting deed/district   → probable (rename), reported as name_changed
  - same id whose name now belongs to another id              → id conflict (renumbering suspected)
Every difference is put in exactly one category (output key `categories.<key>`, Arabic names in `category_labels`):

  new_period_performance  fields of a period that the old run did not have
  scope_change            records present in one run only (never read as a sale or a purchase)
  valuation_change        val.* fields that changed
  data_correction         same-period fields (or names) that changed — a re-upload of the same year is NOT growth
  definition_change       metric definitions / field meta / record classification that changed

`metric_deltas[].affected_by` credits each category with the records behind a metric's difference; a switch of the
field a metric reads (e.g. excl → incl VAT) or of its definition is credited to definition_change and explained in
`definition_causes`. Old recommendations are listed for follow-up with the default status «لم يبدأ» (vocabulary
«لم يبدأ|قيد التنفيذ|منجز|متعثر»), marked as unverified; nothing is ever marked done automatically. An Arabic
summary is printed to stdout. Exit 0 on success, 1 on unusable input. `--self-test` runs the built-in regression cases.
"""
from __future__ import annotations

import os
import sys

sys.dont_write_bytecode = True
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
# Follow-up vocabulary shared with 06_narrative.json (narrative.schema.json) and the supporting workbook.
FOLLOWUP_STATUSES = ("لم يبدأ", "قيد التنفيذ", "منجز", "متعثر")
FOLLOWUP_STATUS = FOLLOWUP_STATUSES[0]
FOLLOWUP_NOTE = ("حالة افتراضية لم يُتحقق منها: تُحدَّث («قيد التنفيذ» أو «منجز» أو «متعثر») بحسب ما تثبته البيانات أو "
                 "إفادة المستخدم فقط؛ لا تُعد التوصية منجزة لمجرد تحسن رقم أو تغير في الملف")
VAT_AR = {"excl": "بدون ضريبة", "incl": "شامل الضريبة", "none_evidenced": "لا ضريبة مثبتة", "unknown": "أساس الضريبة غير معروف"}
BASIS_AR = {"actual": "فعلي", "recorded": "مسجل", "expected": "متوقع", "contract": "تعاقدي"}
COMPLETENESS_AR = {"full_year": "سنة كاملة", "partial": "فترة جزئية"}
GLOBAL_CONFIG_KEYS = ("base_where", "valuation", "filters")
ROLE_AR = {"field": "", "fields": "", "num": " للبسط", "den": " للمقام", "a": " للمطروح منه", "b": " للمطروح",
           "prev": " للفترة السابقة", "curr": " للفترة الحالية"}
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
def _name_index(recs):
    idx = defaultdict(list)
    for r in recs:
        if nk(r):
            idx[nk(r)].append(r["id"])
    return idx


def match_records(prev: Run, curr: Run):
    pairs, id_conflicts, ambiguous = [], [], []
    up = OrderedDict(prev.by_id)
    uc = OrderedDict(curr.by_id)
    names_p, names_c = _name_index(prev.records), _name_index(curr.records)

    def conflict(rid, p, c, note):
        id_conflicts.append(OrderedDict([("id", rid), ("prev_name", p.get("name")), ("curr_name", c.get("name")),
                                         ("note", note)]))

    # pass 1 — same id. A deed known on both sides decides; otherwise the name confirms it, and a different name is
    # read as a rename unless that name belongs to another id (renumbering) or the district contradicts it.
    for rid in [i for i in prev.by_id if i in curr.by_id]:
        p, c = prev.by_id[rid], curr.by_id[rid]
        dp, dc = deed_of(p), deed_of(c)
        same_name = nk(p) == nk(c)
        if dp and dc and dp != dc:
            conflict(rid, p, c, "المعرّف نفسه برقمي صك مختلفين — لا يُطابق بالمعرّف؛ تُجرى المطابقة برقم الصك إن أمكن")
            continue
        if dp and dp == dc:
            pairs.append((p, c, "exact", "المعرّف ورقم الصك" + ("" if same_name else " (مع تغير الاسم)")))
        elif same_name:
            pairs.append((p, c, "exact", "المعرّف والاسم" if nk(p) else "المعرّف (دون اسم في النسختين)"))
        else:
            # the old name now sits under another id that did not carry it before, or the new name was carried by
            # another id that no longer does: the ids were renumbered (two parcels that always shared a name are not)
            moved = sorted({i for i in names_c.get(nk(p), [])
                            if i != rid and (i not in prev.by_id or nk(prev.by_id[i]) != nk(p))} |
                           {j for j in names_p.get(nk(c), [])
                            if j != rid and (j not in curr.by_id or nk(curr.by_id[j]) != nk(c))})
            d_p, d_c = district_of(p), district_of(c)
            if moved:
                conflict(rid, p, c, "المعرّف نفسه باسم مختلف، والاسم نفسه مسجل بمعرّف آخر (" + "، ".join(moved) +
                         ") — يُرجّح إعادة ترقيم؛ لا يُطابق بالمعرّف")
                continue
            if d_p and d_c and d_p != d_c:
                conflict(rid, p, c, "المعرّف نفسه باسم وحي مختلفين — قد يكون إعادة ترقيم؛ لا يُطابق تلقائيًا")
                continue
            if difflib.SequenceMatcher(None, nk(p), nk(c)).ratio() >= NAME_SIMILARITY:
                pairs.append((p, c, "probable", "المعرّف نفسه واسم متقارب"))
            else:
                pairs.append((p, c, "probable", "المعرّف نفسه باسم مختلف دون صك أو حي يناقضه — إعادة تسمية محتملة "
                                               "تحتاج تأكيدًا"))
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
    left_c = _name_index(uc.values())
    for r in up.values():
        if nk(r) and left_c.get(nk(r)):
            ambiguous.append({"key": "name_only", "prev_ids": [r["id"]], "curr_ids": left_c[nk(r)],
                              "note": "تطابق الاسم فقط دون حي أو صك — تشابه الاسم وحده لا يكفي للمطابقة"})

    # how each side of an id conflict ended up (matched by deed / name + district, or left unmatched)
    to_curr = {p["id"]: (c["id"], conf) for p, c, conf, _ in pairs}
    from_prev = {c["id"]: (p["id"], conf) for p, c, conf, _ in pairs}
    for x in id_conflicts:
        x["prev_matched_to"] = (to_curr.get(x["id"]) or (None,))[0]
        x["curr_matched_from"] = (from_prev.get(x["id"]) or (None,))[0]
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

    def add(sub, key, a, b, note, **extra):
        item = OrderedDict([("category", "definition_change"), ("subtype", sub), ("key", key),
                            ("prev_value", _short(a)), ("curr_value", _short(b)), ("note", note)])
        item.update(extra)
        out.append(item)
    if prev.config and curr.config:
        mp = {m.get("id"): m for m in prev.config.get("metrics") or [] if isinstance(m, dict)}
        mc = {m.get("id"): m for m in curr.config.get("metrics") or [] if isinstance(m, dict)}
        for mid in sorted(set(mp) - set(mc)):
            add("metric_removed", mid, mp[mid].get("label"), None, "مؤشر لم يعد معرّفًا في النسخة الجديدة", metric_id=mid)
        for mid in sorted(set(mc) - set(mp)):
            add("metric_added", mid, None, mc[mid].get("label"), "مؤشر جديد في النسخة الجديدة", metric_id=mid)
        for mid in sorted(set(mp) & set(mc)):
            for k in SPEC_KEYS:
                if json.dumps(mp[mid].get(k), sort_keys=True, ensure_ascii=False) != \
                        json.dumps(mc[mid].get(k), sort_keys=True, ensure_ascii=False):
                    sub = "metric_label" if k in ("label", "short_label", "definition", "basis", "scope_text") else "metric_definition"
                    add(sub, f"{mid}.{k}", mp[mid].get(k), mc[mid].get(k),
                        "تغير تعريف المؤشر أو نطاقه — الفرق في قيمته ليس تغيرًا في الأداء" if sub == "metric_definition"
                        else "تغيرت تسمية المؤشر أو وصفه", metric_id=mid, attr=k)
        for k in GLOBAL_CONFIG_KEYS:
            if json.dumps(prev.config.get(k), sort_keys=True, ensure_ascii=False) != \
                    json.dumps(curr.config.get(k), sort_keys=True, ensure_ascii=False):
                add("config", k, prev.config.get(k), curr.config.get(k), "تغير إعداد عام يؤثر في المؤشرات", attr=k)
        pp = {str(p.get("id")): p for p in prev.config.get("periods") or [] if isinstance(p, dict)}
        pc = {str(p.get("id")): p for p in curr.config.get("periods") or [] if isinstance(p, dict)}
        for pid in sorted(set(pp) & set(pc)):
            for k in ("rev_field", "exp_field", "prev_rev_field"):
                if pp[pid].get(k) != pc[pid].get(k):
                    fa, fb = pp[pid].get(k), pc[pid].get(k)
                    add("period_field", f"periods[{pid}].{k}", fa, fb,
                        "تغير الحقل المعتمد للفترة نفسها" + _switch_text(prev, curr, fa, fb), period=pid, attr=k)
            sa = (pp[pid].get("comparable_with_prev") or {}).get("status")
            sb = (pc[pid].get("comparable_with_prev") or {}).get("status")
            if sa != sb:
                add("period_comparability", f"periods[{pid}].comparable_with_prev", sa, sb, "تغير حكم قابلية المقارنة",
                    period=pid)
    for f in sorted(set(prev.meta) & set(curr.meta)):
        a, b = prev.meta.get(f) or {}, curr.meta.get(f) or {}
        for k in META_KEYS:
            if a.get(k) != b.get(k):
                add("field_meta", f"{f}.{k}", a.get(k), b.get(k),
                    "تغير وصف الحقل (الضريبة/الأساس/الاكتمال/السنة) — يؤثر في قابلية المقارنة", field=f, attr=k)
    return out


def _field_desc(meta: dict, f) -> str:
    """Short Arabic description of a field's VAT / basis / completeness, from field_meta."""
    m = (meta.get(f) or {}) if isinstance(f, str) else {}
    parts = [VAT_AR.get(m.get("vat")), BASIS_AR.get(m.get("basis")), COMPLETENESS_AR.get(m.get("completeness"))]
    return "، ".join(p for p in parts if p)


def _switch_text(prev: Run, curr: Run, fa, fb) -> str:
    """«: rev.2025.excl_vat (بدون ضريبة) ← rev.2025.incl_vat (شامل الضريبة)» for a switched field, else ""."""
    if not (isinstance(fa, str) or isinstance(fb, str)):
        return ""
    da, db = _field_desc(prev.meta, fa), _field_desc(curr.meta, fb)
    left = f"{fa or '—'}" + (f" ({da})" if da else "")
    right = f"{fb or '—'}" + (f" ({db})" if db else "")
    return f": {left} ← {right}"


def _resolved(res: dict) -> set:
    out = set()
    for v in (res.get("resolved_fields") or {}).values():
        for f in (v if isinstance(v, list) else [v]):
            if isinstance(f, str):
                out.add(f)
    return out


def _resolved_map(res: dict) -> dict:
    rf = res.get("resolved_fields")
    return rf if isinstance(rf, dict) else {}


def _definition_causes(prev: Run, curr: Run, mid: str, a: dict, b: dict, fields: set, def_changes: list) -> list:
    """Arabic reasons why a metric's definition differs between the two runs (empty when it does not)."""
    causes = []
    ra, rb = _resolved_map(a), _resolved_map(b)
    for role in sorted(set(ra) | set(rb)):
        fa, fb = ra.get(role), rb.get(role)
        if json.dumps(fa, sort_keys=True, ensure_ascii=False) != json.dumps(fb, sort_keys=True, ensure_ascii=False):
            if isinstance(fa, list) or isinstance(fb, list):
                causes.append(f"الحقول المعتمدة{ROLE_AR.get(role, f' ({role})')} تغيرت: {_short(fa, 120)} ← {_short(fb, 120)}")
            else:
                causes.append(f"الحقل المعتمد{ROLE_AR.get(role, f' ({role})')} تغير" + _switch_text(prev, curr, fa, fb))
    for d in def_changes:
        sub = d.get("subtype")
        if sub == "metric_definition" and d.get("metric_id") == mid:
            causes.append(f"تعريف المؤشر تغير ({d.get('attr')})")
        elif sub == "config" and d.get("attr") in ("base_where", "valuation"):  # `filters` only lists dashboard filters
            causes.append(f"إعداد عام تغير ({d.get('attr')}) — يؤثر في كل المؤشرات")
        elif sub == "field_meta" and d.get("field") in fields and d.get("attr") != "label":
            causes.append(f"وصف الحقل {d.get('field')} تغير ({d.get('attr')}): {d.get('prev_value')} ← {d.get('curr_value')}")
    return list(OrderedDict.fromkeys(causes))


def metric_deltas(prev: Run, curr: Run, changes: list, def_changes: list = ()):
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
        causes = _definition_causes(prev, curr, mid, a, b, fields, def_changes)
        if causes:  # a switched field or definition touches every record the metric reads
            changed_ids.setdefault("definition_change", set()).update(inc_ids)
        same_period = a.get("period") == b.get("period")
        va, vb = a.get("value"), b.get("value")
        d = vb - va if _num(va) and _num(vb) else None
        if same_period:
            if d is None:
                note = ("القيمة غير متاحة في النسختين" if va is None and vb is None
                        else "الفترة نفسها: إحدى القيمتين غير متاحة؛ لا يُحسب فرق")
            elif d == 0:
                note = "لا فرق في القيمة" + (" رغم تغير التعريف (" + "؛ ".join(causes) + ")" if causes else "")
            elif causes:
                note = ("الفترة نفسها: الفرق يجمع أثر تغير التعريف — " + "؛ ".join(causes) +
                        " — مع أي تصحيح بيانات أو تغير نطاق أو تقييم؛ ليس نموًا ولا تحسنًا في الأداء")
            else:
                note = "الفترة نفسها: الفرق ناتج عن تصحيح بيانات أو تغير نطاق أو تقييم أو تعريف — ليس نموًا"
        else:
            note = ("فترتان مختلفتان: لا يُقرأ الفرق نموًا؛ استخدم مؤشر revenue_growth بشروط قابلية المقارنة" +
                    ("؛ وتعريف المؤشر تغير أيضًا: " + "؛ ".join(causes) if causes else ""))
        affected = OrderedDict((cat, sorted(changed_ids[cat])) for cat in CATEGORIES if changed_ids.get(cat))
        out.append(OrderedDict([
            ("id", mid), ("label", b.get("label") or a.get("label")),
            ("prev_period", a.get("period")), ("curr_period", b.get("period")), ("same_period", same_period),
            ("prev_value", va), ("curr_value", vb), ("prev_status", a.get("status")), ("curr_status", b.get("status")),
            ("difference", d), ("unit", b.get("unit")),
            ("prev_n_included", a.get("n_included")), ("curr_n_included", b.get("n_included")),
            ("affected_by", affected), ("definition_causes", causes), ("note", note),
        ]))
    return out


def followup_template(prev: Run, changes: list, pairs: list, only_prev: list):
    recs = [r for r in (prev.narrative.get("recommendations") or []) if isinstance(r, dict)]
    ids = set(prev.by_id)
    where_now = {p["id"]: (c["id"], conf) for p, c, conf, _ in pairs}
    gone = {r["id"] for r in only_prev}
    # Old recommendations cite the OLD run's ids: index each record-level change once, by its previous id only (a
    # change whose prev and curr ids are equal must not be listed twice, and a renumbered curr id must not be read
    # as the old record that used to carry it).
    by_rec = defaultdict(list)
    for ch in changes:
        if ch.get("record_id_prev"):
            by_rec[ch["record_id_prev"]].append(ch)
    out = []
    for r in recs:
        text = " ".join(str(x) for x in [r.get("target"), r.get("problem")] + list(r.get("evidence") or []))
        mentioned = sorted({i for i in ids if re.search(rf"(?<![\w-]){re.escape(i)}(?![\w-])", text)})
        signals, seen = [], set()
        for i in mentioned:
            for ch in by_rec.get(i, []):
                key = (i, ch["category"], ch.get("subtype"), ch.get("field"))
                if key in seen:
                    continue
                seen.add(key)
                if sum(1 for s in signals if s["record_id"] == i) >= 10:
                    break
                sig = OrderedDict([("record_id", i), ("category", ch["category"]), ("subtype", ch.get("subtype")),
                                   ("field", ch.get("field")), ("delta", ch.get("delta"))])
                if ch.get("record_id_curr") and ch["record_id_curr"] != i:
                    sig["record_id_curr"] = ch["record_id_curr"]
                signals.append(sig)
        matches = []
        for i in mentioned:
            if i in where_now:
                matches.append(OrderedDict([("prev_id", i), ("curr_id", where_now[i][0]), ("confidence", where_now[i][1])]))
            elif i in gone:
                matches.append(OrderedDict([("prev_id", i), ("curr_id", None), ("note", GONE_RECORD_TEXT)]))
        item = OrderedDict([("rec_id", r.get("id"))])
        if r.get("title"):
            item["title"] = r.get("title")
        item.update([
            ("target", r.get("target")), ("first_action", r.get("first_action")),
            ("success_kpi", r.get("success_kpi")), ("status", FOLLOWUP_STATUS), ("evidence", []),
            ("records", matches), ("data_signals", signals), ("note", FOLLOWUP_NOTE),
        ])
        out.append(item)
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
    def_changes = definition_changes(prev, curr)
    changes += def_changes

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
    renumbered = [(p["id"], c["id"]) for p, c, _, _ in pairs if p["id"] != c["id"]]
    renamed = [(p["id"], c["id"]) for p, c, conf, _ in pairs if conf == "probable" and p["id"] == c["id"] and nk(p) != nk(c)]
    match_counts["id_changed"] = len(renumbered)
    counts = OrderedDict((k, len(v)) for k, v in grouped.items())
    followup = followup_template(prev, changes, pairs, only_prev)

    summary = [
        f"مقارنة نسختين: {p_info['source_file'] or prev.dir} ← {c_info['source_file'] or curr.dir}",
        f"- العلاقة بين الفترتين: {growth_note}",
        f"- مطابقة السجلات: {match_counts['exact']} تامة، {match_counts['deed']} بالصك، {match_counts['probable']} محتملة؛ "
        f"{match_counts['only_prev']} في السابقة فقط، {match_counts['only_curr']} في الجديدة فقط",
    ]
    if id_conflicts:
        resolved = sum(1 for x in id_conflicts if x.get("prev_matched_to") or x.get("curr_matched_from"))
        summary.append(f"- تعارض معرّفات (المعرّف نفسه لصك مختلف أو لاسم صار بمعرّف آخر): {len(id_conflicts)} — لم تُطابق "
                       f"بالمعرّف" + (f"؛ طوبق منها بالصك أو بالاسم والحي: {resolved}" if resolved else ""))
    if renumbered:
        summary.append("- سجلات تغير معرّفها وطوبقت بالصك أو بالاسم والحي (راجعها): " +
                       "، ".join(f"{a} ← {b}" for a, b in renumbered[:8]) + ("…" if len(renumbered) > 8 else ""))
    if renamed:
        summary.append("- سجلات بالمعرّف نفسه واسم مختلف، عُدّت إعادة تسمية محتملة تحتاج تأكيدًا: " +
                       "، ".join(a for a, _ in renamed[:8]) + ("…" if len(renamed) > 8 else ""))
    if ambiguous:
        summary.append(f"- حالات غامضة تحتاج مراجعة يدوية: {len(ambiguous)}")
    for k, label in CATEGORIES.items():
        extra = ""
        if k == "scope_change" and counts[k]:
            extra = (f" ({match_counts['only_curr']} «{NEW_RECORD_TEXT}»، {match_counts['only_prev']} «{GONE_RECORD_TEXT}»)")
        summary.append(f"- {label}: {counts[k]}{extra}")
    if followup:
        summary.append(f"- متابعة التوصيات السابقة (عددها {len(followup)}): الحالة الافتراضية «{FOLLOWUP_STATUS}» ولم يُتحقق "
                       "منها؛ تُحدَّث بالدليل أو إفادة المستخدم فقط")
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
        ("categories", grouped),
        ("metric_deltas", metric_deltas(prev, curr, changes, def_changes)),
        ("recommendation_followup", followup),
        ("summary_ar", summary),
    ])


# ---------------------------------------------------------------------------
# Self-test (synthetic records only; run by tests/run_tests.sh or by hand)
# ---------------------------------------------------------------------------
def _st_rec(rid, name, deed=None, district=None, **fields):
    f = {k.replace("__", "."): {"v": v, "s": "value" if v else "zero_recorded", "ref": f"S!A{rid[-1]}"} for k, v in fields.items()}
    if deed:
        f["deed_no"] = {"v": deed, "s": "value", "ref": "S!Z1"}
    r = {"id": rid, "name": name, "record_type": "asset", "scope": "in_scope", "dims": {}, "fields": f,
         "source": {"source_id": "S", "sheet": "S", "row": 1}}
    if district:
        r["dims"]["district"] = district
    return r


def _st_run(root, name, records, metrics=None, narrative=None):
    d = os.path.join(root, name)
    meta = {"rev.2025.excl_vat": {"group": "revenue", "period": "2025", "vat": "excl", "basis": "recorded"},
            "rev.2025.incl_vat": {"group": "revenue", "period": "2025", "vat": "incl", "basis": "recorded"},
            "val.2024.total": {"group": "valuation", "year": "2024"}}
    save_json(os.path.join(d, "03_records.json"), {"contract_version": CONTRACT_VERSION, "source": {"file": "t.xlsx"},
                                                   "field_meta": meta, "records": records})
    if metrics:
        save_json(os.path.join(d, "05_metrics.json"), {"metrics": metrics, "default_filters": {"period": "2025"}})
    if narrative:
        save_json(os.path.join(d, "06_narrative.json"), narrative)
    return Run(d)


def self_test() -> int:
    import tempfile
    fails = []

    def check(name, ok, detail=""):
        print(("  PASS  " if ok else "  FAIL  ") + name + ("" if ok else f" — {detail}"))
        if not ok:
            fails.append(name)
    with tempfile.TemporaryDirectory() as root:
        # D2: two parcels with the same name whose ids were swapped (deeds tell them apart)
        a = [_st_rec("P-1", "أرض الواحة", "900001", rev__2025__excl_vat=30000, val__2024__total=600000),
             _st_rec("P-2", "أرض الواحة", "900002", val__2024__total=650000)]
        b = [_st_rec("P-2", "أرض الواحة", "900001", rev__2025__excl_vat=30000, val__2024__total=600000),
             _st_rec("P-1", "أرض الواحة", "900002", val__2024__total=650000)]
        r = compare(_st_run(root, "a", a), _st_run(root, "b", b))
        pairs = {(x["prev_id"], x["curr_id"], x["confidence"]) for x in r["matching"]["pairs"]}
        check("D2 swapped ids with different deeds are matched by deed, not as exact",
              pairs == {("P-1", "P-2", "deed"), ("P-2", "P-1", "deed")}, pairs)
        check("D2 no bogus corrections or valuation changes", sum(r["counts"].values()) == 0, r["counts"])
        check("D2 id conflicts record how they were resolved",
              {x["id"]: x["prev_matched_to"] for x in r["matching"]["id_conflicts"]} == {"P-1": "P-2", "P-2": "P-1"},
              r["matching"]["id_conflicts"])
        # D14: same id, renamed, no deed -> probable rename (one name_changed), not gone + new
        a = [_st_rec("M-1", "عمارة الياسمين (ملك للغير)", rev__2025__excl_vat=100)]
        b = [_st_rec("M-1", "مبنى الياسمين السكني (ملك للغير)", rev__2025__excl_vat=100)]
        r = compare(_st_run(root, "c", a), _st_run(root, "d", b))
        check("D14 same-id rename without deed is a probable match",
              [(x["prev_id"], x["confidence"]) for x in r["matching"]["pairs"]] == [("M-1", "probable")], r["matching"])
        check("D14 rename is one data correction and no scope change",
              r["counts"]["scope_change"] == 0 and [c["subtype"] for c in r["categories"]["data_correction"]] == ["name_changed"],
              r["counts"])
        # renumbering: the old name of an id now sits under another id -> conflict, then name+district match
        a = [_st_rec("R-1", "عمارة الهدى", district="النخيل"), _st_rec("R-2", "عمارة الندى", district="الملز")]
        b = [_st_rec("R-1", "عمارة الندى", district="الملز"), _st_rec("R-3", "عمارة الهدى", district="النخيل")]
        r = compare(_st_run(root, "e", a), _st_run(root, "f", b))
        pairs = {(x["prev_id"], x["curr_id"]) for x in r["matching"]["pairs"]}
        check("renumbered id is not read as a rename", pairs == {("R-1", "R-3"), ("R-2", "R-1")}, pairs)
        # two parcels that always shared a name: renaming one is still a rename, not a renumbering
        a = [_st_rec("L-1", "أرض المنار"), _st_rec("L-2", "أرض المنار")]
        b = [_st_rec("L-1", "أرض المنار الشمالية"), _st_rec("L-2", "أرض المنار")]
        r = compare(_st_run(root, "e2", a), _st_run(root, "f2", b))
        pairs = {(x["prev_id"], x["curr_id"], x["confidence"]) for x in r["matching"]["pairs"]}
        check("rename of one of two same-named parcels", pairs == {("L-1", "L-1", "probable"), ("L-2", "L-2", "exact")},
              pairs)
        # D7 + D14: follow-up signals once per change; VAT switch credited to definition_change; vocabulary
        a = [_st_rec("A-1", "عمارة النور", rev__2025__excl_vat=0, rev__2025__incl_vat=0)]
        b = [_st_rec("A-1", "عمارة النور", rev__2025__excl_vat=50000, rev__2025__incl_vat=57500)]
        m_a = {"property_revenue": {"value": 0, "period": "2025", "unit": "SAR", "included_ids": ["A-1"],
                                    "resolved_fields": {"field": "rev.2025.excl_vat"}}}
        m_b = {"property_revenue": {"value": 57500, "period": "2025", "unit": "SAR", "included_ids": ["A-1"],
                                    "resolved_fields": {"field": "rev.2025.incl_vat"}}}
        nar = {"recommendations": [{"id": "R1", "target": "A-1", "problem": "إيراد صفري مسجل", "evidence": ["A-1"]}]}
        r = compare(_st_run(root, "g", a, m_a, nar), _st_run(root, "h", b, m_b))
        sig = [(s["record_id"], s["category"], s["field"]) for s in r["recommendation_followup"][0]["data_signals"]]
        check("D7 data signals are not duplicated", len(sig) == len(set(sig)) == 2, sig)
        check("follow-up status uses the plan vocabulary",
              all(f["status"] in FOLLOWUP_STATUSES for f in r["recommendation_followup"]), r["recommendation_followup"])
        md = r["metric_deltas"][0]
        check("D14 VAT switch is credited to definition_change",
              md["affected_by"].get("definition_change") == ["A-1"] and "شامل الضريبة" in " ".join(md["definition_causes"]),
              md)
        check("D7 five categories with Arabic labels",
              list(r["categories"]) == list(CATEGORIES) == list(r["category_labels"]), list(r["categories"]))
    print("RESULT:", "FAIL" if fails else "PASS")
    return 1 if fails else 0


def main(argv=None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    if "--self-test" in argv:
        return self_test()
    ap = argparse.ArgumentParser(description="مقارنة مجلدي تشغيل لنسختين من ملف المحفظة وتصنيف الفروق")
    ap.add_argument("--prev", required=True, help="مجلد التشغيل السابق")
    ap.add_argument("--curr", required=True, help="مجلد التشغيل الحالي")
    ap.add_argument("--out", required=True, help="ملف JSON الناتج (مثل RUN/version_comparison.json)")
    ap.add_argument("--self-test", action="store_true", help="تشغيل حالات الاختبار المضمنة (بيانات وهمية)")
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
