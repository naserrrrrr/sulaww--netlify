#!/usr/bin/env python3
"""build_dashboard.py — assemble RUN/analysis.json and the self-contained RUN/dashboard.html.

Usage:
    python3 "$SKILL_DIR/scripts/build_dashboard.py" --run-dir "$RUN" [--draft] [--publish]
    python3 "$SKILL_DIR/scripts/build_dashboard.py" --self-test   (token grammar, privacy scrub, meta, fonts, hints)

Reads (data-contract.md §0):
    01_workbook_map.json  (optional; summary only — sheets, anomalies, PII column locations)
    02_mapping.json       (optional; entity, data period, dates, sheet decisions, links)
    03_records.json       (required)
    04_metrics_config.json, 05_metrics.json (required)
    06_narrative.json     (required unless --draft)
    07_validation.json    (required unless --draft; blocking errors stop the build unless --draft)
Writes:
    RUN/analysis.json     (contract §7, PII-free)
    RUN/dashboard.html    (assets/dashboard_template.html + inlined assets/metric_engine.js + data + the Arabic web font
                           of assets/fonts/ embedded base64 in @font-face, so screen and print share one typeface offline)

Exit codes: 0 ok · 1 missing/invalid input · 2 blocking validation errors (without --draft).

Narrative tokens rendered in every narrative string (TOKEN_RE / FACT_TOKEN_RE — validate_analysis.py uses the same grammar):
    {{m:metric_id}}             -> formatted value + unit from 05_metrics.json («غير قابل للاحتساب» when null)
    {{m:metric_id.field}}       -> a field of that result (n_included, abs_change, coverage_num_share, ...)
    {{m:metric_id.field.sub}}   -> a key of a dict field or an index of a list field (comparability.reason,
                                   same_records.n, ...); metric_id may also be `trend` or a breakdown id
    {{f:key}}                   -> a fact outside 05 from facts.compute_facts(RUN) (records.isolated, wb.WB-004.recomputed ...)
Formatting matches the dashboard: Western digits, thousands separators, «ر.س» for SAR,
percent with 2 decimals below 10 and 1 decimal otherwise, half-away-from-zero rounding.

Board-facing meta (analysis.json `meta`, read by the template):
    study_date        the preparation date the file states (a cell or title, 02 study_date); the workbook's last-modified
                      property is never presented as one: then study_date = null, study_date_stated = false and
                      file_modified carries that timestamp's date as context («غير مذكور في الملف» on the page)
    confidentiality   «سري — للاستخدام الداخلي» by default, in the header and on every printed page; 04
                      dashboard.confidentiality (or 06 x_confidentiality) replaces the text, false or "" removes it
Facts: every {{f:key}} fact is embedded except the per-record values rec.<record_id>.<field_id> (C2), of which only
those the narrative cites are embedded (the records are embedded anyway).
Printed next-step hints name the scripts by absolute, quoted paths built from this script's own location.

Shareable build (--publish, or 04 dashboard.publish: true) for a page that will be shared as a link:
    records outside the study scope (isolated / excluded) are withheld from the page — the page lists one summary row
    per source sheet (count, reason) and the supporting workbook keeps every row; data notes that touch withheld records
    only are counted, not listed; title-deed numbers (field_meta group "deed") are masked to their last 3 digits in every
    text of the page. Local file paths (05 inputs) are reduced to file names in every build.

Privacy (defence in depth; normalize.py already drops PII columns): keys matching apa_common.PII_KEY_RE are dropped
and block-level personal-data patterns of assets/pii_patterns.json are masked with apa_common.mask_pii in every string;
review-level hits (e.g. a 10-digit number that may be a national ID or a deed) are kept and reported.
"""
from __future__ import annotations

import sys

sys.dont_write_bytecode = True

import argparse  # noqa: E402
import base64  # noqa: E402
import copy  # noqa: E402
import datetime as _dt  # noqa: E402
import json  # noqa: E402
import os  # noqa: E402
import re  # noqa: E402
from decimal import Decimal, ROUND_HALF_UP  # noqa: E402

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import apa_common as ac  # noqa: E402
import facts as facts_mod  # noqa: E402

SKILL_DIR = ac.SKILL_DIR
SCRIPTS_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(SKILL_DIR, "assets", "dashboard_template.html")
ENGINE_PATH = os.path.join(SKILL_DIR, "assets", "metric_engine.js")
FONTS_DIR = os.path.join(SKILL_DIR, "assets", "fonts")
ENGINE_MARK = "/*__APA_ENGINE__*/"
DATA_MARK = "/*__APA_DATA__*/"
FONT_MARK = "/*__APA_FONTS__*/"   # optional (0 or 1 in the template): replaced by the embedded @font-face rules
CONFIDENTIAL_DEFAULT = "سري — للاستخدام الداخلي"

NOT_COMPUTABLE = "غير قابل للاحتساب"
UNAVAILABLE = "غير متاح"

# {{m:id}} · {{m:id.field}} · {{m:id.field.sub}}  (groups: id, field, sub)
TOKEN_RE = re.compile(r"\{\{\s*m:\s*([A-Za-z0-9_\-]+)(?:\.([A-Za-z0-9_]+))?(?:\.([A-Za-z0-9_]+))?\s*\}\}")
# {{f:key}}  (group: key) — same regex object as facts.FACT_TOKEN_RE
FACT_TOKEN_RE = facts_mod.FACT_TOKEN_RE
# Keys that must never reach the dashboard: the shared field-key pattern of assets/pii_patterns.json.
PII_KEY_RE = ac.PII_KEY_RE
_AR_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")

MONEY_FIELDS = {"num", "den", "num_known_total", "prev_sum", "curr_sum", "abs_change", "total", "top_sum",
                "a_sum", "b_sum", "start_change"}
SHARE_FIELDS = {"coverage_num_share", "share"}
PCT_FIELDS = {"pct"}
COUNT_FIELDS = {"n_scope", "n_included", "n_unknown", "n_known", "top", "n", "n_not_applicable"}


def n_ar(n: int, one: str, two: str, few: str, many: str) -> str:
    """Arabic counted noun: 1 and 2 by their own forms, 3–10 plural, 11+ singular accusative."""
    if n == 1:
        return one
    if n == 2:
        return two
    return f"{n} {few if 3 <= n <= 10 else many}"


def eprint(*a):
    print(*a, file=sys.stderr)


# ---------------------------------------------------------------------------
# Number formatting (mirrors the template's fmt* helpers)
# ---------------------------------------------------------------------------
def _round_half_away(v: float, d: int) -> Decimal:
    q = Decimal(1).scaleb(-d)
    dv = Decimal(repr(float(v)))
    r = abs(dv).quantize(q, rounding=ROUND_HALF_UP)
    return -r if dv < 0 and r != 0 else r


def fmt_number(v: float, d: int) -> str:
    r = _round_half_away(v, d)  # never negative zero
    return f"{r:,.{d}f}"


def currency_label(cur) -> str:
    return "ر.س" if (cur or "SAR").upper() == "SAR" else str(cur)


def fmt_money(v: float, cur: str) -> str:
    d = 2 if abs(v) < 10 and float(v) != round(v) else 0
    return f"{fmt_number(v, d)} {currency_label(cur)}"


def fmt_pct(v: float) -> str:
    return fmt_number(v, 2 if abs(v) < 10 else 1) + "%"


def fmt_plain(v: float) -> str:
    return fmt_number(v, 0) if float(v) == round(v) else fmt_number(v, 2)


def fmt_by_unit(v, unit, cur) -> str:
    if v is None:
        return NOT_COMPUTABLE
    if isinstance(v, bool):
        return "نعم" if v else "لا"
    if not isinstance(v, (int, float)):
        return str(v)
    if unit == "SAR":
        return fmt_money(v, cur)
    if unit == "%":
        return fmt_pct(v)
    return fmt_plain(v)


# ---------------------------------------------------------------------------
# Narrative tokens
# ---------------------------------------------------------------------------
class TokenRenderer:
    def __init__(self, metrics: dict, records_by_id: dict, cur: str, metrics_doc: dict = None, facts: dict = None):
        self.metrics = metrics or {}
        doc = metrics_doc or {}
        self.trend = doc.get("trend") if isinstance(doc.get("trend"), dict) else None
        self.breakdowns = doc.get("breakdowns") if isinstance(doc.get("breakdowns"), dict) else {}
        self.facts = facts or {}
        self.records = records_by_id
        self.cur = cur
        self.unknown = []
        self.unknown_facts = []
        self.count = 0

    def _names(self, ids) -> str:
        out = []
        for x in ids:
            if isinstance(x, str):
                r = self.records.get(x)
                out.append((r or {}).get("name") or x)
            elif isinstance(x, dict) and x.get("id"):
                r = self.records.get(x["id"])
                out.append((r or {}).get("name") or x["id"])
        return "، ".join(out) if out else "لا يوجد"

    def _result(self, mid: str):
        if mid in self.metrics and isinstance(self.metrics.get(mid), dict):
            return self.metrics[mid]
        if mid == "trend" and self.trend is not None:
            return self.trend
        if isinstance(self.breakdowns.get(mid), dict):
            return self.breakdowns[mid]
        return None

    def _fmt(self, key: str, v, unit, ref: str):
        """Format one value found under `key` (a result field or a sub-key)."""
        if v is None:
            return NOT_COMPUTABLE
        if isinstance(v, bool):
            return "نعم" if v else "لا"
        if isinstance(v, str):
            return v
        if isinstance(v, list):
            if key == "excluded":
                return fmt_plain(len(v))
            if key == "per_field":
                return "، ".join(f"{p.get('n_known', 0)}/{p.get('n', 0)}" for p in v if isinstance(p, dict)) or UNAVAILABLE
            if all(isinstance(x, (str, dict)) for x in v):
                return self._names(v)
            self.unknown.append(ref)
            return UNAVAILABLE
        if isinstance(v, dict):
            if key in ("comparability", "comparable"):
                return v.get("reason") or v.get("status") or UNAVAILABLE
            self.unknown.append(ref)
            return UNAVAILABLE
        if not isinstance(v, (int, float)):
            return str(v)
        if key in SHARE_FIELDS:
            return fmt_pct(v * 100)
        if key in PCT_FIELDS:
            return fmt_pct(v)
        if key in COUNT_FIELDS or key.startswith("n_"):
            return fmt_plain(v)
        if key == "value":
            return fmt_by_unit(v, unit, self.cur)
        if key in MONEY_FIELDS:
            return fmt_money(v, self.cur) if unit in ("SAR", "%", None) else fmt_plain(v)
        if key == "hhi":
            return fmt_number(v, 0)
        return fmt_plain(v)

    def value(self, mid: str, field, sub=None):
        ref = mid + ("." + field if field else "") + ("." + sub if sub else "")
        res = self._result(mid)
        if res is None:
            self.unknown.append(ref)
            return UNAVAILABLE
        unit = res.get("unit")
        if (field is None or field == "value") and sub is None:
            if "value" not in res:
                self.unknown.append(ref)
                return UNAVAILABLE
            v = res.get("value")
            if v is None:
                if res.get("status") == "not_comparable":
                    return "غير قابل للمقارنة"
                if res.get("revenue_start"):
                    return "بدء إيراد"
                return NOT_COMPUTABLE
            comp = res.get("comparability") if isinstance(res.get("comparability"), dict) else {}
            if res.get("kind") == "growth" and comp.get("status") != "confirmed":
                # metrics.md §5: an unconfirmed growth reads as its change in riyals, never as a percentage
                ac = res.get("abs_change")
                return (fmt_money(ac, self.cur) if isinstance(ac, (int, float)) and not isinstance(ac, bool)
                        else NOT_COMPUTABLE) + " (تغير بالريال؛ لا نسبة)"
            return fmt_by_unit(v, unit, self.cur)
        if field not in res:
            self.unknown.append(ref)
            return UNAVAILABLE
        v = res.get(field)
        if sub is None:
            return self._fmt(field, v, unit, ref)
        if isinstance(v, dict) and sub in v:
            return self._fmt(sub, v.get(sub), unit, ref)
        if isinstance(v, list) and sub.isdigit() and int(sub) < len(v):
            item = v[int(sub)]
            if isinstance(item, (str, dict)) and (field.endswith("_ids") or field in ("excluded", "only_curr", "only_prev")):
                return self._names([item])   # a record: its name, as for the whole list
            return self._fmt(field, item, unit, ref)
        self.unknown.append(ref)
        return UNAVAILABLE

    def fact(self, key: str) -> str:
        f = self.facts.get(key)
        if not isinstance(f, dict) or not isinstance(f.get("value"), (int, float)) or isinstance(f.get("value"), bool):
            self.unknown_facts.append(key)
            return UNAVAILABLE
        v = f["value"]
        if f.get("unit") == "SAR":
            return fmt_money(v, self.cur)
        if f.get("unit") == "pct":
            return fmt_pct(v)
        return fmt_plain(v)

    def render(self, s: str) -> str:
        def rep(m):
            self.count += 1
            return self.value(m.group(1), m.group(2), m.group(3))

        def rep_f(m):
            self.count += 1
            return self.fact(m.group(1))
        return FACT_TOKEN_RE.sub(rep_f, TOKEN_RE.sub(rep, s))

    def walk(self, x):
        if isinstance(x, str):
            return self.render(x) if "{{" in x else x
        if isinstance(x, list):
            return [self.walk(y) for y in x]
        if isinstance(x, dict):
            return {k: self.walk(v) for k, v in x.items()}
        return x


# ---------------------------------------------------------------------------
# Privacy scrub (defence in depth)
# ---------------------------------------------------------------------------
def raw_is_plain_number(raw: str, v) -> bool:
    """True when `raw` is just the printed form of the numeric value v (e.g. raw "512345678" of a valuation).

    Such raw strings are not scanned (a 9-digit amount starting with 5 would look like a mobile number), but a raw
    with a leading 0 or + ("0501234567", "+966…") is a phone written in a numeric cell and is scanned."""
    t = raw.translate(_AR_DIGITS).strip()
    if re.match(r"^(?:\+|00|0\d)", t):
        return False
    try:
        return float(t.replace(",", "").replace("٬", "")) == float(v)
    except ValueError:
        return False


_ORPHAN_BRACKET_RE = re.compile(r"\(\s*" + re.escape(ac.REDACTED))


_DEED_KEY_RE = re.compile(r"deed|صك", re.IGNORECASE)


class Scrubber:
    """Masks personal data in every string of the embedded data: block-level patterns always; review-level ones (a bare
    10-digit number that may be a national ID) too, except under a deed key (deed_no …) where 10 digits are a deed."""
    def __init__(self):
        self.keys_dropped = 0
        self.values_masked = 0
        self.kinds = set()
        self.review_hits = 0

    def text(self, s: str, review: bool = True) -> str:
        out, kinds = ac.mask_pii(s, ("block", "review") if review else ("block",))  # Arabic-Indic digits are translated
        if kinds:
            self.values_masked += 1
            self.kinds.update(kinds)
            # "(050) 123-4567" is matched from the first digit: drop the opening bracket left in front of the mask
            out = _ORPHAN_BRACKET_RE.sub(ac.REDACTED, out)
        if ac.pii_value_hits(out, ("review",)):
            self.review_hits += 1  # a deed number kept under a deed key; reported to the analyst
        return out

    def walk(self, x, key=None, parent=None, review=True):
        if isinstance(x, str):
            if key in ("sha256", "sha256_short"):
                return x
            if key == "raw" and isinstance(parent, dict) and isinstance(parent.get("v"), (int, float)) \
                    and not isinstance(parent.get("v"), bool) and raw_is_plain_number(x, parent["v"]):
                return x
            return self.text(x, review)
        if isinstance(x, list):
            return [self.walk(y, key, parent, review) for y in x]
        if isinstance(x, dict):
            out = {}
            for k, v in x.items():
                if isinstance(k, str) and PII_KEY_RE.search(k):
                    self.keys_dropped += 1
                    continue
                out[k] = self.walk(v, k, x, review and not (isinstance(k, str) and _DEED_KEY_RE.search(k)))
            return out
        return x


def pii_fields(field_meta: dict) -> set:
    out = set()
    for f, m in (field_meta or {}).items():
        m = m or {}
        label = str(m.get("label") or "")
        label_hit = bool(PII_KEY_RE.search(label) or ac.is_pii_header(label)) \
            and m.get("group") in (None, "text", "other", "pii")
        if m.get("group") == "pii" or m.get("kind") == "pii" or PII_KEY_RE.search(f) or label_hit:
            out.add(f)
    return out


def clean_records(records: list, drop: set) -> list:
    out = []
    for r in records or []:
        r = copy.deepcopy(r)
        fields = r.get("fields") or {}
        r["fields"] = {k: v for k, v in fields.items() if k not in drop}
        for key in ("dims", "dims_evidence"):
            if isinstance(r.get(key), dict):
                r[key] = {k: v for k, v in r[key].items() if k not in drop}
        out.append(r)
    return out


def strip_local_paths(metrics_doc: dict) -> dict:
    """05 `inputs` hold the absolute paths of 03/04 on the analyst's machine: the page keeps their file names only (a
    shared page must not reveal the workspace layout)."""
    out = copy.copy(metrics_doc or {})
    inp = out.get("inputs")
    if isinstance(inp, dict):
        out["inputs"] = {k: (os.path.basename(v) if isinstance(v, str) and (os.sep in v or "/" in v) else v)
                         for k, v in inp.items()}
    return out


ABS_PATH_RE = re.compile(r"^(?:/|[A-Za-z]:[\\/]|file://)\S*$")


def strip_abs_paths(x):
    """Every string of the page that is an absolute local path (a run folder in a version comparison, an input file)
    becomes its last component: the page names files, never the analyst's workspace."""
    if isinstance(x, str):
        if ABS_PATH_RE.match(x) and ("/" in x[1:] or "\\" in x):
            return os.path.basename(x.rstrip("/\\")) or x
        return x
    if isinstance(x, list):
        return [strip_abs_paths(y) for y in x]
    if isinstance(x, dict):
        return {k: strip_abs_paths(v) for k, v in x.items()}
    return x


LOCAL_PATH_RE = re.compile(r"(?:file://|/home/|/tmp/|/Users/|[A-Za-z]:\\Users\\)")


def deed_fields(field_meta: dict) -> set:
    """Fields holding a title-deed number: field_meta group "deed" (data-contract §3), or a text field whose id is
    deed_no / ends with .deed_no."""
    out = set()
    for f, m in (field_meta or {}).items():
        m = m or {}
        if m.get("group") == "deed" or f == "deed_no" or f.endswith(".deed_no"):
            out.add(f)
    return out


DEED_DIGITS_RE = re.compile(r"\d{5,}")


def mask_deed(text: str) -> str:
    """«1234567890» -> «•••890»: every run of 5+ digits keeps its last 3 digits only."""
    return DEED_DIGITS_RE.sub(lambda m: "•••" + m.group(0)[-3:], str(text))


def publish_view(analysis: dict, deed_f: set, mapping: dict = None) -> dict:
    """The shareable page (--publish): withhold records outside the study scope and mask deed numbers.
    Returns a summary of what was withheld (also stored in analysis["withheld"])."""
    recs = analysis.get("records") or []
    keep = [r for r in recs if isinstance(r, dict) and r.get("scope") == "in_scope"]
    gone = [r for r in recs if isinstance(r, dict) and r.get("scope") != "in_scope"]
    gone_ids = {r.get("id") for r in gone}
    sheet_reason = {}
    for sd in (mapping or {}).get("sheets") or []:
        if isinstance(sd, dict) and sd.get("sheet") is not None:
            sheet_reason[str(sd["sheet"]).strip()] = sd.get("reason")
    groups = {}
    for r in gone:
        src = r.get("source") or {}
        key = (str(src.get("sheet") or "").strip(), r.get("scope"))
        g = groups.setdefault(key, {"sheet": key[0], "scope": key[1], "count": 0, "record_types": {}, "reasons": {}})
        g["count"] += 1
        rt = r.get("record_type") or "other"
        g["record_types"][rt] = g["record_types"].get(rt, 0) + 1
        why = str(r.get("scope_reason") or "").strip()
        if why:
            g["reasons"][why] = g["reasons"].get(why, 0) + 1
    rows = []
    for g in groups.values():
        reasons = sorted(g["reasons"].items(), key=lambda kv: -kv[1])
        rows.append({"sheet": g["sheet"], "scope": g["scope"], "count": g["count"], "record_types": g["record_types"],
                     "reason": reasons[0][0] if reasons else (sheet_reason.get(g["sheet"]) or None)})
    rows.sort(key=lambda x: (x["scope"] != "isolated", -x["count"], x["sheet"]))
    issues = analysis.get("issues") or []
    kept_issues, n_issues_gone = [], 0
    for x in issues:
        ids = [i for i in (x.get("record_ids") or []) if isinstance(i, str)] if isinstance(x, dict) else []
        if ids and all(i in gone_ids for i in ids):
            n_issues_gone += 1
            continue
        if ids and any(i in gone_ids for i in ids):
            x = dict(x)
            x["record_ids"] = [i for i in ids if i not in gone_ids]
        kept_issues.append(x)
    analysis["records"] = keep
    analysis["issues"] = kept_issues
    # the largest line of a withheld source («sources.<id>.<field>.max») is one row's amount: dropped unless the
    # narrative cites it (sums and counts of the source stay: they are aggregates the scope decision needs)
    gone_src = {(r.get("source") or {}).get("source_id") for r in gone} - {(r.get("source") or {}).get("source_id") for r in keep}
    facts = analysis.get("facts")
    if isinstance(facts, dict) and gone_src:
        cited = json.dumps(analysis.get("narrative") or {}, ensure_ascii=False)
        for k in list(facts):
            m = re.match(r"^sources\.([^.]+)\..+\.(max|max_share)$", str(k))
            if m and m.group(1) in gone_src and ("f:" + k) not in cited:
                del facts[k]
    wbs = analysis.get("workbook_map_summary")
    n_links_gone = 0
    if isinstance(wbs, dict) and isinstance(wbs.get("links"), list):
        # a possible link to a withheld record or to a line of an isolated/excluded sheet quotes that line's name in its
        # basis: kept in the supporting workbook only
        iso_sheets = {str(d.get("sheet") or "").strip() for d in (mapping or {}).get("sheets") or []
                      if isinstance(d, dict) and d.get("decision") in ("isolated", "excluded")}
        iso_sheets |= {str((r.get("source") or {}).get("sheet") or "").strip() for r in gone}
        iso_sheets.discard("")

        def withheld_link(l):
            if not isinstance(l, dict):
                return False
            other = str(l.get("other_ref") or "").split("!")[0].strip()
            return l.get("record") in gone_ids or other in iso_sheets
        n_links_gone = sum(1 for l in wbs["links"] if withheld_link(l))
        wbs["links"] = [l for l in wbs["links"] if not withheld_link(l)]
    # deed numbers: masked on the record and in every text that quotes one (an issue naming a shared deed)
    deeds = set()
    for r in keep:
        for f in deed_f:
            t = (r.get("fields") or {}).get(f)
            if isinstance(t, dict):
                for k in ("v", "raw"):
                    v = t.get(k)
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        v = str(int(v)) if float(v).is_integer() else str(v)
                        t[k] = v
                    if isinstance(v, str) and DEED_DIGITS_RE.search(v):
                        deeds.update(DEED_DIGITS_RE.findall(v))
                        t[k] = mask_deed(v)
    for r in gone:
        for f in deed_f:
            t = (r.get("fields") or {}).get(f)
            if isinstance(t, dict):
                for k in ("v", "raw"):
                    v = t.get(k)
                    if v is not None:
                        deeds.update(DEED_DIGITS_RE.findall(str(int(v)) if isinstance(v, float) and v.is_integer() else str(v)))
    deeds = sorted(deeds, key=len, reverse=True)
    # names of withheld lines quoted in other texts (a workbook note on a duplicated name, a name match between sheets):
    # replaced, unless an in-scope record carries the same name
    kept_names = {str(r.get("name") or "").strip() for r in keep} | {str(r.get("display_name") or "").strip() for r in keep}
    hidden = sorted({str(r.get("name") or "").strip() for r in gone} - kept_names, key=len, reverse=True)
    hidden = [x for x in hidden if len(x) >= 4 and not any(x in k for k in kept_names if k)]
    n_names = 0

    def walk(x):
        nonlocal n_names
        if isinstance(x, str):
            for d in deeds:
                if d in x:
                    x = re.sub(r"(?<!\d)" + re.escape(d) + r"(?!\d)", "•••" + d[-3:], x)
            for nm in hidden:
                if nm in x:
                    n_names += x.count(nm)
                    x = x.replace(nm, "سطر معزول")
            return x
        if isinstance(x, list):
            return [walk(y) for y in x]
        if isinstance(x, dict):
            return {k: walk(v) for k, v in x.items()}
        return x
    if deeds or hidden:
        for key in list(analysis.keys()):
            analysis[key] = walk(analysis[key])
    summary = {"records": rows, "n_records": len(gone), "n_issues": n_issues_gone, "n_links": n_links_gone,
               "n_deeds_masked": len(deeds), "n_names_masked": n_names,
               "note": "نسخة للمشاركة: السجلات خارج نطاق الدراسة ملخصة هنا دون أسماء أو مبالغ، وأرقام الصكوك مقنّعة "
                       "بآخر ثلاث خانات؛ التفصيل الكامل في المرفق الداعم (للمراجعة الداخلية، لا يُنشر)."}
    analysis["withheld"] = summary
    meta = analysis.get("meta")
    if isinstance(meta, dict):
        meta["publish"] = True
    return summary


# ---------------------------------------------------------------------------
# Summaries
# ---------------------------------------------------------------------------
def workbook_summary(wb: dict, mapping: dict) -> dict:
    out = {"available": bool(wb)}
    if wb:
        f = wb.get("file") or {}
        props = f.get("properties") or {}
        out["file"] = {"name": f.get("name"), "size": f.get("size"),
                       "sha256_short": (f.get("sha256") or "")[:12] or None,
                       "created": props.get("created"), "modified": props.get("modified")}
        out["summary"] = wb.get("summary") or {}
        sheets = []
        for s in wb.get("sheets") or []:
            sheets.append({
                "name": s.get("name"), "index": s.get("index"), "state": s.get("state"), "kind": s.get("kind"),
                "role_hint": s.get("role_hint") or s.get("role"), "dimensions": s.get("dimensions"),
                "max_row": s.get("max_row"), "max_col": s.get("max_col"),
                "tables": [{"name": t.get("name"), "ref": t.get("ref")} for t in (s.get("tables") or []) if isinstance(t, dict)],
                "header_row": s.get("primary_header_row"),
                "formulas": (s.get("formulas") or {}).get("count") if isinstance(s.get("formulas"), dict) else None,
            })
        out["sheets"] = sheets
        # record_ids when 01 attaches an anomaly to records: it then counts among that record's data notes (C3)
        out["anomalies"] = [{"id": a.get("id"), "type": a.get("type"), "severity": a.get("severity"),
                             "where": a.get("where"), "detail": a.get("detail"),
                             "record_ids": a.get("record_ids") if isinstance(a.get("record_ids"), list) else None}
                            for a in (wb.get("anomalies") or []) if isinstance(a, dict)]
        out["pii_columns"] = [{"sheet": p.get("sheet"), "col": p.get("col"), "header": p.get("header"),
                               "reason": p.get("reason")} for p in (wb.get("pii_columns") or []) if isinstance(p, dict)]
    m = mapping or {}
    out["decisions"] = [{k: s.get(k) for k in ("sheet", "role", "decision", "period", "reliability",
                                                "relation_to_register", "reason")}
                        for s in (m.get("sheets") or []) if isinstance(s, dict)]
    out["links"] = [{k: l.get(k) for k in ("record", "other_ref", "match", "basis", "used", "reason")}
                    for l in (m.get("links") or []) if isinstance(l, dict)]
    out["derived_fields"] = [{k: d.get(k) for k in ("field", "formula", "when", "reason")}
                             for d in (m.get("derived_fields") or []) if isinstance(d, dict)]
    out["valuation_dates"] = m.get("valuation_dates") or []
    return out


_MSG_TOKEN_RE = re.compile(r"\{\{\s*([mf]):([A-Za-z0-9_\-.]+)\s*\}\}")
_MSG_STATUS_AR = {"unverified": "غير متحققة", "not_comparable": "غير قابلة للمقارنة", "confirmed": "متحققة",
                  "partial": "جزئي", "empty": "فارغ", "not_computable": "غير قابل للاحتساب", "complete": "مكتمل"}


def _msg_text(t: str) -> str:
    """Board-facing wording of a validator message: tokens become «name» (a literal {{m:x}} would read as a broken token)
    and English status words in «…» become Arabic."""
    t = _MSG_TOKEN_RE.sub(lambda m: f"«{m.group(2)}»", t)
    return re.sub(r"«(unverified|not_comparable|confirmed|partial|empty|not_computable|complete)»",
                  lambda m: f"«{_MSG_STATUS_AR[m.group(1)]}»", t)


def _msg(x) -> str:
    """A validator finding as one Arabic line: its Arabic check title (not the English check id) and its detail."""
    if isinstance(x, str):
        return _msg_text(x)
    if isinstance(x, dict):
        for k in ("message", "detail", "text", "reason"):
            if isinstance(x.get(k), str) and x.get(k):
                pre = x.get("title") or x.get("check") or x.get("name") or x.get("id")
                if pre and re.fullmatch(r"[a-z0-9_]+", str(pre)):
                    pre = None  # an English check id is not shown to the board
                return _msg_text(f"{pre}: {x[k]}" if pre else x[k])
        return json.dumps(x, ensure_ascii=False)
    return _msg_text(str(x))


def blocking_items(v: dict) -> list:
    items = v.get("blocking")
    if items is None:
        items = v.get("errors")
    if items is None:
        items = [c for c in (v.get("checks") or []) if isinstance(c, dict) and not c.get("passed", True)
                 and (c.get("blocking") or _level(c) in ("blocking", "error"))]
    return items if isinstance(items, list) else []


def _level(c: dict) -> str:
    """validate_analysis.py writes `level` (blocking | warning | acknowledged | info); older files `severity`."""
    return str(c.get("level") or c.get("severity") or "")


def blocking_count(v: dict) -> int:
    for src in (v, v.get("summary") or {}):
        bc = src.get("blocking_count")
        if isinstance(bc, int):
            return bc
    return len(blocking_items(v))


def validation_summary(v) -> dict:
    if not isinstance(v, dict):
        return {"available": False, "passed": None, "blocking_count": None, "warning_count": None,
                "blocking": [], "warnings": []}
    warns = v.get("warnings")
    failed = [c for c in (v.get("checks") or []) if isinstance(c, dict) and not c.get("passed", True)]
    if warns is None:
        warns = [c for c in failed if _level(c) in ("warning", "warn")]
    warns = warns if isinstance(warns, list) else []
    # warnings the narrative acknowledged with a reason (06 acknowledged_warnings): listed apart, not counted
    acks = [c for c in failed if _level(c) == "acknowledged"]
    wc = v.get("warning_count")
    if not isinstance(wc, int):
        wc = (v.get("summary") or {}).get("warning_count")
    if not isinstance(wc, int):
        wc = len(warns)
    ac_n = v.get("acknowledged_count")
    if not isinstance(ac_n, int):
        ac_n = len(acks)
    bc = blocking_count(v)
    return {"available": True, "passed": v.get("passed", bc == 0) if bc == 0 else False,
            "blocking_count": bc, "warning_count": wc, "acknowledged_count": ac_n,
            "blocking": [_msg(x) for x in blocking_items(v)][:50], "warnings": [_msg(x) for x in warns][:50],
            "acknowledged": [_msg(x) for x in acks][:50],
            "checked_at": v.get("checked_at") or v.get("generated_at")}


# ---------------------------------------------------------------------------
# Meta helpers: study date, confidentiality, printed hints, fonts, facts
# ---------------------------------------------------------------------------
# study_date_meta lives in apa_common (shared with export_workbook.py «ملخص»)
study_date_meta = ac.study_date_meta


def confidentiality_label(config: dict, narrative: dict):
    """Confidentiality marking for the header and every printed page: on by default (CONFIDENTIAL_DEFAULT);
    04 dashboard.confidentiality, else 06 x_confidentiality, replaces the text (a string) or removes it (false / "")."""
    dash = (config or {}).get("dashboard")
    for holder, key in ((dash, "confidentiality"), (narrative, "x_confidentiality")):
        if isinstance(holder, dict) and key in holder:
            v = holder.get(key)
            if isinstance(v, str):
                return v.strip() or None
            return CONFIDENTIAL_DEFAULT if v is True else None
    return CONFIDENTIAL_DEFAULT


def shell_path(p: str) -> str:
    """A path for a printed shell command: absolute and double-quoted (spaces, quotes, $ and backticks escaped)."""
    return '"' + re.sub(r'(["\\$`])', r"\\\1", os.path.abspath(p)) + '"'


def script_hint(interp: str, script: str, *args: str) -> str:
    """«python3 "/abs/skill/scripts/x.py" …»: the script by its absolute path from this script's own folder."""
    return " ".join([interp, shell_path(os.path.join(SCRIPTS_DIR, script))] + list(args))


def font_face_css(fonts_dir: str = FONTS_DIR):
    """(css, problem): @font-face rules embedding the woff2 files listed in fonts_dir/fonts.json as base64 data URIs,
    or ("", why) when the fonts are missing (the page then falls back to the system fonts of its font stack)."""
    man_p = os.path.join(fonts_dir, "fonts.json")
    try:
        man = ac.load_json(man_p)
    except (OSError, ValueError) as exc:
        return "", f"تعذر قراءة {man_p}: {exc}"
    family = re.sub(r"[^A-Za-z0-9 \-]", "", str(man.get("family") or ""))
    ranges = man.get("unicode_ranges") or {}
    rules = []
    for face in man.get("faces") or []:
        path = os.path.join(fonts_dir, str(face.get("file") or ""))
        try:
            with open(path, "rb") as fh:
                data = fh.read()
        except OSError:
            return "", f"ملف الخط غير موجود: {path}"
        if data[:4] != b"wOF2":
            return "", f"ملف الخط ليس WOFF2: {path}"
        rng = ranges.get(face.get("subset"))
        rules.append("@font-face{font-family:'%s';font-style:normal;font-weight:%d;font-display:swap;"
                     "src:url(data:font/woff2;base64,%s) format('woff2');%s}"
                     % (family, int(face.get("weight") or 400), base64.b64encode(data).decode("ascii"),
                        ("unicode-range:%s;" % re.sub(r"[^0-9A-Fa-fU+,\-]", "", str(rng))) if rng else ""))
    if not family or not rules:
        return "", f"لا توجد خطوط معرّفة في {man_p}"
    return "\n".join(rules), None


_CITED_REC_FACT_RE = re.compile(r"f:\s*(rec\.[A-Za-z0-9_.\-]+)")


def facts_for_page(facts, narrative) -> dict:
    """Facts embedded in the page: every fact except the per-record values rec.<id>.<field> (C2), of which only the
    ones the narrative cites ({{f:rec.…}} tokens, f:rec.… evidence refs); the page holds the records themselves."""
    blob = json.dumps(narrative, ensure_ascii=False) if narrative else ""
    keep = [k for k in facts.keys() if not str(k).startswith("rec.")]
    keep += sorted({k.rstrip(".") for k in _CITED_REC_FACT_RE.findall(blob)} - set(keep))
    out = {}
    for k in keep:
        f = facts.get(k)
        if isinstance(f, dict):
            out[k] = {"value": f.get("value"), "unit": f.get("unit"), "label": f.get("label")}
    return out


def build_meta(mapping: dict, rec_doc: dict, config: dict, wb: dict, draft: bool, records: list,
               narrative: dict = None) -> dict:
    m = mapping or {}
    ent = m.get("entity") or {}
    dp = m.get("data_period") or {}
    src = rec_doc.get("source") or {}
    wbf = (wb or {}).get("file") or {}
    sha = src.get("sha256") or wbf.get("sha256") or ""
    fname = src.get("name") or wbf.get("name") or (os.path.basename(src.get("file")) if src.get("file") else None)
    counts = {"total": len(records)}
    for sc in ac.SCOPES:
        counts[sc] = sum(1 for r in records if r.get("scope") == sc)
    val = config.get("valuation") or {}
    out = {
        "entity": ent.get("name") if isinstance(ent, dict) else (ent or None),
        "entity_evidence": ent.get("evidence") if isinstance(ent, dict) else None,
        "data_period": dp.get("label") if isinstance(dp, dict) else (dp or None),
        "data_period_evidence": dp.get("evidence") if isinstance(dp, dict) else None,
        "analysis_date": m.get("analysis_date") or _dt.date.today().isoformat(),
    }
    out.update(study_date_meta(m, wb))
    out.update({
        "confidentiality": confidentiality_label(config, narrative),
        # one line under the page and on paper: what the page is (e.g. an analytical study, not an official document)
        "disclaimer": ((config.get("dashboard") or {}).get("disclaimer") or None) if isinstance(config.get("dashboard"), dict) else None,
        "currency": config.get("currency") or rec_doc.get("currency") or m.get("currency") or "SAR",
        "file_name": fname,
        "sha256_short": sha[:12] if sha else None,
        "valuation": {k: val.get(k) for k in ("field", "year", "basis", "date_text", "label", "note", "reason")},
        "valuation_dates": m.get("valuation_dates") or [],
        "counts": counts,
        "draft": bool(draft),
        "built_at": _dt.datetime.now().replace(microsecond=0).isoformat(),
        "generator": "build_dashboard.py",
    })
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def load_opt(path):
    if not os.path.exists(path):
        return None
    return ac.load_json(path)


def self_test() -> int:
    """Token grammar and privacy scrub on synthetic values (no run folder needed). Exit 0 = ok, 1 = a case failed."""
    fails = []

    def eq(name, got, exp):
        if got != exp:
            fails.append(f"{name}: {got!r} != {exp!r}")

    metrics = {"m1": {"unit": "SAR", "value": 1234567.4, "status": "complete", "n_included": 3,
                      "comparability": {"status": "unverified", "reason": "سبب"}, "per_field": [{"n_known": 2, "n": 3}],
                      "top_ids": ["A-1"], "start_change": -50000},
               "g": {"unit": "%", "value": None, "status": "not_comparable"}}
    doc = {"trend": {"same_records": {"n": 4}, "points": [{"value": 10}]}, "breakdowns": {"b1": {"total": 900}}}
    facts = {"records.isolated": {"value": 7, "unit": "count"}, "wb.WB-1.recorded": {"value": 2500, "unit": "SAR"}}
    tr = TokenRenderer(metrics, {"A-1": {"name": "عمارة"}}, "SAR", metrics_doc=doc, facts=facts)
    eq("value", tr.render("{{m:m1}}"), "1,234,567 ر.س")
    eq("field", tr.render("{{m:m1.n_included}}"), "3")
    eq("money field", tr.render("{{m:m1.start_change}}"), "-50,000 ر.س")
    eq("nested dict", tr.render("{{m:m1.comparability.status}}"), "unverified")
    eq("ids", tr.render("{{m:m1.top_ids}}"), "عمارة")
    eq("list index", tr.render("{{m:m1.top_ids.0}}"), "عمارة")
    eq("trend point", tr.render("{{m:trend.points.0}}"), "غير متاح")   # a dict: not a printable value
    eq("not comparable", tr.render("{{m:g}}"), "غير قابل للمقارنة")
    eq("trend nested", tr.render("{{m:trend.same_records.n}}"), "4")
    eq("breakdown", tr.render("{{m:b1.total}}"), "900 ر.س")
    eq("fact count", tr.render("{{f:records.isolated}}"), "7")
    eq("fact money", tr.render("{{f:wb.WB-1.recorded}}"), "2,500 ر.س")
    eq("unknown", (tr.render("{{m:nope}} {{f:nope}}"), tr.unknown[-1:], tr.unknown_facts), ("غير متاح غير متاح", ["nope"], ["nope"]))
    try:
        ac._self_test_pii()  # value and header cases of assets/pii_patterns.json
    except AssertionError as exc:
        fails.append(f"pii_patterns.json: {exc}")
    sc = Scrubber()
    for x in (ac.PII_PATTERNS.get("tests") or {}).get("must_match", []):
        out = sc.text("قبل " + x + " بعد")
        if ac.pii_value_hits(x, ("block",)) and (ac.REDACTED not in out or ac.pii_value_hits(out, ("block",))):
            fails.append(f"block pattern not masked: {x[:3]}…")
    for x in (ac.PII_PATTERNS.get("tests") or {}).get("must_not_match", []):
        eq(f"kept {x}", sc.text(x), x)
    eq("bracket", sc.text("جوال (050) 123-4567"), "جوال " + ac.REDACTED)
    eq("raw amount kept", sc.walk({"v": 512345678, "s": "value", "raw": "512345678"}), {"v": 512345678, "s": "value", "raw": "512345678"})
    eq("raw phone masked", sc.walk({"v": 501234567, "s": "value", "raw": "0501234567"})["raw"], ac.REDACTED)
    eq("pii key dropped", sc.walk({"supervisor_phone": "x", "name": "y"}), {"name": "y"})
    # «تاريخ إعداد الدراسة»: the workbook's last-modified date is context, never the study date (fictional dates)
    wbp = {"file": {"properties": {"modified": "2030-04-09T08:15:00"}}}
    sdm = study_date_meta({"study_date": {"value": "2030-04-09", "basis": "خصائص المصنف", "evidence": "file.properties.modified"}}, wbp)
    eq("study date from file properties", (sdm["study_date"], sdm["study_date_stated"], sdm["file_modified"]), (None, False, "2030-04-09"))
    sdm = study_date_meta({"study_date": {"value": "2030-04-09", "basis": None, "evidence": "—"}}, wbp)
    eq("study date equal to last-modified", (sdm["study_date"], sdm["file_modified"]), (None, "2030-04-09"))
    sdm = study_date_meta({"study_date": {"value": "2030-01-15", "basis": "عنوان البيان", "evidence": "ورقة تجريبية!B2"}}, wbp)
    eq("study date stated in a cell", (sdm["study_date"], sdm["study_date_stated"], sdm["file_modified"]), ("2030-01-15", True, "2030-04-09"))
    eq("study date absent", study_date_meta({}, None)["study_date"], None)
    # confidentiality: on by default; the 04 dashboard setting replaces or removes it
    eq("confidential default", confidentiality_label({}, {}), CONFIDENTIAL_DEFAULT)
    eq("confidential off", confidentiality_label({"dashboard": {"confidentiality": False}}, {}), None)
    eq("confidential text", confidentiality_label({}, {"x_confidentiality": "داخلي"}), "داخلي")
    # printed hints: absolute, quoted paths
    hint = script_hint("node", "check_dashboard.mjs", "--out", shell_path("مجلد تجريبي/run"))
    eq("hint quoted absolute", hint.startswith('node "' + os.path.join(SCRIPTS_DIR, "check_dashboard.mjs") + '"')
       and ('"' + os.path.abspath("مجلد تجريبي/run") + '"') in hint, True)
    eq("hint escapes", shell_path('/tmp/a"b$c'), '"/tmp/a\\"b\\$c"')
    # embedded font: six base64 @font-face rules (assets/fonts)
    css, why = font_face_css()
    eq("fonts embedded", (why, css.count("@font-face"), "data:font/woff2;base64," in css), (None, 6, True))
    # per-record facts: only the cited ones are embedded
    fx = {"records.isolated": {"value": 1, "unit": "count", "label": "x"},
          "rec.A-1.val.demo": {"value": 5, "unit": "SAR", "label": "y"}, "rec.A-2.val.demo": {"value": 6, "unit": "SAR", "label": "z"}}
    eq("rec facts cited only", sorted(facts_for_page(fx, {"t": "قيمة {{f:rec.A-1.val.demo}}"})), ["rec.A-1.val.demo", "records.isolated"])
    # local paths reduced to file names; the shareable view withholds out-of-scope records and masks deeds (fictional)
    eq("abs paths in any string", strip_abs_paths({"prev": {"run_dir": "/srv/x/run_a/"}, "t": "نص / عادي", "u": "/"}),
       {"prev": {"run_dir": "run_a"}, "t": "نص / عادي", "u": "/"})
    eq("paths stripped", strip_local_paths({"inputs": {"records": "/home/x/run/03_records.json", "source_sha256": "ab"}})["inputs"],
       {"records": "03_records.json", "source_sha256": "ab"})
    an = {"meta": {}, "records": [
        {"id": "A-1", "scope": "in_scope", "source": {"sheet": "س"}, "fields": {"deed_no": {"v": 9876543210, "s": "value", "raw": "9876543210"}}},
        {"id": "B-1", "scope": "isolated", "source": {"sheet": "مخفية"}, "record_type": "asset", "scope_reason": "لم يثبت",
         "name": "عقار تجريبي", "fields": {"deed_no": {"v": "1112223334", "s": "value", "raw": "1112223334"}}}],
        "issues": [{"id": "DQ-1", "record_ids": ["B-1"], "detail": "x"}, {"id": "DQ-2", "record_ids": ["A-1"], "detail": "صك 9876543210 مشترك"}]}
    smy = publish_view(an, {"deed_no"})
    eq("publish records", [r["id"] for r in an["records"]], ["A-1"])
    eq("publish summary", (smy["records"][0]["sheet"], smy["records"][0]["count"], smy["n_issues"]), ("مخفية", 1, 1))
    eq("publish deed masked", (an["records"][0]["fields"]["deed_no"]["raw"], an["issues"][0]["detail"]), ("•••210", "صك •••210 مشترك"))
    eq("publish no names", "عقار تجريبي" in json.dumps(an, ensure_ascii=False), False)
    for f in fails:
        eprint("فشل: " + f)
    print(f"build_dashboard --self-test: {'ناجح' if not fails else 'فشل ' + str(len(fails))}")
    return 0 if not fails else 1


def main(argv=None) -> int:
    if argv is None:
        argv = sys.argv[1:]
    if "--self-test" in argv:
        return self_test()
    ap = argparse.ArgumentParser(description="بناء analysis.json وdashboard.html من مجلد التشغيل")
    ap.add_argument("--run-dir", required=True, help="مجلد التشغيل (RUN)")
    ap.add_argument("--draft", action="store_true", help="بناء مسودة رغم الأخطاء المانعة أو غياب السرد/التحقق (تظهر لافتة «مسودة»)")
    ap.add_argument("--publish", action="store_true",
                    help="نسخة للمشاركة برابط: تُلخَّص السجلات خارج النطاق دون أسماء أو مبالغ وتُقنَّع أرقام الصكوك "
                         "(يكافئ dashboard.publish: true في 04)")
    args = ap.parse_args(argv)
    run = args.run_dir
    if not os.path.isdir(run):
        eprint(f"خطأ: مجلد التشغيل غير موجود: {run}")
        return 1

    def need(name):
        p = os.path.join(run, name)
        if not os.path.exists(p):
            eprint(f"خطأ: الملف المطلوب غير موجود: {p}")
            return None
        try:
            return ac.load_json(p)
        except (OSError, ValueError) as exc:
            eprint(f"خطأ: تعذر قراءة {p}: {exc}")
            return None

    rec_doc = need("03_records.json")
    config = need("04_metrics_config.json")
    metrics = need("05_metrics.json")
    if rec_doc is None or config is None or metrics is None:
        return 1
    try:
        wb = load_opt(os.path.join(run, "01_workbook_map.json"))
        mapping = load_opt(os.path.join(run, "02_mapping.json"))
        narrative = load_opt(os.path.join(run, "06_narrative.json"))
        validation = load_opt(os.path.join(run, "07_validation.json"))
    except (OSError, ValueError) as exc:
        eprint(f"خطأ: ملف JSON غير صالح في مجلد التشغيل: {exc}")
        return 1

    draft = bool(args.draft)
    if narrative is None:
        if not draft:
            eprint("خطأ: 06_narrative.json غير موجود. اكتب السرد أولًا، أو استخدم --draft لمعاينة مسودة.")
            return 1
        eprint("تنبيه: 06_narrative.json غير موجود؛ تُبنى مسودة دون سرد.")
        narrative = {"contract_version": ac.CONTRACT_VERSION}
    if validation is None:
        if not draft:
            eprint("خطأ: 07_validation.json غير موجود. شغّل أولًا:\n  "
                   + script_hint("python3", "validate_analysis.py", "--run-dir", shell_path(run))
                   + "\nأو استخدم --draft لمعاينة مسودة.")
            return 1
        eprint("تنبيه: 07_validation.json غير موجود؛ تُبنى مسودة لم تجتز التحقق.")
    else:
        bc = blocking_count(validation)
        if bc > 0 and not draft:
            eprint(f"توقف البناء: يوجد {bc} خطأ مانع في 07_validation.json. أصلحها ثم أعد التحقق والبناء، "
                   "أو استخدم --draft لمعاينة مسودة لا تُسلَّم:")
            for m in [_msg(x) for x in blocking_items(validation)][:10]:
                eprint(f"  • {m}")
            return 2
        if bc > 0:
            eprint(f"تنبيه: {bc} خطأ مانع في التحقق؛ تُبنى مسودة.")
    if isinstance(narrative, dict) and narrative.get("draft"):
        draft = True
    if validation is None or blocking_count(validation or {}) > 0:
        draft = True

    for p in (TEMPLATE_PATH, ENGINE_PATH):
        if not os.path.exists(p):
            eprint(f"خطأ: ملف المهارة غير موجود: {p}")
            return 1
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as fh:
        template = fh.read()
    with open(ENGINE_PATH, "r", encoding="utf-8") as fh:
        engine = fh.read()
    for mark in (ENGINE_MARK, DATA_MARK):
        n = template.count(mark)
        if n != 1:
            eprint(f"خطأ: القالب يجب أن يحتوي العلامة {mark} مرة واحدة (وُجدت {n}).")
            return 1
    if template.count(FONT_MARK) > 1:
        eprint(f"خطأ: القالب يحتوي العلامة {FONT_MARK} أكثر من مرة.")
        return 1

    vocab = ac.load_vocab()
    field_meta = rec_doc.get("field_meta") or (mapping or {}).get("field_meta") or {}
    drop = pii_fields(field_meta)
    records = clean_records(rec_doc.get("records") or [], drop)
    field_meta = {k: v for k, v in field_meta.items() if k not in drop}
    rec_by_id = {r.get("id"): r for r in records if isinstance(r, dict)}

    cur = config.get("currency") or rec_doc.get("currency") or "SAR"
    try:
        facts = facts_mod.compute_facts(run)
    except (OSError, ValueError) as exc:
        eprint(f"تنبيه: تعذر حساب الحقائق ({{{{f:…}}}}) من مجلد التشغيل: {exc}")
        facts = {}
    tr = TokenRenderer(metrics.get("metrics") or {}, rec_by_id, cur, metrics_doc=metrics, facts=facts)
    narrative_r = tr.walk(narrative)

    issues = list(rec_doc.get("issues") or [])
    for x in (narrative_r.get("data_issues_extra") or []) if isinstance(narrative_r, dict) else []:
        if isinstance(x, dict):
            y = dict(x)
            y.setdefault("severity", "medium")
            y.setdefault("record_ids", [])
            y["origin"] = "narrative"
            issues.append(y)

    analysis = {
        "contract_version": ac.CONTRACT_VERSION,
        "meta": build_meta(mapping, rec_doc, config, wb, draft, records, narrative),
        "vocab": vocab,
        "field_meta": field_meta,
        "records": records,
        "excluded_rows": rec_doc.get("excluded_rows") or [],
        "config": config,
        "metrics": strip_local_paths(metrics),
        "narrative": narrative_r,
        "workbook_map_summary": workbook_summary(wb, mapping),
        "issues": issues,
        "validation": validation_summary(validation),
        # counts outside 05 (isolated records, hidden sheets …) — the {{f:key}} vocabulary, read by the template;
        # per-record values (rec.*) only when the narrative cites them
        "facts": facts_for_page(facts, narrative),
    }
    publish = bool(args.publish or (isinstance(config.get("dashboard"), dict) and config["dashboard"].get("publish") is True))
    withheld = publish_view(analysis, deed_fields(field_meta), mapping) if publish else None
    if not publish and isinstance(analysis.get("meta"), dict):
        analysis["meta"]["publish"] = False
    analysis = strip_abs_paths(analysis)
    sc = Scrubber()
    analysis = sc.walk(analysis)

    out_json = os.path.join(run, "analysis.json")
    ac.save_json(out_json, analysis)

    data_js = json.dumps(analysis, ensure_ascii=False, separators=(",", ":"))
    data_js = data_js.replace("</", "<\\/").replace("<!--", "<\\u0021--")
    if re.search(r"</script|<!--", engine, re.IGNORECASE):
        engine = engine.replace("</", "<\\/").replace("<!--", "<\\!--")
    fonts_css, font_problem = font_face_css() if FONT_MARK in template else ("", None)
    html = template.replace(FONT_MARK, fonts_css, 1)
    html = html.replace(ENGINE_MARK, engine, 1)
    # the data block is inserted last so that nothing inside the data can be mistaken for a marker
    html = html.replace(DATA_MARK, "const DATA = " + data_js + ";", 1)
    out_html = os.path.join(run, "dashboard.html")
    with open(out_html, "w", encoding="utf-8") as fh:
        fh.write(html)
    lp = LOCAL_PATH_RE.findall(data_js)
    if lp:
        eprint("تنبيه خصوصية: بيانات اللوحة تحوي " + n_ar(len(lp), "مسارًا محليًا واحدًا", "مسارين محليين", "مسارات محلية", "مسارًا محليًا")
               + " (/home/ أو /tmp/ أو file://)؛ راجع مصدره قبل مشاركة اللوحة.")

    if tr.unknown:
        eprint("تنبيه: رموز {{m:...}} تشير إلى مؤشرات أو حقول غير موجودة في 05_metrics.json وعُرضت «غير متاح»: "
               + "، ".join(sorted(set(tr.unknown))))
    if tr.unknown_facts:
        eprint("تنبيه: رموز {{f:...}} تشير إلى حقائق غير موجودة (يسردها "
               + script_hint("python3", "facts.py", "--run-dir", shell_path(run)) + ") "
               "وعُرضت «غير متاح»: " + "، ".join(sorted(set(tr.unknown_facts))))
    if font_problem:
        eprint("تنبيه: لم يُضمَّن الخط العربي في اللوحة (" + font_problem + ")؛ تُستخدم خطوط النظام.")
    if drop:
        eprint(f"حُذفت حقول بيانات شخصية من السجلات: {len(drop)} حقل.")
    if sc.keys_dropped or sc.values_masked:
        parts = []
        if sc.keys_dropped:
            parts.append("حُذف " + n_ar(sc.keys_dropped, "مفتاح واحد", "مفتاحان", "مفاتيح", "مفتاحًا"))
        if sc.values_masked:
            parts.append("حُجب نص بيانات شخصية في " + n_ar(sc.values_masked, "قيمة واحدة", "قيمتين", "قيم", "قيمة")
                         + f" ({'، '.join(sorted(sc.kinds)) or '—'}) واستُبدل بـ«{ac.REDACTED}»")
        eprint("حماية الخصوصية: " + "، و".join(parts) + ".")
    if sc.review_hits:
        eprint("تنبيه خصوصية: " + n_ar(sc.review_hits, "نص واحد يحوي", "نصّان يحويان", "نصوص تحوي", "نصًا يحوي")
               + " رقمًا من 10 خانات يبدأ بـ1 أو 2 (قد يكون رقم هوية أو صكًا)؛ لم يُحجب. راجعه، واحذفه من الربط أو السرد إن كان رقم هوية.")
    kb = lambda p: f"{os.path.getsize(p) / 1024:,.0f} ك.ب"  # noqa: E731
    if withheld is not None:
        print("  نسخة للمشاركة: حُجب " + n_ar(withheld["n_records"], "سجل واحد", "سجلان", "سجلات", "سجلًا")
              + " خارج النطاق (ملخصة حسب الورقة) و" + n_ar(withheld["n_issues"], "ملاحظة بيانات واحدة", "ملاحظتا بيانات", "ملاحظات بيانات", "ملاحظة بيانات")
              + " تخصها وحدها، وقُنّع " + n_ar(withheld["n_deeds_masked"], "رقم صك واحد", "رقما صك", "أرقام صكوك", "رقم صك") + ".")
    state = "مسودة لم تجتز التحقق الآلي" if draft else "نسخة نهائية"
    print(f"تم البناء — {state}: {out_json} ({kb(out_json)}) · {out_html} ({kb(out_html)})")
    print(f"  السجلات في الصفحة: {len(analysis.get('records') or [])} من {len(records)} · المؤشرات: {len((metrics.get('metrics') or {}))} · رموز السرد المعالجة: {tr.count}")
    print("  الخطوة التالية: " + script_hint("node", "check_dashboard.mjs", "--html", shell_path(out_html), "--out", shell_path(run)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
