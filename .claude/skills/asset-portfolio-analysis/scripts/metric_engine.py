"""Metric DSL engine (Python twin of assets/metric_engine.js).

Implements data-contract.md §5 literally. The JS engine embedded in the dashboard
must return the same `value`, `status` and `n_included` for every metric and filter
combination (checked by the parity cases in 05_metrics.json), so this module avoids
anything that is not in the spec. Implementation choices where the spec is silent are
listed in IMPLEMENTATION_NOTES and must be mirrored by the JS twin.

Pure stdlib; no I/O. Imported by compute_metrics.py and tests/test_metric_engine.py.
"""
from __future__ import annotations

import os
import sys
from typing import Any, Dict, Iterable, List, Optional, Tuple

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from apa_common import KNOWN_STATUSES  # noqa: E402

IMPLEMENTATION_NOTES = [
    "R (records in scope of a metric) is ordered by id ascending (plain string order); every sum is accumulated in that order.",
    "Check order for every metric: (1) a placeholder resolving to null -> not_computable; (2) growth with comparable_with_prev.status == not_comparable -> not_comparable; (3) |R| == 0 -> empty (count: value 0, others: value null); (4) kind rules.",
    "Early-return statuses (not_computable by placeholder, not_comparable) have n_included 0, included_ids [], excluded [].",
    "n_included: count |R|; sum |K|; ratio/diff/growth |I|; concentration |K| (known and v > 0); coverage |R|.",
    "eq/in use strict equality (no bool/number coercion; lists and objects never equal). has: list attr -> membership (strict), scalar attr -> strict equality.",
    "A field missing from a record has status null: known() false, status() never matches; sum excluded reason 'absent'.",
    "diff: not_computable with reason «لا توجد بيانات مصروفات» when no record of R has the key b in fields; value = sum over I (id order) of (a.v - b.v).",
    "growth: excluded = R minus I with reasons prev_unknown|curr_unknown|both_unknown; status partial if excluded non-empty; value null when prev_sum <= 0.",
    "concentration: K sorted by v desc then id asc; total summed in id order; top sum in rank order; status partial only if some record of R has an unknown field (known zeros/negatives are excluded with reason 'nonpositive' but do not make it partial).",
    "coverage: value = min(n_known_f / |R|) * 100; empty when |R| == 0 (value null).",
    "breakdown: rows sorted by value desc (null last) then key (string) asc; list dims give one row per element (empty list -> '__none__'); a group with no known value has value null.",
    "Dashboard filters: every key except 'period' whose value is not null/'' adds {eq:[dim,value]}, or {has:[dim,value]} when any record holds a list for that dim.",
    "Optional spec key text_by_period[period] overrides label/short_label/definition/scope_text/basis/period_text/valuation_date/source_text/unavailable_reason for that period.",
    "growth spec keys prev/curr default to '{REV_PREV}'/'{REV}'.",
]

TEXT_KEYS = ("label", "short_label", "definition", "scope_text", "basis", "valuation_date",
             "period_text", "source_text", "unavailable_reason")
PLACEHOLDERS = ("{REV}", "{REV_PREV}", "{EXP}", "{VAL}")
NO_EXPENSE_REASON = "لا توجد بيانات مصروفات"
DEFAULT_UNAVAILABLE = "الحقل غير متاح لهذه الفترة"
NONE_KEY = "__none__"
OVERLAP_NOTE = "قد تتداخل الفئات"


class Unresolved(Exception):
    """Raised internally when a placeholder resolves to null."""

    def __init__(self, placeholder: str):
        super().__init__(placeholder)
        self.placeholder = placeholder


# ---------------------------------------------------------------------------
# Record access
# ---------------------------------------------------------------------------
def get_attr(rec: dict, attr: str) -> Any:
    if attr in ("id", "record_type", "scope"):
        return rec.get(attr)
    dims = rec.get("dims") or {}
    return dims.get(attr)


def get_tv(rec: dict, field: str) -> Optional[dict]:
    fields = rec.get("fields") or {}
    tv = fields.get(field)
    return tv if isinstance(tv, dict) else None


def field_status(rec: dict, field: str) -> Optional[str]:
    tv = get_tv(rec, field)
    return None if tv is None else tv.get("s")


def is_known(rec: dict, field: str) -> bool:
    return field_status(rec, field) in KNOWN_STATUSES


def fv(rec: dict, field: str) -> Any:
    """The numeric value of a known field (caller checks is_known first)."""
    return get_tv(rec, field)["v"]


def _is_num(x: Any) -> bool:
    return isinstance(x, (int, float)) and not isinstance(x, bool)


def strict_eq(a: Any, b: Any) -> bool:
    """Equality without Python's bool/number coercion (mirrors JS ===)."""
    if isinstance(a, (list, dict)) or isinstance(b, (list, dict)):
        return False
    if isinstance(a, bool) or isinstance(b, bool):
        return isinstance(a, bool) and isinstance(b, bool) and a == b
    if _is_num(a) or _is_num(b):
        return _is_num(a) and _is_num(b) and a == b
    if a is None or b is None:
        return a is None and b is None
    return type(a) is type(b) and a == b


# ---------------------------------------------------------------------------
# Placeholders
# ---------------------------------------------------------------------------
def placeholder_context(config: dict, period_id: Optional[str]) -> Dict[str, Optional[str]]:
    period = None
    for p in config.get("periods") or []:
        if p.get("id") == period_id:
            period = p
            break
    val = config.get("valuation") or {}
    return {
        "{REV}": (period or {}).get("rev_field"),
        "{REV_PREV}": (period or {}).get("prev_rev_field"),
        "{EXP}": (period or {}).get("exp_field"),
        "{VAL}": val.get("field"),
    }


def resolve_field(name: Any, ctx: Dict[str, Optional[str]]) -> str:
    """Resolve a field reference. Whole-string placeholders only; null -> Unresolved."""
    if isinstance(name, str) and name.startswith("{") and name.endswith("}"):
        resolved = ctx.get(name)
        if not resolved:
            raise Unresolved(name)
        return resolved
    if not isinstance(name, str) or not name:
        raise Unresolved(str(name))
    return name


def resolve_where(where: Any, ctx: Dict[str, Optional[str]]) -> Any:
    """Return a copy of a Where with field placeholders resolved (raises Unresolved)."""
    if where is None:
        return None
    if not isinstance(where, dict) or len(where) != 1:
        raise ValueError(f"Where غير صالح: {where!r}")
    (op, arg), = where.items()
    if op in ("all", "any"):
        return {op: [resolve_where(w, ctx) for w in arg]}
    if op == "not":
        return {"not": resolve_where(arg, ctx)}
    if op == "known":
        return {"known": resolve_field(arg, ctx)}
    if op in ("status", "gt"):
        return {op: [resolve_field(arg[0], ctx), arg[1]]}
    if op in ("eq", "in", "has"):
        return {op: list(arg)}
    raise ValueError(f"عامل Where غير معروف: {op}")


# ---------------------------------------------------------------------------
# Where predicate
# ---------------------------------------------------------------------------
def eval_where(where: Any, rec: dict) -> bool:
    """Evaluate a (resolved) Where predicate on one record. None means True."""
    if where is None:
        return True
    (op, arg), = where.items()
    if op == "all":
        return all(eval_where(w, rec) for w in arg)
    if op == "any":
        return any(eval_where(w, rec) for w in arg)
    if op == "not":
        return not eval_where(arg, rec)
    if op == "eq":
        return strict_eq(get_attr(rec, arg[0]), arg[1])
    if op == "in":
        val = get_attr(rec, arg[0])
        return any(strict_eq(val, x) for x in (arg[1] or []))
    if op == "has":
        val = get_attr(rec, arg[0])
        if isinstance(val, list):
            return any(strict_eq(x, arg[1]) for x in val)
        return strict_eq(val, arg[1])
    if op == "known":
        return is_known(rec, arg)
    if op == "status":
        st = field_status(rec, arg[0])
        return st is not None and st in (arg[1] or [])
    if op == "gt":
        if not is_known(rec, arg[0]):
            return False
        return fv(rec, arg[0]) > arg[1]
    raise ValueError(f"عامل Where غير معروف: {op}")


def and_where(*parts: Any) -> Optional[dict]:
    items = [p for p in parts if p is not None]
    if not items:
        return None
    if len(items) == 1:
        return items[0]
    return {"all": items}


# ---------------------------------------------------------------------------
# Engine
# ---------------------------------------------------------------------------
class Engine:
    def __init__(self, config: dict, records: Iterable[dict]):
        self.config = config
        self.records: List[dict] = sorted(list(records), key=lambda r: r["id"])
        self.periods = [p for p in (config.get("periods") or [])]
        self.period_ids = [p.get("id") for p in self.periods]
        self.default_period = self.period_ids[0] if self.period_ids else None
        self.list_dims = set()
        for r in self.records:
            for k, v in (r.get("dims") or {}).items():
                if isinstance(v, list):
                    self.list_dims.add(k)

    # -- filters ---------------------------------------------------------
    def normalize_filters(self, filters: Optional[dict]) -> dict:
        out: Dict[str, Any] = {}
        filters = dict(filters or {})
        period = filters.pop("period", None)
        if period in (None, ""):
            period = self.default_period
        if self.period_ids and period not in self.period_ids:
            raise ValueError(f"الفترة غير معروفة: {period!r}؛ المتاح: {self.period_ids}")
        out["period"] = period
        for k in sorted(filters):
            v = filters[k]
            if v is None or v == "":
                continue
            out[k] = v
        return out

    def filter_where(self, filters: dict) -> Optional[dict]:
        clauses = []
        for k in sorted(filters):
            if k == "period":
                continue
            op = "has" if k in self.list_dims else "eq"
            clauses.append({op: [k, filters[k]]})
        if not clauses:
            return None
        return {"all": clauses}

    def _select(self, where: Optional[dict]) -> List[dict]:
        return [r for r in self.records if eval_where(where, r)]

    def filter_records(self, filters: Optional[dict] = None) -> List[dict]:
        f = self.normalize_filters(filters)
        where = and_where(self.config.get("base_where"), self.filter_where(f))
        return self._select(where)

    def context(self, period_id: Optional[str]) -> Dict[str, Optional[str]]:
        return placeholder_context(self.config, period_id)

    def _period(self, period_id: Optional[str]) -> dict:
        for p in self.periods:
            if p.get("id") == period_id:
                return p
        return {}

    # -- metrics ---------------------------------------------------------
    def compute(self, spec: dict, filters: Optional[dict] = None) -> dict:
        f = self.normalize_filters(filters)
        ctx = self.context(f["period"])
        res = _base_result(spec, f["period"])
        fw = self.filter_where(f)
        base = self.config.get("base_where")
        # R without the metric's own placeholders is needed for n_scope in early returns
        try:
            where = resolve_where(spec.get("where"), ctx)
        except Unresolved as exc:
            r_scope = self._select(and_where(base, fw))
            res.update(status="not_computable", n_scope=len(r_scope),
                       reason=res.get("unavailable_reason") or DEFAULT_UNAVAILABLE,
                       unresolved=exc.placeholder)
            return res
        R = self._select(and_where(base, where, fw))
        res["n_scope"] = len(R)
        kind = spec.get("kind")
        try:
            fn = _KINDS[kind]
        except KeyError:
            raise ValueError(f"نوع مؤشر غير معروف: {kind!r} ({spec.get('id')})")
        try:
            fn(self, spec, R, ctx, res, f)
        except Unresolved as exc:
            res.update(status="not_computable", value=None, n_included=0, included_ids=[], excluded=[],
                       reason=res.get("unavailable_reason") or DEFAULT_UNAVAILABLE,
                       unresolved=exc.placeholder)
        return res

    def compute_breakdown(self, spec: dict, filters: Optional[dict] = None) -> dict:
        f = self.normalize_filters(filters)
        ctx = self.context(f["period"])
        res = _base_result(spec, f["period"])
        res.update(kind="breakdown", by=spec.get("by"), rows=[], total=None, n_known=0, n_unknown=0,
                   overlap=False, note=None)
        base = self.config.get("base_where")
        fw = self.filter_where(f)
        try:
            field = resolve_field(spec.get("field"), ctx)
            where = resolve_where(spec.get("where"), ctx)
        except Unresolved as exc:
            r_scope = self._select(and_where(base, fw))
            res.update(status="not_computable", n_scope=len(r_scope),
                       reason=res.get("unavailable_reason") or DEFAULT_UNAVAILABLE, unresolved=exc.placeholder)
            return res
        res["resolved_fields"] = {"field": field}
        R = self._select(and_where(base, where, fw))
        res["n_scope"] = len(R)
        by = spec.get("by")
        if not R:
            res["status"] = "empty"
            return res
        K = [r for r in R if is_known(r, field)]
        U = [r for r in R if not is_known(r, field)]
        groups: Dict[Any, dict] = {}
        overlap = False

        def keys_of(rec: dict) -> List[Any]:
            nonlocal overlap
            val = get_attr(rec, by)
            if isinstance(val, list):
                overlap = True
                ks = [NONE_KEY if x is None else x for x in val]
                return ks if ks else [NONE_KEY]
            return [NONE_KEY if val is None else val]

        def group(key: Any) -> dict:
            gk = _key_str(key)
            if gk not in groups:
                groups[gk] = {"key": key, "value": None, "n_known": 0, "n_unknown": 0, "ids": [], "_sum": 0}
            return groups[gk]

        total = 0
        for r in K:
            total = total + fv(r, field)
            seen = set()
            for k in keys_of(r):
                if _key_str(k) in seen:
                    continue
                seen.add(_key_str(k))
                g = group(k)
                g["_sum"] = g["_sum"] + fv(r, field)
                g["n_known"] += 1
                g["ids"].append(r["id"])
        for r in U:
            seen = set()
            for k in keys_of(r):
                if _key_str(k) in seen:
                    continue
                seen.add(_key_str(k))
                g = group(k)
                g["n_unknown"] += 1
                g["ids"].append(r["id"])
        rows = []
        for g in groups.values():
            value = g["_sum"] if g["n_known"] > 0 else None
            share = (value / total) if (value is not None and total != 0) else None
            rows.append({"key": g["key"], "value": value, "share": share, "n_known": g["n_known"],
                         "n_unknown": g["n_unknown"], "ids": sorted(g["ids"])})
        rows.sort(key=lambda x: (x["value"] is None, -(x["value"] or 0), _key_str(x["key"])))
        res.update(rows=rows, total=total if K else None, n_known=len(K), n_unknown=len(U),
                   overlap=overlap, note=OVERLAP_NOTE if overlap else None,
                   n_included=len(K), included_ids=[r["id"] for r in K],
                   excluded=[{"id": r["id"], "reason": field_status(r, field) or "absent"} for r in U])
        if not K:
            res["status"] = "not_computable"
            res["reason"] = "لا توجد قيم معروفة للحقل ضمن النطاق"
        elif U:
            res["status"] = "partial"
        else:
            res["status"] = "complete"
        return res

    def compute_trend(self, filters: Optional[dict] = None) -> dict:
        f = self.normalize_filters(filters)
        tcfg = self.config.get("trend") or {}
        comparable = tcfg.get("comparable") or {"status": "unverified", "reason": None}
        out = {"id": "trend", "label": tcfg.get("label"), "status": None, "points": [], "comparable": comparable,
               "not_comparable": comparable.get("status") == "not_comparable", "n_scope": 0, "reason": None,
               "unit": tcfg.get("unit", "SAR")}
        if not tcfg or not tcfg.get("enabled") or not tcfg.get("fields"):
            out.update(status="not_computable", reason="التطور الزمني غير مفعّل أو لا توجد فترات")
            return out
        base = self.config.get("base_where")
        ctx = self.context(f["period"])
        try:
            where = resolve_where(tcfg.get("where"), ctx)
        except Unresolved:
            where = None
        R = self._select(and_where(base, where, self.filter_where(f)))
        out["n_scope"] = len(R)
        if not R:
            out["status"] = "empty"
            out["points"] = [{"period": t.get("period"), "field": t.get("field"), "value": None, "n_known": 0, "n": 0}
                             for t in tcfg["fields"]]
            return out
        any_known = False
        any_partial = False
        for t in tcfg["fields"]:
            field = t.get("field")
            s = 0
            n_known = 0
            for r in R:
                if is_known(r, field):
                    s = s + fv(r, field)
                    n_known += 1
            if n_known:
                any_known = True
            if n_known < len(R):
                any_partial = True
            out["points"].append({"period": t.get("period"), "field": field,
                                  "value": s if n_known else None, "n_known": n_known, "n": len(R)})
        if not any_known:
            out["status"] = "not_computable"
            out["reason"] = "لا توجد قيم إيراد معروفة ضمن النطاق"
        elif any_partial:
            out["status"] = "partial"
        else:
            out["status"] = "complete"
        return out

    def compute_all(self, filters: Optional[dict] = None) -> dict:
        f = self.normalize_filters(filters)
        metrics = {}
        for spec in self.config.get("metrics") or []:
            metrics[spec["id"]] = self.compute(spec, f)
        breakdowns = {}
        for spec in self.config.get("breakdowns") or []:
            breakdowns[spec["id"]] = self.compute_breakdown(spec, f)
        return {"filters": f, "metrics": metrics, "breakdowns": breakdowns, "trend": self.compute_trend(f),
                "records": [r["id"] for r in self.filter_records(f)]}


# ---------------------------------------------------------------------------
# Result helpers
# ---------------------------------------------------------------------------
def _key_str(k: Any) -> str:
    if isinstance(k, bool):
        return "true" if k else "false"
    return str(k)


def _texts(spec: dict, period: Optional[str]) -> dict:
    out = {k: spec.get(k) for k in TEXT_KEYS}
    tbp = (spec.get("text_by_period") or {}).get(period) if period is not None else None
    if isinstance(tbp, dict):
        for k in TEXT_KEYS:
            if tbp.get(k) is not None:
                out[k] = tbp[k]
    return out


def _base_result(spec: dict, period: Optional[str]) -> dict:
    res = {
        "id": spec.get("id"), "kind": spec.get("kind"), "period": period,
        "value": None, "status": None, "unit": spec.get("unit"),
        "n_scope": 0, "n_included": 0, "included_ids": [], "excluded": [],
        "resolved_fields": {}, "reason": None,
        "card": bool(spec.get("card", False)), "card_order": spec.get("card_order"),
        "card_group": spec.get("card_group"), "kind_tag": spec.get("kind_tag"),
    }
    res.update(_texts(spec, period))
    return res


def _finish_status(res: dict, R: list, included: list, excluded: list, empty_reason: str) -> None:
    res["n_included"] = len(included)
    res["included_ids"] = [r["id"] for r in included]
    res["excluded"] = excluded
    if not R:
        res["status"] = "empty"
    elif not included:
        res["status"] = "not_computable"
        res["reason"] = res.get("reason") or empty_reason
    elif excluded:
        res["status"] = "partial"
    else:
        res["status"] = "complete"


def _k_count(eng: Engine, spec: dict, R: list, ctx: dict, res: dict, f: dict) -> None:
    res["value"] = len(R)
    res["n_included"] = len(R)
    res["included_ids"] = [r["id"] for r in R]
    res["status"] = "complete" if R else "empty"


def _k_sum(eng: Engine, spec: dict, R: list, ctx: dict, res: dict, f: dict) -> None:
    field = resolve_field(spec.get("field"), ctx)
    res["resolved_fields"] = {"field": field}
    K = [r for r in R if is_known(r, field)]
    U = [r for r in R if not is_known(r, field)]
    total = 0
    for r in K:
        total = total + fv(r, field)
    excluded = [{"id": r["id"], "reason": field_status(r, field) or "absent"} for r in U]
    _finish_status(res, R, K, excluded, "لا توجد قيم معروفة للحقل ضمن النطاق")
    res["n_unknown"] = len(U)
    res["value"] = total if (R and K) else None


def _k_ratio(eng: Engine, spec: dict, R: list, ctx: dict, res: dict, f: dict) -> None:
    num_f = resolve_field(spec.get("num"), ctx)
    den_f = resolve_field(spec.get("den"), ctx)
    scale = spec.get("scale", 1)
    if scale is None:
        scale = 1
    res["resolved_fields"] = {"num": num_f, "den": den_f}
    I, excluded = [], []
    for r in R:
        if not is_known(r, num_f):
            excluded.append({"id": r["id"], "reason": "num_unknown"})
        elif not is_known(r, den_f):
            excluded.append({"id": r["id"], "reason": "den_unknown"})
        elif not fv(r, den_f) > 0:
            excluded.append({"id": r["id"], "reason": "den_nonpositive"})
        else:
            I.append(r)
    num_sum = 0
    den_sum = 0
    for r in I:
        num_sum = num_sum + fv(r, num_f)
        den_sum = den_sum + fv(r, den_f)
    num_known_total = 0
    for r in R:
        if is_known(r, num_f):
            num_known_total = num_known_total + fv(r, num_f)
    _finish_status(res, R, I, excluded, "لا توجد سجلات يتوافر لها البسط والمقام معًا")
    res["num"] = num_sum
    res["den"] = den_sum
    res["value"] = (num_sum / den_sum * scale) if I else None
    res["num_known_total"] = num_known_total
    res["coverage_num_share"] = (num_sum / num_known_total) if num_known_total != 0 else None
    res["scale"] = scale


def _k_diff(eng: Engine, spec: dict, R: list, ctx: dict, res: dict, f: dict) -> None:
    a_f = resolve_field(spec.get("a"), ctx)
    b_f = resolve_field(spec.get("b"), ctx)
    res["resolved_fields"] = {"a": a_f, "b": b_f}
    if not R:
        res["status"] = "empty"
        return
    if not any(b_f in (r.get("fields") or {}) for r in R):
        res.update(status="not_computable", reason=NO_EXPENSE_REASON)
        return
    I, excluded = [], []
    for r in R:
        if not is_known(r, a_f):
            excluded.append({"id": r["id"], "reason": "a_unknown"})
        elif not is_known(r, b_f):
            excluded.append({"id": r["id"], "reason": "b_unknown"})
        else:
            I.append(r)
    value = 0
    a_sum = 0
    b_sum = 0
    for r in I:
        value = value + (fv(r, a_f) - fv(r, b_f))
        a_sum = a_sum + fv(r, a_f)
        b_sum = b_sum + fv(r, b_f)
    _finish_status(res, R, I, excluded, "لا توجد سجلات يتوافر لها الإيراد والمصروف معًا")
    res["value"] = value if I else None
    res["a_sum"] = a_sum
    res["b_sum"] = b_sum


def _k_growth(eng: Engine, spec: dict, R: list, ctx: dict, res: dict, f: dict) -> None:
    prev_f = resolve_field(spec.get("prev", "{REV_PREV}"), ctx)
    curr_f = resolve_field(spec.get("curr", "{REV}"), ctx)
    res["resolved_fields"] = {"prev": prev_f, "curr": curr_f}
    period = eng._period(f["period"])
    comp = period.get("comparable_with_prev") or {"status": "unverified", "reason": None}
    res["comparability"] = comp
    res["prev_period"] = period.get("prev_id")
    if comp.get("status") == "not_comparable":
        res.update(status="not_comparable", value=None, reason=comp.get("reason") or "الفترتان غير قابلتين للمقارنة")
        return
    if not R:
        res["status"] = "empty"
        return
    I, excluded, only_curr, only_prev = [], [], [], []
    for r in R:
        kp, kc = is_known(r, prev_f), is_known(r, curr_f)
        if kp and kc:
            I.append(r)
        elif kc:
            only_curr.append(r["id"])
            excluded.append({"id": r["id"], "reason": "prev_unknown"})
        elif kp:
            only_prev.append(r["id"])
            excluded.append({"id": r["id"], "reason": "curr_unknown"})
        else:
            excluded.append({"id": r["id"], "reason": "both_unknown"})
    prev_sum = 0
    curr_sum = 0
    for r in I:
        prev_sum = prev_sum + fv(r, prev_f)
        curr_sum = curr_sum + fv(r, curr_f)
    _finish_status(res, R, I, excluded, "لا توجد سجلات لها إيراد معروف في الفترتين")
    abs_change = curr_sum - prev_sum
    pct = (abs_change / prev_sum * 100) if prev_sum > 0 else None
    res.update(prev_sum=prev_sum if I else None, curr_sum=curr_sum if I else None,
               abs_change=abs_change if I else None, pct=pct if I else None,
               revenue_start=bool(I) and prev_sum == 0 and curr_sum > 0,
               only_curr=only_curr, only_prev=only_prev,
               value=pct if I else None)


def _k_concentration(eng: Engine, spec: dict, R: list, ctx: dict, res: dict, f: dict) -> None:
    field = resolve_field(spec.get("field"), ctx)
    top = int(spec.get("top", 3) or 3)
    res["resolved_fields"] = {"field": field}
    res["top"] = top
    if not R:
        res["status"] = "empty"
        return
    K = [r for r in R if is_known(r, field) and fv(r, field) > 0]
    excluded = []
    any_unknown = False
    for r in R:
        if not is_known(r, field):
            any_unknown = True
            excluded.append({"id": r["id"], "reason": field_status(r, field) or "absent"})
        elif not fv(r, field) > 0:
            excluded.append({"id": r["id"], "reason": "nonpositive"})
    total = 0
    for r in K:  # id order
        total = total + fv(r, field)
    ranked = sorted(K, key=lambda r: (-fv(r, field), r["id"]))
    top_recs = ranked[:top]
    top_sum = 0
    for r in top_recs:
        top_sum = top_sum + fv(r, field)
    res["n_included"] = len(K)
    res["included_ids"] = [r["id"] for r in K]
    res["excluded"] = excluded
    res["top_ids"] = [r["id"] for r in top_recs]
    res["total"] = total
    res["top_sum"] = top_sum
    if total != 0:
        res["value"] = top_sum / total * 100
        h = 0
        for r in ranked:
            h = h + (fv(r, field) / total) ** 2
        res["hhi"] = h * 10000
        res["top_shares"] = [{"id": r["id"], "value": fv(r, field), "share": fv(r, field) / total * 100}
                             for r in top_recs]
    else:
        res["value"] = None
        res["hhi"] = None
        res["top_shares"] = []
    if not K or total == 0:
        res["status"] = "not_computable"
        res["reason"] = "لا توجد قيم موجبة معروفة ضمن النطاق"
        res["value"] = None
    elif any_unknown:
        res["status"] = "partial"
    else:
        res["status"] = "complete"


def _k_coverage(eng: Engine, spec: dict, R: list, ctx: dict, res: dict, f: dict) -> None:
    fields = [resolve_field(x, ctx) for x in (spec.get("fields") or [])]
    res["resolved_fields"] = {"fields": fields}
    n = len(R)
    per_field = []
    for fld in fields:
        nk = 0
        for r in R:
            if is_known(r, fld):
                nk += 1
        per_field.append({"field": fld, "n_known": nk, "n": n})
    res["per_field"] = per_field
    res["n_included"] = n
    res["included_ids"] = [r["id"] for r in R]
    if n == 0:
        res["status"] = "empty"
        res["value"] = None
        return
    ratios = [pf["n_known"] / n for pf in per_field]
    res["value"] = (min(ratios) * 100) if ratios else None
    res["status"] = "complete"


_KINDS = {
    "count": _k_count,
    "sum": _k_sum,
    "ratio": _k_ratio,
    "diff": _k_diff,
    "growth": _k_growth,
    "concentration": _k_concentration,
    "coverage": _k_coverage,
}
METRIC_KINDS = tuple(_KINDS)


def compact(result: dict) -> dict:
    """The fields compared by the dashboard parity checks."""
    return {"value": result.get("value"), "status": result.get("status"), "n_included": result.get("n_included")}


def check_records(records: List[dict], fields: Optional[Iterable[str]] = None) -> List[str]:
    """Structural sanity checks the engine relies on. Returns Arabic error strings.

    Ids must be unique non-empty strings. When `fields` is given, every known-status value of
    those (numeric) fields must be a number. Text fields legitimately hold strings.
    """
    fields = set(fields or [])
    errs = []
    seen = set()
    for r in records:
        rid = r.get("id")
        if not isinstance(rid, str) or not rid:
            errs.append(f"سجل بلا معرّف صالح: {r.get('source')}")
            continue
        if rid in seen:
            errs.append(f"معرّف مكرر: {rid}")
        seen.add(rid)
        for fname, tv in (r.get("fields") or {}).items():
            if fname in fields and isinstance(tv, dict) and tv.get("s") in KNOWN_STATUSES and not _is_num(tv.get("v")):
                errs.append(f"{rid}: الحقل {fname} حالته {tv.get('s')} لكن قيمته ليست رقمًا")
    return errs
