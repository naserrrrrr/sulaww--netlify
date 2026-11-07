# Data contract (v1.0) — files exchanged between the analyst (Claude) and the scripts

> Technical spec, kept in English for precision. Every script and every hand-authored JSON follows it.
> Rule of thumb: **Claude decides meaning (mapping, classification, narrative); scripts do arithmetic, checks and rendering.** Claude never types a computed number into the dashboard or narrative by hand — numbers are referenced by metric id (`{{m:metric_id}}`) or copied from `05_metrics.json`.

## 0. Run folder layout

All outputs of one run go to `outputs/portfolio-analysis/<file-stem>_<YYYYMMDD-HHMM>/` (relative to the working directory). The source workbook is **never modified**; its path and SHA-256 are recorded.

| File | Written by | Purpose |
|---|---|---|
| `01_workbook_map.json` / `.md` | `inspect_workbook.py` | Every sheet (incl. hidden), tables, formulas, anomalies, PII columns |
| `02_mapping.json` | **Claude** | Scope decisions, column roles, classification of every record |
| `03_records.json` | `normalize.py` | Typed, source-referenced records + auto-detected data issues |
| `04_metrics_config.json` | `compute_metrics.py --init-config` then **Claude** edits | Metric definitions (DSL), periods, valuation basis, breakdowns |
| `05_metrics.json` | `compute_metrics.py` | Authoritative metric results + parity cases for the dashboard |
| `06_narrative.json` | **Claude** | Executive summary, assessment, gaps, recommendations, plan, issues, assumptions, questions |
| `07_validation.json` / `.md` | `validate_analysis.py` | Blocking errors + warnings. Build refuses when blocking errors exist |
| `analysis.json` | `build_dashboard.py` | Single bundle embedded in the dashboard (no PII) |
| `dashboard.html` | `build_dashboard.py` | Self-contained, offline, RTL dashboard |
| `08_dashboard_check.json` + `screens/` | `check_dashboard.mjs` | Browser test results (parity, filters, mobile, print, PII) |
| `supporting_workbook.xlsx` | `export_workbook.py` | Reviewable Excel: assets, calculations, exclusions, data issues, plan |
| `09_message.md` | **Claude** | The short delivery message |

## 1. `01_workbook_map.json` (inspect_workbook.py)

```jsonc
{
  "contract_version": "1.0",
  "file": {"path": "...", "name": "...", "sha256": "...", "size": 40988,
           "properties": {"creator": "...", "last_modified_by": "<redacted:name>", "created": "...", "modified": "..."}},
  "sheets": [{
    "name": "عقارات الجمعية ",            // exact, may contain trailing spaces
    "index": 0, "state": "visible|hidden|veryHidden",
    "dimensions": "A1:AM43", "max_row": 43, "max_col": 39,
    "merged": ["A1:AI1"], "hidden_rows": [], "hidden_cols": [],
    "tables": [{"name": "الجدول4", "ref": "A2:AJ4", "header_row": 2, "data_rows": [3, 3], "totals_row": 4}],
    "header_candidates": [{"row": 2, "score": 0.97, "headers": {"A": "...", "B": "..."}}],
    "title_cells": [{"ref": "A1", "text": "بيان بأصول الجمعية ... لعام 2025م"}],
    "columns": [{"col": "C", "header": "الايراد بدون ضريبة 2025", "years_in_header": [2025],
                 "profile": {"numeric": 1, "zero": 0, "text": 0, "blank": 0, "filler": 0, "error": 0, "formula": 0, "date": 0},
                 "samples": ["144450"], "mixed_types": false, "pii": false}],
    "row_kinds": {"3": "data", "4": "total_formula", "28-43": "filler"},
    "formulas": {"count": 36, "cached_missing": 0, "errors": 0},
    "comments": [{"ref": "...", "text": "..."}],
    "data_validations": [...]
  }],
  "anomalies": [{
    "id": "WB-001", "type": "formula_year_mismatch|cached_value_mismatch|total_mismatch|hardcoded_total|excel_error|cached_missing|mixed_types|text_in_numeric|duplicate_name|shared_deed|serial_gap|pii_column|partial_period_header|garbled_header|filler_block|derived_from_total|count_mismatch",
    "severity": "high|medium|low",
    "where": "عقارات الجمعية !AA3",
    "detail": "Arabic explanation with numbers",
    "evidence": {"formula": "...", "cached": 1100000, "recomputed": 823356}
  }],
  "pii_columns": [{"sheet": "...", "col": "AH", "header": "رقم التواصل", "reason": "header|pattern"}]
}
```
PII **values are never written** to any output; only the column location.

## 2. Typed value (`TV`) — used in every record field

```jsonc
{"v": 144450, "s": "value", "raw": "144450", "ref": "عقارات الجمعية !C3", "note": null}
```
`s` ∈ `value | derived | zero_recorded | zero_confirmed | zero_as_missing | blank | filler | text | error | not_applicable | not_available`.
**Known** statuses = `value, derived, zero_recorded, zero_confirmed` (v is a number; zero statuses have v = 0).
Everything else has `v = null` and is *unknown* — never summed as zero.
Zero policy per column (mapping `zero_policy`): `recorded` (default for revenue, expense, units, counts) → `zero_recorded`; `missing` (default for `val.*`, areas) → `zero_as_missing`.
Dates: `{"v": "1982-12-14", "s": "value", "raw": "1403/2/29هـ", "calendar": "hijri", "hijri": "1403-02-29"}`.
Text fields: `{"v": "منفوحة", "s": "value", "raw": "...", "ref": "..."}`.

## 3. `02_mapping.json` (authored by Claude)

```jsonc
{
  "contract_version": "1.0",
  "source_file": "path/to/file.xlsx",
  "entity": {"name": null, "evidence": "اسم الجهة غير مذكور في الملف"},
  "currency": "SAR",
  "data_period": {"label": "2025", "evidence": "عقارات الجمعية !A1"},
  "study_date": {"value": "2025-12-02", "basis": "آخر تعديل للملف (خصائص المصنف)", "evidence": "file.properties.modified"},
  "analysis_date": "2026-11-07",
  "valuation_dates": [{"year": "2024", "basis": "تقييم مسجل في الملف، الأساس غير موضح", "evidence": "..."}],
  "sheets": [{"sheet": "...", "role": "asset_register|activity_register|allocation|managed_properties|lookup|notes|other",
              "decision": "primary|supplementary|isolated|excluded", "period": "2025", "reliability": "high|medium|low",
              "relation_to_register": "...", "reason": "..."}],
  "sources": [{
    "id": "REG",                                   // short code, used in record ids
    "sheet": "عقارات الجمعية ",
    "header_row": 2,
    "data_rows": {"from": 3, "to": 3},             // inclusive
    "exclude_rows": [{"rows": [4], "reason": "صف إجمالي (SUBTOTAL)"}],
    "skip_if_blank": "B",                           // skip rows whose name cell is blank/filler
    "default_record_type": "asset",
    "default_scope": "in_scope",                    // isolated for sheets whose scope is unverified
    "id_rule": {"column": "A", "prefix": "REG-"},   // fallback when blank: prefix + "r" + row
    "name_column": "B",
    "columns": {
      "C": {"field": "rev.2025.excl_vat", "kind": "money"},
      "N": {"field": "units", "kind": "count_mixed"},
      "V": {"field": "val.2022.land", "kind": "money", "zero_policy": "missing"},
      "AE": {"field": "ownership_date", "kind": "date", "calendar": "auto"},
      "E": {"field": "district", "kind": "text", "dim": "district"},
      "K": {"field": "asset_type_text", "kind": "text", "dim": "asset_type"},
      "AG": {"field": "supervisor", "kind": "pii"}
    }
  }],
  "field_meta": {                                   // one entry per field name used above
    "rev.2025.excl_vat": {"label": "الإيراد المسجل 2025 بدون ضريبة", "group": "revenue", "period": "2025",
                          "vat": "excl|incl|none_evidenced|unknown", "basis": "actual|expected|recorded|contract",
                          "cash_basis": "collected|accrued|unknown", "completeness": "full_year|partial|unknown",
                          "unit": "SAR", "note": "..."},
    "val.2024.total": {"label": "تقييم 2024 (إجمالي)", "group": "valuation", "year": "2024",
                       "valuation_basis": "market|book|cost|unknown", "valuation_date": "2024 (اليوم غير محدد)", "unit": "SAR"}
  },
  "dim_rules": [                                    // optional: derive dims from a text field
    {"dim": "use", "from_field": "allocation_text", "map": {"استثماري": "investment", "غير استثماري": "service", "اداري": "admin"}, "default": "unknown"}
  ],
  "records": {                                      // per-record overrides (win over dim_rules)
    "REG-1": {"record_type": "asset", "scope": "in_scope", "scope_reason": null,
              "dims": {"use": "investment", "ownership": ["owned"], "waqf": "unknown", "income_nature": "rent",
                       "asset_type": "عمارة", "location": "جنوب الرياض", "managed_for_others": false},
              "dims_evidence": {"use": "عقارات الجمعية !M3", "ownership": "عقارات الجمعية !J3 (شراء)"},
              "field_overrides": {"rev.2025.excl_vat": {"s": "zero_confirmed", "evidence": "..."}},
              "notes": "..."}
  },
  "links": [{"record": "REG-1", "other_ref": "تخصيص الايرادات!B12", "match": "probable|confirmed|rejected",
             "basis": "تشابه الاسم فقط؛ لا معرّف مشترك", "used": false, "reason": "..."}],
  "derived_fields": [                               // optional, explicit derivations (never silent)
    {"field": "val.2023.total", "formula": "val.2023.land + val.2023.building", "when": "always|if_unknown",
     "reason": "الإجمالي المدون يستخدم أعمدة 2022 (خطأ صيغة)"}
  ]
}
```
Rules: dims vocabularies are in `assets/vocab.json`. `ownership` is a **list** (non-exclusive). `waqf` is separate. Fields that hold labels (e.g. "صافي الأصول" containing «غير مقيد») must be mapped as `kind: "text"`, never `money`.

## 4. `03_records.json` (normalize.py)

```jsonc
{
  "contract_version": "1.0",
  "source": {"file": "...", "sha256": "..."},
  "field_meta": { ... copied from mapping ... },
  "records": [{
    "id": "REG-1", "name": "عمارة ...", "source": {"source_id": "REG", "sheet": "...", "row": 3, "ref_name": "...!B3"},
    "record_type": "asset", "scope": "in_scope", "scope_reason": null,
    "dims": {"asset_type": "عمارة", "location": "...", "district": "منفوحة", "use": "investment", "ownership": ["owned"],
             "waqf": "unknown", "income_nature": "rent", "managed_for_others": false, "rent_allocation": "...", "restriction": "..."},
    "dims_evidence": {...},
    "fields": {"rev.2025.excl_vat": TV, "val.2024.total": TV, "units": TV, ...},
    "issue_ids": ["DQ-004"]
  }],
  "issues": [{"id": "DQ-001", "type": "...", "severity": "high|medium|low", "where": "ref", "record_ids": ["REG-1"],
              "detail": "...", "impact": "which metric is affected and how", "treatment": "how it was handled"}],
  "stats": {"records": 1, "by_type": {...}, "by_scope": {...}, "pii_columns_dropped": 2}
}
```
Auto checks in normalize: duplicate ids, possibly-duplicate names (`name_key`), shared deed numbers, total-like rows inside data range, `val.Y.total` vs `land+building` mismatch, `incl_vat < excl_vat`, text in numeric fields, zero valuations, mixed `units`, missing classification (`unknown` dims) for in-scope records, Hijri/Gregorian conversion notes. PII columns are dropped (count only).

## 5. `04_metrics_config.json` — metric DSL

```jsonc
{
  "contract_version": "1.0",
  "currency": "SAR",
  "periods": [                                       // drives the period filter; first = default
    {"id": "2025", "label": "2025 — بدون ضريبة", "rev_field": "rev.2025.excl_vat", "exp_field": null,
     "prev_id": "2024", "prev_rev_field": "rev.2024.actual",
     "comparable_with_prev": {"status": "confirmed|unverified|not_comparable", "reason": "..."}}
  ],
  "valuation": {"field": "val.2024.total", "year": "2024", "basis": "...", "date_text": "...", "label": "تقييم 2024"},
  "base_where": {"all": [{"eq": ["scope", "in_scope"]}]},   // ANDed into every metric
  "metrics": [ MetricSpec, ... ],
  "breakdowns": [{"id": "value_by_type", "label": "توزيع القيمة حسب نوع الأصل", "field": "{VAL}", "by": "asset_type",
                  "where": {...}, "unit": "SAR"}],
  "trend": {"enabled": true, "label": "...", "fields": [{"period": "2021", "field": "rev.2021.actual"}],
            "where": {...}, "comparable": {"status": "confirmed|unverified|not_comparable", "reason": "..."}},
  "filters": ["asset_type", "location", "use"],     // dims offered as dashboard filters (+ period always)
  "parity_max_values_per_dim": 3
}
```

### MetricSpec
```jsonc
{
  "id": "gross_yield", "label": "...", "short_label": "...", "card": true, "card_order": 5,
  "kind": "count|sum|ratio|diff|growth|concentration|coverage",
  "field": "{VAL}",                 // sum, concentration
  "num": "{REV}", "den": "{VAL}",   // ratio
  "a": "{REV}", "b": "{EXP}",       // diff (a - b per record)
  "fields": ["{REV}", "{VAL}"],     // coverage
  "top": 3,                          // concentration
  "scale": 100,                      // ratio multiplier
  "where": Where,                    // ANDed with base_where and with dashboard filters
  "unit": "SAR|%|count|asset",
  "definition": "Arabic definition", "scope_text": "Arabic scope", "basis": "Arabic basis (e.g. قبل المصروفات)",
  "valuation_date": "...", "period_text": "...", "source_text": "...",
  "unavailable_reason": "Arabic reason shown when a placeholder resolves to null",
  "kind_tag": "fact|derived"         // statement kind of the result (counts of file rows = fact; computations = derived)
}
```
Placeholders: `{REV}` → `periods[p].rev_field`; `{REV_PREV}` → `periods[p].prev_rev_field`; `{EXP}` → `periods[p].exp_field`; `{VAL}` → `valuation.field`. A placeholder resolving to `null` ⇒ status `not_computable`, `reason = unavailable_reason || "الحقل غير متاح لهذه الفترة"`.

### Where (predicate on a record)
- `{"all": [W...]}`, `{"any": [W...]}`, `{"not": W}`
- `{"eq": [attr, value]}`, `{"in": [attr, [values]]}`, `{"has": [attr, value]}` (attr is a list, e.g. ownership)
- `{"known": field}` — field status ∈ known; `{"status": [field, [statuses]]}`; `{"gt": [field, number]}` (known and v > n)
- `attr` resolution: `id`, `record_type`, `scope` at top level; otherwise `dims[attr]`. Missing attr ⇒ `null` (eq to null is allowed).
- Dashboard filters add `{"eq": [dim, value]}` for each selected dim (multi-valued dims use `has`).

### Semantics (Python `compute_metrics.py` and JS `metric_engine.js` MUST be identical)
Let `R` = records satisfying `base_where ∧ where ∧ filters`. `known(f)` uses the record's field status. Sums use `v`.
- **count**: `value = |R|`; `status = complete` (or `empty` if |R| = 0, value 0).
- **sum**: `K = {r ∈ R : known(f)}`, `U = R \ K`. `value = Σ_K v`. status: `empty` if |R|=0; `not_computable` if |K|=0 (value null); `partial` if |U|>0; else `complete`. `excluded = U` with reason = field status.
- **ratio** (pairwise): `I = {r ∈ R : known(num) ∧ known(den) ∧ den.v > 0}`. `num_sum = Σ_I num.v`, `den_sum = Σ_I den.v`. `value = num_sum / den_sum × scale` (null if I empty). `excluded = R \ I` with reason ∈ `num_unknown | den_unknown | den_nonpositive` (first failing). status: `empty` | `not_computable` (I empty) | `partial` (excluded non-empty) | `complete`. Extra: `num_known_total = Σ_{r∈R, known(num)} num.v`; `coverage_num_share = num_sum / num_known_total` (null if total 0). **Never** report coverage by value.
- **diff**: `I = {r ∈ R : known(a) ∧ known(b)}`; `value = Σ_I (a.v − b.v)`; same status rules; if no record has field `b` at all ⇒ `not_computable` with reason «لا توجد بيانات مصروفات».
- **growth**: period p must define `prev_rev_field`; else `not_computable`. If `comparable_with_prev.status = not_comparable` ⇒ `not_comparable`, value null. `I = {r ∈ R : known(prev) ∧ known(curr)}`. `prev_sum`, `curr_sum`, `abs_change = curr_sum − prev_sum`, `pct = abs_change / prev_sum × 100` if `prev_sum > 0` else null and `revenue_start = (prev_sum == 0 ∧ curr_sum > 0)`. `value = pct`. `only_curr` / `only_prev` = records known in one period only (reported separately, not in I). `comparability` copied from config.
- **concentration**: `K = {r ∈ R : known(f) ∧ v > 0}` sorted by v desc (tie → id asc). `total = Σ_K v`; `top_ids` = first `top`; `value = Σ_top v / total × 100` (null if total 0). Also `hhi = Σ (v/total)² × 10000`.
- **coverage**: for each field f in `fields`: `n_known_f / |R|`. `value = min over fields` × 100; `per_field = [{field, n_known, n}]`. status `complete` unless |R| = 0.
- **breakdown**: group `K = {r ∈ R : known(field)}` by `dims[by]` (list dims: one row per element; label it «قد تتداخل الفئات»); `unknown` groups by `null` → key `"__none__"`. Each group: `value`, `n_known`; plus `n_unknown` per group from `R \ K`. Shares relative to `Σ_K v`.
- Rounding: never round stored values; the dashboard formats only for display.

### Result object (`05_metrics.json` → `metrics[id]`)
```jsonc
{"id": "...", "kind": "...", "label": "...", "value": 1.23, "status": "partial", "unit": "%",
 "n_scope": 10, "n_included": 7, "included_ids": [...], "excluded": [{"id": "REG-3", "reason": "den_unknown"}],
 "num": 123, "den": 456, "num_known_total": 130, "coverage_num_share": 0.946,
 "prev_sum": ..., "curr_sum": ..., "abs_change": ..., "revenue_start": false, "only_curr": [], "only_prev": [],
 "top_ids": [...], "hhi": ..., "per_field": [...],
 "resolved_fields": {"num": "rev.2025.excl_vat", "den": "val.2024.total"},
 "reason": null, "definition": "...", "scope_text": "...", "basis": "...", "valuation_date": "...", "period_text": "..."}
```
`05_metrics.json` = `{"contract_version", "default_filters": {"period": "2025"}, "metrics": {...}, "breakdowns": {...}, "trend": {...}, "parity_cases": [{"filters": {...}, "metrics": {id: {"value", "status", "n_included"}}}]}`.
Parity cases: default; each period; for each filter dim up to `parity_max_values_per_dim` values.

## 6. `06_narrative.json` (authored by Claude)

```jsonc
{
  "contract_version": "1.0",
  "executive_summary": "120–180 Arabic words. Numbers via {{m:metric_id}} or {{m:metric_id.field}} tokens, rendered by the build.",
  "assessment": [{"dimension": "income_generation|income_stability|asset_utilization|data_quality",
                  "judgment": "وصف دون درجات", "criterion": null, "evidence": ["m:property_revenue", "DQ-002"], "kind": "inference"}],
  "strengths": [{"text": "...", "evidence": ["..."], "kind": "fact|derived|inference"}],
  "gaps": [{"id": "G1", "title": "...", "scope": "REG-1 | portfolio | group:use=investment",
            "evidence": [{"ref": "m:zero_rev_investment_value | DQ-003 | WB-002 | عقارات الجمعية !C3", "kind": "fact|derived"}],
            "impact": {"metric": "m:zero_rev_investment_value", "text": "..."},
            "proven_cause": null, "possible_causes": ["..."], "verification_needed": ["..."],
            "proposed_action": "...", "severity": "high|medium|low", "is_proven_problem": false}],
  "recommendations": [{"id": "R1", "gap_ids": ["G1"], "target": "...", "problem": "...", "evidence": ["..."],
                       "first_action": "...", "alternatives": ["..."], "recommended_option": "...", "rationale": "...",
                       "owner_role": "...", "duration": "...", "cost": "... أو متطلبات تقديرها", "expected_impact": "...",
                       "success_kpi": "...", "approval_required": "... | null",
                       "ranking": {"impact": "high|medium|low", "urgency": "...", "evidence_strength": "strong|moderate|weak",
                                   "feasibility": "...", "cost_level": "..."}, "priority": 1}],
  "top_gaps": ["G1", "G2", "G3"], "top_actions": ["R1", "R2", "R3"],
  "plan": {"d30": [{"rec_id": "R1", "action": "...", "deliverable": "...", "owner_role": "..."}], "d60": [], "d90": []},
  "scenarios": {"status": "not_prepared", "reason": "...", "required_inputs": ["..."]},
  "data_issues_extra": [{"id": "DQ-X1", "where": "...", "detail": "...", "impact": "...", "treatment": "..."}],
  "exclusions": [{"what": "...", "where": "...", "reason": "..."}],
  "assumptions": [{"id": "A1", "text": "...", "kind": "inference|assumption", "basis": "..."}],
  "data_requests": [{"item": "...", "enables": ["..."], "priority": 1}],
  "questions": ["≤ 3"],
  "sources": [{"type": "file|web|user", "ref": "...", "date": "...", "limits": "..."}],
  "version_comparison": null,
  "recommendation_followup": [],
  "methodology_notes": ["..."]
}
```
Scenario items (only when `status = "prepared"`): `{"id", "opportunity", "asset_ids", "assumptions": [{"text", "source"}], "conservative|base|optimistic": {"extra_revenue", "extra_expense", "net_effect", "capex", "start", "year1_effect", "run_rate_effect"}, "recurring": true, "overlaps_with": []}`.

## 7. `analysis.json` (build_dashboard.py → embedded in HTML)

```jsonc
{"contract_version": "1.0", "meta": {"entity", "data_period", "analysis_date", "study_date", "currency", "file_name", "sha256_short", "valuation"},
 "vocab": {...}, "field_meta": {...}, "records": [... in_scope + isolated + excluded, PII-free ...],
 "config": {...04...}, "metrics": {...05...}, "narrative": {...06 with {{m:..}} rendered...},
 "workbook_map_summary": {...}, "issues": [...03 issues + narrative extra...], "validation": {...07 summary...}}
```

## 8. Dashboard engine API (`assets/metric_engine.js`, exposed as `window.APA`)
- `APA.data` — the embedded `analysis.json`.
- `APA.compute(metricSpec, filters)` → result object (same shape as §5).
- `APA.computeAll(filters)` → `{metrics: {id: result}, breakdowns: {id: ...}, trend: {...}, records: [ids in R for base_where ∧ filters]}`.
- `APA.filterRecords(filters)` → records passing `base_where ∧ filters`.
- `filters` = `{"period": "2025", "asset_type": null|"...", "location": null|"...", "use": null|"..."}`.
The UI renders **only** from `APA.computeAll(currentFilters)`; cards, charts and table share that single result.
