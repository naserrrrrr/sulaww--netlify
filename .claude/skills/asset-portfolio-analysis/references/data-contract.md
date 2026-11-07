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
| `07_validation.json` / `.md` | `validate_analysis.py` | Blocking errors + warnings (+ acknowledged warnings, totals reconciliation). Build refuses when blocking errors exist. Re-run after `export_workbook.py` and after writing `09_message.md` (or pass `--xlsx-out`) so the supporting workbook and the message are checked before delivery |
| `analysis.json` | `build_dashboard.py` | Single bundle embedded in the dashboard (no PII) |
| `dashboard.html` | `build_dashboard.py` | Self-contained, offline, RTL dashboard |
| `08_dashboard_check.json` + `screens/` | `check_dashboard.mjs` | Browser test results (parity, filters, mobile, print, PII) |
| `supporting_workbook.xlsx` | `export_workbook.py` | Reviewable Excel: assets, calculations, exclusions, data issues, plan; plus «المطابقة مع الإجماليات» (when 03 keeps total-row values), «متابعة التوصيات» and «مقارنة الإصدارات» (when 06 holds them). Exit 2 and the file is deleted if the post-save privacy scan finds contact/IBAN data |
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
    "id": "WB-001", "type": "formula_year_mismatch|cached_value_mismatch|total_mismatch|row_total_mismatch|hardcoded_total|excel_error|cached_missing|mixed_types|text_in_numeric|numeric_text|duplicate_name|shared_deed|serial_gap|pii_column|partial_period_header|garbled_header|filler_block|derived_from_total|count_mismatch|vat_identical|vat_inconsistent|hijri_date_text|zero_valuation|unlabeled_column|outside_table_rows|hidden_sheet|hidden_rows|hidden_columns|cross_sheet_name_match|mixed_calendar|formula_blank_inputs",
    "severity": "high|medium|low|info",
    "where": "عقارات الجمعية !AA3",
    "detail": "Arabic explanation with numbers",
    "evidence": {"formula": "...", "cached": 1100000, "recomputed": 823356}
  }],
  "pii_columns": [{"sheet": "...", "col": "AH", "header": "رقم التواصل", "reason": "header|pattern"}]
}
```
Notes on `01`:
- **Header row**: the first header candidate from the top wins; a later candidate only starts a new block (blank/title/
  section/total row right above it and enough contrast with the rows below), so a plain range is never re-headed by a data row.
- `columns[].hints` may contain `category` (headers with طبيعة/نوع/تصنيف/حالة or type/category/status/nature): such a column
  is a classification, never money/revenue/valuation.
- `row_kinds` values: `data | header | title | section_label | total_formula | total_label | filler | blank`. A row whose only
  text is a total label («المجموع»، «جملة القسم») with Excel errors or n/a marks in the other cells is a total row.
- `formula_blank_inputs` (medium): a formula whose cached value is 0 while every referenced cell is blank/n/a (COUNT-type
  formulas excluded) — the inspect counterpart of normalize's `zero_from_blank_inputs`.
- `mixed_calendar` (medium): one column holds both Hijri text and Gregorian text/Excel dates; `hijri_date_text` examples come
  from Hijri cells only. `parse_date` adds the note «تاريخ ملتبس: اليوم والشهر ≤ 12؛ قُرئ يوم/شهر» when day and month are both ≤ 12.
- `cross_sheet_name_match` (info) compares the name column of every sub-table of every non-lookup sheet (hidden included);
  a token subset counts (all words of the shorter name, ≥ 2 words). Evidence adds `matches[].equal_amounts` (equal
  amount-like cells ≥ 1000 that are not years, serials, counts, areas, dates, deeds or personal data) and `with_equal_amounts`.
- A personal-data column is never chosen as a name/serial/deed key. `pii_columns[].reason = "pattern"` also covers e-mail,
  IBAN and whole-cell national-ID-like values (not for columns whose header gives the numbers another meaning).
  Header detection uses `apa_common.is_pii_header` (shared pattern, tatweel/hamza folded, plural person nouns).

PII **values are never written** to any output; only the column location. Redacted samples and person-like file
properties (creator / last_modified_by) appear as «<محجوب>». The map is an internal working file (it shows asset
names): only `workbook_map_summary` goes into `analysis.json`. Extra keys (`register_candidates`, `regions`,
`key_columns`, `summary`, …) are additive hints — Claude still decides the mapping.

**Converted inputs (.xls / .ods / .xlsb):** `file.converted.path` (RUN/_converted/<stem>.xlsx) is the copy every later
step reads (`normalize.py --xlsx`, `validate_analysis.py --xlsx`); the original stays untouched. LibreOffice may have
recalculated its formulas — say so in `methodology_notes`.

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

Required context (schema, exit 2 when missing): `entity`, `data_period`, `study_date`, `currency`, `valuation_dates`,
`sheets`. Values may be null, but `entity`, `data_period`, `study_date` and every `valuation_dates[]` item need a non-empty
`evidence`; `sheets[]` has ≥ 1 item and must list **every** sheet of 01 (hidden and veryHidden included, none that 01 lacks —
`mapping_context` blocks otherwise). These keys plus `analysis_date` are copied to 03 under `context`.

More rules:
- `data_rows.to` is a required integer; `from > to` exits 2, and a source that yields no row exits 2.
- **Controlled dimensions** `use`, `income_nature`, `ownership` (list), `waqf`, `managed_for_others` (bool) take only codes
  from `assets/vocab.json`. A literal outside the vocabulary in `dim_rules` (`map`/`default`) or `records[].dims` exits 2;
  map Arabic cell text through `dim_rules`. A dim-column value outside the vocabulary is stored as `unknown` with a high
  `dim_out_of_vocabulary` issue (blocking for in-scope records in validation until a rule or override sets it).
- **Scope per sheet**: each source's `default_scope` must agree with its sheet's `decision` (primary/supplementary →
  in_scope, isolated → isolated, excluded → excluded); a missing `default_scope` takes the decision's scope. A source on a
  hidden/veryHidden sheet must state `default_scope` **and** `scope_reason`, else `sheet_scope_mismatch` (high) and the
  records default to isolated. In-scope records of a `managed_properties` sheet need `managed_for_others: true`
  (`managed_unflagged`, high).
- **Joins** (a confirmed link carries figures into existing records): a source may set
  `"join": {"into": "<source_id>", "on": "id|deed|name_key", "key_column": "<col>", "link_id": "<links[].id>"}`. It needs a
  `links[]` entry `{"id": "<link_id>", "status": "confirmed"}` (or `match: "confirmed"`), else exit 2. Matched rows merge
  their fields into the target record (each TV keeps its own `ref` plus `joined_from`; the target gets `source.joined[]`
  `{source_id, sheet, row, link_id}`). A field the target already holds with a different value is kept (`join_conflict`,
  medium); a key matching several targets is not merged (`join_ambiguous`, high); unmatched rows become isolated records
  (`join_unmatched`, medium). A join never creates in-scope records; join sources run after plain sources.
- **Totals**: total rows are auto-excluded (see §4). To keep a row the scripts read as a total, override it in
  `records/<id>` with `record_type` or `scope` **and** `scope_reason`/`notes`; its `total_like_row` issue then carries
  `mapping_override: true` and validation warns instead of blocking (unless its amounts equal the sum of its siblings).
- Duplicate ids, units/contracts beside their parent asset and overlapping sources block in validation unless the rows are
  excluded or overridden in `records/<id>` with a reason.

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
  "excluded_rows": [{"source_id": "REG", "sheet": "...", "row": 18, "reason": "صف إجمالي (SUBTOTAL)", "auto": true,
                     "values": {"rev.2025.excl_vat": TV, ...}}],
  "context": {"entity": {...}, "data_period": {...}, "study_date": {...}, "analysis_date": "...", "currency": "SAR",
              "valuation_dates": [...], "sheets": [...]},
  "stats": {"records": 1, "by_type": {...}, "by_scope": {...}, "pii_columns_dropped": 2, "rows_skipped_blank": 0,
            "rows_skipped_with_content": 0, "rows_excluded": 1, "rows_total_excluded": 0, "rows_joined": 0}
}
```
Auto checks in normalize: duplicate ids, possibly-duplicate names (`name_key`), shared deed numbers, total-like rows inside data range, `val.Y.total` vs `land+building` mismatch, `incl_vat < excl_vat`, text in numeric fields, zero valuations, mixed `units`, missing classification (`unknown` dims) for in-scope records, Hijri/Gregorian conversion notes. PII columns are dropped (count only).

Additional behaviour (issue types in brackets):
- **Total rows never become records** [`total_like_row`, high]: a total label in the name/id cell (`apa_common.is_total_label`),
  a short total label (≤ 4 words, no digits) in another mapped text cell, a SUBTOTAL/SUM-range formula in a numeric cell, or
  a `total*` row kind in `RUN/01_workbook_map.json` (used only when its sha256 matches the workbook). Such rows go to
  `excluded_rows` with `auto: true`, `values` and an issue whose `record_ids` is `[]` (plus `source_id`, `row`).
  Excluded total rows (auto or listed in `exclude_rows`) keep `values` `{field: TV}` for mapped numeric fields; separator
  and header rows do not. `facts.py` reads them as `excluded.<sid>.r<row>.<field>`.
- Blank rows [`blank_rows_skipped`, low, with `count`]; a row skipped by `skip_if_blank` that still holds mapped content
  [`row_skipped_with_content`, medium + an `excluded_rows` entry]. Name/id/key cells inside a vertical merge read the merge
  anchor and the record gets `source.merge_note`.
- A cached 0 of a formula whose referenced cells are all blank/unknown becomes `not_available`
  [`zero_from_blank_inputs`, medium; TV key `zero_from_blank_inputs`].
- '%' text or a percent-formatted number in a money/area/count column becomes `text` [`percent_in_amount`, medium; TV key
  `percent_in_amount`]. «غير متاح» tokens in text/url/date columns become `not_available`, «لا ينطبق» tokens `not_applicable`
  (raw kept); they are never dimension categories.
- Vocabulary [`dim_out_of_vocabulary`, high], joins [`join_conflict`, `join_unmatched`, `join_ambiguous`], sheet scope
  [`sheet_scope_mismatch`, `managed_unflagged`, high], the same sheet row read into more than one record by non-join
  sources [`overlapping_source_rows`, high], a PII column used as name/id [`pii_header_not_marked`, high].
- **Issue order and impact**: issues touching in-scope records (or none) come first, then those on isolated/excluded records
  only, each group by severity; DQ ids follow that order. An issue whose records are all isolated/excluded gets the impact
  «لا أثر على المؤشرات (سجل معزول)» (or the excluded/plural variant).

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
  "parity_max_values_per_dim": 3,
  "init_warnings": [{"code": "revenue_incl_vat|yield_periods_limited|no_in_scope_records", "period": "2025",
                     "field": "rev.2025.incl_vat", "text": "Arabic"}]   // written by --init-config; review before editing
}
```
`--init-config` rules: periods, the revenue/expense field of each period and the valuation field are chosen only from fields
with known values on **in-scope** records (isolated/excluded data never drives them; with no in-scope record there are no
periods and `init_warnings` has `no_in_scope_records`). A VAT-inclusive fallback revenue field makes every revenue label
say «شامل الضريبة» (`revenue_incl_vat`). `value_by_*`/`revenue_by_*` breakdowns are created only when their own records hold
≥ 2 distinct classified values of the dim. Revenue metrics, revenue breakdowns and the trend also require
`record_type ∈ {asset, land_parcel, operating_activity}` (unit/contract rows are never added to their asset).
`compute_metrics.py` exits 2 when a period or the trend is not marked `not_comparable` although field_meta shows a known
difference (see growth below), and on `eq`/`in` against a dim that holds a list in any record (use `has`).
To trim periods: delete the period object (and its id from every `valid_periods`), set `prev_id: null` on the next period
(its `prev_rev_field` stays valid for growth).

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
  "kind_tag": "fact|derived",        // statement kind of the result (counts of file rows = fact; computations = derived)
  "valid_periods": ["2025"],         // optional: other periods give not_comparable with invalid_period_reason
  "invalid_period_reason": "Arabic",
  "empty_text": "Arabic",            // shown when nothing in the file qualifies (empty_cause scope); also in text_by_period
  "empty_as_zero": true              // sum/count only: empty with cause scope -> value 0, status complete, note (count default true)
}
```
`--init-config` sets `valid_periods` on `gross_yield_investment`, `gross_yield_owned` and `gross_yield_income_producing` to the
periods whose year is the valuation year or the year after, `empty_as_zero` + `empty_text` «لا توجد أصول استثمارية بإيراد
صفري مسجل» on `zero_rev_investment_value`, and `empty_text` on `activity_revenue` and the count metrics.
Placeholders: `{REV}` → `periods[p].rev_field`; `{REV_PREV}` → `periods[p].prev_rev_field`; `{EXP}` → `periods[p].exp_field`; `{VAL}` → `valuation.field`. A placeholder resolving to `null` ⇒ status `not_computable`, `reason = unavailable_reason || "الحقل غير متاح لهذه الفترة"`.

### Where (predicate on a record)
- `{"all": [W...]}`, `{"any": [W...]}`, `{"not": W}`
- `{"eq": [attr, value]}`, `{"in": [attr, [values]]}`, `{"has": [attr, value]}` (attr is a list, e.g. ownership)
- `{"known": field}` — field status ∈ known; `{"status": [field, [statuses]]}`; `{"gt": [field, number]}` (known and v > n)
- `attr` resolution: `id`, `record_type`, `scope` at top level; otherwise `dims[attr]`. Missing attr ⇒ `null` (eq to null is allowed).
- Dashboard filters add `{"eq": [dim, value]}` for each selected dim (multi-valued dims use `has`).

### Semantics (Python `compute_metrics.py` and JS `metric_engine.js` MUST be identical)
Let `R` = records satisfying `base_where ∧ where ∧ filters`. `known(f)` uses the record's field status. Sums use `v`.
Check order for every metric: (0) `valid_periods` given and the period outside it ⇒ `not_comparable` (reason
`invalid_period_reason`); (1) unresolved placeholder ⇒ `not_computable`; (2) growth with `comparable_with_prev.status =
not_comparable` ⇒ `not_comparable`; (3) `|R| = 0` ⇒ `empty` with `empty_cause` (`filters` when a dashboard dim filter alone
empties a non-empty scope, else `scope`) and `reason` («لا تطابق التصفية الحالية أي سجل في نطاق هذا المؤشر», or
`empty_text`, or «لا توجد سجلات مؤهلة لهذا المؤشر في الملف»); sum/count with cause `scope` and `empty_as_zero` ⇒ value 0,
status `complete`, `note`; (4) the kind rules below.
- **count**: `value = |R|`; `status = complete` (or `empty` if |R| = 0, value 0).
- **sum**: `K = {r ∈ R : known(f)}`, `U = R \ K`. `value = Σ_K v`. status: `empty` if |R|=0; `not_computable` if |K|=0 (value null); `partial` if |U|>0; else `complete`. `excluded = U` with reason = field status.
- **ratio** (pairwise): `I = {r ∈ R : known(num) ∧ known(den) ∧ den.v > 0}`. `num_sum = Σ_I num.v`, `den_sum = Σ_I den.v`. `value = num_sum / den_sum × scale` (null if I empty). `excluded = R \ I` with reason ∈ `num_unknown | den_unknown | den_nonpositive` (first failing). status: `empty` | `not_computable` (I empty) | `partial` (excluded non-empty) | `complete`. Extra: `num_known_total = Σ_{r∈R, known(num)} num.v`; `coverage_num_share = num_sum / num_known_total` (null if total 0). **Never** report coverage by value.
- **diff**: `I = {r ∈ R : known(a) ∧ known(b)}`; `value = Σ_I (a.v − b.v)`; same status rules; if no record has field `b` at all ⇒ `not_computable` with reason «لا توجد بيانات مصروفات».
- **growth**: period p must define `prev_rev_field`; else `not_computable`. If `comparable_with_prev.status = not_comparable` ⇒ `not_comparable`, value null. A record known in both periods with prev 0 and curr > 0 is a **revenue start**: excluded with reason `revenue_start`, listed in `revenue_start_ids`, its current revenue summed into `start_change` (outside `prev_sum`, `curr_sum` and the %); `revenue_start = true` when any. `I` = the other records known in both periods. `prev_sum`, `curr_sum`, `abs_change = curr_sum − prev_sum`, `pct = abs_change / prev_sum × 100`. `I` empty with only revenue starts ⇒ `not_computable` («بدء إيراد»); `I` non-empty with `prev_sum ≤ 0` ⇒ `not_computable` (sums kept). `value = pct`. `only_curr` / `only_prev` = records known in one period only (reported separately, not in I). `comparability` copied from config.
  Comparability (`compute_metrics.comparability`, mirrored by the validator): `not_comparable` when either period is partial, a
  field's own period differs from the period id it stands for, or a key is known on both sides in different classes — VAT
  (`excl`/`none_evidenced` vs `incl`), basis (`actual`/`recorded` vs `expected`/`contract`), accrual (`collected` vs
  `accrued`); `confirmed` only when vat, basis and completeness are known and equal; otherwise `unverified`. The trend checks
  every pair of its periods.
- **concentration**: `K = {r ∈ R : known(f) ∧ v > 0}` sorted by v desc (tie → id asc). `total = Σ_K v`; `top_ids` = first `top`; `value = Σ_top v / total × 100` (null if total 0). Also `hhi = Σ s·s × 10000` with `s = v/total` (same arithmetic in both engines).
- **coverage**: for each field f in `fields`: records whose value is «لا ينطبق» (`not_applicable`) leave that field's denominator: `n_f = |R| − n_not_applicable_f`, share `n_known_f / n_f`. `value = min over fields with n_f > 0` × 100; `per_field = [{field, n_known, n, n_not_applicable}]`. No applicable field ⇒ `not_computable`. status `complete` unless |R| = 0.
- **trend**: points over R per field, plus `same_records = {n, ids, points: [{period, field, value, n_known}]}` summed over the records known in every trend period, and `n_known_varies` (true when the set of known records changes between periods). Show the same-records series when it differs.
- **breakdown**: group `K = {r ∈ R : known(field)}` by `dims[by]` (list dims: one row per element; label it «قد تتداخل الفئات»); `unknown` groups by `null` → key `"__none__"`. Each group: `value`, `n_known`; plus `n_unknown` per group from `R \ K`. Shares relative to `Σ_K v`.
- Rounding: never round stored values; the dashboard formats only for display.

### Result object (`05_metrics.json` → `metrics[id]`)
```jsonc
{"id": "...", "kind": "...", "label": "...", "value": 1.23, "status": "partial", "unit": "%",
 "n_scope": 10, "n_included": 7, "included_ids": [...], "excluded": [{"id": "REG-3", "reason": "den_unknown"}],
 "num": 123, "den": 456, "num_known_total": 130, "coverage_num_share": 0.946,
 "prev_sum": ..., "curr_sum": ..., "abs_change": ..., "revenue_start": false, "revenue_start_ids": [], "start_change": null,
 "only_curr": [], "only_prev": [], "top_ids": [...], "hhi": ..., "per_field": [...],
 "empty_cause": "filters|scope|null", "note": null,          // note: e.g. the empty_as_zero text
 "resolved_fields": {"num": "rev.2025.excl_vat", "den": "val.2024.total"},
 "reason": null, "definition": "...", "scope_text": "...", "basis": "...", "valuation_date": "...", "period_text": "..."}
```
`05_metrics.json` = `{"contract_version", "default_filters": {"period": "2025"}, "metrics": {...}, "breakdowns": {...}, "trend": {...}, "parity_cases": [{"filters": {...}, "metrics": {id: {"value", "status", "n_included"}}}]}`.
Parity cases: default; each period; for each filter dim up to `parity_max_values_per_dim` values.

## 6. `06_narrative.json` (authored by Claude)

```jsonc
{
  "contract_version": "1.0",
  "executive_summary": "120–180 Arabic words (counted as rendered). Numbers via {{m:metric_id}}, {{m:metric_id.field}}, {{m:metric_id.field.sub}} or {{f:fact_key}} tokens, rendered by the build.",
  "assessment": [{"dimension": "income_generation|income_stability|asset_utilization|data_quality",
                  "judgment": "وصف دون درجات", "criterion": null, "evidence": ["m:property_revenue", "DQ-002"], "kind": "inference"}],
  "strengths": [{"text": "...", "evidence": ["..."], "kind": "fact|derived|inference"}],
  "gaps": [{"id": "G1", "title": "...", "scope": "REG-1 | portfolio | group:use=investment",
            "evidence": [{"ref": "m:zero_rev_investment_value | DQ-003 | WB-002 | عقارات الجمعية !C3", "kind": "fact|derived"}],
            "impact": {"metric": "m:zero_rev_investment_value", "text": "..."},
            "proven_cause": null, "possible_causes": ["..."], "verification_needed": ["..."],
            "proposed_action": "...", "severity": "high|medium|low", "is_proven_problem": false}],
  "recommendations": [{"id": "R1", "title": "≤ 60 chars, optional heading", "gap_ids": ["G1"], "target": "...", "problem": "...", "evidence": ["..."],
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
  "sources": [{"type": "file|web|user", "ref": "...", "date": "...", "limits": "..."}],   // web: date and limits required
  "acknowledged_warnings": [{"check_id": "growth_comparability", "reason": "≥ 8 words: why the warning is accepted"}],
  "version_comparison": null,
  "recommendation_followup": [],
  "methodology_notes": ["..."]
}
```
Strict rules enforced by `schemas/narrative.schema.json` + `validate_analysis.py`:
- `assessment` has exactly 4 items, one per dimension. `criterion` null ⇒ descriptive judgment, no score.
- Gap `evidence` entries are objects `{ref, kind}`; refs are `m:<metric_id>` (also `m:<id>.<field>`, `m:<breakdown id>`,
  `m:trend`), `DQ-…` (03 issues), `WB-…` (01 anomalies), record ids, or a cell `Sheet!A1` (exact sheet name, incl. trailing
  spaces; quoting allowed). Unresolved gap refs are **blocking**; elsewhere they are warnings.
- `ranking.urgency|feasibility|cost_level` ∈ high|medium|low. Plan items may carry `status`
  («لم يبدأ» | «قيد التنفيذ» | «منجز» | «متعثر»); `recommendation_followup` items are `{rec_id, status, evidence, note}` —
  copy them from `compare_versions.py` output and mark «منجز» only with evidence or the user's word. `sources` has ≥ 1 item.
- Unknown top-level keys are rejected; deliberate extensions use the `x_` prefix.
- Token grammar (validator and build use the same regexes): `{{m:id}}`, `{{m:id.field}}`, `{{m:id.field.sub}}` (dict key or
  list index; `id` may be a metric, a breakdown or `trend`), `{{f:key}}` (a fact from `facts.compute_facts(run_dir)`; an
  unknown key is blocking, `fact_tokens`). Deeper tokens are malformed. The summary length is counted on rendered words
  (money 2, «غير قابل للاحتساب»/«غير قابل للمقارنة» 3, «بدء إيراد»/«غير متاح» 2, others 1).
- `acknowledged_warnings`: a **warning** whose check id is listed with a reason of ≥ 8 words becomes level `acknowledged`
  (not counted in `warning_count`, listed apart). Blocking checks can never be acknowledged.
- When recommendations exist, `plan.d30/d60/d90` each have items, `top_actions` (and `top_gaps` when there are gaps) is
  non-empty, and every top action appears in the plan (`plan_completeness`, blocking). Priorities are unique and consecutive.
- Numeric scores («82 من 100»، «8/10»، «درجة 4») are blocking anywhere (`narrative_scores`); a criterion must cite a
  `sources[]` entry or resolvable evidence. A drastic action (sell, exit, replace, demolish, change of use, liquidate) needs a
  study/verification step first and a reason beyond yield (`drastic_without_study`).
- Gap `scope`: a record id, `portfolio`, or `group:<dim>=<value>`; record ids anywhere in gaps/recommendations are
  auto-linked in the dashboard. Set `opportunity_type` on scenario items so overlaps are detected.
- Any ratio label/short_label/basis shown on a card must contain «قبل المصروفات»; affirmative «صافي» or «العائد
  الاستثماري الكلي» without the required inputs is blocking.

Scenario items (only when `status = "prepared"`, listed under `scenarios.items`): `{"id", "opportunity", "asset_ids", "assumptions": [{"text", "source"}], "conservative|base|optimistic": {"extra_revenue", "extra_expense", "net_effect", "capex", "start", "year1_effect", "run_rate_effect"}, "recurring": true, "overlaps_with": []}`.

### `07_validation.json` (validate_analysis.py)
`{"blocking_count", "warning_count", "acknowledged_count", "info_count", "warning_ids", "acknowledged_ids", "checks": [{"id",
"level": "blocking|warning|acknowledged|info", "passed", "detail", "where", "title"}], "reconciliation": [...], "inputs":
{..., "xlsx_out", "message"}}`. `reconciliation[]` (one per excluded total row and field with values):
`{source_id, sheet, row, ref, field, recorded, sum_in_scope, n_in_scope_known, n_in_scope_unknown, difference,
sum_all_records, matched_sum, status, status_ar}` with `status ∈ match | match_section | match_all_records |
match_section_all_records | match_without_text_numbers | mismatch | recorded_unknown` (same wording as the export sheet
«المطابقة مع الإجماليات»); only `mismatch` raises the `totals_reconciliation` warning.
Data-integrity checks (blocking): `total_like_in_scope`, `unit_double_count`, `duplicate_ids_in_scope`, `overlapping_rows`,
`dims_vocabulary`, `sheet_scope`, `mapping_context`, `fact_tokens`, `pii_export_patterns`; (warning): `total_like_records`,
`unit_levels_mixed`, `duplicate_ids_kept`, `overlapping_rows_other`, `hidden_sheet_in_scope`, `totals_reconciliation`,
`pii_review_patterns`, `message_numbers`. Narrative-quality checks (blocking): `period_fields`,
`growth_comparability_basis`, `yield_zero_revenue_scope`, `label_vat_basis`, `label_years`, `label_roi`,
`narrative_metric_naming`, `narrative_scores`, `drastic_without_study`, `plan_completeness`; (warning):
`yield_vat_inclusive`, `partial_unnamed`, `growth_token_caveat`, `promised_outcomes`, `scenario_consistency`, `plan_phasing`,
`top_lists_count`, `recommendation_quality`, `gap_completeness`, `gap_unsupported_claims`. Every value scan uses the shared
patterns of `assets/pii_patterns.json` (block level blocks, review level — national-ID-like numbers — warns) over 02, 03, 06,
`analysis.json`, the dashboard text and embedded data, `supporting_workbook.xlsx` (or `--xlsx-out`) and `09_message.md`
(whose hand-typed numbers are also matched against 05/facts).

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
