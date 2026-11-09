# Data contract (v1.0) — files exchanged between the analyst (Claude) and the scripts

> Technical spec, kept in English for precision. Every script and every hand-authored JSON follows it.
> Rule of thumb: **Claude decides meaning (mapping, classification, narrative); scripts do arithmetic, checks and rendering.** Claude never types a computed number into the dashboard or narrative by hand — numbers are referenced by metric id (`{{m:metric_id}}`), by fact key (`{{f:key}}`, §6) or copied from `05_metrics.json`.

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
| `08_dashboard_check.json` + `screens/` | `check_dashboard.mjs` | Browser test results (parity, filters, mobile, print, PII) and `warnings`; `screens/desktop.png` (full page), `desktop_01.png…` (1440×900 viewport crops), `details_open.png` (every details section open), `mobile.png`, `mobile_full.png`, `mobile_360.png`, `mobile_detail.png`, `mobile_table.png` (≤ 600 px asset cards), `print.pdf` |
| `version_comparison.json` | `compare_versions.py` | Only when a previous run really exists (§6 `version_comparison`) |
| `supporting_workbook.xlsx` | `export_workbook.py` | Reviewable Excel (RTL, filters, gray = not available with the original text of a non-numeric cell, red = proven problem only). Sheets: «ملخص»، «الأصول»، «الحسابات»، «المستبعدات»، «المطابقة مع الإجماليات» (when 03 keeps total-row values)، «مشكلات البيانات»، «الفجوات والتوصيات»، «الخطة التشغيلية»، «متابعة التوصيات» (when 06 has `recommendation_followup`; status dropdown «لم يبدأ/قيد التنفيذ/منجز/متعثر», «منجز» without evidence flagged)، «مقارنة الإصدارات» (when 06 has `version_comparison`; detail from `version_comparison.json` if 06 holds a summary)، «الافتراضات والمصادر»، «قوائم الاختيار» (the lists behind the plan's dropdowns). «الحسابات» recomputes every sum/ratio/diff/growth metric with live formulas over its included records (a literal 0 with the row «لا سجلات مشمولة…» when nothing is included), beside a row «القيمة المعروضة في الداشبورد» (numbers) and a row «التحقق» (=IF(…,"مطابق","فرق") per field and result); each «ملخص» KPI value is a live link to its recomputed cell when the two agree; every formula also stores its computed value, so a preview that does not recalculate shows numbers. A metric whose spec carries `review` reads «يحتاج تحقق — <reason>» as its status in «ملخص» and «الحسابات»; an unverified growth shows its change in riyals only (no % in any cell). «المطابقة مع الإجماليات» has «الفرق عن مجموع النطاق» (=D−E, live), «مجموع المقارنة» (the sum the total matched, or the comparison base on a mismatch), «أساس مجموع المقارنة» and «الفرق عن مجموع المقارنة» (=D−I, live), with a status text consistent with both differences. «الخطة التشغيلية» adds «المدة», «مؤشر النجاح» and «الاعتماد المطلوب», with list validation on «المرحلة» (strict), «المسؤول» (free text accepted) and «الحالة», and date validation on «تاريخ المتابعة»; the recommendations table has a «الحالة» dropdown. «الأصول» carries «عدد ملاحظات البيانات» and «ملاحظات البيانات» from the same list as the dashboard (03 issues on or naming the record, 06 `data_issues_extra`, 01 anomalies with its `record_ids`); recommendations appear in ascending priority with «الأولوية n» everywhere. Board wording: field and metric keys, issue types, gap scopes («portfolio»), evidence tokens («m:…») and check ids are shown by their Arabic labels; the keys appear only in code-styled «المرجع التقني» columns or rows (above the «الأصول» headers). «ملخص» lists the top gaps and the top actions as two separate tables (an action names the gaps it addresses), and «المطابقة مع الإجماليات» lists for each total the records of the sum it was compared with (a section total: its own section). Personal-data fields are dropped and block-level values masked; exit 2 and the file is deleted if the post-save scan still finds contact/IBAN data, or if a formula refers to its own cell (a circular reference). `--out` sets another path; `--self-test` runs built-in cases |
| `09_message.md` | **Claude** | The short delivery message |

## 1. `01_workbook_map.json` (inspect_workbook.py)

```jsonc
{
  "contract_version": "1.0",
  "file": {"path": "...", "name": "...", "sha256": "...", "size": 40988,
           "properties": {"creator": "<محجوب>", "last_modified_by": "<محجوب>", "created": "...", "modified": "..."}},
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
                 "samples": ["187500"], "mixed_types": false, "pii": false}],
    "row_kinds": {"3": "data", "4": "total_formula", "28-43": "filler"},
    "formulas": {"count": 36, "cached_missing": 0, "errors": 0},
    "comments": [{"ref": "...", "text": "..."}],
    "data_validations": [...]
  }],
  "anomalies": [{
    "id": "WB-001", "type": "formula_year_mismatch|cached_value_mismatch|total_mismatch|row_total_mismatch|hardcoded_total|excel_error|cached_missing|mixed_types|text_in_numeric|numeric_text|duplicate_name|shared_deed|serial_gap|pii_column|partial_period_header|garbled_header|filler_block|derived_from_total|count_mismatch|vat_identical|vat_inconsistent|hijri_date_text|zero_valuation|unlabeled_column|outside_table_rows|hidden_sheet|hidden_rows|hidden_columns|cross_sheet_name_match|mixed_calendar|formula_blank_inputs",
    "severity": "high|medium|low|info",
    "where": "سجل الأصول!N18",
    "detail": "Arabic explanation with numbers",
    "evidence": {"formula": "...", "cached": 2150000, "recomputed": 2388000},
    "record_ids": ["REG-3"]                         // optional: records the anomaly concerns (build keeps them; C3 count)
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

PII **values are never written** to any output; only the column location. Redacted samples and the file's author
properties (creator / last_modified_by) appear as «<محجوب>»: author metadata names a person or an account (even
initials), so both are masked unless the value is the name of the software that wrote the file (openpyxl, LibreOffice …).
The personal-data patterns (header words, field keys,
value patterns with a `block` or `review` level) live in one file, `assets/pii_patterns.json`, shared by every Python
script (`apa_common.is_pii_header`, `PII_KEY_RE`, `pii_value_hits`, `mask_pii`) and by `check_dashboard.mjs`;
values are matched after Arabic-Indic digits are translated. Block level = mobile, landline, e-mail, IBAN (never in an
output); review level = a 10-digit number starting with 1 or 2 (may be a national ID or a deed: reported, not masked). The map is an internal working file (it shows asset
names): only `workbook_map_summary` goes into `analysis.json`. Extra keys (`register_candidates`, `regions`,
`key_columns`, `summary`, …) are additive hints — Claude still decides the mapping.

**Converted inputs (.xls / .ods / .xlsb):** `file.converted.path` (RUN/_converted/<stem>.xlsx) is the copy every later
step reads (`normalize.py --xlsx`, `validate_analysis.py --xlsx`); the original stays untouched. LibreOffice may have
recalculated its formulas — say so in `methodology_notes`.

## 2. Typed value (`TV`) — used in every record field

```jsonc
{"v": 187500, "s": "value", "raw": "187500", "ref": "عقارات الجمعية !C3", "note": null}
```
`s` ∈ `value | derived | zero_recorded | zero_confirmed | zero_as_missing | blank | filler | text | error | not_applicable | not_available`.
**Known** statuses = `value, derived, zero_recorded, zero_confirmed` (v is a number; zero statuses have v = 0).
Everything else has `v = null` and is *unknown* — never summed as zero.
N/A tokens (`apa_common.NA_TOKENS`: «غير متوفر»، «غير متاح»، «لا يوجد»، «غير معروف»، «-»، «—»، «–»، «؟»، «n/a») give
`not_available` and «لا ينطبق»/«غير منطبق» give `not_applicable`, in numeric, text, url and date columns alike (raw kept).
Optional TV keys: `formula`, `no_cache`, `zero_from_blank_inputs`, `percent_in_amount`, `ambiguous_thousands`,
`foreign_currency`, `redacted`, `joined_from`, `rollup` (`{from, evidence}`), `override` (`{from_s, from_v, evidence}`),
`calendar`, `hijri`.
Zero policy per column (mapping `zero_policy`): `recorded` (default for revenue, expense, units, counts) → `zero_recorded`; `missing` (default for `val.*`, areas) → `zero_as_missing`.
Dates: `{"v": "1989-12-10", "s": "value", "raw": "1410/5/12هـ", "calendar": "hijri", "hijri": "1410-05-12"}`.
Text fields: `{"v": "حي النموذج", "s": "value", "raw": "...", "ref": "..."}`.

## 3. `02_mapping.json` (authored by Claude)

```jsonc
{
  "contract_version": "1.0",
  "source_file": "path/to/file.xlsx",
  "entity": {"name": null, "evidence": "اسم الجهة غير مذكور في الملف"},
  "currency": "SAR",
  "data_period": {"label": "2025", "evidence": "سجل الأصول!A1"},
  "study_date": {"value": null, "basis": "لا يذكر الملف تاريخ إعداد؛ آخر تعديل للملف (خصائص المصنف)", "evidence": "file.properties.modified"},
  // a date only the file properties give (last modified) is not the study date: build shows «غير مذكور في الملف» with
  // «آخر تعديل للملف المصدر: <date>» as context (meta.study_date_stated = false, meta.file_modified)
  "analysis_date": "2031-05-20",
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
  },
  {"id": "EXP", "sheet": "مصروفات 2025", "header_row": 2, "data_rows": {"from": 3, "to": 9},
   "default_scope": "in_scope", "scope_reason": "...", "default_record_type": "asset",
   "join": {"into": "REG", "on": "id", "key_column": "A", "link_id": "L1"},   // merge into REG records, see "Joins"
   "columns": {"A": {"field": "serial", "kind": "text"}, "C": {"field": "exp.2025.opex", "kind": "money"}}}],
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
              "notes": "..."},
    "W-1": {"rollup": {"from": ["W-2", "W-3", "W-4"], "fields": ["rev.2024.recorded"],   // units carry the revenue,
                       "evidence": "الإيراد مدون على الوحدات H4:H6 لا على صف العمارة"}}   // the building only its valuation
  },
  "links": [{"record": "REG-1", "other_ref": "تخصيص الايرادات!B12", "match": "probable|confirmed|rejected",
             "basis": "تشابه الاسم فقط؛ لا معرّف مشترك", "used": false, "reason": "..."},
            {"id": "L1", "status": "confirmed", "basis": "المسلسل «م» نفسه في الورقتين ومبالغه مطابقة", "used": true,
             "reason": "مصروفات تشغيل أصول السجل"}],   // a link needs record+match or id+status; join needs id+confirmed
  "derived_fields": [                               // optional, explicit derivations (never silent)
    {"field": "val.2023.total", "formula": "val.2023.land + val.2023.building", "when": "always|if_unknown",
     "reason": "الإجمالي المدون يستخدم أعمدة 2022 (خطأ صيغة)"}
  ]
}
```
Rules: dims vocabularies are in `assets/vocab.json`. `ownership` is a **list** (non-exclusive). `waqf` is separate. Fields that hold labels (e.g. "صافي الأصول" containing «غير مقيد») must be mapped as `kind: "text"`, never `money`.
Column `kind` ∈ `money | number | area | count | count_mixed | percent | date | text | pii | url | ignore`; a column whose
header looks personal (`apa_common.is_pii_header`) but whose values are not (e.g. «المالك» holding only the entity's own
name) takes `"pii_ok": true`,
otherwise it is dropped like `pii` with a high `pii_header_not_marked` issue. `record_type` ∈
`asset | land_parcel | unit | contract | operating_activity | total | other` and `scope` ∈ `in_scope | isolated | excluded`
(schema enums). When `managed_for_others` is not set, normalize derives it from `ownership` (contains
`managed_for_others` → true; empty/unknown → null; otherwise false).

Required context (schema, exit 2 when missing): `entity`, `data_period`, `study_date`, `currency`, `valuation_dates`,
`sheets`. Values may be null, but `entity`, `data_period`, `study_date` and every `valuation_dates[]` item need a non-empty
`evidence`; `sheets[]` has ≥ 1 item and must list **every** sheet of 01 (hidden and veryHidden included, none that 01 lacks —
`mapping_context` blocks otherwise). These keys plus `analysis_date` are copied to 03 under `context`.

More rules:
- `data_rows.from` and `data_rows.to` are required integers (inclusive); `from > to` exits 2, and a source that yields no
  record (every row blank, excluded or a total) exits 2 with an Arabic message. Rows after the sheet's last used row are
  counted as blank. A row whose mapped cells (name and data columns; an id or join key alone does not count) are all
  blank/filler is skipped and counted (`blank_rows_skipped`).
- `skip_if_blank: "<col>"` skips a row whose cell in that column is blank; if the row still holds other mapped content it
  becomes an `excluded_rows` entry and a medium `row_skipped_with_content` issue (never a silent drop).
- **Controlled dimensions** `use`, `income_nature`, `ownership` (list), `waqf`, `managed_for_others` (bool) take only codes
  from `assets/vocab.json`. A literal outside the vocabulary in `dim_rules` (`map`/`default`) or `records[].dims` exits 2;
  map Arabic cell text through `dim_rules`. A dim-column value outside the vocabulary is stored as `unknown` with a high
  `dim_out_of_vocabulary` issue (blocking for in-scope records in validation until a rule or override sets it).
- **Scope per sheet**: each source's `default_scope` must agree with its sheet's `decision` (primary/supplementary →
  in_scope, isolated → isolated, excluded → excluded); a missing `default_scope` takes the decision's scope (a sheet
  missing from `sheets[]`: isolated when hidden, in_scope when visible). Each of these raises a high
  `sheet_scope_mismatch` issue on the source's records (blocking in validation through `sheet_scope`): the sheet is not in
  `sheets[]`; `default_scope` differs from the decision; a source on a hidden/veryHidden sheet without **both**
  `default_scope` and `scope_reason`; a join source whose sheet decision is not primary/supplementary. The records keep
  the scope applied (it is not silently changed). In-scope records of a `managed_properties` sheet need
  `managed_for_others: true` (`managed_unflagged`, high, blocking). In-scope records from a hidden sheet raise the
  `hidden_sheet_in_scope` warning in validation.
- **Joins** (a confirmed link carries figures into existing records): a source may set
  `"join": {"into": "<source_id>", "on": "id|deed|name_key", "key_column": "<col>", "link_id": "<links[].id>"}`. It needs a
  `links[]` entry `{"id": "<link_id>", "status": "confirmed"}` (or `match: "confirmed"`), else exit 2. Matched rows merge
  their fields into the target record (each TV keeps its own `ref` plus `joined_from`; the target gets `source.joined[]`
  `{source_id, sheet, row, link_id}`). A field the target already holds with a different value is kept (`join_conflict`,
  medium); a key matching several targets is not merged (`join_ambiguous`, high); unmatched rows become isolated records
  (`join_unmatched`, medium). A join never creates in-scope records; join sources run after plain sources.
  Exit 2 also when `into` is missing, is the source itself or is another join source; `on: "id"` needs
  `id_rule.column` on the target source, `on: "name_key"` its `name_column`, `on: "deed"` a `deed_no` field; the
  `key_column` must not be a `pii` column. The key column's own field is not merged; blank cells never overwrite; a
  dimension value is merged only when the target's is unknown (a different known value is a `join_conflict`).
  Unmatched/ambiguous rows take the join source's `default_record_type` (or `other`). Joined rows are not counted as
  overlapping rows, and their values are left out of the totals reconciliation.
- **Totals**: total rows are auto-excluded (see §4). To keep a row the scripts read as a total, override it in
  `records/<id>` with `record_type` or `scope` **and** `scope_reason`/`notes`; its `total_like_row` issue then carries
  `mapping_override: true` and validation warns instead of blocking (unless its amounts equal the sum of its siblings).
- Duplicate ids, units/contracts beside their parent asset and overlapping sources block in validation unless the rows are
  excluded or overridden in `records/<id>` with a reason (an override counts as evidenced when it sets `record_type` or
  `scope` **and** a non-empty `scope_reason` or `notes`; the check then warns). Normalize keeps both copies of a duplicate
  id (the second becomes `<id>-r<row>`, high `duplicate_id`); `duplicate_ids_in_scope` blocks while both are in scope.
  A unit/contract is tied to a parent asset by an id branch (`REG-1.2` under `REG-1`), the same deed, or a name that
  starts with the parent's name (≥ 2 words); in-scope units/contracts with revenue/valuation amounts and no parent only
  warn (`unit_levels_mixed`).
- Lines of an isolated sheet that name properties (allocation lines, a second register) are typed by what they describe
  (`asset`, `land_parcel` or `unit`) with `scope: isolated`, not `other`: they then count in the fact
  `records.isolated_assetlike` and the coverage card says the register may be incomplete. The fact and the card use
  one rule (facts.py): isolated records typed `asset`/`land_parcel`/`unit`, or read from a sheet whose `sheets[]` role
  is `asset_register`, `managed_properties` or `allocation`. `other` is for section headers, notes and lines that
  describe no property.

## 4. `03_records.json` (normalize.py)

```jsonc
{
  "contract_version": "1.0",
  "source": {"file": "...", "name": "...", "sha256": "...", "mapping_source_file": "..."},
  "currency": "SAR",
  "field_meta": { ... copied from mapping ... },
  "records": [{
    "id": "REG-1", "name": "عمارة ...",
    "source": {"source_id": "REG", "sheet": "...", "row": 3, "ref_name": "...!B3", "ref_id": "...!A3",
               "merge_note": "...",                       // only when a name/id cell was read from a vertical merge
               "joined": [{"source_id": "EXP", "sheet": "...", "row": 5, "link_id": "L1"}]},   // only after a join
    "record_type": "asset", "scope": "in_scope", "scope_reason": null,
    "dims": {"asset_type": "عمارة", "location": "...", "district": "حي النموذج", "use": "investment", "ownership": ["owned"],
             "waqf": "unknown", "income_nature": "rent", "managed_for_others": false, "rent_allocation": "...", "restriction": "..."},
    "dims_evidence": {...},
    "fields": {"rev.2025.excl_vat": TV, "val.2024.total": TV, "units": TV, ...},
    "issue_ids": ["DQ-004"]
  }],
  "issues": [{"id": "DQ-001", "type": "...", "severity": "high|medium|low", "where": "ref", "record_ids": ["REG-1"],
              "detail": "...", "impact": "which metric is affected and how", "treatment": "how it was handled",
              "field": "...", "source_id": "...", "row": 18, "rows": [...], "count": 3, "dim": "...",
              "mapping_override": true, "dev_note": "..."}],   // extra keys only where the issue type has them
  "excluded_rows": [{"source_id": "REG", "sheet": "...", "row": 18, "reason": "صف إجمالي (SUBTOTAL)", "auto": true,
                     "values": {"rev.2025.excl_vat": TV, ...}},
                    {"source_id": "REG", "sheet": "...", "row": 20, "auto": true, "after_data_rows": true,
                     "reason": "صف إجمالي بعد نطاق البيانات — للمطابقة فقط: …", "values": {...}}],
  "context": {"entity": {...}, "data_period": {...}, "study_date": {...}, "analysis_date": "...", "currency": "SAR",
              "valuation_dates": [...], "sheets": [...]},
  "stats": {"records": 1, "by_type": {...}, "by_scope": {...}, "by_source": {...}, "pii_columns_dropped": 2,
            "pii_values_dropped": 0, "rows_skipped_blank": 0, "rows_skipped_with_content": 0, "rows_excluded": 1,
            "rows_total_excluded": 0, "rows_total_after_range": 1, "rows_joined": 0, "issues": 9,
            "issues_by_severity": {...}}
}
```
`excluded_rows[].reason` is the `exclude_rows` reason, «صف إجمالي مكتشف آليًا: …» (`auto: true`) or «تم تخطي الصف لأن خانة …
(skip_if_blank) فارغة رغم وجود بيانات في: …»; only total rows carry `values`.
Issue texts are board-facing (dashboard «سجل مشكلات البيانات», workbook «مشكلات البيانات»): `detail`/`impact`/`treatment`
name fields by their `field_meta` label in «…» and metrics by Arabic names, never `rev.2025.excl_vat`, `gross_yield_*` or
engine words (`zero_as_missing`, `status text`); counts agree with the noun through `apa_common.ar_count`/`ar_n` («سجل
واحد»، «سجلان»، «4 سجلات»، «15 سجلًا»، «105 صفوف»، «لا نصوص»; `case="gen"` after a preposition: «بين سجلين»), in
normalize and in the inspect anomaly texts alike. Instructions for the analyst (which
part of `02_mapping.json` to change: `derived_fields`, `records`, `dim_rules`, `pii` …) go to `dev_note`, which no output
shows; the key stays in `field`. A field `<f>__original` (the stored value a documented `derived_fields` override replaced)
has no `field_meta` entry: outputs label it «<label of f> كما ورد (قبل الاشتقاق)» and flag it «قيمة مخزنة معيبة — غير مستخدمة».
Auto checks in normalize (issue types): `duplicate_id` (high), `possible_duplicate_name`, `shared_deed`, `total_like_row`,
`valuation_total_mismatch`, `valuation_component_formula`, `vat_incl_lt_excl`, `vat_equal`, `text_in_numeric`, `excel_error`,
`zero_as_missing`, `numeric_text_converted`, `count_mixed_text`, `formula_no_cache`, `hijri_conversion`, `date_note`,
`date_unparsed`, `missing_classification` (unknown `use`/`income_nature` on in-scope records), `pii_redacted` (high: a
block-level pattern of `assets/pii_patterns.json` — phone, e-mail, IBAN, a national ID next to «هوية/سجل مدني/إقامة/ID» — in a
text field; the personal span is replaced by «<محجوب>» and the rest of the text kept, a value that is only a phone is
dropped), `ambiguous_thousands_separator` (medium: «12.500» in a money/area/count column stays text — thousands separator or
decimal), `foreign_currency` (high: «USD 100,000»، «$»، «دولار» in a money column of another currency stays text),
`percent_scale_mixed` (medium: one percent field holds both fractions ≤ 1 and percentages > 1), `possible_total_row`,
`section_label_skipped` (low), `rollup_applied` (low), `rollup_incomplete`/`rollup_conflict` (medium),
`rollup_unknown_record` (high),
`pii_header_not_marked`, `derived_field`, `field_override`, `mapping_unknown_record`, `mapping_sheet_name`, plus the
types below. PII columns are dropped (count only).

Additional behaviour (issue types in brackets):
- **Total rows never become records** [`total_like_row`, high]: a total label in the name/id cell (`apa_common.is_total_label`),
  a short total label (≤ 4 words, no digits) in another mapped text cell, a SUBTOTAL/SUM-range formula in a numeric cell, or
  a `total*` row kind in `RUN/01_workbook_map.json` (used only when its sha256 matches the workbook). Such rows go to
  `excluded_rows` with `auto: true`, `values` and an issue whose `record_ids` is `[]` (plus `source_id`, `row`).
  A row with a real id **and** a non-total name is a record even when only the weak signs fire (a «إجمالي/صافي» rent-basis
  column, a building row whose formula sums its unit rows, 01's row kind): it is kept with [`possible_total_row`, medium]
  for review, never dropped silently. Inspect does not label such a row `total_label` either.
- **Totals after the data rows** (inspect's `data_rows` usually stop before them): 01 `regions[].total_rows` below
  `data_rows.to` inside the source's region (without 01: the next rows whose name/id/text cell is a total label) are added to
  `excluded_rows` with `auto: true`, `after_data_rows: true` and their `values`, for reconciliation only.
  Excluded total rows keep `values` `{field: TV}` for mapped numeric fields: auto-detected ones, and rows listed in
  `exclude_rows` that read as a total or whose reason names one («إجمالي»، «مجموع»، «جملة»، total، sum); separator and
  header rows do not. `facts.py` reads them as `excluded.<sid>.r<row>.<field>`, and validation reconciles them
  (`totals_reconciliation`, §6).
- Blank rows [`blank_rows_skipped`, low, with `count`]; a row whose only content is a label in the id/key column («Jeddah»)
  or that 01 classed `section_label` [`section_label_skipped`, low, with `rows`]; a row skipped by `skip_if_blank` that still holds mapped content
  [`row_skipped_with_content`, medium + an `excluded_rows` entry]. Name/id/key cells inside a vertical merge read the merge
  anchor and the record gets `source.merge_note`.
- A cached 0 of a formula whose referenced cells are all blank/unknown becomes `not_available`
  [`zero_from_blank_inputs`, medium; TV key `zero_from_blank_inputs`].
- '%' text or a percent-formatted number in a money/area/count column becomes `text` [`percent_in_amount`, medium; TV key
  `percent_in_amount`]; in a `percent`/`number` column the text «12%» is read as 12, and in a `percent` column a
  percent-formatted cell (0.12) is stored × 100.
  N/A tokens (§2) become `not_available`/`not_applicable` in every column kind (raw kept); they are never dimension
  categories.
- Numbers: in a money/area/count column «12.500» (one dot + exactly three digits) is ambiguous and stays text
  [`ambiguous_thousands_separator`]; an area «500 m2»/«٨٠٠ م²»/«300 متر مربع» is read without its unit (note kept); a money text
  in another currency than `currency` stays text [`foreign_currency`, high]; «−» (U+2212) is a minus.
- **Rollup** (`records/<parent>.rollup`): when revenue sits on unit rows and the building row holds only the valuation, the
  children's known values of `fields` are summed into the parent's unknown field (status `derived`, `ref` = the children's
  cells, TV key `rollup`), and the children go to `scope: excluded` unless their own override states a scope. Any unknown
  child value leaves the parent unknown [`rollup_incomplete`]; a known parent value is kept [`rollup_conflict`].
- Vocabulary [`dim_out_of_vocabulary`, high], joins [`join_conflict`, `join_unmatched`, `join_ambiguous`; the join index
  prefers targets that are not excluded and are asset/land_parcel/operating_activity after `records/` overrides, so units
  sharing their building's deed do not make every deed ambiguous], sheet scope
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
  "breakdowns_reason": "...",                       // optional: board wording of why no distribution is drawn
  "parity_max_values_per_dim": 3,
  "init_warnings": [{"code": "revenue_incl_vat|yield_periods_limited|no_in_scope_records|expense_fields_partial", "period": "2025",
                     "field": "rev.2025.incl_vat", "text": "Arabic"}]   // written by --init-config; review before editing
}
```
`--init-config` rules: periods, the revenue/expense field of each period and the valuation field are chosen only from fields
with known values on **in-scope** records (isolated/excluded data never drives them; with no in-scope record there are no
periods and `init_warnings` has `no_in_scope_records`). A VAT-inclusive fallback revenue field makes every revenue label
say «شامل الضريبة» (`revenue_incl_vat`). `value_by_*`/`revenue_by_*` breakdowns are created only when their own records hold
≥ 2 distinct classified values of the dim. Revenue metrics, revenue breakdowns and the trend also require
`record_type ∈ {asset, land_parcel, operating_activity}` (unit/contract rows are never added to their asset; use a
`rollup` in 02 when the amounts are recorded only on the unit rows). NOI subtracts one expense field per period — a field
named `exp.<P>.total` (e.g. a `derived_fields` sum of the components) is preferred; when other expense fields of the period
are left out, `init_warnings` gets `expense_fields_partial` and the NOI source text says «لا يشمل: …». `activity_revenue` and
`operating_records_count` are not cards when no operating activity is in scope.
`compute_metrics.py` exits 2 (Arabic message, nothing written) on a schema error and on these semantic errors: a period
or the trend not marked `not_comparable` although field_meta shows a known difference (see growth below); `eq`/`in`
against a dim that holds a list in any record (use `has`), in `base_where`, a metric, a breakdown or the trend; an
`eq`/`in`/`has` literal on `use`, `income_nature`, `record_type`, `scope`, `ownership` or `waqf` that is not a key of
`assets/vocab.json` (e.g. «Investment»); a
`prev_id` that is not a period id; duplicate period, metric or breakdown ids; a `valid_periods` entry that is not a
period id; `empty_as_zero` on a kind other than sum/count.
To trim periods: delete the period object (and its id from every `valid_periods`), set `prev_id: null` on the next period
(its `prev_rev_field` stays valid for growth). The oldest period gets `comparable_with_prev = {"status": "unverified",
"reason": "لا توجد فترة سابقة بحقل إيراد"}` and no `prev_rev_field` (growth → not_computable).

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
  "card_group": "counts",            // cards sharing a group are drawn as one card (≤ 8 groups)
  "prev": "{REV_PREV}", "curr": "{REV}",   // growth (defaults)
  "text_by_period": {"2024": {"label": "...", "period_text": "..."}},   // per-period overrides of the text keys
  "valid_periods": ["2025"],         // optional: other periods give not_comparable with invalid_period_reason
  "invalid_period_reason": "Arabic",
  "empty_text": "Arabic",            // shown when nothing in the file qualifies (empty_cause scope); also in text_by_period
  "empty_as_zero": true,             // sum/count only: empty with cause scope -> value 0, status complete, note (count default true)
  "review": {"status": "needs_verification", "reason": "Arabic sentence, 6–25 words, no digits"}   // optional analyst flag
}
```
`review` (optional) marks a metric whose value needs verification before the board relies on it: an input disputed by
another source in the file (a second figure for the same period's revenue in another statement), or a value that covers
only part of the portfolio while it looks complete (100% coverage of a one-asset register). Both engines copy it unchanged
into every result of the metric (`05_metrics.json` → `metrics[id].review`, and `APA.compute` in the dashboard), so it holds
under every period and filter. The dashboard shows the amber chip «يحتاج تحقق» instead of the green «مكتمل…» chip with
the reason as a caveat line (screen and print); the workbook shows «يحتاج تحقق — <reason>» as the metric's status.
`status` ∈ `needs_verification` (schema); `validate_analysis.py` blocks (`metric_review`) an unknown status, a reason
outside 6–25 words or with digits, and a 05 whose `review` differs from 04 (re-run `compute_metrics.py`).
`--init-config` sets `valid_periods` on `gross_yield_investment`, `gross_yield_owned` and `gross_yield_income_producing` to the
periods whose year is the valuation year or the year after (warning `yield_periods_limited` when that drops periods),
`empty_as_zero` + `empty_text` «لا توجد أصول استثمارية بإيراد صفري مسجل» on `zero_rev_investment_value`, and `empty_text`
on `activity_revenue` and the count metrics (`asset_count`, `operating_records_count`, `managed_records_count`).
`managed_records_count` (records with `managed_for_others = true`) is a card in the `counts` group only when the scope holds
such records. `asset_count`, `recorded_value`, `data_coverage` and `value_concentration_top3` take asset/land_parcel
records that are not managed for others.
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
- **count**: `value = |R|`; `status = complete`. |R| = 0: cause `scope` (count default `empty_as_zero: true`) ⇒ value 0,
  `complete`, `note`; cause `filters` ⇒ `empty`, value 0.
- **sum**: `K = {r ∈ R : known(f)}`, `U = R \ K`. `value = Σ_K v`. status: `empty` if |R|=0; `not_computable` if |K|=0 (value null); `partial` if |U|>0; else `complete`. `excluded = U` with reason = field status.
- **ratio** (pairwise): `I = {r ∈ R : known(num) ∧ known(den) ∧ den.v > 0}`. `num_sum = Σ_I num.v`, `den_sum = Σ_I den.v`. `value = num_sum / den_sum × scale` (null if I empty). `excluded = R \ I` with reason ∈ `num_unknown | den_unknown | den_nonpositive` (first failing). status: `empty` | `not_computable` (I empty) | `partial` (excluded non-empty) | `complete`. Extra: `num_known_total = Σ_{r∈R, known(num)} num.v`; `coverage_num_share = num_sum / num_known_total` (null if total 0). **Never** report coverage by value.
- **diff**: `I = {r ∈ R : known(a) ∧ known(b)}`; `value = Σ_I (a.v − b.v)`; same status rules; if no record has field `b` at all ⇒ `not_computable` with reason «لا توجد بيانات مصروفات».
- **growth**: period p must define `prev_rev_field`; else `not_computable`. If `comparable_with_prev.status = not_comparable` ⇒ `not_comparable`, value null. A record known in both periods with prev 0 and curr > 0 is a **revenue start**: excluded with reason `revenue_start`, listed in `revenue_start_ids`, its current revenue summed into `start_change` (outside `prev_sum`, `curr_sum` and the %); `revenue_start = true` when any. `I` = the other records known in both periods. `prev_sum`, `curr_sum`, `abs_change = curr_sum − prev_sum`, `pct = abs_change / prev_sum × 100`. `I` empty with only revenue starts ⇒ `not_computable` («بدء إيراد»); `I` non-empty with `prev_sum ≤ 0` ⇒ `not_computable` (sums kept). `value = pct`. `only_curr` / `only_prev` = records known in one period only (reported separately, not in I). Other `excluded` reasons: `prev_unknown`, `curr_unknown`, `both_unknown`; status `partial` when anything is excluded. `comparability` copied from config and `prev_period` = `prev_id`.
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
 "comparability": {"status": "...", "reason": "..."}, "prev_period": "2024",   // growth only
 "n_unknown": 0,                                             // sum only
 "unresolved": "{EXP}",                                      // only when a placeholder resolved to null
 "review": {"status": "needs_verification", "reason": "..."},  // only when the spec has one (copied unchanged)
 "scope_ids": ["REG-1"],                                     // only for not_computable with nothing included or excluded:
                                                             // the scope's records, counted and listed with the reason
 "resolved_fields": {"num": "rev.2025.excl_vat", "den": "val.2024.total"},
 "reason": null, "definition": "...", "scope_text": "...", "basis": "...", "valuation_date": "...", "period_text": "..."}
```
`05_metrics.json` = `{"contract_version", "default_filters": {"period": "2025"}, "metrics": {...}, "breakdowns": {...}, "trend": {...}, "parity_cases": [{"filters": {...}, "metrics": {id: {"value", "status", "n_included"}}}]}`.
The trend result holds `points[{period, field, value, n_known, n}]`, `comparable`, `same_records`, `n_known_varies`,
`empty_cause`; breakdowns and the trend take `empty_cause` like metrics but are never zero-filled.
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
  "version_comparison": null,                    // or the compare_versions.py output / a summary of it (see below)
  "recommendation_followup": [{"rec_id": "R2", "status": "لم يبدأ", "evidence": null, "note": "..."}],
  "methodology_notes": ["..."]
}
```
Strict rules enforced by `schemas/narrative.schema.json` + `validate_analysis.py`:
- `assessment` has exactly 4 items, one per dimension. `criterion` null ⇒ descriptive judgment, no score.
- Gap `evidence` entries are objects `{ref, kind}`; refs are `m:<metric_id>` (also `m:<id>.<field>`, `m:<breakdown id>`,
  `m:trend`), `f:<fact key>` (a fact of `facts.py`, `f:rec.<record_id>.<field_id>` included), `DQ-…` (03 issues), `WB-…`
  (01 anomalies), record ids, or a cell `Sheet!A1` (exact sheet name, incl. trailing spaces; quoting allowed). Unresolved
  gap refs are **blocking**; elsewhere they are warnings.
- `ranking.urgency|feasibility|cost_level` ∈ high|medium|low. Plan items may carry `status`
  («لم يبدأ» | «قيد التنفيذ» | «منجز» | «متعثر»); `recommendation_followup` items need `rec_id` and `status` (same four
  values; `evidence` and `note` optional) — copy them from `compare_versions.py` output (its extra keys `title`, `target`,
  `first_action`, `success_kpi`, `records`, `data_signals` are allowed and shown) and mark «منجز» only with evidence or
  the user's word (`followup_status` warns otherwise). `sources` has ≥ 1 item.
- `version_comparison` is an object (or null). The dashboard and the export read the `compare_versions.py` shape:
  `categories.<key>` lists with `category_labels`, `counts`, `prev`/`curr`, `period_relation`, `growth_note`,
  `matching`, `metric_deltas`; a hand summary may give `counts` + `examples[{category, ...}]` instead. When 06 holds only
  a summary, the export takes the detail from `RUN/version_comparison.json`; when 06 has no `recommendation_followup`,
  the dashboard shows the one inside `version_comparison`.
- Recommendation `title` (optional, 1–60 chars) is the heading of the recommendation (else `target`) and of its
  top-action card (else `first_action`, then `target`); the follow-up table and sheet use it too.
- Unknown top-level keys are rejected; deliberate extensions use the `x_` prefix.
- Token grammar (validator and build use the same regexes): `{{m:id}}`, `{{m:id.field}}`, `{{m:id.field.sub}}` (dict key or
  list index; `id` may be a metric, a breakdown or `trend`; e.g. `{{m:revenue_growth.comparability.status}}`,
  `{{m:trend.same_records.n}}`), `{{f:key}}` (a fact from `facts.compute_facts(run_dir)`; an unknown key is blocking,
  `fact_tokens`; an unknown metric is blocking, `metric_tokens`). Deeper tokens are malformed. The summary length is
  counted on rendered words (money 2, «غير قابل للاحتساب»/«غير قابل للمقارنة» 3, «بدء إيراد»/«غير متاح» 2, others 1).
- **Facts** (`scripts/facts.py`; list a run's keys with `python3 "$SKILL_DIR/scripts/facts.py" --run-dir "$RUN"`): numbers the board needs
  that are not metrics. Keys: `records.total|in_scope|isolated|excluded|excluded_rows|isolated_assetlike`,
  `records.in_scope.<record_type>`, `sources.<source_id>.records|isolated`, `sheets.total|hidden`,
  `issues.total|high|medium|low`, `anomalies.total|high`, `wb.<WB-id>.<evidence_key>` (numeric evidence of a 01 anomaly,
  e.g. `wb.WB-004.recomputed`), `excluded.<source_id>.r<row>.<field>` (values kept on excluded total rows),
  `sources.<source_id>.<field>.sum|n_known|n_zero|max|max_share` (a revenue/expense field over **every** record of a source,
  isolated included; `max_share` is the largest line as a % of the sum, unit `pct`) and `sections.<source_id>.r<row>.records`
  (records between the previous total row and row `<row>`) — to describe an isolated statement apart from the metrics (lines
  per section, its total, its largest line, its zero lines) without adding it to them. Use them
  instead of counting isolated records with a metric (that trips `isolated_in_metrics`).
  `rec.<record_id>.<field_id>` is the known numeric value of one field of **any** record of 03 (in scope, isolated or
  excluded), formatted like the other facts (revenue/expense/valuation fields as money): e.g.
  `{{f:rec.STM-r7.rev.2025.recorded}}` states the figure another statement gives beside `{{m:property_revenue}}`. An
  unknown record, an unknown field, or a value that is not a known number (blank, text, error, «غير متاح») is an
  unresolvable token (`fact_tokens`, blocking) and an unresolved `f:` evidence ref. `facts.compute_facts` — the single
  resolver of `build_dashboard.py` and `validate_analysis.py` — adds the `rec.*` keys that 06 cites (tokens and `f:`
  evidence refs), so `analysis.json.facts` holds exactly those; `facts.py --run-dir` prints the `rec.*` pattern with
  one example instead of every record and field. A record id with characters outside `[A-Za-z0-9_.-]` cannot be cited.
- `acknowledged_warnings`: a **warning** whose check id is listed with a reason of ≥ 8 words becomes level `acknowledged`
  (not counted in `warning_count`, listed apart). Blocking checks can never be acknowledged; an unknown id, a blocking id or
  a short reason adds an `acknowledged_warnings` warning, and an id with no current warning adds an info entry.
- When recommendations exist, `plan.d30/d60/d90` each have items, `top_actions` (and `top_gaps` when there are gaps) is
  non-empty, and every top action appears in the plan (`plan_completeness`, blocking). Recommendation priorities are
  unique positive integers and `top_actions` lists recommendation ids in **ascending priority**
  (`top_actions_priority`, blocking), so «الأولوية n» is one number everywhere: the dashboard top-3 and full list, the
  workbook «ملخص» and «الخطة التشغيلية», and the 09 message. Priorities should be consecutive (1…n), consistent with
  the ranking, and the top actions should be the recommendations with the lowest numbers (`recommendation_quality`,
  warning).
- Board-facing narrative (summary, gaps, recommendations, plan, 09) names assets by their names, not record codes
  (codes stay in evidence chips, tables and `scope`); a figure disputed by another source in the file is stated with
  both figures as tokens (`{{m:…}}` beside `{{f:rec.…}}`, or the ratio and change each would give), never as an
  approximate share («بنحو الثلث»).
- Numeric scores («82 من 100»، «8/10»، «درجة 4») are blocking anywhere (`narrative_scores`); a criterion must cite a
  `sources[]` entry or resolvable evidence. A drastic action (sell, exit, replace, demolish, change of use, liquidate) needs a
  study/verification step first and a reason beyond yield (`drastic_without_study`).
- Gap `scope`: a record id, `portfolio`, or `group:<dim>=<value>`; record ids anywhere in gaps/recommendations are
  auto-linked in the dashboard. Two scenario items sharing an asset must name each other in `overlaps_with` (else
  `scenarios` blocks, whatever their `opportunity_type`); numeric assumptions cite a `sources[]` entry or a file ref,
  placeholder sources («افتراض»، «تقدير»، «المحلل») block, and `net_effect` must equal `extra_revenue − extra_expense`.
- Any ratio label/short_label/basis shown on a card must contain «قبل المصروفات»; affirmative «صافي» or «العائد
  الاستثماري الكلي» without the required inputs is blocking.

Scenario items (only when `status = "prepared"`, listed under `scenarios.items`): `{"id", "opportunity", "asset_ids", "assumptions": [{"text", "source"}], "conservative|base|optimistic": {"extra_revenue", "extra_expense", "net_effect", "capex", "start", "year1_effect", "run_rate_effect"}, "recurring": true, "overlaps_with": []}`.

### `07_validation.json` (validate_analysis.py)
`{"generator", "generated_at", "run_dir", "passed", "blocking_count", "warning_count", "acknowledged_count", "info_count",
"blocking_ids", "warning_ids", "acknowledged_ids", "checks": [{"id", "level": "blocking|warning|acknowledged|info", "passed",
"detail", "where", "title", "skipped"?, "acknowledged_reason"?}], "reconciliation": [...], "inputs": {<run file>: sha16,
"xlsx", "html", "xlsx_out", "message"}}`. Exit 0 = no blocking finding, 2 = blocking findings, 1 = missing run folder. A
check that cannot run (missing input) is an `info` entry with `skipped: true`, never a pass; a crash is a failure at the
check's level. `acknowledged_warnings` entries are added outside the registry (the `scenario_overlap_cross_type` hook still
exists but no longer emits: every undeclared same-asset overlap is now blocking under `scenarios`).
`reconciliation[]` (one per excluded total row and field with values):
`{source_id, sheet, row, ref, field, recorded, compared_sum, difference, difference_basis, matched_sum, sum_in_scope,
n_in_scope_known, n_in_scope_unknown, sum_all_records, status, status_ar}` with `difference` = recorded −
`compared_sum`, the sum named by `difference_basis` ∈ `in_scope | in_scope_section | all_records | all_records_section`:
the in-scope records (or those of the total's own section for a section total), or **all** the source's records when
none is in scope or the match is on all records (an isolated statement never shows «recorded − 0»). The difference is
never taken against the matched sum, so a 0 never stands beside two different numbers: `matched_sum` (shown in its own
column) is the sum that equals the recorded total when one does — e.g. `match_original` shows the in-scope sum, a
non-zero difference and the pre-derivation sum that matches; `match_without_text_numbers` shows the difference made by
numbers stored as text. The export sheet «المطابقة مع الإجماليات» follows the same definition. `status ∈ match | match_section | match_all_records |
match_section_all_records | match_without_text_numbers | match_original | mismatch | recorded_unknown` (same wording as
the export sheet «المطابقة مع الإجماليات»; `match_original`: the total equals the sum of `<field>__original`, the values
before a documented `derived_fields` override); only `mismatch` raises the `totals_reconciliation` warning.
Data-integrity checks (blocking): `total_like_in_scope`, `unit_double_count`, `duplicate_ids_in_scope`, `overlapping_rows`,
`dims_vocabulary`, `sheet_scope`, `mapping_context`, `fact_tokens`, `pii_export_patterns`, `metric_review` (§5 `review`),
`top_actions_priority` (§6); (warning): `total_like_records`,
`unit_levels_mixed`, `duplicate_ids_kept`, `overlapping_rows_other`, `hidden_sheet_in_scope`, `totals_reconciliation`,
`pii_review_patterns`, `message_numbers`, `export_growth_pct` (the supporting workbook shows a growth % whose comparability
is not confirmed: a value cell or a «x.x%» text outside the review note; skipped while the workbook is older than 05/06), `totals_not_reconciled` (01 lists total rows after a source's `data_rows.to`
that neither `excluded_rows` nor a record covers), `partial_ownership_full_amount` (an in-scope record with ownership
`partial` or a share field below 100 whose full valuation/revenue enters `recorded_value` or a revenue metric, without a
share-based field — e.g. `derived_fields` `val.<Y>.share` — or an `assumptions` entry naming it). `total_like_records`
also warns (not blocks) when a record with a real id and a non-total name holds a total word in another field.
Narrative-quality checks (blocking): `period_fields`,
`growth_comparability_basis`, `yield_zero_revenue_scope`, `label_vat_basis`, `label_years`, `label_roi`,
`narrative_metric_naming`, `narrative_scores`, `drastic_without_study`, `plan_completeness`; (warning):
`yield_vat_inclusive`, `partial_unnamed`, `growth_token_caveat`, `promised_outcomes`, `scenario_consistency`, `plan_phasing`,
`top_lists_count`, `recommendation_quality`, `gap_completeness`, `gap_unsupported_claims` (also «أداء منخفض»/«دون
المستهدف» on investment assets without a `sources[]` target or a comparison metric in the evidence), `annualised_text`
(«الإيراد السنوي الكامل»، «على أساس سنوي»، «يعادل سنويًا» in 06 or 09 while the linked revenue field is not documented
`full_year`, unless the sentence says «تقدير» and 06 has `assumptions`), `card_metric_empty` (a card metric empty without a
filter: a mistyped where or an empty scope), `exec_summary_hygiene` (the executive summary is for the board: a
`Sheet!Cell` reference, a cell range or an Excel formula in it, or a numeral doubled by a number word once tokens are
rendered — «1 أصلًا واحدًا»، «2 أصلان» — is flagged; the evidence stays in the evidence chips and the details, and a count
of one or two is written in words or as a token followed by its noun, never both). `narrative_scores` also blocks scores that need no context word («75/100»،
«82 من 100»، «8/10»، «4/5»; a following count noun — عقود، مستأجرين، أصول … — makes it a count); `assessment_scores` warns on
«N من 5/10/100» in an assessment judgment; `exec_summary_elements` warns when «أبرز الفجوات»/«الأولويات» are announced but
fewer than 3 items are listed or named (asset id, evidence token or title words of `top_gaps`/`top_actions`);
`scenario_consistency` warns when an amount has no `start`, `year1_effect` or `run_rate_effect`; `label_vat_basis` ignores
a rejected alternative («… لا الإيراد مع الضريبة»). Every value scan uses the shared
patterns of `assets/pii_patterns.json` (block level blocks, review level — national-ID-like numbers — warns) over 02, 03, 06,
`analysis.json`, the dashboard text and embedded data, `supporting_workbook.xlsx` (or `--xlsx-out`; text cells, and numbers
under a personal-data header; a workbook older than 03/06 is flagged as possibly stale) and `09_message.md`. The
`message_numbers` warning matches every hand-typed number of `09_message.md` (percentages and «ألف/مليون/مليار» amounts
scaled; years, list markers, «30/60/90 يومًا», dates, times, paths and code spans ignored) against 05, facts, the file's
numbers, narrative list counts, 07/08 counts and shares × 100.

## 7. `analysis.json` (build_dashboard.py → embedded in HTML)

```jsonc
{"contract_version": "1.0", "meta": {"entity", "entity_evidence", "data_period", "data_period_evidence", "analysis_date",
          "study_date", "study_date_basis", "study_date_stated", "file_modified", "confidentiality", "currency",
          "file_name", "sha256_short", "valuation", "valuation_dates", "counts", "draft", "built_at", "generator"},
 "vocab": {...}, "field_meta": {...}, "records": [... in_scope + isolated + excluded, PII-free ...], "excluded_rows": [...],
 "config": {...04...}, "metrics": {...05...}, "narrative": {...06 with {{m:..}} and {{f:..}} rendered...},
 "workbook_map_summary": {... anomalies[] keep record_ids ...}, "issues": [...03 issues + narrative extra...],
 "validation": {...07 summary...}, "facts": {key: {"value", "unit", "label"}}}   // rec.* facts only when 06 cites them
```
- `meta.study_date` is null and `study_date_stated` false when 02 `study_date` is only the file's last-modified date (its
  evidence or basis names the file properties, or it equals the workbook's modified date without a cell reference);
  `file_modified` then holds that date for the caption «آخر تعديل للملف المصدر».
- `meta.confidentiality` defaults to «سري — للاستخدام الداخلي» (header badge and every printed page footer). Optional 04
  `"dashboard": {"confidentiality": "<text>" | false}` (or 06 `x_confidentiality`) replaces or removes it.
`build_dashboard.py` exits 2 while 07 has blocking errors (and 1 when 06/07 are missing) unless `--draft`; a draft shows a
«مسودة» banner and fails the browser check `draft` unless `--allow-draft`. Every string is scrubbed: keys matching the
shared personal-data key pattern are dropped and block-level values are masked.

## 8. Dashboard engine API (`assets/metric_engine.js`, exposed as `window.APA`)
- `APA.data` — the embedded `analysis.json`.
- `APA.compute(metricSpec, filters)` → result object (same shape as §5).
- `APA.computeAll(filters)` → `{metrics: {id: result}, breakdowns: {id: ...}, trend: {...}, records: [ids in R for base_where ∧ filters]}`.
- `APA.filterRecords(filters)` → records passing `base_where ∧ filters`.
- `filters` = `{"period": "2025", "asset_type": null|"...", "location": null|"...", "use": null|"..."}`.
The UI renders **only** from `APA.computeAll(currentFilters)`; cards, charts and table share that single result.
