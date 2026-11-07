#!/usr/bin/env bash
# End-to-end regression test of the asset-portfolio-analysis skill on a FICTIONAL workbook.
#
#   bash tests/run_tests.sh            # full run (needs python3 + openpyxl + jsonschema; node + playwright for the browser check)
#   APA_SKIP_BROWSER=1 bash tests/run_tests.sh   # skip the browser check (reported as SKIP, not PASS)
#
# Everything is written under tests/_build/ (git-ignored). The real sample workbook is never used here.
# Exit code: 0 = every step passed (skips allowed only when requested), 1 = at least one failure.
set -euo pipefail

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SD="$(dirname "$TESTS_DIR")"
S="$SD/scripts"
# APA_BUILD_DIR lets parallel runs (or CI) use a private build folder instead of tests/_build.
BUILD="${APA_BUILD_DIR:-$TESTS_DIR/_build}"
RUN="$BUILD/run"
XLSX="$BUILD/synthetic_portfolio.xlsx"
LOGS="$BUILD/logs"
NEG_A="$BUILD/neg_short_summary"
NEG_B="$BUILD/neg_text_as_money"
PY=(python3 -I)

rm -rf "$BUILD"
mkdir -p "$RUN" "$LOGS"

PASSES=()
FAILS=()
SKIPS=()
N=0

say() { printf '%s\n' "$*"; }
pass() { PASSES+=("$1"); say "  PASS  $1"; }
fail() {
  FAILS+=("$1"); say "  FAIL  $1"
  if [[ -n "${2:-}" && -f "$2" ]]; then sed -e 's/^/        | /' "$2" | tail -n 15; fi
}
skip() { SKIPS+=("$1"); say "  SKIP  $1"; }

# run_step NAME EXPECTED_EXIT CMD... — runs CMD, logs output, compares the exit code
run_step() {
  local name="$1" want="$2"; shift 2
  N=$((N + 1))
  local log; log="$LOGS/$(printf '%02d' "$N")_$(printf '%s' "$name" | tr -c 'A-Za-z0-9_.-' '_').log"
  local rc=0
  "$@" >"$log" 2>&1 || rc=$?
  if [[ "$rc" == "$want" ]]; then
    pass "$name (exit $rc)"
    return 0
  fi
  fail "$name — exit $rc, expected $want (log: ${log#"$SD"/})" "$log"
  return 1
}

# have_script FILE STEP — a missing script is a clear failure, never a silent pass
have_script() {
  if [[ -f "$1" ]]; then return 0; fi
  fail "$2 — السكربت غير موجود: ${1#"$SD"/}"
  return 1
}

# have_files STEP FILE... — skip a step whose inputs were not produced (the earlier failure is already reported)
have_files() {
  local step="$1"; shift
  local f
  for f in "$@"; do
    if [[ ! -f "$f" ]]; then
      fail "$step — لم يُنفذ لغياب المدخل ${f#"$BUILD"/} (فشل في خطوة سابقة)"
      return 1
    fi
  done
  return 0
}

say "== اختبار المهارة على مصنف تجريبي وهمي =="
say "   المجلد: ${BUILD}"

# --- 0. prerequisites -----------------------------------------------------------------------
say "[0] المتطلبات"
if "${PY[@]}" -c "import openpyxl, jsonschema" >/dev/null 2>&1; then pass "python3 + openpyxl + jsonschema"
else fail "python3 + openpyxl + jsonschema غير متاحة (pip install -r requirements.txt)"; fi
"${PY[@]}" -c "import hijridate" >/dev/null 2>&1 && pass "hijridate" || say "  تنبيه: hijridate غير مثبتة؛ قد تختلف نتائج تحويل التواريخ الهجرية"

# --- 1. fixture ----------------------------------------------------------------------------------
say "[1] المصنف التجريبي والنتائج المتوقعة"
if have_script "$TESTS_DIR/make_synthetic_fixture.py" "generate fixture"; then
  run_step "generate fixture" 0 "${PY[@]}" "$TESTS_DIR/make_synthetic_fixture.py" --out "$XLSX" || true
  run_step "expected.json matches generator data" 0 "${PY[@]}" "$TESTS_DIR/make_synthetic_fixture.py" \
    --check-expected "$TESTS_DIR/synthetic/expected.json" || true
fi

# --- 2. main chain --------------------------------------------------------------------------------
say "[2] السلسلة الكاملة"
if have_script "$S/inspect_workbook.py" "inspect_workbook" && have_files "inspect_workbook" "$XLSX"; then
  run_step "inspect_workbook" 0 "${PY[@]}" "$S/inspect_workbook.py" "$XLSX" --out "$RUN" || true
fi
cp "$TESTS_DIR/synthetic/02_mapping.json" "$RUN/02_mapping.json"
if have_script "$S/normalize.py" "normalize" && have_files "normalize" "$XLSX"; then
  run_step "normalize" 0 "${PY[@]}" "$S/normalize.py" --xlsx "$XLSX" --mapping "$RUN/02_mapping.json" \
    --out "$RUN/03_records.json" || true
fi
if have_script "$S/compute_metrics.py" "compute_metrics --init-config" && have_files "init-config" "$RUN/03_records.json"; then
  run_step "compute_metrics --init-config" 0 "${PY[@]}" "$S/compute_metrics.py" --init-config \
    --records "$RUN/03_records.json" --out "$RUN/04_metrics_config.json" || true
fi
if have_script "$S/compute_metrics.py" "compute_metrics" && have_files "compute_metrics" "$RUN/04_metrics_config.json"; then
  run_step "compute_metrics" 0 "${PY[@]}" "$S/compute_metrics.py" --records "$RUN/03_records.json" \
    --config "$RUN/04_metrics_config.json" --out "$RUN/05_metrics.json" --print || true
  run_step "compute_metrics --filters use=investment --print" 0 "${PY[@]}" "$S/compute_metrics.py" \
    --records "$RUN/03_records.json" --config "$RUN/04_metrics_config.json" --filters '{"use":"investment"}' --print || true
fi
cp "$TESTS_DIR/synthetic/06_narrative.json" "$RUN/06_narrative.json"
if have_script "$S/validate_analysis.py" "validate_analysis" && have_files "validate_analysis" "$RUN/05_metrics.json"; then
  run_step "validate_analysis (before build)" 0 "${PY[@]}" "$S/validate_analysis.py" --run-dir "$RUN" || true
fi
if have_script "$S/build_dashboard.py" "build_dashboard" && have_files "build_dashboard" "$RUN/07_validation.json"; then
  run_step "build_dashboard" 0 "${PY[@]}" "$S/build_dashboard.py" --run-dir "$RUN" || true
fi
if have_script "$S/validate_analysis.py" "validate_analysis --html --xlsx" && have_files "validate --html" "$RUN/dashboard.html"; then
  run_step "validate_analysis --html --xlsx (after build)" 0 "${PY[@]}" "$S/validate_analysis.py" --run-dir "$RUN" \
    --xlsx "$XLSX" --html "$RUN/dashboard.html" || true
fi
if [[ "${APA_SKIP_BROWSER:-0}" == "1" ]]; then
  skip "check_dashboard (APA_SKIP_BROWSER=1)"
elif ! command -v node >/dev/null 2>&1; then
  fail "check_dashboard — node غير مثبت (ثبّت Node 18+ وplaywright، أو شغّل بـ APA_SKIP_BROWSER=1)"
elif have_script "$S/check_dashboard.mjs" "check_dashboard" && have_files "check_dashboard" "$RUN/dashboard.html"; then
  run_step "check_dashboard (browser)" 0 node "$S/check_dashboard.mjs" --html "$RUN/dashboard.html" --out "$RUN" || true
fi
if have_script "$S/export_workbook.py" "export_workbook" && have_files "export_workbook" "$RUN/05_metrics.json"; then
  run_step "export_workbook" 0 "${PY[@]}" "$S/export_workbook.py" --run-dir "$RUN" || true
fi
# re-validate once the supporting workbook exists: pii_export_patterns scans every cell of it (C24)
if have_script "$S/validate_analysis.py" "validate_analysis (after export)" && have_files "validate after export" \
    "$RUN/supporting_workbook.xlsx" "$RUN/dashboard.html"; then
  run_step "validate_analysis (after export: supporting workbook scanned)" 0 "${PY[@]}" "$S/validate_analysis.py" \
    --run-dir "$RUN" --xlsx "$XLSX" --html "$RUN/dashboard.html" || true
  run_step "validate after export: pii_export_patterns ran and passed" 0 "${PY[@]}" -c '
import json, sys
v = json.load(open(sys.argv[1], encoding="utf-8"))
c = [x for x in v["checks"] if x["id"] == "pii_export_patterns"]
print(c)
sys.exit(0 if c and all(x["passed"] and not x.get("skipped") for x in c) else 1)' "$RUN/07_validation.json" || true
fi
if have_script "$S/compare_versions.py" "compare_versions" && have_files "compare_versions" "$RUN/03_records.json"; then
  run_step "compare_versions (run vs itself)" 0 "${PY[@]}" "$S/compare_versions.py" --prev "$RUN" --curr "$RUN" \
    --out "$RUN/version_comparison.json" || true
fi

# --- 3. assertions against hand-computed expectations ------------------------------------------
say "[3] المطابقة مع النتائج المتوقعة المحسوبة يدويًا"
ASSERT_ARGS=(--run-dir "$RUN" --expected "$TESTS_DIR/synthetic/expected.json" --xlsx "$XLSX")
[[ "${APA_SKIP_BROWSER:-0}" == "1" ]] && ASSERT_ARGS+=(--skip-dashboard)
if have_script "$TESTS_DIR/assert_expected.py" "assert_expected"; then
  run_step "assert_expected (01/03/04/05/07/08/versions/privacy)" 0 "${PY[@]}" "$TESTS_DIR/assert_expected.py" \
    "${ASSERT_ARGS[@]}" || true
fi

# --- 4. negative tests -------------------------------------------------------------------------------
say "[4] اختبارات سلبية"
# (a) a 60-word executive summary must be a blocking validation error (exit 2)
if have_files "neg (a) short summary" "$RUN/05_metrics.json" "$RUN/01_workbook_map.json"; then
  mkdir -p "$NEG_A"
  for f in 01_workbook_map.json 02_mapping.json 03_records.json 04_metrics_config.json 05_metrics.json; do
    cp "$RUN/$f" "$NEG_A/$f"
  done
  "${PY[@]}" - "$RUN/06_narrative.json" "$NEG_A/06_narrative.json" <<'PYEOF'
import json, sys
n = json.load(open(sys.argv[1], encoding="utf-8"))
n["executive_summary"] = " ".join(n["executive_summary"].split()[:60])
json.dump(n, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False, indent=2)
PYEOF
  if have_script "$S/validate_analysis.py" "neg (a)"; then
    run_step "neg (a): 60-word summary -> validate exits 2" 2 "${PY[@]}" "$S/validate_analysis.py" --run-dir "$NEG_A" || true
    run_step "neg (a): exec_summary_length is the blocking finding" 0 "${PY[@]}" -c '
import json, sys
v = json.load(open(sys.argv[1], encoding="utf-8"))
bad = [c["id"] for c in v["checks"] if not c["passed"] and c["level"] == "blocking"]
print("blocking:", bad)
sys.exit(0 if "exec_summary_length" in bad else 1)' "$NEG_A/07_validation.json" || true
  fi
  # (c) build without --draft must refuse while validation has blocking errors (exit 2); --draft still builds
  if have_script "$S/build_dashboard.py" "neg (c)" && have_files "neg (c)" "$NEG_A/07_validation.json"; then
    run_step "neg (c): build without --draft on blocking validation -> exit 2" 2 "${PY[@]}" "$S/build_dashboard.py" \
      --run-dir "$NEG_A" || true
    if [[ -f "$NEG_A/dashboard.html" ]]; then fail "neg (c): dashboard.html written despite blocking errors"
    else pass "neg (c): no dashboard.html written"; fi
    run_step "neg (c): build --draft still produces a draft" 0 "${PY[@]}" "$S/build_dashboard.py" --run-dir "$NEG_A" --draft || true
    # D18: the browser check fails a draft unless --allow-draft is passed
    if [[ "${APA_SKIP_BROWSER:-0}" != "1" ]] && command -v node >/dev/null 2>&1 && [[ -f "$NEG_A/dashboard.html" ]]; then
      run_step "neg (c): check_dashboard on a draft -> exit 1" 1 node "$S/check_dashboard.mjs" --html "$NEG_A/dashboard.html" \
        --out "$NEG_A/check_draft" || true
      run_step "neg (c): check_dashboard --allow-draft -> exit 0" 0 node "$S/check_dashboard.mjs" --html "$NEG_A/dashboard.html" \
        --out "$NEG_A/check_draft_allowed" --allow-draft || true
    fi
  fi
fi
# (d) growth marked «confirmed» while field_meta does not prove equal VAT/basis/completeness -> blocking (B1)
# (e) a numeric score («82 من 100») in the summary -> blocking (B6)
if have_files "neg (d)/(e)" "$RUN/05_metrics.json" "$RUN/01_workbook_map.json" && have_script "$S/validate_analysis.py" "neg (d)/(e)"; then
  for case in d e; do
    NEG_X="$BUILD/neg_$case"
    mkdir -p "$NEG_X"
    for f in 01_workbook_map.json 02_mapping.json 03_records.json 04_metrics_config.json 05_metrics.json 06_narrative.json; do
      cp "$RUN/$f" "$NEG_X/$f"
    done
  done
  "${PY[@]}" - "$BUILD/neg_d/04_metrics_config.json" <<'PYEOF'
import json, sys
c = json.load(open(sys.argv[1], encoding="utf-8"))
p = next(p for p in c["periods"] if p.get("prev_rev_field"))
p["comparable_with_prev"] = {"status": "confirmed", "reason": "تأكيد يدوي دون دليل"}
json.dump(c, open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False, indent=2)
PYEOF
  "${PY[@]}" - "$BUILD/neg_e/06_narrative.json" <<'PYEOF'
import json, sys
n = json.load(open(sys.argv[1], encoding="utf-8"))
n["executive_summary"] += " وتحصل المحفظة على تقييم عام 82 من 100."
json.dump(n, open(sys.argv[1], "w", encoding="utf-8"), ensure_ascii=False, indent=2)
PYEOF
  for pair in "d:growth_comparability_basis:growth marked confirmed without evidence" "e:narrative_scores:numeric score in the summary"; do
    case="${pair%%:*}"; rest="${pair#*:}"; cid="${rest%%:*}"; what="${rest#*:}"
    run_step "neg ($case): $what -> validate exits 2" 2 "${PY[@]}" "$S/validate_analysis.py" --run-dir "$BUILD/neg_$case" || true
    run_step "neg ($case): $cid is a blocking finding" 0 "${PY[@]}" -c '
import json, sys
v = json.load(open(sys.argv[1], encoding="utf-8"))
bad = [c["id"] for c in v["checks"] if not c["passed"] and c["level"] == "blocking"]
print("blocking:", bad)
sys.exit(0 if sys.argv[2] in bad else 1)' "$BUILD/neg_$case/07_validation.json" "$cid" || true
  done
fi
# (b) mapping «صافي الأصول» (text labels) as money -> text_in_numeric issue and no number
if have_script "$S/normalize.py" "neg (b)" && have_files "neg (b)" "$XLSX"; then
  mkdir -p "$NEG_B"
  "${PY[@]}" - "$TESTS_DIR/synthetic/02_mapping.json" "$NEG_B/02_mapping.json" <<'PYEOF'
import json, sys
m = json.load(open(sys.argv[1], encoding="utf-8"))
reg = next(s for s in m["sources"] if s["id"] == "REG")
reg["columns"]["H"] = {"field": "net_assets.amount", "kind": "money"}
m["field_meta"]["net_assets.amount"] = {"label": "صافي الأصول (ربط خاطئ كمبلغ)", "group": "other", "unit": "SAR"}
json.dump(m, open(sys.argv[2], "w", encoding="utf-8"), ensure_ascii=False, indent=2)
PYEOF
  run_step "neg (b): normalize with «صافي الأصول» mapped as money" 0 "${PY[@]}" "$S/normalize.py" --xlsx "$XLSX" \
    --mapping "$NEG_B/02_mapping.json" --out "$NEG_B/03_records.json" || true
  if [[ -f "$NEG_B/03_records.json" ]]; then
    run_step "neg (b): text_in_numeric reported and no number stored" 0 "${PY[@]}" -c '
import json, sys
d = json.load(open(sys.argv[1], encoding="utf-8"))
f = "net_assets.amount"
iss = [i for i in d["issues"] if i["type"] == "text_in_numeric" and i.get("field") == f]
tvs = [r["fields"][f] for r in d["records"] if f in r.get("fields", {})]
nums = [t for t in tvs if t.get("v") is not None or t.get("s") != "text"]
print("issues:", len(iss), "| values:", len(tvs), "| non-text values:", len(nums))
sys.exit(0 if iss and tvs and not nums else 1)' "$NEG_B/03_records.json" || true
  fi
fi

# --- 4b. versions + follow-up in the dashboard (U1) and the delivery message checks (U4) -------------------------
say "[4b] مقارنة الإصدارات ومتابعة التوصيات ورسالة التسليم"
if [[ -f "$TESTS_DIR/test_versions_message.py" ]] && have_files "versions/message" "$RUN/05_metrics.json" "$RUN/06_narrative.json"; then
  VM_ARGS=(--run "$RUN" --work "$BUILD/versions_message")
  [[ "${APA_SKIP_BROWSER:-0}" == "1" ]] && VM_ARGS+=(--skip-browser)
  run_step "tests/test_versions_message.py (versions section, follow-up table, 09 message)" 0 "${PY[@]}" \
    "$TESTS_DIR/test_versions_message.py" "${VM_ARGS[@]}" || true
fi

# --- 5. unit tests ---------------------------------------------------------------------------------------
say "[5] اختبارات الوحدة والانحدار"
if [[ -f "$TESTS_DIR/test_metric_engine.py" ]]; then
  run_step "tests/test_metric_engine.py" 0 "${PY[@]}" "$TESTS_DIR/test_metric_engine.py" || true
else skip "tests/test_metric_engine.py (غير موجود)"; fi
if [[ -f "$TESTS_DIR/test_metric_engine_js.mjs" ]]; then
  if command -v node >/dev/null 2>&1; then
    run_step "tests/test_metric_engine_js.mjs" 0 node "$TESTS_DIR/test_metric_engine_js.mjs" || true
  else fail "tests/test_metric_engine_js.mjs — node غير مثبت"; fi
else skip "tests/test_metric_engine_js.mjs (غير موجود)"; fi
if [[ -f "$TESTS_DIR/test_normalize.py" ]]; then
  run_step "tests/test_normalize.py" 0 "${PY[@]}" "$TESTS_DIR/test_normalize.py" || true
else skip "tests/test_normalize.py (غير موجود)"; fi
if [[ -f "$TESTS_DIR/test_review_round3.py" ]]; then
  run_step "tests/test_review_round3.py (isolated totals, export growth, issue wording, facts)" 0 "${PY[@]}" \
    "$TESTS_DIR/test_review_round3.py" || true
else skip "tests/test_review_round3.py (غير موجود)"; fi
if [[ -f "$TESTS_DIR/robustness/test_inspect_robustness.py" ]]; then
  # honours APA_BUILD_DIR (writes to $APA_BUILD_DIR/robustness), else a temp dir removed afterwards
  run_step "tests/robustness/test_inspect_robustness.py" 0 env APA_BUILD_DIR="$BUILD" "${PY[@]}" \
    "$TESTS_DIR/robustness/test_inspect_robustness.py" || true
else skip "tests/robustness/test_inspect_robustness.py (غير موجود)"; fi
for st in compare_versions export_workbook build_dashboard; do
  if have_script "$S/$st.py" "$st --self-test"; then
    run_step "$st --self-test" 0 "${PY[@]}" "$S/$st.py" --self-test || true
  fi
done
if command -v node >/dev/null 2>&1; then
  if have_script "$S/check_dashboard.mjs" "check_dashboard --self-test"; then
    run_step "check_dashboard --self-test" 0 node "$S/check_dashboard.mjs" --self-test || true
  fi
else fail "check_dashboard --self-test — node غير مثبت"; fi

# --- summary -------------------------------------------------------------------------------------------------
say ""
say "== الملخص =="
say "  ناجح: ${#PASSES[@]} · فاشل: ${#FAILS[@]} · متخطى: ${#SKIPS[@]}"
if ((${#FAILS[@]})); then
  for f in "${FAILS[@]}"; do say "  ✗ $f"; done
  say "RESULT: FAIL"
  exit 1
fi
if ((${#SKIPS[@]})); then for s in "${SKIPS[@]}"; do say "  – $s"; done; fi
say "RESULT: PASS"
exit 0
