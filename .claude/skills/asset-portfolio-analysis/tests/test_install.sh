#!/usr/bin/env bash
# Regression test of install.sh in a scratch HOME. The real ~/.claude is never touched.
#
#   bash tests/test_install.sh        # exit 0 = every case passed, 1 = at least one failed
#
# It builds a small FICTIONAL package (asset-portfolio-agent/.claude/{skills,agents}) around the real
# install.sh, in folders whose names contain spaces, and runs it with a minimal PATH:
#   - without rsync (the cp fallback) and, when rsync is available on this machine, with rsync;
#   - fresh install, re-install over a stale copy, junk exclusion (tests/_build*, __pycache__);
#   - install.sh run from the installed copy (must change nothing and exit 0, never delete it);
#   - an install that is a symlink to the package, an agent file that is the same file as the
#     package's, an install that is a symlink to another folder (replaced, never written through);
#   - pip failing (warning with the venv hint, exit 0).
# Python is replaced by a stub that only records the pip call, so the test needs no network.
# Work folder: $APA_BUILD_DIR/install when set (kept for inspection), else a temp dir removed afterwards.
set -uo pipefail

TESTS_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
SD="$(dirname "$TESTS_DIR")"
INSTALL_SH="$SD/install.sh"
BASH_BIN="$(command -v bash)"

if [[ -n "${APA_BUILD_DIR:-}" ]]; then
  WORK="$APA_BUILD_DIR/install"; rm -rf "$WORK"; mkdir -p "$WORK"
else
  WORK="$(mktemp -d)"; trap 'rm -rf "$WORK"' EXIT
fi

FAILS=0; PASSES=0; SKIPS=0
ok() { PASSES=$((PASSES + 1)); printf '  PASS  %s\n' "$1"; }
ko() { FAILS=$((FAILS + 1)); printf '  FAIL  %s\n' "$1"; [[ -n "${2:-}" && -f "$2" ]] && sed -e 's/^/        | /' "$2" | tail -n 12; }
check() { local name="$1"; shift; if "$@"; then ok "$name"; else ko "$name" "${LAST_LOG:-}"; fi; }

# manifest DIR — one "crc size ./path" line per file or symlink, sorted (empty output if DIR is missing)
manifest() {
  [[ -d "$1" ]] || return 0
  (cd "$1" && find . \( -type f -o -type l \) -exec cksum {} + | LC_ALL=C sort)
}

# ---- tool shims: only what install.sh needs, plus a python3 stub that records the pip call
SHIM="$WORK/shim bin"; mkdir -p "$SHIM"
for t in bash env dirname mkdir rm cp find mv; do
  p="$(command -v "$t")" || { echo "missing tool: $t"; exit 1; }
  ln -s "$p" "$SHIM/$t"
done
PIPLOG="$WORK/pip calls.log"
cat >"$SHIM/python3" <<EOF
#!/bin/sh
printf '%s\n' "\$*" >>"$PIPLOG"
exit "\${APA_FAKE_PIP_RC:-0}"
EOF
chmod +x "$SHIM/python3"
RSYNC_DIR=""
if RS="$(command -v rsync 2>/dev/null)"; then RSYNC_DIR="$WORK/rsync bin"; mkdir -p "$RSYNC_DIR"; ln -s "$RS" "$RSYNC_DIR/rsync"; fi

# ---- the fictional package
PKG="$WORK/pkg dir/asset-portfolio-agent"
PSKILL="$PKG/.claude/skills/asset-portfolio-analysis"
PAGENT="$PKG/.claude/agents/asset-portfolio-analyst.md"
mkdir -p "$PSKILL/scripts/__pycache__" "$PSKILL/tests/_build/run" "$PSKILL/tests/_build_extra" \
         "$PSKILL/tests/robustness/__pycache__" "$PSKILL/references" "$(dirname "$PAGENT")"
cp "$INSTALL_SH" "$PSKILL/install.sh"
printf -- '---\nname: asset-portfolio-analysis\ndescription: fictional stub for the install test\n---\n' >"$PSKILL/SKILL.md"
printf 'openpyxl>=3.1\n' >"$PSKILL/requirements.txt"
printf 'print("stub")\n' >"$PSKILL/scripts/tool.py"
printf 'junk' >"$PSKILL/scripts/__pycache__/tool.cpython-312.pyc"
printf 'junk' >"$PSKILL/tests/robustness/__pycache__/t.cpython-312.pyc"
printf 'junk' >"$PSKILL/tests/_build/run/dashboard.html"
printf 'junk' >"$PSKILL/tests/_build_extra/x.log"
printf '_build*/\n' >"$PSKILL/tests/.gitignore"
printf 'echo stub\n' >"$PSKILL/tests/run_tests.sh"
printf '# مرجع وهمي\n' >"$PSKILL/references/notes.md"
printf -- '---\nname: asset-portfolio-analyst\ndescription: fictional stub agent\n---\nstub\n' >"$PAGENT"
EXPECTED="$WORK/expected.manifest"
manifest "$PSKILL" | grep -v -e '/__pycache__/' -e ' \./tests/_build' >"$EXPECTED"
PKG_SKILL_MANIFEST="$WORK/pkg.manifest"; manifest "$PSKILL" >"$PKG_SKILL_MANIFEST"

# run_install LOG HOME PATHDIRS SCRIPT [ENV=VAL...] — run SCRIPT from HOME with a clean environment; sets RC
run_install() {
  local log="$1" home="$2" pathdirs="$3" script="$4"; shift 4
  LAST_LOG="$log"; RC=0
  (cd "$home" && env -i HOME="$home" PATH="$pathdirs" LANG=C.UTF-8 "$@" "$BASH_BIN" "$script") >"$log" 2>&1 || RC=$?
}
same_manifest() { diff -q <(manifest "$1") "$2" >/dev/null; }
no_stage_left() { ! compgen -G "$1/.claude/skills/.asset-portfolio-analysis.*" >/dev/null; }

run_mode() {
  local mode="$1" pathdirs="$2"
  local base="$WORK/$mode"; mkdir -p "$base"
  local H="$base/home with space" rc=0
  mkdir -p "$H"
  local DEST="$H/.claude/skills/asset-portfolio-analysis" AG="$H/.claude/agents/asset-portfolio-analyst.md"
  echo "[$mode]"
  : >"$PIPLOG"

  # A. fresh install from the package
  run_install "$base/a.log" "$H" "$pathdirs" "$PSKILL/install.sh"; rc=$RC
  check "$mode: fresh install exits 0 (got $rc)" test "$rc" = 0
  check "$mode: installed tree = package minus tests/_build* and __pycache__" same_manifest "$DEST" "$EXPECTED"
  check "$mode: agent file installed" cmp -s "$PAGENT" "$AG"
  check "$mode: no staging folder left behind" no_stage_left "$H"
  check "$mode: requirements installed with python3 -m pip from the installed copy" \
    grep -qxF -- "-m pip install -q -r $DEST/requirements.txt" "$PIPLOG"

  # B. re-install over a stale, edited copy
  printf 'stale' >"$DEST/stale.txt"; printf 'edited' >>"$DEST/SKILL.md"; mkdir -p "$DEST/tests/_build"
  run_install "$base/b.log" "$H" "$pathdirs" "$PSKILL/install.sh"; rc=$RC
  check "$mode: re-install exits 0 (got $rc)" test "$rc" = 0
  check "$mode: re-install replaces the stale copy exactly" same_manifest "$DEST" "$EXPECTED"
  check "$mode: no staging folder left after re-install" no_stage_left "$H"

  # C. install.sh run from the installed copy (relative path from HOME, as an old README said)
  local before="$base/c.before"; manifest "$DEST" >"$before"; cp "$AG" "$base/agent.before"
  run_install "$base/c.log" "$H" "$pathdirs" ".claude/skills/asset-portfolio-analysis/install.sh"; rc=$RC
  check "$mode: run from the installed copy exits 0 (got $rc)" test "$rc" = 0
  check "$mode: installed copy is left intact" same_manifest "$DEST" "$before"
  check "$mode: agent file is left intact" cmp -s "$AG" "$base/agent.before"
  check "$mode: says it is already installed and to run from the package folder" \
    grep -q 'asset-portfolio-agent' "$base/c.log"
  # ... and by absolute path
  run_install "$base/c2.log" "$H" "$pathdirs" "$DEST/install.sh"; rc=$RC
  check "$mode: absolute path to the installed copy exits 0 and changes nothing" \
    eval 'test "$rc" = 0 && same_manifest "$DEST" "$before"'

  # D. the install is a symlink to the package itself: nothing may be copied or deleted
  local H2="$base/home linked"; mkdir -p "$H2/.claude/skills"
  ln -s "$PSKILL" "$H2/.claude/skills/asset-portfolio-analysis"
  run_install "$base/d.log" "$H2" "$pathdirs" "$PSKILL/install.sh"; rc=$RC
  check "$mode: install symlinked to the package exits 0 (got $rc)" test "$rc" = 0
  check "$mode: package folder untouched" same_manifest "$PSKILL" "$PKG_SKILL_MANIFEST"
  check "$mode: symlink kept" test -L "$H2/.claude/skills/asset-portfolio-analysis"

  # E. the agent destination is already the package's agent file (symlink): copy skipped, no error
  local H3="$base/home agent link"; mkdir -p "$H3/.claude/agents"
  ln -s "$PAGENT" "$H3/.claude/agents/asset-portfolio-analyst.md"
  run_install "$base/e.log" "$H3" "$pathdirs" "$PSKILL/install.sh"; rc=$RC
  check "$mode: agent already the same file exits 0 (got $rc)" test "$rc" = 0
  check "$mode: skill installed although the agent copy was skipped" \
    same_manifest "$H3/.claude/skills/asset-portfolio-analysis" "$EXPECTED"
  check "$mode: agent symlink kept and package agent unchanged" \
    eval 'test -L "$H3/.claude/agents/asset-portfolio-analyst.md" && grep -qx stub "$PAGENT"'

  # F. the install is a symlink to ANOTHER folder: it is replaced, never written through
  local H4="$base/home other link" OTHER="$base/other checkout/asset-portfolio-analysis"
  mkdir -p "$H4/.claude/skills" "$OTHER"; printf 'keep me' >"$OTHER/marker.txt"
  ln -s "$OTHER" "$H4/.claude/skills/asset-portfolio-analysis"
  run_install "$base/f.log" "$H4" "$pathdirs" "$PSKILL/install.sh"; rc=$RC
  check "$mode: install over a symlink to another folder exits 0 (got $rc)" test "$rc" = 0
  check "$mode: the other folder is untouched" eval 'test "$(ls -A "$OTHER")" = marker.txt'
  check "$mode: the symlink is replaced by a real copy" \
    eval 'test ! -L "$H4/.claude/skills/asset-portfolio-analysis" && same_manifest "$H4/.claude/skills/asset-portfolio-analysis" "$EXPECTED"'

  # G. pip fails (e.g. externally-managed-environment): warning with the venv hint, still exit 0
  run_install "$base/g.log" "$H" "$pathdirs" "$PSKILL/install.sh" APA_FAKE_PIP_RC=1; rc=$RC
  check "$mode: pip failure still exits 0 (got $rc)" test "$rc" = 0
  check "$mode: pip failure prints the manual command and the venv hint" \
    eval 'grep -q "python3 -m pip install -r" "$base/g.log" && grep -q "python3 -m venv" "$base/g.log"'
}

run_mode "no-rsync" "$SHIM"
if [[ -n "$RSYNC_DIR" ]]; then
  run_mode "rsync" "$SHIM:$RSYNC_DIR"
else
  SKIPS=$((SKIPS + 1)); printf '  SKIP  rsync cases (rsync is not installed on this machine)\n'
fi

printf 'test_install: passed %d · failed %d · skipped %d\n' "$PASSES" "$FAILS" "$SKIPS"
[[ "$FAILS" == 0 ]]
