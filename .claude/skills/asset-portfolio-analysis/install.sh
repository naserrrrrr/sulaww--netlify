#!/usr/bin/env bash
# Install the skill + subagent at user level (~/.claude) so they work in any project.
# Run it from the package folder:  cd asset-portfolio-agent && bash .claude/skills/asset-portfolio-analysis/install.sh
# Run from the installed copy itself it changes nothing: it says so and exits 0.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"   # as reached (finds the agent file next to it)
HERE_P="$(cd "$HERE" && pwd -P)"                       # physical path (symlinks resolved)
AGENT_SRC="$HERE/../../agents/asset-portfolio-analyst.md"
SKILLS="$HOME/.claude/skills"
AGENTS="$HOME/.claude/agents"
DEST="$SKILLS/asset-portfolio-analysis"
AGENT_DEST="$AGENTS/asset-portfolio-analyst.md"
mkdir -p "$SKILLS" "$AGENTS"

# Same folder (run from the installed copy, or the install is a symlink to this folder): copying it
# onto itself would delete it, so stop here.
DEST_P="$(cd "$SKILLS" && pwd -P)/asset-portfolio-analysis"
[ -d "$DEST" ] && DEST_P="$(cd "$DEST" && pwd -P)"
if [ "$HERE_P" = "$DEST_P" ]; then
  echo "المهارة مثبتة في هذا المجلد أصلًا ($DEST)، فلم يُنسخ شيء ولم يتغير شيء."
  echo "لإعادة التثبيت أو التحديث شغّل install.sh من مجلد الحزمة لا من النسخة المثبتة:"
  echo "  cd asset-portfolio-agent && bash .claude/skills/asset-portfolio-analysis/install.sh"
  exit 0
fi

# Build the new copy in a staging folder next to DEST, then swap it in with mv, so a failed copy
# never leaves a half-installed skill. An existing DEST (even a symlink) is moved aside, never written through.
STAGE="$SKILLS/.asset-portfolio-analysis.install.$$"
rm -rf "$STAGE"; mkdir "$STAGE"
trap 'rm -rf "$STAGE"' EXIT
if command -v rsync >/dev/null 2>&1; then
  rsync -a --exclude 'tests/_build*' --exclude '__pycache__' "$HERE_P/" "$STAGE/new/"
else  # no rsync (e.g. Git Bash on Windows): plain copy, then drop the build folders and caches
  cp -R "$HERE_P" "$STAGE/new"
  rm -rf "$STAGE/new"/tests/_build*
  find "$STAGE/new" -name __pycache__ -type d -prune -exec rm -rf {} +
fi
[ -f "$STAGE/new/SKILL.md" ] || { echo "تعذر نسخ المهارة (SKILL.md غير موجود في النسخة)؛ لم يتغير التثبيت السابق." >&2; exit 1; }
if [ -e "$DEST" ] || [ -L "$DEST" ]; then mv "$DEST" "$STAGE/old"; fi
if ! mv "$STAGE/new" "$DEST"; then
  if [ -e "$STAGE/old" ] || [ -L "$STAGE/old" ]; then mv "$STAGE/old" "$DEST"; fi
  echo "تعذر وضع النسخة الجديدة في $DEST؛ أُعيد التثبيت السابق كما كان." >&2
  exit 1
fi

if [ ! -f "$AGENT_SRC" ]; then
  echo "تنبيه: ملف الوكيل غير موجود بجوار المهارة ($AGENT_SRC)؛ ثُبّتت المهارة وحدها."
elif [ "$AGENT_SRC" -ef "$AGENT_DEST" ]; then
  echo "ملف الوكيل في $AGENT_DEST هو الملف نفسه في الحزمة؛ لم يُنسخ."
else
  cp "$AGENT_SRC" "$AGENT_DEST.tmp.$$" && mv -f "$AGENT_DEST.tmp.$$" "$AGENT_DEST"
fi

if ! python3 -m pip install -q -r "$DEST/requirements.txt"; then
  echo "تنبيه: تعذر تثبيت متطلبات Python تلقائيًا. ثبّتها يدويًا:"
  echo "  python3 -m pip install -r \"$DEST/requirements.txt\""
  echo "وإن رفض pip ذلك (externally-managed-environment) فأنشئ بيئة افتراضية وفعّلها قبل فتح Claude Code، ثم أعد الأمر:"
  echo "  python3 -m venv .venv && . .venv/bin/activate"
fi
echo "تم التثبيت: $DEST و $AGENT_DEST"
