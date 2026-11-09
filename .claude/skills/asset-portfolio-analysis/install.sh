#!/usr/bin/env bash
# Install the skill + subagent at user level (~/.claude) so they work in any project.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_CLAUDE="$(cd "$HERE/../.." && pwd)"
mkdir -p "$HOME/.claude/skills" "$HOME/.claude/agents"
DEST="$HOME/.claude/skills/asset-portfolio-analysis"
if command -v rsync >/dev/null 2>&1; then
  rsync -a --delete --exclude 'tests/_build*' --exclude '__pycache__' "$HERE/" "$DEST/"
else  # no rsync (e.g. Git Bash on Windows): replace the folder with a plain copy
  rm -rf "$DEST" && mkdir -p "$DEST" && cp -R "$HERE/." "$DEST/"
  rm -rf "$DEST"/tests/_build* && find "$DEST" -name __pycache__ -type d -prune -exec rm -rf {} +
fi
cp "$REPO_CLAUDE/agents/asset-portfolio-analyst.md" "$HOME/.claude/agents/"
python3 -m pip install -q -r "$HERE/requirements.txt" || echo "تنبيه: تعذر تثبيت المتطلبات تلقائيًا؛ ثبّتها يدويًا: pip install -r requirements.txt"
echo "تم التثبيت: ~/.claude/skills/asset-portfolio-analysis و ~/.claude/agents/asset-portfolio-analyst.md"
