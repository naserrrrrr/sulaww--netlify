#!/usr/bin/env bash
# Install the skill + subagent at user level (~/.claude) so they work in any project.
set -euo pipefail
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO_CLAUDE="$(cd "$HERE/../.." && pwd)"
mkdir -p "$HOME/.claude/skills" "$HOME/.claude/agents"
rsync -a --delete --exclude 'tests/_build' "$HERE/" "$HOME/.claude/skills/asset-portfolio-analysis/"
cp "$REPO_CLAUDE/agents/asset-portfolio-analyst.md" "$HOME/.claude/agents/"
python3 -m pip install -q -r "$HERE/requirements.txt" || echo "تنبيه: تعذر تثبيت المتطلبات تلقائيًا؛ ثبّتها يدويًا: pip install -r requirements.txt"
echo "تم التثبيت: ~/.claude/skills/asset-portfolio-analysis و ~/.claude/agents/asset-portfolio-analyst.md"
