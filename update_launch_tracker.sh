#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────
# update_launch_tracker.sh
# Run manually twice daily (13:00 Europe/Skopje, 13:00 US/Pacific)
#
# What it does:
#   1. Downloads the latest Excel from SharePoint (via OneDrive CLI or manual copy)
#   2. Runs extract_launch_data.py to parse Excel → CSV → inject into index.html
#   3. Commits and pushes to GitHub Pages
#
# Usage:
#   bash update_launch_tracker.sh
# ─────────────────────────────────────────────────────────────

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

EXCEL_FILE="data/Kumo Team July-October Launch tracker.xlsx"
PYLIBS="$HOME/.personal-assistant/workspace/tmp/pylibs"

echo "🚀 Kumo Launch Tracker — Update $(date '+%Y-%m-%d %H:%M %Z')"
echo "─────────────────────────────────────────────"

# Step 1: Check Excel file exists
if [ ! -f "$EXCEL_FILE" ]; then
  echo "⚠️  Excel file not found at: $EXCEL_FILE"
  echo "   Please download the latest from SharePoint into data/"
  echo "   URL: https://nvidia-my.sharepoint.com/:x:/r/personal/hraghavan_nvidia_com/_layouts/15/doc2.aspx?sourcedoc=%7BBE0E9C14-7BAD-4F75-8932-DA226AA32C51%7D"
  exit 1
fi

echo "📊 Excel found: $EXCEL_FILE"
echo "   Size: $(du -h "$EXCEL_FILE" | cut -f1)"
echo ""

# Step 2: Extract data
echo "⚙️  Running extract_launch_data.py..."
PYTHONPATH="${PYLIBS}:${PYTHONPATH:-}" python3 extract_launch_data.py
echo ""

# Step 3: Git commit & push
echo "📤 Committing and pushing to GitHub..."
git add -A
TIMESTAMP="$(date '+%Y-%m-%d %H:%M')"
if git diff --cached --quiet; then
  echo "ℹ️  No changes detected — skipping commit"
else
  git commit -m "📊 Data update: ${TIMESTAMP}"
  git push origin main
  echo "✅ Pushed to GitHub Pages!"
fi

echo ""
echo "🎉 Done! Dashboard will update at:"
echo "   https://aleksssokolovski.github.io/kumo-launch-tracker/"
