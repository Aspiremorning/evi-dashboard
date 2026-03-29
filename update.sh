#!/bin/bash
# ─────────────────────────────────────────────────────────────────
#  update.sh  —  Build dashboard + commit + push to GitHub Pages
#  Usage:  ./update.sh
#  Or:     ./update.sh "optional commit message"
# ─────────────────────────────────────────────────────────────────
set -e

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

TODAY=$(date +"%d %b %Y")
MSG="${1:-EVI update $TODAY}"

echo ""
echo "════════════════════════════════════════════"
echo "  EVI Dashboard — Daily Update"
echo "  $TODAY"
echo "════════════════════════════════════════════"
echo ""

# 1. Build the dashboard
echo "▶  Building dashboard..."
python3 scripts/build.py
echo ""

# 2. Stage changed files
echo "▶  Staging changes..."
git add data/EVI_2025-26.xlsx docs/index.html

# 3. Commit (skip if nothing changed)
if git diff --cached --quiet; then
  echo "   No changes to commit."
else
  git commit -m "$MSG"
  echo "   Committed: $MSG"
fi

# 4. Push to GitHub
echo ""
echo "▶  Pushing to GitHub..."
git push origin main

echo ""
echo "════════════════════════════════════════════"
echo "  ✅  Done! Dashboard live in ~30 seconds:"
echo "  🌐  https://$(git remote get-url origin | sed 's/.*github.com[:/]//' | sed 's/\.git//' | awk -F'/' '{print $1".github.io/"$2}')"
echo "════════════════════════════════════════════"
echo ""
