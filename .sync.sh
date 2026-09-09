#!/bin/bash
REPO="$HOME/akij-commodities-control-tower"
SRC_DIR="/Users/apple/Documents/Default Project"
SRC_HTML="$SRC_DIR/AKIJ-Commodities-Control-Tower.html"
SRC_DATA="$SRC_DIR/ctr-data.js"

cd "$REPO" || exit 1

git pull --rebase origin main >/dev/null 2>&1 || { git rebase --abort >/dev/null 2>&1 || true; }

changed=0
if [ -f "$SRC_HTML" ] && { [ ! -f "$REPO/index.html" ] || [ "$SRC_HTML" -nt "$REPO/index.html" ]; }; then
  cp "$SRC_HTML" "$REPO/index.html"
  changed=1
fi
if [ -f "$SRC_DATA" ] && { [ ! -f "$REPO/ctr-data.js" ] || [ "$SRC_DATA" -nt "$REPO/ctr-data.js" ]; }; then
  cp "$SRC_DATA" "$REPO/ctr-data.js"
  changed=1
fi

git add -A >/dev/null 2>&1
if ! git diff --cached --quiet; then
  git -c user.name="Mulham Hasan" -c user.email="mulhamhasan@akijconsulting.com" \
      commit -m "Auto-sync: source file updated $(date '+%Y-%m-%d %H:%M:%S')" >/dev/null 2>&1
  git push origin main >/dev/null 2>&1
fi
