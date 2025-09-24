#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"
echo "== Pull from cloud =="
clasp pull
echo "== Quick sanity =="
ls -1 | wc -l
OUT_DIR="_dist"
mkdir -p "$OUT_DIR"
STAMP=$(date +"%Y%m%d_%H%M%S")
ZIP="${OUT_DIR}/appsScript_${STAMP}.zip"
echo "== Zipping =="
zip -rq "$ZIP" . -x "*.git*" -x "*node_modules*" -x "*_dist*"
echo "Done: $ZIP"
