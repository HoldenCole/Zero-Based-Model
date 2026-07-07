#!/usr/bin/env bash
# Rebuild the committed desk workbook from the current data/ files.
# Run this after editing anything under data/, then commit dist/US_Naphtha_Model.xlsx.
#
#   ./build_workbook.sh            # simple 3-sheet model (default)
#   ./build_workbook.sh --full     # extended workbook (balances, dashboard, checks)
set -euo pipefail
cd "$(dirname "$0")"

OUT="dist/US_Naphtha_Model.xlsx"
mkdir -p dist

python3 -m naphtha_model export --out "$OUT" "$@"
echo "Built $OUT"
