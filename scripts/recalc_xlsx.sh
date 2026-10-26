#!/usr/bin/env sh
# Recalcula un .xlsx con Microsoft Excel en macOS (AppleScript) y lo guarda en el sitio.
# Uso: ./scripts/recalc_xlsx.sh <archivo.xlsx>
set -e
FILE="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
osascript -e 'tell application "Microsoft Excel"' \
  -e "open POSIX file \"$FILE\"" \
  -e 'calculate full' \
  -e 'save active workbook' \
  -e 'close active workbook saving no' \
  -e 'end tell'
echo "Recalculado: $FILE"
