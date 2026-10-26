#!/usr/bin/env sh
# Renderiza docx/pptx/xlsx a PDF con Microsoft Office en macOS (AppleScript).
# Uso: ./scripts/render_preview.sh <archivo> [salida.pdf]
set -e
FILE="$(cd "$(dirname "$1")" && pwd)/$(basename "$1")"
# Por defecto la salida va a la carpeta temporal, nunca junto al archivo de origen.
TMPD="${TMPDIR:-/tmp}/document-workflow"; mkdir -p "$TMPD"
BASE="$(basename "$FILE")"; OUT="${2:-$TMPD/${BASE%.*}.pdf}"
case "$FILE" in
  *.docx|*.dotx) APP="Microsoft Word"
    osascript -e "tell application \"$APP\"" -e "open POSIX file \"$FILE\"" \
      -e "save as active document file name \"$OUT\" file format format PDF" \
      -e "close active document saving no" -e "end tell" ;;
  *.pptx|*.potx) APP="Microsoft PowerPoint"
    osascript -e "tell application \"$APP\"" -e "open POSIX file \"$FILE\"" \
      -e "save active presentation in POSIX file \"$OUT\" as save as PDF" \
      -e "close active presentation saving no" -e "end tell" ;;
  *.xlsx) APP="Microsoft Excel"
    osascript -e "tell application \"$APP\"" -e "open POSIX file \"$FILE\"" \
      -e "save active workbook in POSIX file \"$OUT\" as PDF file format" \
      -e "close active workbook saving no" -e "end tell" ;;
  *) echo "Formato no soportado" >&2; exit 2 ;;
esac
echo "$OUT"
