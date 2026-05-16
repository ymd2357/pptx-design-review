#!/usr/bin/env bash
# PPTX -> PDF -> PNG (1 file per slide) using Microsoft PowerPoint Mac + sips.
# Based on vscode-pptx-viewer/tools/capture-pp-mac.sh.
#
# Usage: scripts/capture-pp-mac.sh <pptx-path> <out-prefix>
#   <out-prefix> = path/to/<name>
#   PNGs are written as <out-prefix>-slide-01.png, ...
#
# IMPORTANT: <out-prefix>'s parent directory must be a folder PowerPoint Mac
# has already been granted access to (via a previously approved "ファイル
# アクセスを許可" dialog). Creating a brand-new sibling subfolder will
# trigger a fresh sandbox prompt and PowerPoint's response to a cancelled
# prompt is buggy (it stalls and returns -1712 / -9074). So reuse a known-
# good directory and rely on filename prefixes to namespace runs.
set -e
PPTX="${1:?pptx required}"
OUT_PREFIX="${2:?out prefix required}"
OUT_DIR="$(dirname "$OUT_PREFIX")"
OUT_BASE="$(basename "$OUT_PREFIX")"
mkdir -p "$OUT_DIR"
PDF="$OUT_DIR/.${OUT_BASE}.capture.pdf"
rm -f "$PDF"

osascript - <<OSAEOF
tell application "Microsoft PowerPoint"
  activate
  open POSIX file "$PPTX"
  delay 5
  set thePres to active presentation
  save thePres in (POSIX file "$PDF") as save as PDF
  delay 2
  close thePres saving no
end tell
OSAEOF

PAGES=$(mdls -name kMDItemNumberOfPages -raw "$PDF" 2>/dev/null || echo 1)
if [ "$PAGES" = "1" ] || [ -z "$PAGES" ]; then
  sips -s format png -Z 1920 "$PDF" --out "${OUT_PREFIX}-slide-01.png" >/dev/null
else
  TMP=$(mktemp -d)
  pdfseparate "$PDF" "$TMP/page-%d.pdf"
  for f in "$TMP"/page-*.pdf; do
    n=$(basename "$f" .pdf | sed 's/page-//')
    nn=$(printf '%02d' "$n")
    sips -s format png -Z 1920 "$f" --out "${OUT_PREFIX}-slide-${nn}.png" >/dev/null
  done
  rm -rf "$TMP"
fi
rm -f "$PDF"
echo "Captured slides to ${OUT_PREFIX}-slide-*.png"
ls "$OUT_DIR" | grep "^${OUT_BASE}-slide-" | head
