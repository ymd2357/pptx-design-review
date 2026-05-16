#!/usr/bin/env bash
# PPTX -> PDF -> PNG (1 file per slide) using Microsoft PowerPoint Mac + CoreGraphics sips.
# Based on vscode-pptx-viewer/tools/capture-pp-mac.sh.
#
# Usage: scripts/capture-pp-mac.sh <pptx-path> <out-dir>
#   out-dir will be filled with slide-01.png, slide-02.png, ...
set -e
PPTX="${1:?pptx required}"
OUT_DIR="${2:?out dir required}"
mkdir -p "$OUT_DIR"
# PowerPoint sandbox blocks /private/tmp; keep intermediate PDF inside the
# user-supplied OUT_DIR so PowerPoint can write to it without a per-folder
# permission dialog (~/Documents, repo dirs, ~/Desktop are typically allowed).
PDF="$OUT_DIR/.capture.pdf"
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
  sips -s format png -Z 1920 "$PDF" --out "$OUT_DIR/slide-01.png" >/dev/null
else
  TMP=$(mktemp -d)
  pdfseparate "$PDF" "$TMP/page-%d.pdf"
  for f in "$TMP"/page-*.pdf; do
    n=$(basename "$f" .pdf | sed 's/page-//')
    nn=$(printf '%02d' "$n")
    sips -s format png -Z 1920 "$f" --out "$OUT_DIR/slide-${nn}.png" >/dev/null
  done
  rm -rf "$TMP"
fi
rm -f "$PDF"
echo "Captured slides to $OUT_DIR/"
ls "$OUT_DIR" | head
