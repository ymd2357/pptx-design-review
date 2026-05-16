#!/usr/bin/env python3
"""Build tmp/review-snapshot/<deck>/rev-<rev>/meta.json from before/after PNGs.

For each slide-NN.png that exists in BOTH images/before/ and images/after/,
runs `magick compare -metric AE` to count the pixels that differ, and emits
a `meta.json` with the list of changed slides (count > threshold).

Usage:
    python3 scripts/build-compare-meta.py --deck 260329-seminar-curriculum-proposal --rev 019 [--threshold 1000]
"""

from __future__ import annotations

import argparse
import json
import re
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SNAPSHOT_ROOT = REPO_ROOT / "tmp" / "review-snapshot"


def compare_ae(a: Path, b: Path) -> int:
    proc = subprocess.run(
        ["magick", "compare", "-metric", "AE", str(a), str(b), "null:"],
        capture_output=True,
    )
    text = proc.stderr.decode(errors="replace").strip()
    match = re.search(r"^(\d+)", text)
    return int(match.group(1)) if match else 0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--deck", required=True)
    ap.add_argument("--rev", required=True)
    ap.add_argument("--threshold", type=int, default=1000,
                    help="AE pixel count below this is considered unchanged (default: 1000)")
    args = ap.parse_args()

    base = SNAPSHOT_ROOT / args.deck / f"rev-{args.rev}"
    before_dir = base / "images" / "before"
    after_dir = base / "images" / "after"
    if not before_dir.is_dir() or not after_dir.is_dir():
        print(f"missing before/after dirs under {base}", file=sys.stderr)
        return 2

    slide_diffs: list[dict] = []
    changed_slides: list[int] = []
    for before in sorted(before_dir.glob("slide-*.png")):
        match = re.match(r"slide-(\d+)\.png$", before.name)
        if not match:
            continue
        slide_no = int(match.group(1))
        after = after_dir / before.name
        if not after.is_file():
            continue
        ae = compare_ae(before, after)
        changed = ae > args.threshold
        slide_diffs.append({"slide_no": slide_no, "ae_pixels": ae, "changed": changed})
        if changed:
            changed_slides.append(slide_no)

    meta = {
        "deck": args.deck,
        "rev": args.rev,
        "threshold_ae_pixels": args.threshold,
        "changed_slides": changed_slides,
        "slide_diffs": slide_diffs,
    }
    out = base / "meta.json"
    out.write_text(json.dumps(meta, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(f"wrote {out.relative_to(REPO_ROOT)}")
    print(f"changed slides ({len(changed_slides)}): {changed_slides}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
