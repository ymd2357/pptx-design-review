#!/usr/bin/env python3
"""Build tmp/review-snapshot/<deck>/rev-<rev>/meta.json from before/after PNGs.

For each slide-NN.png that exists in BOTH images/before/ and images/after/,
computes how many pixels differ between before and after **above a luminance
threshold** (= anti-aliasing noise is ignored), writes a two-color diff PNG
into `images/diff/`, and emits a `meta.json` with the list of changed slides.

The diff PNG shows REMOVED content in red (= content that was in `before`
but not in `after`) and ADDED content in green (= content present only in
`after`). The background is a faded version of `before` to keep context.

Usage:
    python3 scripts/build-compare-meta.py --deck 260329-seminar-curriculum-proposal --rev 019 [--threshold 1000]
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

import numpy as np
from PIL import Image

REPO_ROOT = Path(__file__).resolve().parent.parent
SNAPSHOT_ROOT = REPO_ROOT / "tmp" / "review-snapshot"

# Per-pixel luminance threshold for treating a pixel as "actually different"
# vs "anti-aliasing noise". Empirically PowerPoint render noise is in the
# single-digit luminance range; 20/255 leaves clear margin.
PIXEL_DIFF_LUMINANCE_THRESHOLD = 20

# Colors for the two-tone diff (RGB).
REMOVED_COLOR = (220, 60, 60)   # red = content was in `before` but not `after`
ADDED_COLOR = (60, 180, 60)     # green = content present only in `after`
BASE_FADE_RATIO = 0.4           # base.png contributes 40 %; remaining 60 % white


def compare_and_maybe_write_diff(
    before_path: Path,
    after_path: Path,
    diff_path: Path,
    threshold: int,
) -> tuple[int, bool]:
    """Count differing pixels (luminance-thresholded) and write a red/green
    diff PNG **only when the count exceeds `threshold`**. Returns
    `(diff_pixel_count, wrote_diff)`. Unchanged slides skip the diff write
    entirely (= no compute + storage waste on slides whose fix didn't touch
    them).
    """
    base_rgb = np.array(Image.open(before_path).convert("RGB"))
    after_rgb = np.array(Image.open(after_path).convert("RGB"))
    if base_rgb.shape != after_rgb.shape:
        # Resize after to before's resolution if PowerPoint rendered them
        # slightly differently.
        after_img = Image.open(after_path).convert("RGB").resize(
            (base_rgb.shape[1], base_rgb.shape[0]), Image.BILINEAR
        )
        after_rgb = np.array(after_img)

    base_luma = base_rgb.mean(axis=2).astype(np.int16)
    after_luma = after_rgb.mean(axis=2).astype(np.int16)

    removed_mask = base_luma < (after_luma - PIXEL_DIFF_LUMINANCE_THRESHOLD)
    added_mask = base_luma > (after_luma + PIXEL_DIFF_LUMINANCE_THRESHOLD)
    diff_count = int(np.count_nonzero(removed_mask | added_mask))

    if diff_count <= threshold:
        return diff_count, False

    faded = (base_rgb * BASE_FADE_RATIO + 255 * (1 - BASE_FADE_RATIO)).astype(np.uint8)
    faded[removed_mask] = REMOVED_COLOR
    faded[added_mask] = ADDED_COLOR
    Image.fromarray(faded).save(diff_path)
    return diff_count, True


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
    diff_dir = base / "images" / "diff"
    if not before_dir.is_dir() or not after_dir.is_dir():
        print(f"missing before/after dirs under {base}", file=sys.stderr)
        return 2
    diff_dir.mkdir(parents=True, exist_ok=True)

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
        diff_path = diff_dir / before.name
        ae, wrote_diff = compare_and_maybe_write_diff(
            before, after, diff_path, args.threshold,
        )
        changed = ae > args.threshold
        slide_diffs.append({"slide_no": slide_no, "ae_pixels": ae, "changed": changed})
        if changed:
            changed_slides.append(slide_no)
        else:
            # Unchanged: before is the only image we need. Drop any stale
            # after / diff PNGs so the snapshot doesn't carry redundant
            # bytes (= REV-033 onwards saves ~17/20 PNG slots on a typical
            # deck where most slides are untouched).
            after.unlink(missing_ok=True)
            if not wrote_diff:
                diff_path.unlink(missing_ok=True)

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
