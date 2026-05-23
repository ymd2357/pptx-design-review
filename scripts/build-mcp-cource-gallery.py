#!/usr/bin/env python3
"""Build a single-page HTML gallery comparing before vs 3 fix outputs for
mcp_cource: box_canvas_clip / text_box_resize / text_canvas_reflow.

Layout: one row per slide, 4 columns (before, fix-A, fix-B, fix-C). Slides
that changed (any of the 3 fixes produced a diff over the threshold) are
highlighted in the row header.
"""

from __future__ import annotations

import json
import shutil
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
WORK = REPO_ROOT / "tmp" / "review" / "mcp-cource"
SNAPSHOT_ROOT = REPO_ROOT / "tmp" / "review-snapshot" / "mcp-cource"
OUT_DIR = SNAPSHOT_ROOT / "gallery"

FIX_VARIANTS = [
    ("box-canvas-clip", "box_canvas_clip", "rev-003"),
    ("text-box-resize", "text_box_resize", "rev-004"),
    ("text-canvas-reflow", "text_canvas_reflow", "rev-005"),
]


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    images_dir = OUT_DIR / "images"
    images_dir.mkdir(exist_ok=True)
    before_src = WORK / "render-before"
    slides = sorted(int(p.stem.split("-")[1]) for p in before_src.glob("slide-*.png"))

    changed_slides: dict[str, set[int]] = {}
    for slug, _label, rev in FIX_VARIANTS:
        meta_path = SNAPSHOT_ROOT / rev / "meta.json"
        if meta_path.exists():
            meta = json.loads(meta_path.read_text(encoding="utf-8"))
            changed_slides[slug] = set(meta.get("changed_slides") or [])
        else:
            changed_slides[slug] = set()

    # Copy images: before + 3 variants
    for slide_no in slides:
        name = f"slide-{slide_no:02d}.png"
        for src_label, dest_prefix in [
            ("render-before", "before"),
            *((f"render-{slug}", slug) for slug, _, _ in FIX_VARIANTS),
        ]:
            src = WORK / src_label / name
            if src.exists():
                shutil.copyfile(src, images_dir / f"{dest_prefix}-{name}")

    rows_html: list[str] = []
    for slide_no in slides:
        name = f"slide-{slide_no:02d}.png"
        cells = [
            f'<td><img src="images/before-{name}" loading="lazy"></td>'
        ]
        for slug, _label, _rev in FIX_VARIANTS:
            after_path = images_dir / f"{slug}-{name}"
            changed = slide_no in changed_slides[slug]
            marker = '<span class="badge">CHANGED</span>' if changed else ""
            if after_path.exists():
                cells.append(
                    f'<td class="{"changed" if changed else ""}">'
                    f'<img src="images/{slug}-{name}" loading="lazy">{marker}'
                    f'</td>'
                )
            else:
                cells.append('<td class="missing">—</td>')
        rows_html.append(f"<tr><th>slide {slide_no}</th>{''.join(cells)}</tr>")

    head_html = (
        '<tr><th></th>'
        '<th>before</th>'
        '<th>box_canvas_clip<br><small>auto_fix</small></th>'
        '<th>text_box_resize<br><small>judgement / shrink_font</small></th>'
        '<th>text_canvas_reflow<br><small>judgement / enable_wrap</small></th>'
        '</tr>'
    )

    counts = ", ".join(
        f"{slug}: {len(changed_slides[slug])} changed" for slug, _, _ in FIX_VARIANTS
    )

    html = f"""<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>mcp-cource overflow fixes gallery</title>
<style>
  body {{ font-family: -apple-system, "Segoe UI", sans-serif; margin: 16px; }}
  h1 {{ font-size: 18px; }}
  .meta {{ color: #555; margin-bottom: 8px; font-size: 13px; }}
  table {{ border-collapse: collapse; }}
  th, td {{ border: 1px solid #ddd; padding: 6px; vertical-align: top; text-align: center; }}
  th {{ background: #f6f6f6; font-size: 13px; white-space: nowrap; }}
  td.changed {{ background: #fff4f4; position: relative; }}
  td.missing {{ color: #aaa; font-size: 20px; }}
  img {{ width: 320px; height: auto; display: block; border: 1px solid #eee; }}
  .badge {{
    position: absolute; top: 6px; right: 6px;
    background: #d63b3b; color: white; font-size: 10px;
    padding: 2px 6px; border-radius: 3px;
  }}
  small {{ color: #777; font-weight: normal; }}
</style>
</head>
<body>
<h1>mcp-cource overflow fixes (DS-OVERFLOW-001) ギャラリー</h1>
<p class="meta">3 つの新 fix rule を mcp_cource.pptx に個別適用した結果。{counts}</p>
<table>
{head_html}
{''.join(rows_html)}
</table>
</body>
</html>
"""
    (OUT_DIR / "index.html").write_text(html, encoding="utf-8")
    print(f"wrote {OUT_DIR / 'index.html'}")
    print(f"  {len(slides)} slides × 4 columns")
    for slug, _, _ in FIX_VARIANTS:
        print(f"  {slug}: {len(changed_slides[slug])} changed")


if __name__ == "__main__":
    main()
