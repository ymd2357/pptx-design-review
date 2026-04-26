#!/usr/bin/env python3
"""Generate slide images for before/after review (side-by-side HTML).

Why
- PPTX cannot be rendered in a browser.
- For visual design review and question/answer with evidence, export slides to images.

Pipeline
1) PPTX -> PDF via LibreOffice (soffice)
2) PDF -> per-page PNG via poppler (pdftoppm)

Output
- <outdir>/before/slide-001.png ...
- <outdir>/after/slide-001.png ...
- <outdir>/index.html (side-by-side)

Notes
- This script does NOT modify the input PPTX files.
- It generates an HTML viewer and normalized filenames for easy comparison.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import List, Tuple


def die(msg: str, code: int = 1) -> None:
    print(msg, file=sys.stderr)
    raise SystemExit(code)


def which(cmd: str) -> str | None:
    from shutil import which as _which

    return _which(cmd)


_num_re = re.compile(r"(\d+)")


def _sort_key(p: Path) -> Tuple[int, str]:
    m = _num_re.search(p.stem)
    n = int(m.group(1)) if m else 10**9
    return (n, p.name)


def export_pdf(soffice: str, pptx: Path, outdir: Path) -> Path:
    outdir.mkdir(parents=True, exist_ok=True)
    cmd = [
        soffice,
        "--headless",
        "--invisible",
        "--norestore",
        "--nodefault",
        "--nolockcheck",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(outdir),
        str(pptx),
    ]
    r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    if r.returncode != 0:
        die("Failed to export PDF via soffice. Output:\n" + r.stdout)

    pdfs = sorted(outdir.glob("*.pdf"), key=_sort_key)
    if not pdfs:
        die("No PDF was exported. soffice output:\n" + r.stdout)

    # Usually exactly one PDF
    return pdfs[0]


def export_pngs_from_pdf(pdftoppm: str, pdf: Path, outdir: Path, dpi: int) -> List[Path]:
    outdir.mkdir(parents=True, exist_ok=True)
    prefix = outdir / "page"

    # pdftoppm writes files like page-1.png, page-2.png, ...
    cmd = [pdftoppm, "-png", "-r", str(dpi), str(pdf), str(prefix)]
    r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    if r.returncode != 0:
        die("Failed to export PNGs via pdftoppm. Output:\n" + r.stdout)

    pngs = sorted(outdir.glob("page-*.png"), key=_sort_key)
    if not pngs:
        die("No PNGs were exported from PDF.")
    return pngs


def normalize_pngs(pngs: List[Path], outdir: Path, limit: int | None) -> List[Path]:
    outdir.mkdir(parents=True, exist_ok=True)

    if limit is not None:
        pngs = pngs[:limit]

    out: List[Path] = []
    for i, src in enumerate(pngs, start=1):
        dst = outdir / f"slide-{i:03d}.png"
        shutil.copyfile(src, dst)
        out.append(dst)
    return out


def write_index(outdir: Path, before: List[Path], after: List[Path]) -> None:
    n = max(len(before), len(after))
    html = [
        "<!doctype html>",
        "<html>",
        "<head>",
        "  <meta charset='utf-8'>",
        "  <meta name='viewport' content='width=device-width,initial-scale=1'>",
        "  <title>PPTX review images</title>",
        "  <style>",
        "    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial; margin:16px;}",
        "    .row{display:grid; grid-template-columns:1fr 1fr; gap:16px; align-items:start; margin:24px 0;}",
        "    .cell{border:1px solid #ddd; padding:12px; border-radius:8px;}",
        "    .label{font-size:12px; color:#555; margin-bottom:8px;}",
        "    img{max-width:100%; height:auto; display:block; background:#f7f7f7;}",
        "    .idx{font-weight:700; margin-top:8px;}",
        "  </style>",
        "</head>",
        "<body>",
        "  <h1>PPTX review images</h1>",
        "  <p>Before/After exported by LibreOffice (PPTX→PDF) + poppler (PDF→PNG).</p>",
    ]

    for i in range(1, n + 1):
        b = f"before/slide-{i:03d}.png" if i <= len(before) else ""
        a = f"after/slide-{i:03d}.png" if i <= len(after) else ""
        html.extend(
            [
                f"  <div class='idx'>Slide {i}</div>",
                "  <div class='row'>",
                "    <div class='cell'>",
                "      <div class='label'>Before</div>",
                (f"      <img src='{b}' alt='before slide {i}'>" if b else "      <div>(missing)</div>"),
                "    </div>",
                "    <div class='cell'>",
                "      <div class='label'>After</div>",
                (f"      <img src='{a}' alt='after slide {i}'>" if a else "      <div>(missing)</div>"),
                "    </div>",
                "  </div>",
            ]
        )

    html.extend(["</body>", "</html>"])
    (outdir / "index.html").write_text("\n".join(html) + "\n", encoding="utf-8")


def main(argv: List[str]) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--before", required=True, help="Before PPTX path")
    ap.add_argument("--after", required=True, help="After PPTX path")
    ap.add_argument("--outdir", required=True, help="Output directory")
    ap.add_argument("--slides", type=int, default=0, help="Limit number of slides (0 = all)")
    ap.add_argument("--dpi", type=int, default=200, help="PNG DPI (higher = sharper, larger files)")
    args = ap.parse_args(argv)

    soffice = which("soffice")
    if not soffice:
        die("LibreOffice (soffice) not found in PATH.")

    pdftoppm = which("pdftoppm")
    if not pdftoppm:
        die("pdftoppm not found in PATH (install poppler).")

    before_pptx = Path(args.before)
    after_pptx = Path(args.after)
    outdir = Path(args.outdir)

    if not before_pptx.exists():
        die(f"Before PPTX not found: {before_pptx}")
    if not after_pptx.exists():
        die(f"After PPTX not found: {after_pptx}")

    limit = None if args.slides == 0 else args.slides

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        b_dir = tmp / "before"
        a_dir = tmp / "after"

        b_pdf = export_pdf(soffice, before_pptx, b_dir)
        a_pdf = export_pdf(soffice, after_pptx, a_dir)

        b_pngs_raw = export_pngs_from_pdf(pdftoppm, b_pdf, b_dir / "png", dpi=args.dpi)
        a_pngs_raw = export_pngs_from_pdf(pdftoppm, a_pdf, a_dir / "png", dpi=args.dpi)

        b_norm = normalize_pngs(b_pngs_raw, outdir / "before", limit)
        a_norm = normalize_pngs(a_pngs_raw, outdir / "after", limit)

    write_index(outdir, b_norm, a_norm)
    print(f"Wrote: {outdir/'index.html'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
