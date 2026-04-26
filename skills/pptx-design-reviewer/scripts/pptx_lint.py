#!/usr/bin/env python3
"""PPTX lint MVP for slide-guideline-v1 compliance.

Checks
- overflow             (error)   element extends beyond slide canvas (1440x810pt)
- safe_text_area       (warning) text-bearing element outside safe text area
- text_autofit_disabled(error)   text frame auto-size is not NONE
- font_family          (warning) font name not in allowlist
- font_size_scale      (warning) font size not in allowed scale
- text_color_allowlist (warning) explicit text color not in allowlist
- background_color_palette (warning) explicit shape fill color not in palette
- animation_present    (error)   slide contains <p:transition> or <p:timing>

Source of truth for thresholds is doc/slide-guideline-v1.yml.
Constants below mirror that file; keep them in sync if the guideline changes.

Usage
    python3 pptx_lint.py DECK.pptx
    python3 pptx_lint.py DECK.pptx --json
    python3 pptx_lint.py DECK.pptx --severity error

Exit code
    0 = no errors (warnings allowed)
    1 = at least one error
    2 = invocation error (file missing, etc.)
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Iterable, List, Optional

from pptx import Presentation
from pptx.dml.color import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE

# ---- Thresholds (mirror doc/slide-guideline-v1.yml) ------------------------

SLIDE_W_PT = 1440
SLIDE_H_PT = 810

SAFE_MARGIN_LEFT_PT = 81
SAFE_MARGIN_RIGHT_PT = 81
SAFE_MARGIN_TOP_PT = 40
SAFE_MARGIN_BOTTOM_PT = 80

SAFE_TEXT_AREA_PT = (
    SAFE_MARGIN_LEFT_PT,
    SAFE_MARGIN_TOP_PT,
    SLIDE_W_PT - SAFE_MARGIN_LEFT_PT - SAFE_MARGIN_RIGHT_PT,  # 1278
    SLIDE_H_PT - SAFE_MARGIN_TOP_PT - SAFE_MARGIN_BOTTOM_PT,  # 690
)

ALLOWED_FONT_FAMILIES = ("Noto Sans JP", "Calibri")
ALLOWED_FONT_SIZES_PT = {80, 56, 36, 32, 24, 20}

ALLOWED_TEXT_COLORS_HEX = {
    "#333333",  # text.primary
    "#666666",  # text.muted
    "#1E112D",  # text.alt_dark
    "#FFFFFF",  # text.inverse
    "#A51E6D",  # accent.magenta, allowed for emphasis text
    "#0072B2",  # state.info, allowed for link-like text
}

ALLOWED_FILL_COLORS_HEX = {
    "#FFFFFF",  # background.base / neutral.n0
    "#F7F7F7",  # background.muted / neutral.n100
    "#EEEEEE",  # neutral.n200
    "#DDDDDD",  # neutral.n300 / border.default
    "#A51E6D",  # accent.magenta
    "#1E112D",  # data series / dark emphasis
    "#0072B2",  # state.info / data series
    "#009E73",  # state.success / data series
    "#E69F00",  # state.warning / data series
    "#D55E00",  # state.danger / data series
    "#56B4E9",  # data series
    "#CC79A7",  # data series
}

# Tolerance for fp comparisons in pt. Geometry rounding policy is 0.5pt.
TOL_PT = 0.5

PML_NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}

# ---- Result types ----------------------------------------------------------


@dataclass
class Finding:
    severity: str  # "error" | "warning"
    check: str
    slide_index: int  # 1-based
    slide_id: Optional[int]
    shape_id: Optional[int]
    shape_name: Optional[str]
    message: str
    detail: dict = field(default_factory=dict)


# ---- Helpers ---------------------------------------------------------------


def emu_to_pt(emu: int) -> float:
    return emu / 12700.0


def iter_shapes(shapes) -> Iterable:
    """Recursively iterate shapes, descending into groups."""
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from iter_shapes(s.shapes)
        else:
            yield s


def animation_markers(slide) -> List[str]:
    markers: List[str] = []
    for tag in ("transition", "timing"):
        if slide.element.find(f"p:{tag}", PML_NS) is not None:
            markers.append(f"p:{tag}")
    return markers


def has_animation(slide) -> bool:
    return bool(animation_markers(slide))


def _font_name_allowed(name: str) -> bool:
    """Match family even when the run carries a weight suffix.

    Some source decks emit names like "Noto Sans JP Medium" or
    "Noto Sans JP Bold". Treat those as the same family as "Noto Sans JP".
    """
    for family in ALLOWED_FONT_FAMILIES:
        if name == family or name.startswith(family + " "):
            return True
    return False


def _rgb_hex(color_format) -> Optional[str]:
    """Return #RRGGBB for explicit RGB colors; inherited/theme colors are skipped."""
    try:
        if color_format.type != MSO_COLOR_TYPE.RGB:
            return None
        rgb = color_format.rgb
    except (AttributeError, TypeError, ValueError):
        return None
    if rgb is None:
        return None
    return f"#{str(rgb).upper()}"


def _shape_fill_rgb_hex(shape) -> Optional[str]:
    try:
        fore_color = shape.fill.fore_color
    except (AttributeError, TypeError, ValueError):
        return None
    return _rgb_hex(fore_color)


def shape_bbox_pt(shape) -> Optional[tuple]:
    if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
        return None
    return (
        emu_to_pt(shape.left),
        emu_to_pt(shape.top),
        emu_to_pt(shape.width),
        emu_to_pt(shape.height),
    )


def make_finding(severity, check, slide_idx, slide_id, shape, message, detail=None) -> Finding:
    return Finding(
        severity=severity,
        check=check,
        slide_index=slide_idx,
        slide_id=slide_id,
        shape_id=getattr(shape, "shape_id", None) if shape is not None else None,
        shape_name=getattr(shape, "name", None) if shape is not None else None,
        message=message,
        detail=detail or {},
    )


# ---- Checks ----------------------------------------------------------------


def check_overflow(slide_idx, slide_id, shape, bbox, findings):
    x, y, w, h = bbox
    right, bottom = x + w, y + h
    out: List[tuple] = []
    if x < -TOL_PT:
        out.append(("left", -x))
    if y < -TOL_PT:
        out.append(("top", -y))
    if right > SLIDE_W_PT + TOL_PT:
        out.append(("right", right - SLIDE_W_PT))
    if bottom > SLIDE_H_PT + TOL_PT:
        out.append(("bottom", bottom - SLIDE_H_PT))
    if not out:
        return
    msg = "outside slide canvas: " + ", ".join(f"{s}+{a:.1f}pt" for s, a in out)
    findings.append(
        make_finding(
            "error", "overflow", slide_idx, slide_id, shape, msg,
            {"bbox_pt": [round(v, 2) for v in bbox]},
        )
    )


def check_safe_text_area(slide_idx, slide_id, shape, bbox, findings):
    if not shape.has_text_frame:
        return
    if not shape.text_frame.text.strip():
        return
    x, y, w, h = bbox
    sx, sy, sw, sh = SAFE_TEXT_AREA_PT
    out: List[tuple] = []
    if x < sx - TOL_PT:
        out.append(("left", sx - x))
    if y < sy - TOL_PT:
        out.append(("top", sy - y))
    if x + w > sx + sw + TOL_PT:
        out.append(("right", (x + w) - (sx + sw)))
    if y + h > sy + sh + TOL_PT:
        out.append(("bottom", (y + h) - (sy + sh)))
    if not out:
        return
    msg = "text outside safe text area: " + ", ".join(f"{s}+{a:.1f}pt" for s, a in out)
    findings.append(
        make_finding(
            "warning", "safe_text_area", slide_idx, slide_id, shape, msg,
            {"bbox_pt": [round(v, 2) for v in bbox]},
        )
    )


def check_autofit(slide_idx, slide_id, shape, findings):
    if not shape.has_text_frame:
        return
    af = shape.text_frame.auto_size
    if af is None or af == MSO_AUTO_SIZE.NONE:
        return
    findings.append(
        make_finding(
            "error", "text_autofit_disabled", slide_idx, slide_id, shape,
            f"text auto-size is {af} (must be NONE)",
        )
    )


def check_font(slide_idx, slide_id, shape, findings):
    if not shape.has_text_frame:
        return
    bad_fonts: dict = {}
    bad_sizes: dict = {}
    for para in shape.text_frame.paragraphs:
        for run in para.runs:
            if not run.text:
                continue
            name = run.font.name
            if name and not _font_name_allowed(name):
                bad_fonts[name] = bad_fonts.get(name, 0) + 1
            size = run.font.size
            if size is not None:
                pt = round(size.pt, 1)
                if pt != int(pt) or int(pt) not in ALLOWED_FONT_SIZES_PT:
                    bad_sizes[pt] = bad_sizes.get(pt, 0) + 1
    for name, count in bad_fonts.items():
        findings.append(
            make_finding(
                "warning", "font_family", slide_idx, slide_id, shape,
                f"font '{name}' not in allowlist {list(ALLOWED_FONT_FAMILIES)} ({count} run(s))",
                {"font": name, "runs": count},
            )
        )
    for pt, count in bad_sizes.items():
        findings.append(
            make_finding(
                "warning", "font_size_scale", slide_idx, slide_id, shape,
                f"font size {pt}pt not in scale {sorted(ALLOWED_FONT_SIZES_PT)} ({count} run(s))",
                {"size_pt": pt, "runs": count},
            )
        )


def check_color(slide_idx, slide_id, shape, findings):
    if shape.has_text_frame:
        bad_text_colors: dict = {}
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                if not run.text:
                    continue
                hex_color = _rgb_hex(run.font.color)
                if hex_color and hex_color not in ALLOWED_TEXT_COLORS_HEX:
                    bad_text_colors[hex_color] = bad_text_colors.get(hex_color, 0) + 1
        for hex_color, count in bad_text_colors.items():
            findings.append(
                make_finding(
                    "warning", "text_color_allowlist", slide_idx, slide_id, shape,
                    (
                        f"text color {hex_color} not in allowlist "
                        f"{sorted(ALLOWED_TEXT_COLORS_HEX)} ({count} run(s))"
                    ),
                    {"color_hex": hex_color, "runs": count},
                )
            )

    fill_hex = _shape_fill_rgb_hex(shape)
    if fill_hex and fill_hex not in ALLOWED_FILL_COLORS_HEX:
        findings.append(
            make_finding(
                "warning", "background_color_palette", slide_idx, slide_id, shape,
                f"fill color {fill_hex} not in palette {sorted(ALLOWED_FILL_COLORS_HEX)}",
                {"color_hex": fill_hex},
            )
        )


# ---- Driver ----------------------------------------------------------------


def lint_pptx(path: Path) -> List[Finding]:
    prs = Presentation(str(path))
    findings: List[Finding] = []

    actual_w = emu_to_pt(prs.slide_width)
    actual_h = emu_to_pt(prs.slide_height)
    if abs(actual_w - SLIDE_W_PT) > 1 or abs(actual_h - SLIDE_H_PT) > 1:
        findings.append(
            Finding(
                severity="warning",
                check="slide_size",
                slide_index=0,
                slide_id=None,
                shape_id=None,
                shape_name=None,
                message=(
                    f"slide size {actual_w:.1f}x{actual_h:.1f}pt differs from "
                    f"guideline {SLIDE_W_PT}x{SLIDE_H_PT}pt"
                ),
                detail={"actual_pt": [actual_w, actual_h]},
            )
        )

    for idx, slide in enumerate(prs.slides, start=1):
        slide_id = getattr(slide, "slide_id", None)
        markers = animation_markers(slide)
        if markers:
            findings.append(
                make_finding(
                    "error", "animation_present", idx, slide_id, None,
                    "slide contains animation/transition XML: " + ", ".join(markers),
                )
            )
        for shape in iter_shapes(slide.shapes):
            bbox = shape_bbox_pt(shape)
            if bbox is None:
                continue
            check_overflow(idx, slide_id, shape, bbox, findings)
            check_safe_text_area(idx, slide_id, shape, bbox, findings)
            check_autofit(idx, slide_id, shape, findings)
            check_font(idx, slide_id, shape, findings)
            check_color(idx, slide_id, shape, findings)

    return findings


# ---- Consolidation (recurring template-level issues) ----------------------


def _slide_range(slides: List[int]) -> str:
    """Format a sorted list of ints as compact ranges, e.g. [2,3,4,7,9,10] -> '2-4, 7, 9-10'."""
    if not slides:
        return ""
    runs: List[tuple] = []
    start = prev = slides[0]
    for s in slides[1:]:
        if s == prev + 1:
            prev = s
        else:
            runs.append((start, prev))
            start = prev = s
    runs.append((start, prev))
    return ", ".join(str(a) if a == b else f"{a}-{b}" for a, b in runs)


def _group_key(f: Finding) -> tuple:
    """Return a grouping key that survives shape_name drift across slides.

    PowerPoint assigns shape_name like "Freeform 7" based on insertion order, so
    the same logical template element can be "Freeform 6" on one slide and
    "Freeform 7" on the next. When a finding carries a bbox in its detail, use
    that as the stable identifier; otherwise fall back to shape_name.
    """
    bbox = f.detail.get("bbox_pt") if f.detail else None
    if bbox:
        bbox_sig = ",".join(f"{round(v):d}" for v in bbox)
        return (f.check, f"bbox:{bbox_sig}", f.message)
    return (f.check, f.shape_name or "", f.message)


def consolidate_recurring(findings: List[Finding], min_slides: int = 3) -> List[Finding]:
    """Collapse identical findings that appear on many slides into one deck-level finding.

    When the same key (see _group_key) appears on >= min_slides distinct slides,
    replace the per-slide findings with a single deck-level entry that lists the
    affected slide range.
    """
    deck_level = [f for f in findings if f.slide_index == 0]
    per_slide = [f for f in findings if f.slide_index != 0]

    groups: dict = defaultdict(list)
    for f in per_slide:
        groups[_group_key(f)].append(f)

    consolidated: List[Finding] = []
    individual: List[Finding] = []
    for key, group in groups.items():
        slides = sorted({f.slide_index for f in group})
        if len(slides) >= min_slides:
            example = group[0]
            consolidated.append(
                Finding(
                    severity=example.severity,
                    check=example.check,
                    slide_index=0,
                    slide_id=None,
                    shape_id=None,
                    shape_name=example.shape_name,
                    message=(
                        f"{example.message} "
                        f"(recurring on {len(slides)} slides: {_slide_range(slides)})"
                    ),
                    detail={
                        "affected_slides": slides,
                        "occurrences": len(group),
                        "example_shape_id": example.shape_id,
                    },
                )
            )
        else:
            individual.extend(group)

    consolidated.sort(key=lambda f: (f.check, f.shape_name or ""))
    return deck_level + consolidated + individual


# ---- Output ----------------------------------------------------------------


def format_text(findings: List[Finding]) -> str:
    if not findings:
        return "OK: no issues found.\n"
    err_n = sum(1 for f in findings if f.severity == "error")
    warn_n = sum(1 for f in findings if f.severity == "warning")
    by_slide: dict = {}
    for f in findings:
        by_slide.setdefault(f.slide_index, []).append(f)
    lines = [f"Found {len(findings)} issues ({err_n} errors, {warn_n} warnings)", ""]
    for idx in sorted(by_slide):
        header = "Deck-level" if idx == 0 else f"Slide {idx}"
        lines.append(f"--- {header} ---")
        for f in by_slide[idx]:
            tag = "[ERR ]" if f.severity == "error" else "[WARN]"
            if f.shape_name and f.shape_id is not None:
                loc = f"{f.shape_name} (id={f.shape_id})"
            elif f.shape_name:
                loc = f.shape_name
            else:
                loc = "-"
            lines.append(f"{tag} {f.check}: {f.message}  [{loc}]")
        lines.append("")
    return "\n".join(lines)


def filter_by_severity(findings: List[Finding], min_sev: str) -> List[Finding]:
    if min_sev == "all":
        return findings
    if min_sev == "error":
        return [f for f in findings if f.severity == "error"]
    if min_sev == "warning":
        return findings  # warning is the lower bar; include errors too
    return findings


# ---- CLI -------------------------------------------------------------------


def main(argv: List[str]) -> int:
    ap = argparse.ArgumentParser(description="PPTX lint (v1 guideline)")
    ap.add_argument("pptx", type=Path, help="path to .pptx")
    ap.add_argument("--json", action="store_true", help="emit findings as JSON")
    ap.add_argument(
        "--severity",
        choices=["all", "warning", "error"],
        default="all",
        help="minimum severity to report (default: all)",
    )
    ap.add_argument(
        "--no-consolidate",
        action="store_true",
        help="disable deck-level consolidation of recurring identical findings",
    )
    ap.add_argument(
        "--min-recurring-slides",
        type=int,
        default=3,
        help="threshold for consolidating recurring findings (default: 3)",
    )
    args = ap.parse_args(argv)

    if not args.pptx.exists():
        print(f"file not found: {args.pptx}", file=sys.stderr)
        return 2

    findings = lint_pptx(args.pptx)
    if not args.no_consolidate:
        findings = consolidate_recurring(findings, min_slides=args.min_recurring_slides)
    selected = filter_by_severity(findings, args.severity)

    if args.json:
        payload: List[Any] = [asdict(f) for f in selected]
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(format_text(selected), end="")

    return 1 if any(f.severity == "error" for f in findings) else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
