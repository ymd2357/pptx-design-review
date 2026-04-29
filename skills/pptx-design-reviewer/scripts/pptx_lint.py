#!/usr/bin/env python3
"""PPTX lint MVP for slide-guideline-v1 compliance.

Checks
- overflow_text        (error)   text-bearing element extends beyond slide canvas
- overflow_shapes      (error)   shape element extends beyond slide canvas
- overflow_images      (error)   image element extends beyond slide canvas
- safe_text_area_text  (warning) text-bearing element outside safe text area
- text_autofit_disabled(error)   text frame auto-size is not NONE
- font_family          (warning) font name not in allowlist
- font_size_scale      (warning) font size not in allowed scale
- safe_margins        (warning) non-text element outside safe margins
- line_height         (warning) explicit paragraph line spacing not in allowed scale
- alignment_left_top  (warning) explicit text alignment differs from left/top
- geometry_rounding   (warning) geometry is not on integer pt coordinates
- image_upscale_ratio (warning) picture displayed larger than source pixels
- image_aspect_distortion (warning) picture aspect ratio differs from display box
- alt_text_required   (warning) meaningful image-like object has no alt text
- text_color_allowlist (warning) explicit text color not in allowlist
- background_color_palette (warning) explicit shape/table cell fill color not in palette
- text_overlap       (error)   text frames overlap each other
- object_overlap     (error)   non-text object bboxes overlap
- object_gap_too_small (warning) adjacent object gap is below minimum spacing
- alignment_drift    (warning) nearby left/top/center alignment differs
- inner_padding_imbalance (warning) child objects are unbalanced inside a container
- text_vertical_balance (warning) text fits but vertical padding/center balance is unnatural
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
from pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_ALIGN

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

ALLOWED_FONT_FAMILIES = ("Noto Sans JP", "Avenir Next Arabic", "Nunito Sans")
ALLOWED_FONT_SIZES_PT = {80, 64, 56, 48, 40, 36, 32, 28, 24, 22, 20}
FONT_SIZE_TOL_PT = 1.0
ALLOWED_LINE_HEIGHTS_PT = {90, 66, 42, 36, 30, 24}
LINE_HEIGHT_TOL_PT = 2.0
OBJECT_OVERLAP_AREA_PT2_MIN = 1.0
OBJECT_GAP_MIN_PT = 8.0
ALIGNMENT_GROUP_TOL_PT = 24.0
ALIGNMENT_DRIFT_TOL_PT = 2.0
INNER_PADDING_RATIO_MIN = 0.5
INNER_PADDING_RATIO_MAX = 2.0
INNER_PADDING_SIDE_MIN_PT = 4.0
TEXT_VERTICAL_BALANCE_DEAD_SPACE_RATIO_MAX = 0.40
TEXT_VERTICAL_BALANCE_DEAD_SPACE_PT_MAX = 60.0
TEXT_VERTICAL_BALANCE_MIDDLE_MARGIN_ASYMMETRY_PT_MAX = 12.0
TEXT_VERTICAL_BALANCE_CENTER_OFFSET_RATIO_MAX = 0.20
TEXT_VERTICAL_BALANCE_MIN_BOX_HEIGHT_PT = 30.0
DEFAULT_LINE_HEIGHT_MULTIPLIER = 1.2
MAX_IMAGE_UPSCALE_RATIO = 1.0
MAX_IMAGE_ASPECT_DELTA_RATIO = 0.05
DECORATIVE_RASTER_KEYWORDS = (
    "background",
    "bg",
    "decoration",
    "decorative",
    "footer",
    "gradient",
    "header",
)
PX_PER_PT = 96 / 72

ALLOWED_TEXT_COLORS_HEX = {
    "#000000",  # brand.utility.black
    "#333333",  # text.primary
    "#474747",  # brand.black.b800
    "#5C5C5C",  # brand.black.b700
    "#666666",  # text.muted
    "#707070",  # brand.black.b600
    "#858585",  # brand.black.b500
    "#999999",  # brand.black.b400
    "#1E112D",  # text.alt_dark
    "#FEFEFE",  # brand.utility.off_white
    "#FFFFFF",  # text.inverse
    "#A51E6D",  # accent.magenta, allowed for emphasis text
    "#0072B2",  # state.info, allowed for link-like text
}

ALLOWED_FILL_COLORS_HEX = {
    "#FFFFFF",  # background.base / neutral.n0
    "#F7F7F7",  # background.muted / neutral.n100
    "#EEEEEE",  # neutral.n200
    "#EBEBEB",  # brand.black.b50
    "#DDDDDD",  # neutral.n300 / border.default
    "#D6D6D6",  # brand.black.b100
    "#C2C2C2",  # brand.black.b200
    "#ADADAD",  # brand.black.b300
    "#999999",  # brand.black.b400
    "#858585",  # brand.black.b500
    "#707070",  # brand.black.b600
    "#5C5C5C",  # brand.black.b700
    "#474747",  # brand.black.b800
    "#333333",  # brand.black.b900
    "#F6E9F0",  # brand.primary.p50
    "#EDD2E2",  # brand.primary.p100
    "#E4BCD3",  # brand.primary.p200
    "#DBA5C5",  # brand.primary.p300
    "#D28FB6",  # brand.primary.p400
    "#C978A7",  # brand.primary.p500
    "#C06299",  # brand.primary.p600
    "#B74B8A",  # brand.primary.p700
    "#AE357C",  # brand.primary.p800
    "#A51E6D",  # accent.magenta
    "#E6F4F2",  # brand.secondary.s50
    "#CDEAE4",  # brand.secondary.s100
    "#B3DFD7",  # brand.secondary.s200
    "#9AD5C9",  # brand.secondary.s300
    "#81CABC",  # brand.secondary.s400
    "#68BFAE",  # brand.secondary.s500
    "#4FB5A1",  # brand.secondary.s600
    "#35AA93",  # brand.secondary.s700
    "#1CA086",  # brand.secondary.s800
    "#039578",  # brand.secondary.s900
    "#FF5757",  # brand.gradient.red
    "#E6033D",  # brand.gradient.red_deep
    "#32BED2",  # brand.gradient.cyan
    "#8C52FF",  # brand.gradient.purple
    "#673BA0",  # brand.gradient.purple_deep
    "#FFBC2A",  # brand.gold.base
    "#FFE177",  # brand.gold.light
    "#FFF9CF",  # brand.gold.pale
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
GEOMETRY_INTEGER_TOL_PT = 0.01
SLIDE_ASPECT_TOL = 0.01

PML_NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

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


@dataclass(frozen=True)
class LintPolicy:
    check_vertical_text_anchor: bool = False


LINT_PROFILES = {
    "default": LintPolicy(),
    "strict": LintPolicy(check_vertical_text_anchor=True),
}


@dataclass(frozen=True)
class LintContext:
    """Coordinate normalization from actual slide size to guideline base size."""

    actual_w_pt: float
    actual_h_pt: float
    scale_x: float
    scale_y: float
    proportional_to_base: bool
    policy: LintPolicy

    @property
    def font_scale(self) -> float:
        return (self.scale_x + self.scale_y) / 2


@dataclass
class ShapeRecord:
    shape: Any
    actual_bbox_pt: tuple
    bbox_pt: tuple
    kind: str


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


def _run_explicit_typefaces(run) -> list[tuple[str, str]]:
    try:
        r_pr = run._r.rPr
    except AttributeError:
        return []
    if r_pr is None:
        return []
    typefaces: list[tuple[str, str]] = []
    for script in ("latin", "ea"):
        node = r_pr.find(f"{A_NS}{script}")
        if node is None:
            continue
        name = node.get("typeface")
        if name:
            typefaces.append((script, name))
    if not typefaces:
        name = run.font.name
        if name:
            typefaces.append(("font", name))
    return typefaces


def _nearest_allowed_font_size(size_pt: float) -> int:
    return min(ALLOWED_FONT_SIZES_PT, key=lambda allowed: abs(size_pt - allowed))


def _font_size_allowed(size_pt: float) -> bool:
    return abs(size_pt - _nearest_allowed_font_size(size_pt)) <= FONT_SIZE_TOL_PT


def _nearest_allowed_line_height(line_height_pt: float) -> int:
    return min(ALLOWED_LINE_HEIGHTS_PT, key=lambda allowed: abs(line_height_pt - allowed))


def _line_height_allowed(line_height_pt: float) -> bool:
    return abs(line_height_pt - _nearest_allowed_line_height(line_height_pt)) <= LINE_HEIGHT_TOL_PT


def _normalized_length_pt(value, scale: float) -> float:
    if value is None:
        return 0.0
    try:
        return value.pt * scale
    except AttributeError:
        return emu_to_pt(value) * scale


def _text_frame_dominant_font_size_pt(ctx: LintContext, text_frame) -> Optional[float]:
    first_size_pt: Optional[float] = None
    size_weights: dict[float, int] = {}
    for para in text_frame.paragraphs:
        for run in para.runs:
            size = run.font.size
            if size is None:
                continue
            size_pt = round(size.pt * ctx.font_scale, 4)
            if first_size_pt is None:
                first_size_pt = size_pt
            weight = len(run.text.strip()) or 1
            size_weights[size_pt] = size_weights.get(size_pt, 0) + weight
    if size_weights:
        return max(size_weights.items(), key=lambda item: item[1])[0]
    return first_size_pt


def _paragraph_line_height_pt(ctx: LintContext, para, font_size_pt: float) -> float:
    line_spacing = para.line_spacing
    if line_spacing is None or isinstance(line_spacing, float):
        return font_size_pt * DEFAULT_LINE_HEIGHT_MULTIPLIER
    try:
        return line_spacing.pt * ctx.font_scale
    except AttributeError:
        return font_size_pt * DEFAULT_LINE_HEIGHT_MULTIPLIER


def _vertical_anchor_name(anchor) -> str:
    if anchor is None:
        return "NONE"
    if anchor == MSO_VERTICAL_ANCHOR.TOP:
        return "TOP"
    if anchor == MSO_VERTICAL_ANCHOR.MIDDLE:
        return "MIDDLE"
    if anchor == MSO_VERTICAL_ANCHOR.BOTTOM:
        return "BOTTOM"
    return str(anchor)


def _shape_has_visible_text(shape) -> bool:
    if not getattr(shape, "has_text_frame", False):
        return False
    return bool(shape.text_frame.text.strip())


def _shape_kind(shape) -> str:
    if _shape_has_visible_text(shape):
        return "text"
    if getattr(shape, "has_table", False):
        return "table"
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        return "image"
    return "shape"


def _shape_label(record: ShapeRecord) -> str:
    name = getattr(record.shape, "name", None) or record.kind
    shape_id = getattr(record.shape, "shape_id", None)
    if shape_id is None:
        return name
    return f"{name}#{shape_id}"


def _bbox_edges(bbox: tuple) -> tuple[float, float, float, float]:
    x, y, w, h = bbox
    return x, y, x + w, y + h


def _intersection_bbox(a: tuple, b: tuple) -> Optional[tuple]:
    ax1, ay1, ax2, ay2 = _bbox_edges(a)
    bx1, by1, bx2, by2 = _bbox_edges(b)
    x1 = max(ax1, bx1)
    y1 = max(ay1, by1)
    x2 = min(ax2, bx2)
    y2 = min(ay2, by2)
    if x2 <= x1 or y2 <= y1:
        return None
    return (x1, y1, x2 - x1, y2 - y1)


def _axis_gap(a_min: float, a_max: float, b_min: float, b_max: float) -> float:
    if a_max < b_min:
        return b_min - a_max
    if b_max < a_min:
        return a_min - b_max
    return 0.0


def _bbox_contains(outer: tuple, inner: tuple, tolerance: float = TOL_PT) -> bool:
    ox1, oy1, ox2, oy2 = _bbox_edges(outer)
    ix1, iy1, ix2, iy2 = _bbox_edges(inner)
    return (
        ix1 >= ox1 - tolerance
        and iy1 >= oy1 - tolerance
        and ix2 <= ox2 + tolerance
        and iy2 <= oy2 + tolerance
    )


def _shape_record_detail(record: ShapeRecord) -> dict:
    return {
        "shape_id": getattr(record.shape, "shape_id", None),
        "shape_name": getattr(record.shape, "name", None),
        "kind": record.kind,
        "bbox_pt": [round(v, 2) for v in record.bbox_pt],
    }


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


def _table_cell_fill_rgb_hexes(shape) -> dict[str, list[dict[str, int]]]:
    if not getattr(shape, "has_table", False):
        return {}

    colors: dict[str, list[dict[str, int]]] = defaultdict(list)
    for row_idx, row in enumerate(shape.table.rows, start=1):
        for col_idx, cell in enumerate(row.cells, start=1):
            try:
                fore_color = cell.fill.fore_color
            except (AttributeError, TypeError, ValueError):
                continue
            hex_color = _rgb_hex(fore_color)
            if hex_color:
                colors[hex_color].append({"row": row_idx, "col": col_idx})
    return colors


def _iter_text_runs(shape):
    if getattr(shape, "has_text_frame", False):
        for para in shape.text_frame.paragraphs:
            for run in para.runs:
                yield run, None

    if getattr(shape, "has_table", False):
        for row_idx, row in enumerate(shape.table.rows, start=1):
            for col_idx, cell in enumerate(row.cells, start=1):
                location = {"scope": "table_cell", "row": row_idx, "col": col_idx}
                for para in cell.text_frame.paragraphs:
                    for run in para.runs:
                        yield run, location


def _c_nv_pr(shape):
    try:
        matches = shape._element.xpath(".//p:cNvPr")
    except (AttributeError, TypeError, ValueError):
        return None
    return matches[0] if matches else None


def _alt_text_values(shape) -> tuple[str, str]:
    c_nv_pr = _c_nv_pr(shape)
    if c_nv_pr is None:
        return "", ""
    title = (c_nv_pr.get("title") or "").strip()
    descr = (c_nv_pr.get("descr") or "").strip()
    return title, descr


def _decorative_template_raster_reason(shape) -> Optional[str]:
    title, descr = _alt_text_values(shape)
    haystack = " ".join([getattr(shape, "name", ""), title, descr]).lower()
    for keyword in DECORATIVE_RASTER_KEYWORDS:
        if keyword in haystack:
            return f"keyword:{keyword}"
    return None


def shape_bbox_pt(shape) -> Optional[tuple]:
    if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
        return None
    return (
        emu_to_pt(shape.left),
        emu_to_pt(shape.top),
        emu_to_pt(shape.width),
        emu_to_pt(shape.height),
    )


def make_context(
    actual_w: float,
    actual_h: float,
    *,
    policy: LintPolicy,
) -> LintContext:
    scale_x = SLIDE_W_PT / actual_w if actual_w else 1.0
    scale_y = SLIDE_H_PT / actual_h if actual_h else 1.0
    base_aspect = SLIDE_W_PT / SLIDE_H_PT
    actual_aspect = actual_w / actual_h if actual_h else 0
    proportional = abs(actual_aspect - base_aspect) <= SLIDE_ASPECT_TOL
    return LintContext(
        actual_w_pt=actual_w,
        actual_h_pt=actual_h,
        scale_x=scale_x,
        scale_y=scale_y,
        proportional_to_base=proportional,
        policy=policy,
    )


def normalize_bbox(ctx: LintContext, bbox: tuple) -> tuple:
    x, y, w, h = bbox
    return (
        x * ctx.scale_x,
        y * ctx.scale_y,
        w * ctx.scale_x,
        h * ctx.scale_y,
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


def check_overflow(ctx, slide_idx, slide_id, shape, bbox, findings):
    normalized = normalize_bbox(ctx, bbox)
    x, y, w, h = normalized
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
    if shape.has_text_frame and shape.text_frame.text.strip():
        check_id = "overflow_text"
    elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        check_id = "overflow_images"
    else:
        check_id = "overflow_shapes"
    msg = "outside slide canvas: " + ", ".join(f"{s}+{a:.1f}pt" for s, a in out)
    findings.append(
        make_finding(
            "error", check_id, slide_idx, slide_id, shape, msg,
            {
                "bbox_pt": [round(v, 2) for v in normalized],
                "actual_bbox_pt": [round(v, 2) for v in bbox],
            },
        )
    )


def check_safe_text_area(ctx, slide_idx, slide_id, shape, bbox, findings):
    if not shape.has_text_frame:
        return
    if not shape.text_frame.text.strip():
        return
    normalized = normalize_bbox(ctx, bbox)
    x, y, w, h = normalized
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
            "warning", "safe_text_area_text", slide_idx, slide_id, shape, msg,
            {
                "bbox_pt": [round(v, 2) for v in normalized],
                "actual_bbox_pt": [round(v, 2) for v in bbox],
            },
        )
    )


def check_safe_margins(ctx, slide_idx, slide_id, shape, bbox, findings):
    if shape.has_text_frame and shape.text_frame.text.strip():
        return
    normalized = normalize_bbox(ctx, bbox)
    x, y, w, h = normalized
    out: List[tuple] = []
    if x < SAFE_MARGIN_LEFT_PT - TOL_PT:
        out.append(("left", SAFE_MARGIN_LEFT_PT - x))
    if y < SAFE_MARGIN_TOP_PT - TOL_PT:
        out.append(("top", SAFE_MARGIN_TOP_PT - y))
    if x + w > SLIDE_W_PT - SAFE_MARGIN_RIGHT_PT + TOL_PT:
        out.append(("right", (x + w) - (SLIDE_W_PT - SAFE_MARGIN_RIGHT_PT)))
    if y + h > SLIDE_H_PT - SAFE_MARGIN_BOTTOM_PT + TOL_PT:
        out.append(("bottom", (y + h) - (SLIDE_H_PT - SAFE_MARGIN_BOTTOM_PT)))
    if not out:
        return
    msg = "non-text outside safe margins: " + ", ".join(f"{s}+{a:.1f}pt" for s, a in out)
    findings.append(
        make_finding(
            "warning", "safe_margins", slide_idx, slide_id, shape, msg,
            {
                "bbox_pt": [round(v, 2) for v in normalized],
                "actual_bbox_pt": [round(v, 2) for v in bbox],
            },
        )
    )


def check_geometry_rounding(ctx, slide_idx, slide_id, shape, bbox, findings):
    normalized = normalize_bbox(ctx, bbox)
    names = ("x", "y", "w", "h")
    drifted = {
        name: round(value, 3)
        for name, value in zip(names, normalized)
        if abs(value - round(value)) > GEOMETRY_INTEGER_TOL_PT
    }
    if not drifted:
        return
    msg = "geometry is not rounded to integer pt: " + ", ".join(
        f"{name}={value:g}pt" for name, value in drifted.items()
    )
    findings.append(
        make_finding(
            "warning", "geometry_rounding", slide_idx, slide_id, shape, msg,
            {
                "bbox_pt": [round(v, 3) for v in normalized],
                "actual_bbox_pt": [round(v, 3) for v in bbox],
                "drifted": drifted,
            },
        )
    )


def check_image_upscale(ctx, slide_idx, slide_id, shape, bbox, findings):
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return
    try:
        source_w_px, source_h_px = shape.image.size
    except (AttributeError, ValueError):
        return
    if source_w_px <= 0 or source_h_px <= 0:
        return
    decorative_reason = _decorative_template_raster_reason(shape)
    if decorative_reason:
        return
    normalized = normalize_bbox(ctx, bbox)
    _, _, display_w_pt, display_h_pt = normalized
    if display_w_pt <= 0 or display_h_pt <= 0:
        return
    source_aspect = source_w_px / source_h_px
    display_aspect = display_w_pt / display_h_pt
    aspect_delta_ratio = abs(display_aspect / source_aspect - 1)
    if aspect_delta_ratio > MAX_IMAGE_ASPECT_DELTA_RATIO:
        findings.append(
            make_finding(
                "warning", "image_aspect_distortion", slide_idx, slide_id, shape,
                (
                    f"image aspect ratio changed from {source_aspect:.2f} to "
                    f"{display_aspect:.2f} ({aspect_delta_ratio:.0%} delta)"
                ),
                {
                    "source_px": [source_w_px, source_h_px],
                    "source_aspect": round(source_aspect, 4),
                    "display_aspect": round(display_aspect, 4),
                    "aspect_delta_ratio": round(aspect_delta_ratio, 4),
                    "actual_display_pt": [round(bbox[2], 2), round(bbox[3], 2)],
                    "normalized_display_pt": [round(display_w_pt, 2), round(display_h_pt, 2)],
                },
            )
        )
    display_w_px = display_w_pt * PX_PER_PT
    display_h_px = display_h_pt * PX_PER_PT
    ratio = max(display_w_px / source_w_px, display_h_px / source_h_px)
    if ratio <= MAX_IMAGE_UPSCALE_RATIO + 0.01:
        return
    findings.append(
        make_finding(
            "warning", "image_upscale_ratio", slide_idx, slide_id, shape,
            f"image displayed at {ratio:.2f}x source pixel size; maximum is {MAX_IMAGE_UPSCALE_RATIO:.2f}x",
            {
                "source_px": [source_w_px, source_h_px],
                "display_px": [round(display_w_px, 1), round(display_h_px, 1)],
                "actual_display_pt": [round(bbox[2], 2), round(bbox[3], 2)],
                "normalized_display_pt": [round(display_w_pt, 2), round(display_h_pt, 2)],
                "upscale_ratio": round(ratio, 3),
                "source_aspect": round(source_aspect, 4),
                "display_aspect": round(display_aspect, 4),
                "aspect_delta_ratio": round(aspect_delta_ratio, 4),
            },
        )
    )


def check_alt_text(slide_idx, slide_id, shape, findings):
    if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
        return
    title, descr = _alt_text_values(shape)
    if title or descr:
        return
    findings.append(
        make_finding(
            "warning", "alt_text_required", slide_idx, slide_id, shape,
            "picture has no alt text title or description",
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


def check_font(ctx, slide_idx, slide_id, shape, findings):
    if not shape.has_text_frame and not getattr(shape, "has_table", False):
        return
    bad_fonts: dict = {}
    bad_sizes: dict = {}
    bad_font_cells: dict = defaultdict(list)
    bad_size_cells: dict = defaultdict(list)
    for run, location in _iter_text_runs(shape):
        if not run.text:
            continue
        for script, name in _run_explicit_typefaces(run):
            if name and not _font_name_allowed(name):
                key = (script, name)
                bad_fonts[key] = bad_fonts.get(key, 0) + 1
                if location:
                    bad_font_cells[key].append(location)
        size = run.font.size
        if size is not None:
            actual_pt = round(size.pt, 1)
            normalized_pt = round(size.pt * ctx.font_scale, 1)
            if not _font_size_allowed(normalized_pt):
                key = (actual_pt, normalized_pt)
                bad_sizes[key] = bad_sizes.get(key, 0) + 1
                if location:
                    bad_size_cells[key].append(location)
    for (script, name), count in bad_fonts.items():
        findings.append(
            make_finding(
                "warning", "font_family", slide_idx, slide_id, shape,
                f"{script} font '{name}' not in allowlist {list(ALLOWED_FONT_FAMILIES)} ({count} run(s))",
                {
                    "script": script,
                    "font": name,
                    "runs": count,
                    "cells": bad_font_cells.get((script, name), []),
                },
            )
        )
    for (actual_pt, normalized_pt), count in bad_sizes.items():
        findings.append(
            make_finding(
                "warning", "font_size_scale", slide_idx, slide_id, shape,
                (
                    f"font size {normalized_pt}pt normalized from {actual_pt}pt "
                    f"not within {FONT_SIZE_TOL_PT:g}pt of scale "
                    f"{sorted(ALLOWED_FONT_SIZES_PT)} ({count} run(s))"
                ),
                {
                    "size_pt": normalized_pt,
                    "nearest_allowed_size_pt": _nearest_allowed_font_size(normalized_pt),
                    "actual_size_pt": actual_pt,
                    "normalization_scale": round(ctx.font_scale, 4),
                    "tolerance_pt": FONT_SIZE_TOL_PT,
                    "runs": count,
                    "cells": bad_size_cells.get((actual_pt, normalized_pt), []),
                },
            )
        )


def check_line_height(ctx, slide_idx, slide_id, shape, findings):
    if not shape.has_text_frame:
        return
    bad: dict = {}
    for para in shape.text_frame.paragraphs:
        if not para.text.strip():
            continue
        line_spacing = para.line_spacing
        if line_spacing is None:
            continue
        if isinstance(line_spacing, float):
            key = f"{line_spacing:g}x"
        else:
            actual_pt = round(line_spacing.pt, 1)
            normalized_pt = round(line_spacing.pt * ctx.font_scale, 1)
            if _line_height_allowed(normalized_pt):
                continue
            key = f"{normalized_pt:g}pt normalized from {actual_pt:g}pt"
        bad[key] = bad.get(key, 0) + 1
    for value, count in bad.items():
        findings.append(
            make_finding(
                "warning", "line_height", slide_idx, slide_id, shape,
                (
                    f"line spacing {value} not within {LINE_HEIGHT_TOL_PT:g}pt of "
                    f"allowed fixed pt scale {sorted(ALLOWED_LINE_HEIGHTS_PT)} "
                    f"({count} paragraph(s))"
                ),
                {
                    "line_spacing": value,
                    "allowed_line_heights_pt": sorted(ALLOWED_LINE_HEIGHTS_PT),
                    "tolerance_pt": LINE_HEIGHT_TOL_PT,
                    "paragraphs": count,
                },
            )
        )


def check_alignment(ctx, slide_idx, slide_id, shape, findings):
    if not shape.has_text_frame:
        return
    if not shape.text_frame.text.strip():
        return
    vertical = shape.text_frame.vertical_anchor
    if (
        ctx.policy.check_vertical_text_anchor
        and
        vertical is not None
        and vertical != MSO_VERTICAL_ANCHOR.TOP
    ):
        findings.append(
            make_finding(
                "warning", "alignment_left_top", slide_idx, slide_id, shape,
                f"text vertical alignment is {vertical}; expected TOP",
                {"vertical_anchor": str(vertical)},
            )
        )
    bad: dict = {}
    for para in shape.text_frame.paragraphs:
        if not para.text.strip():
            continue
        if para.alignment is not None and para.alignment != PP_ALIGN.LEFT:
            bad[str(para.alignment)] = bad.get(str(para.alignment), 0) + 1
    for alignment, count in bad.items():
        findings.append(
            make_finding(
                "warning", "alignment_left_top", slide_idx, slide_id, shape,
                f"paragraph alignment is {alignment}; expected LEFT ({count} paragraph(s))",
                {"alignment": alignment, "paragraphs": count},
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

    for fill_hex, cells in _table_cell_fill_rgb_hexes(shape).items():
        if fill_hex in ALLOWED_FILL_COLORS_HEX:
            continue
        findings.append(
            make_finding(
                "warning", "background_color_palette", slide_idx, slide_id, shape,
                (
                    f"table cell fill color {fill_hex} not in palette "
                    f"{sorted(ALLOWED_FILL_COLORS_HEX)} ({len(cells)} cell(s))"
                ),
                {"color_hex": fill_hex, "scope": "table_cell", "cells": cells},
            )
        )


def check_object_relationships(slide_idx, slide_id, records: list[ShapeRecord], findings):
    for idx, first in enumerate(records):
        for second in records[idx + 1:]:
            overlap_bbox = _intersection_bbox(first.bbox_pt, second.bbox_pt)
            if overlap_bbox is not None:
                overlap_area_pt2 = overlap_bbox[2] * overlap_bbox[3]
                if overlap_area_pt2 > OBJECT_OVERLAP_AREA_PT2_MIN:
                    if first.kind == "text" and second.kind == "text":
                        check_id = "text_overlap"
                        severity = "error"
                        subject = "text frames"
                    else:
                        check_id = "object_overlap"
                        severity = "error"
                        subject = "objects"
                    findings.append(
                        make_finding(
                            severity, check_id, slide_idx, slide_id, first.shape,
                            (
                                f"{subject} overlap by {overlap_area_pt2:.1f}pt^2 "
                                f"(>{OBJECT_OVERLAP_AREA_PT2_MIN:g}pt^2): "
                                f"{_shape_label(first)} / {_shape_label(second)}"
                            ),
                            {
                                "shape_a": _shape_record_detail(first),
                                "shape_b": _shape_record_detail(second),
                                "overlap_area_pt2": round(overlap_area_pt2, 2),
                                "overlap_bbox_pt": [round(v, 2) for v in overlap_bbox],
                                "threshold_pt2": OBJECT_OVERLAP_AREA_PT2_MIN,
                            },
                        )
                    )
                continue

            ax1, ay1, ax2, ay2 = _bbox_edges(first.bbox_pt)
            bx1, by1, bx2, by2 = _bbox_edges(second.bbox_pt)
            horizontal_gap = _axis_gap(ax1, ax2, bx1, bx2)
            vertical_gap = _axis_gap(ay1, ay2, by1, by2)
            horizontal_overlap = vertical_gap == 0.0
            vertical_overlap = horizontal_gap == 0.0
            gap_candidates: list[tuple[str, float]] = []
            if horizontal_overlap and 0 < horizontal_gap < OBJECT_GAP_MIN_PT:
                gap_candidates.append(("horizontal", horizontal_gap))
            if vertical_overlap and 0 < vertical_gap < OBJECT_GAP_MIN_PT:
                gap_candidates.append(("vertical", vertical_gap))
            for axis, gap in gap_candidates:
                findings.append(
                    make_finding(
                        "warning", "object_gap_too_small", slide_idx, slide_id, first.shape,
                        (
                            f"{axis} gap {gap:.1f}pt is below {OBJECT_GAP_MIN_PT:g}pt: "
                            f"{_shape_label(first)} / {_shape_label(second)}"
                        ),
                        {
                            "shape_a": _shape_record_detail(first),
                            "shape_b": _shape_record_detail(second),
                            "axis": axis,
                            "gap_pt": round(gap, 2),
                            "threshold_pt": OBJECT_GAP_MIN_PT,
                        },
                    )
                )

            x_metrics = {
                "left": (first.bbox_pt[0], second.bbox_pt[0]),
                "center_x": (
                    first.bbox_pt[0] + first.bbox_pt[2] / 2,
                    second.bbox_pt[0] + second.bbox_pt[2] / 2,
                ),
            }
            y_metrics = {
                "top": (first.bbox_pt[1], second.bbox_pt[1]),
                "center_y": (
                    first.bbox_pt[1] + first.bbox_pt[3] / 2,
                    second.bbox_pt[1] + second.bbox_pt[3] / 2,
                ),
            }

            def axis_drift(metrics: dict[str, tuple[float, float]]) -> dict[str, float]:
                deltas = {name: abs(a - b) for name, (a, b) in metrics.items()}
                if any(delta <= ALIGNMENT_DRIFT_TOL_PT for delta in deltas.values()):
                    return {}
                return {
                    name: round(delta, 2)
                    for name, delta in deltas.items()
                    if ALIGNMENT_DRIFT_TOL_PT < delta <= ALIGNMENT_GROUP_TOL_PT
                }

            drifted = axis_drift(x_metrics) | axis_drift(y_metrics)
            if drifted:
                findings.append(
                    make_finding(
                        "warning", "alignment_drift", slide_idx, slide_id, first.shape,
                        (
                            "near-aligned objects drift beyond "
                            f"{ALIGNMENT_DRIFT_TOL_PT:g}pt within "
                            f"{ALIGNMENT_GROUP_TOL_PT:g}pt group: "
                            + ", ".join(f"{name}+{delta:g}pt" for name, delta in drifted.items())
                        ),
                        {
                            "shape_a": _shape_record_detail(first),
                            "shape_b": _shape_record_detail(second),
                            "drifted": drifted,
                            "group_tolerance_pt": ALIGNMENT_GROUP_TOL_PT,
                            "drift_tolerance_pt": ALIGNMENT_DRIFT_TOL_PT,
                        },
                    )
                )


def _is_container_candidate(record: ShapeRecord) -> bool:
    if record.kind in {"text", "image", "table"}:
        return False
    _, _, w, h = record.bbox_pt
    return w > INNER_PADDING_SIDE_MIN_PT * 2 and h > INNER_PADDING_SIDE_MIN_PT * 2


def check_inner_padding_imbalance(slide_idx, slide_id, records: list[ShapeRecord], findings):
    for container in records:
        if not _is_container_candidate(container):
            continue
        children = [
            child
            for child in records
            if child is not container and _bbox_contains(container.bbox_pt, child.bbox_pt)
        ]
        if len(children) < 2:
            continue

        child_left = min(child.bbox_pt[0] for child in children)
        child_top = min(child.bbox_pt[1] for child in children)
        child_right = max(child.bbox_pt[0] + child.bbox_pt[2] for child in children)
        child_bottom = max(child.bbox_pt[1] + child.bbox_pt[3] for child in children)
        container_left, container_top, container_right, container_bottom = _bbox_edges(container.bbox_pt)
        padding = {
            "left": child_left - container_left,
            "right": container_right - child_right,
            "top": child_top - container_top,
            "bottom": container_bottom - child_bottom,
        }
        triggered_rules: list[str] = []
        for side, value in padding.items():
            if value < INNER_PADDING_SIDE_MIN_PT:
                triggered_rules.append(f"{side}_padding_below_min")

        horizontal_ratio = None
        vertical_ratio = None
        if padding["left"] > 0 and padding["right"] > 0:
            horizontal_ratio = padding["left"] / padding["right"]
            if not INNER_PADDING_RATIO_MIN <= horizontal_ratio <= INNER_PADDING_RATIO_MAX:
                triggered_rules.append("horizontal_padding_ratio")
        if padding["top"] > 0 and padding["bottom"] > 0:
            vertical_ratio = padding["top"] / padding["bottom"]
            if not INNER_PADDING_RATIO_MIN <= vertical_ratio <= INNER_PADDING_RATIO_MAX:
                triggered_rules.append("vertical_padding_ratio")

        if not triggered_rules:
            continue

        findings.append(
            make_finding(
                "warning", "inner_padding_imbalance", slide_idx, slide_id, container.shape,
                (
                    f"container padding is imbalanced for {len(children)} child objects: "
                    + ", ".join(triggered_rules)
                ),
                {
                    "container": _shape_record_detail(container),
                    "children": [_shape_record_detail(child) for child in children],
                    "padding_pt": {side: round(value, 2) for side, value in padding.items()},
                    "horizontal_ratio": (
                        round(horizontal_ratio, 3) if horizontal_ratio is not None else None
                    ),
                    "vertical_ratio": (
                        round(vertical_ratio, 3) if vertical_ratio is not None else None
                    ),
                    "triggered_rules": triggered_rules,
                    "thresholds": {
                        "side_min_pt": INNER_PADDING_SIDE_MIN_PT,
                        "ratio_min": INNER_PADDING_RATIO_MIN,
                        "ratio_max": INNER_PADDING_RATIO_MAX,
                    },
                },
            )
        )


def check_text_vertical_balance(ctx, slide_idx, slide_id, shape, bbox, findings):
    if not shape.has_text_frame:
        return
    text_frame = shape.text_frame
    if not text_frame.text.strip():
        return

    normalized = normalize_bbox(ctx, bbox)
    _, box_y_pt, box_w_pt, box_h_pt = normalized
    if box_h_pt < TEXT_VERTICAL_BALANCE_MIN_BOX_HEIGHT_PT:
        return

    font_size_pt = _text_frame_dominant_font_size_pt(ctx, text_frame)
    if font_size_pt is None or font_size_pt <= 0:
        return

    non_empty_paragraphs = [para for para in text_frame.paragraphs if para.text.strip()]
    if not non_empty_paragraphs:
        return
    anchor = text_frame.vertical_anchor
    if (
        anchor is None
        and len(non_empty_paragraphs) == 1
        and box_w_pt >= SAFE_TEXT_AREA_PT[2] - TOL_PT
    ):
        return

    line_heights_pt = [
        _paragraph_line_height_pt(ctx, para, font_size_pt)
        for para in non_empty_paragraphs
    ]
    estimated_text_height_pt = sum(line_heights_pt)
    margin_top_pt = _normalized_length_pt(text_frame.margin_top, ctx.scale_y)
    margin_bottom_pt = _normalized_length_pt(text_frame.margin_bottom, ctx.scale_y)
    inner_height_pt = box_h_pt - margin_top_pt - margin_bottom_pt
    if estimated_text_height_pt <= 0 or inner_height_pt <= 0:
        return
    if estimated_text_height_pt > inner_height_pt + TOL_PT:
        return

    free_space_pt = max(0.0, inner_height_pt - estimated_text_height_pt)
    anchor_name = _vertical_anchor_name(anchor)

    if anchor == MSO_VERTICAL_ANCHOR.MIDDLE:
        anchor_top_offset_pt = free_space_pt / 2
        anchor_bottom_offset_pt = free_space_pt / 2
    elif anchor == MSO_VERTICAL_ANCHOR.BOTTOM:
        anchor_top_offset_pt = free_space_pt
        anchor_bottom_offset_pt = 0.0
    else:
        anchor_top_offset_pt = 0.0
        anchor_bottom_offset_pt = free_space_pt

    actual_top_pad_pt = margin_top_pt + anchor_top_offset_pt
    actual_bottom_pad_pt = margin_bottom_pt + anchor_bottom_offset_pt
    visual_center_y_pt = box_y_pt + margin_top_pt + anchor_top_offset_pt + estimated_text_height_pt / 2
    box_geometric_center_y_pt = box_y_pt + box_h_pt / 2
    visual_center_offset_pt = visual_center_y_pt - box_geometric_center_y_pt
    free_space_ratio = free_space_pt / inner_height_pt
    visual_center_offset_ratio = abs(visual_center_offset_pt) / inner_height_pt

    triggered_rules: list[str] = []
    if (
        free_space_ratio > TEXT_VERTICAL_BALANCE_DEAD_SPACE_RATIO_MAX
        and anchor != MSO_VERTICAL_ANCHOR.MIDDLE
    ):
        triggered_rules.append("dead_space_ratio_non_middle")
    if anchor in (None, MSO_VERTICAL_ANCHOR.TOP) and actual_bottom_pad_pt > TEXT_VERTICAL_BALANCE_DEAD_SPACE_PT_MAX:
        triggered_rules.append("top_anchor_bottom_dead_space")
    if anchor == MSO_VERTICAL_ANCHOR.BOTTOM and actual_top_pad_pt > TEXT_VERTICAL_BALANCE_DEAD_SPACE_PT_MAX:
        triggered_rules.append("bottom_anchor_top_dead_space")
    if (
        anchor == MSO_VERTICAL_ANCHOR.MIDDLE
        and abs(margin_top_pt - margin_bottom_pt)
        > TEXT_VERTICAL_BALANCE_MIDDLE_MARGIN_ASYMMETRY_PT_MAX
    ):
        triggered_rules.append("middle_anchor_margin_asymmetry")
    if visual_center_offset_ratio > TEXT_VERTICAL_BALANCE_CENTER_OFFSET_RATIO_MAX:
        triggered_rules.append("visual_center_offset")

    if not triggered_rules:
        return

    findings.append(
        make_finding(
            "warning", "text_vertical_balance", slide_idx, slide_id, shape,
            (
                "text fits but vertical padding/visual center is imbalanced: "
                f"{', '.join(triggered_rules)}"
            ),
            {
                "anchor": anchor_name,
                "font_size_pt": round(font_size_pt, 2),
                "line_height_pt": round(estimated_text_height_pt / len(non_empty_paragraphs), 2),
                "paragraphs": len(non_empty_paragraphs),
                "estimated_text_height_pt": round(estimated_text_height_pt, 2),
                "inner_height_pt": round(inner_height_pt, 2),
                "free_space_pt": round(free_space_pt, 2),
                "margin_top_pt": round(margin_top_pt, 2),
                "margin_bottom_pt": round(margin_bottom_pt, 2),
                "visual_center_offset_pt": round(visual_center_offset_pt, 2),
                "triggered_rules": triggered_rules,
                "bbox_pt": [round(v, 2) for v in normalized],
                "box_width_pt": round(box_w_pt, 2),
                "actual_padding_pt": {
                    "top": round(actual_top_pad_pt, 2),
                    "bottom": round(actual_bottom_pad_pt, 2),
                },
                "thresholds": {
                    "dead_space_ratio_max": TEXT_VERTICAL_BALANCE_DEAD_SPACE_RATIO_MAX,
                    "top_or_bottom_dead_space_pt_max": TEXT_VERTICAL_BALANCE_DEAD_SPACE_PT_MAX,
                    "middle_anchor_padding_asymmetry_pt_max": (
                        TEXT_VERTICAL_BALANCE_MIDDLE_MARGIN_ASYMMETRY_PT_MAX
                    ),
                    "visual_center_offset_ratio_max": TEXT_VERTICAL_BALANCE_CENTER_OFFSET_RATIO_MAX,
                    "min_box_height_pt": TEXT_VERTICAL_BALANCE_MIN_BOX_HEIGHT_PT,
                },
            },
        )
    )


# ---- Driver ----------------------------------------------------------------


def lint_pptx(
    path: Path,
    *,
    profile: str = "default",
    policy: Optional[LintPolicy] = None,
) -> List[Finding]:
    if policy is None:
        try:
            policy = LINT_PROFILES[profile]
        except KeyError as exc:
            raise ValueError(f"unknown lint profile: {profile}") from exc
    prs = Presentation(str(path))
    findings: List[Finding] = []

    actual_w = emu_to_pt(prs.slide_width)
    actual_h = emu_to_pt(prs.slide_height)
    ctx = make_context(actual_w, actual_h, policy=policy)
    if not ctx.proportional_to_base:
        findings.append(
            Finding(
                severity="warning",
                check="slide_size",
                slide_index=0,
                slide_id=None,
                shape_id=None,
                shape_name=None,
                message=(
                    f"slide size {actual_w:.1f}x{actual_h:.1f}pt is not proportional to "
                    f"guideline base {SLIDE_W_PT}x{SLIDE_H_PT}pt"
                ),
                detail={
                    "actual_pt": [actual_w, actual_h],
                    "base_pt": [SLIDE_W_PT, SLIDE_H_PT],
                    "scale": [round(ctx.scale_x, 4), round(ctx.scale_y, 4)],
                },
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
        records: list[ShapeRecord] = []
        for shape in iter_shapes(slide.shapes):
            bbox = shape_bbox_pt(shape)
            if bbox is None:
                continue
            records.append(
                ShapeRecord(
                    shape=shape,
                    actual_bbox_pt=bbox,
                    bbox_pt=normalize_bbox(ctx, bbox),
                    kind=_shape_kind(shape),
                )
            )
        check_object_relationships(idx, slide_id, records, findings)
        check_inner_padding_imbalance(idx, slide_id, records, findings)
        for record in records:
            shape = record.shape
            bbox = record.actual_bbox_pt
            before_overflow_count = len(findings)
            check_overflow(ctx, idx, slide_id, shape, bbox, findings)
            overflowed = any(f.check == "overflow_text" for f in findings[before_overflow_count:])
            check_safe_text_area(ctx, idx, slide_id, shape, bbox, findings)
            check_safe_margins(ctx, idx, slide_id, shape, bbox, findings)
            check_geometry_rounding(ctx, idx, slide_id, shape, bbox, findings)
            check_image_upscale(ctx, idx, slide_id, shape, bbox, findings)
            check_alt_text(idx, slide_id, shape, findings)
            check_autofit(idx, slide_id, shape, findings)
            check_font(ctx, idx, slide_id, shape, findings)
            check_line_height(ctx, idx, slide_id, shape, findings)
            check_alignment(ctx, idx, slide_id, shape, findings)
            check_color(idx, slide_id, shape, findings)
            if not overflowed:
                check_text_vertical_balance(ctx, idx, slide_id, shape, bbox, findings)

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
    ap.add_argument(
        "--profile",
        choices=sorted(LINT_PROFILES),
        default="default",
        help="lint policy profile (default: default)",
    )
    args = ap.parse_args(argv)

    if not args.pptx.exists():
        print(f"file not found: {args.pptx}", file=sys.stderr)
        return 2

    findings = lint_pptx(args.pptx, profile=args.profile)
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
