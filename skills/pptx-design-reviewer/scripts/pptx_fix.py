#!/usr/bin/env python3
"""PPTX auto-fixer for safe mechanical violations of slide-guideline-v1.

Fixes (mechanical, no semantic decisions):
- autofit   text frame auto-size != NONE -> set to MSO_AUTO_SIZE.NONE
- geometry  shape left/top/width/height in EMU rounded to nearest 1pt,
            but only when the current value is within 0.1pt of an integer
            (catches float drift; preserves intentional sub-pt placements)
- font_size text run size snapped to the nearest allowed scale size only when
            the feature flag is enabled, explicitly requested, and all safety
            checks pass; otherwise reported as manual_required

Out of fixer scope (require human judgment):
- font_family, overflow, safe_text_area, animation_present, slide_size

After --apply, the script re-reads the saved file and re-detects pending
actions; any residual means the change was not durable on disk. Known
causes include corrupted source PPTX (duplicate zip entries -- watch for
"Duplicate name: ppt/slides/slideN.xml" from zipfile) and inherited bodyPr
that python-pptx cannot override at slide level. Residuals are printed as
a warning and exit code becomes 2.

Usage
    python3 pptx_fix.py DECK.pptx                     # dry-run, prints plan
    python3 pptx_fix.py DECK.pptx --apply             # write in-place
    python3 pptx_fix.py DECK.pptx --apply --backup    # write DECK.pptx.bak if absent
    python3 pptx_fix.py DECK.pptx --apply --rules autofit
    python3 pptx_fix.py DECK.pptx --rules font_size
    python3 pptx_fix.py DECK.pptx --json

Exit code
    0 = success (or nothing to fix)
    1 = invocation error (missing file, unknown rule)
    2 = applied but self-check found residual actions on the saved file
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, List, Optional, Sequence

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Pt


EMU_PER_PT = 12700
GEOMETRY_ROUND_TOL_PT = 0.1  # only fix when drift is well below half-pt
SLIDE_W_PT = 1440
SLIDE_H_PT = 810
ALLOWED_FONT_SIZES_PT = {80, 64, 56, 48, 40, 36, 32, 28, 24, 22, 20}
FONT_SIZE_FIX_DELTA_MAX_PT = 1.0
FONT_SIZE_FIXER_ENABLED = False

DEFAULT_RULES = ("autofit", "geometry")
ALL_RULES = ("autofit", "geometry", "font_size")
CHECK_TO_RULE = {
    "text_autofit_disabled": "autofit",
    "geometry_rounding": "geometry",
    "font_size_scale": "font_size",
}


@dataclass
class FixAction:
    rule: str
    slide_index: int  # 1-based
    slide_id: Optional[int]
    shape_id: Optional[int]
    shape_name: Optional[str]
    status: str = "apply"  # "apply" | "manual_required"
    reasons: list[str] = field(default_factory=list)
    before: dict = field(default_factory=dict)
    after: dict = field(default_factory=dict)


# ---- Shape traversal -------------------------------------------------------


def _iter_shapes(shapes) -> Iterable:
    for s in shapes:
        if s.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _iter_shapes(s.shapes)
        else:
            yield s


def _walk(prs):
    for idx, slide in enumerate(prs.slides, start=1):
        slide_id = getattr(slide, "slide_id", None)
        for shape in _iter_shapes(slide.shapes):
            yield shape, idx, slide_id, slide


# ---- Per-rule detectors (pure: no mutation) --------------------------------


def _detect_autofit(shape, slide_idx, slide_id, slide=None) -> Optional[FixAction]:
    if not shape.has_text_frame:
        return None
    af = shape.text_frame.auto_size
    if af is None or af == MSO_AUTO_SIZE.NONE:
        return None
    return FixAction(
        rule="autofit",
        slide_index=slide_idx,
        slide_id=slide_id,
        shape_id=getattr(shape, "shape_id", None),
        shape_name=getattr(shape, "name", None),
        before={"auto_size": str(af)},
        after={"auto_size": "NONE"},
    )


def _round_emu_to_pt(value_emu: int) -> Optional[int]:
    """Return EMU rounded to nearest 1pt if drift < 0.1pt, else None."""
    pt = value_emu / EMU_PER_PT
    nearest = round(pt)
    if abs(pt - nearest) >= GEOMETRY_ROUND_TOL_PT:
        return None
    new_emu = nearest * EMU_PER_PT
    if new_emu == value_emu:
        return None
    return new_emu


def _detect_geometry(shape, slide_idx, slide_id, slide=None) -> Optional[FixAction]:
    coords = ("left", "top", "width", "height")
    raw = {c: getattr(shape, c) for c in coords}
    if any(v is None for v in raw.values()):
        return None
    new_vals = {}
    for c, v in raw.items():
        nv = _round_emu_to_pt(v)
        if nv is not None:
            new_vals[c] = nv
    if not new_vals:
        return None
    before = {c: round(raw[c] / EMU_PER_PT, 4) for c in new_vals}
    after = {c: new_vals[c] / EMU_PER_PT for c in new_vals}
    return FixAction(
        rule="geometry",
        slide_index=slide_idx,
        slide_id=slide_id,
        shape_id=getattr(shape, "shape_id", None),
        shape_name=getattr(shape, "name", None),
        before=before,
        after=after,
    )


def _nearest_allowed_font_size(size_pt: float) -> int:
    return min(ALLOWED_FONT_SIZES_PT, key=lambda allowed: abs(size_pt - allowed))


def _slide_scale(slide) -> float:
    prs = slide.part.package.presentation_part.presentation
    w_pt = prs.slide_width / EMU_PER_PT
    h_pt = prs.slide_height / EMU_PER_PT
    scale_x = SLIDE_W_PT / w_pt if w_pt else 1.0
    scale_y = SLIDE_H_PT / h_pt if h_pt else 1.0
    return (scale_x + scale_y) / 2


def _bbox_pt(shape) -> tuple[float, float, float, float]:
    return (
        shape.left / EMU_PER_PT,
        shape.top / EMU_PER_PT,
        shape.width / EMU_PER_PT,
        shape.height / EMU_PER_PT,
    )


def _overlaps(a: tuple[float, float, float, float], b: tuple[float, float, float, float]) -> bool:
    ax, ay, aw, ah = a
    bx, by, bw, bh = b
    ix = max(0, min(ax + aw, bx + bw) - max(ax, bx))
    iy = max(0, min(ay + ah, by + bh) - max(ay, by))
    return ix * iy > 1.0


def _estimate_text_width_pt(text: str, font_size_pt: float) -> float:
    width = 0.0
    for ch in text:
        code = ord(ch)
        if ch.isspace():
            width += 0.33
        elif 0x3040 <= code <= 0x30FF or 0x3400 <= code <= 0x9FFF:
            width += 1.0
        elif ch.isupper():
            width += 0.62
        else:
            width += 0.54
    return width * font_size_pt


def _text_containers(shape):
    if getattr(shape, "has_text_frame", False):
        yield {
            "kind": "shape",
            "text_frame": shape.text_frame,
            "width_pt": shape.width / EMU_PER_PT,
            "height_pt": shape.height / EMU_PER_PT,
            "offset": None,
        }

    if getattr(shape, "has_table", False):
        table = shape.table
        top = shape.top / EMU_PER_PT
        for row_idx, row in enumerate(table.rows, start=1):
            left = shape.left / EMU_PER_PT
            row_height = row.height / EMU_PER_PT
            for col_idx, col in enumerate(table.columns, start=1):
                cell = table.cell(row_idx - 1, col_idx - 1)
                col_width = col.width / EMU_PER_PT
                yield {
                    "kind": "table_cell",
                    "row": row_idx,
                    "col": col_idx,
                    "text_frame": cell.text_frame,
                    "width_pt": col_width,
                    "height_pt": row_height,
                    "offset": (left, top, col_width, row_height),
                }
                left += col_width
            top += row_height


def _non_empty_paragraphs(text_frame) -> list:
    return [
        para
        for para in text_frame.paragraphs
        if any(run.text for run in para.runs)
    ]


def _detect_font_size(shape, slide_idx, slide_id, slide=None) -> Optional[FixAction]:
    if slide is None or (not getattr(shape, "has_text_frame", False) and not getattr(shape, "has_table", False)):
        return None

    scale = _slide_scale(slide)
    updates: list[dict] = []
    reasons: list[str] = []
    before_sizes: list[dict] = []
    after_sizes: list[dict] = []
    affected_containers: list[dict] = []

    for container in _text_containers(shape):
        paragraphs = _non_empty_paragraphs(container["text_frame"])
        container_updates: list[dict] = []
        container_text = " ".join(
            run.text
            for para in paragraphs
            for run in para.runs
            if run.text
        )
        for p_idx, para in enumerate(container["text_frame"].paragraphs, start=1):
            for r_idx, run in enumerate(para.runs, start=1):
                if not run.text or run.font.size is None:
                    continue
                actual = run.font.size.pt
                normalized = actual * scale
                nearest = _nearest_allowed_font_size(normalized)
                target_actual = nearest / scale
                if abs(normalized - nearest) <= 1.0:
                    continue
                update = {
                    "kind": container["kind"],
                    "paragraph": p_idx,
                    "run": r_idx,
                    "before_size_pt": round(actual, 4),
                    "after_size_pt": round(target_actual, 4),
                    "normalized_size_pt": round(normalized, 4),
                    "target_normalized_size_pt": nearest,
                    "text": run.text[:80],
                }
                if container["kind"] == "table_cell":
                    update["row"] = container["row"]
                    update["col"] = container["col"]
                container_updates.append(update)

        if not container_updates:
            continue

        affected_containers.append({
            k: container[k]
            for k in ("kind", "row", "col")
            if k in container
        })
        updates.extend(container_updates)
        before_sizes.extend(
            {
                "size_pt": u["before_size_pt"],
                "normalized_size_pt": u["normalized_size_pt"],
                "text": u["text"],
            }
            for u in container_updates
        )
        after_sizes.extend(
            {
                "size_pt": u["after_size_pt"],
                "normalized_size_pt": u["target_normalized_size_pt"],
                "text": u["text"],
            }
            for u in container_updates
        )

        if len(paragraphs) != 1:
            reasons.append("not_single_line")
        elif "\n" in container_text or "\v" in container_text:
            reasons.append("not_single_line")

        max_delta = max(abs(u["after_size_pt"] - u["before_size_pt"]) for u in container_updates)
        if max_delta > FONT_SIZE_FIX_DELTA_MAX_PT:
            reasons.append("font_size_delta_too_large")

        target_size = max(u["after_size_pt"] for u in container_updates)
        required_height = target_size * 1.2
        if required_height > container["height_pt"] * 0.85:
            reasons.append("insufficient_box_height")
        if _estimate_text_width_pt(container_text, target_size) > container["width_pt"] * 0.95:
            reasons.append("insufficient_box_width")

        box = container["offset"] or _bbox_pt(shape)
        prs = slide.part.package.presentation_part.presentation
        slide_w = prs.slide_width / EMU_PER_PT
        slide_h = prs.slide_height / EMU_PER_PT
        x, y, w, h = box
        if x < 0 or y < 0 or x + w > slide_w or y + h > slide_h:
            reasons.append("would_overflow_slide")

    if not updates:
        return None

    shape_box = _bbox_pt(shape)
    shape_id = getattr(shape, "shape_id", None)
    for other in _iter_shapes(slide.shapes):
        if other is shape or getattr(other, "shape_id", None) == shape_id:
            continue
        try:
            if _overlaps(shape_box, _bbox_pt(other)):
                reasons.append("shape_bbox_overlaps_other_shape")
                break
        except (AttributeError, TypeError, ValueError):
            continue

    status = "manual_required" if reasons else "apply"
    return FixAction(
        rule="font_size",
        slide_index=slide_idx,
        slide_id=slide_id,
        shape_id=getattr(shape, "shape_id", None),
        shape_name=getattr(shape, "name", None),
        status=status,
        reasons=sorted(set(reasons)),
        before={"runs": before_sizes, "scale": round(scale, 4)},
        after={"runs": after_sizes, "updates": updates},
    )


DETECTORS: dict = {
    "autofit": _detect_autofit,
    "geometry": _detect_geometry,
    "font_size": _detect_font_size,
}

RULE_ENABLED: dict[str, bool] = {
    "autofit": True,
    "geometry": True,
    "font_size": FONT_SIZE_FIXER_ENABLED,
}


def _enabled_rules(rules: Sequence[str]) -> tuple[str, ...]:
    return tuple(rule for rule in rules if RULE_ENABLED.get(rule, True))


def _finding_field(finding: Any, key: str, default: Any = None) -> Any:
    if isinstance(finding, dict):
        return finding.get(key, default)
    return getattr(finding, key, default)


def _finding_detail(finding: Any) -> dict:
    detail = _finding_field(finding, "detail", {})
    return detail if isinstance(detail, dict) else {}


def _finding_detail_value(finding: Any, key: str, default: Any = None) -> Any:
    detail = _finding_detail(finding)
    if key in detail:
        return detail[key]
    return _finding_field(finding, key, default)


def _finding_rule(finding: Any) -> Optional[str]:
    check = _finding_field(finding, "check")
    if check is None:
        return None
    return CHECK_TO_RULE.get(check, check)


def _finding_matches_action(finding: Any, action: FixAction) -> bool:
    if _finding_rule(finding) != action.rule:
        return False

    finding_slide = _finding_field(finding, "slide_index")
    if finding_slide is not None and finding_slide != action.slide_index:
        return False

    finding_shape_id = _finding_field(finding, "shape_id")
    if finding_shape_id is not None and action.shape_id is not None:
        return finding_shape_id == action.shape_id

    finding_shape_name = _finding_field(finding, "shape_name")
    if finding_shape_name is not None and action.shape_name is not None:
        return finding_shape_name == action.shape_name

    return True


def _candidate_values_present(candidate_values: Any) -> bool:
    if candidate_values is None:
        return False
    if isinstance(candidate_values, (list, tuple, set, dict)):
        return len(candidate_values) > 0
    return True


def _finding_reasons(finding: Any) -> list[str]:
    reasons: list[str] = []
    for key in (
        "manual_required_reason",
        "manual_reason",
        "fixability_reason",
        "reason",
        "reasons",
    ):
        value = _finding_detail_value(finding, key)
        if value is None:
            continue
        if isinstance(value, str):
            reasons.append(value)
        elif isinstance(value, (list, tuple, set)):
            reasons.extend(str(item) for item in value if item is not None)
        else:
            reasons.append(str(value))
    return reasons


def _fixability_decision_from_finding(finding: Any) -> Optional[tuple[str, list[str]]]:
    """Return an action status override from evidence-schema findings.

    Legacy findings that do not declare fixability fall back to the existing
    detector-only behavior.
    """
    fixability = _finding_detail_value(finding, "fixability")
    if fixability is None:
        return None

    if fixability == "auto_fix_candidate":
        candidate_values = _finding_detail_value(finding, "candidate_values")
        if _candidate_values_present(candidate_values):
            return "apply", []
        return "manual_required", ["missing_candidate_values"]

    reasons = _finding_reasons(finding)
    if fixability == "manual_required":
        return "manual_required", reasons or ["finding_marked_manual_required"]

    return "manual_required", reasons or [f"fixability_{fixability}"]


def _apply_finding_fixability(action: FixAction, finding: Any) -> FixAction:
    decision = _fixability_decision_from_finding(finding)
    if decision is None:
        return action

    status, reasons = decision
    if status == "manual_required":
        action.status = "manual_required"
        action.reasons = sorted(set(action.reasons + reasons))
    elif action.status != "manual_required":
        action.status = status
    return action


def _apply_matching_finding_fixability(
    action: FixAction,
    findings: Optional[Sequence[Any]],
) -> FixAction:
    if not findings:
        return action

    for finding in findings:
        if _finding_matches_action(finding, action):
            return _apply_finding_fixability(action, finding)
    return action


# ---- Apply -----------------------------------------------------------------


def _backup_once(path: Path) -> None:
    backup_path = Path(str(path) + ".bak")
    if backup_path.exists():
        return
    shutil.copy2(path, backup_path)


def _apply_action(shape, action: FixAction) -> None:
    if action.status != "apply":
        return
    if action.rule == "autofit":
        shape.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    elif action.rule == "geometry":
        for c, pt in action.after.items():
            setattr(shape, c, int(round(pt * EMU_PER_PT)))
    elif action.rule == "font_size":
        for update in action.after.get("updates", []):
            if update["kind"] == "shape":
                tf = shape.text_frame
            else:
                tf = shape.table.cell(update["row"] - 1, update["col"] - 1).text_frame
            run = tf.paragraphs[update["paragraph"] - 1].runs[update["run"] - 1]
            run.font.size = Pt(update["after_size_pt"])


# ---- Driver ----------------------------------------------------------------


def fix_pptx(
    path: Path,
    *,
    apply: bool = False,
    backup: bool = False,
    rules: Sequence[str] = DEFAULT_RULES,
    findings: Optional[Sequence[Any]] = None,
) -> List[FixAction]:
    prs = Presentation(str(path))
    actions: List[FixAction] = []
    active_rules = _enabled_rules(rules)

    for shape, idx, sid, slide in _walk(prs):
        for rule in active_rules:
            det: Optional[Callable] = DETECTORS.get(rule)
            if det is None:
                continue
            action = det(shape, idx, sid, slide)
            if action is None:
                continue
            action = _apply_matching_finding_fixability(action, findings)
            actions.append(action)
            if action.status == "apply":
                _apply_action(shape, action)

    if apply and actions:
        if backup:
            _backup_once(path)
        prs.save(str(path))

    return actions


def verify_pptx(
    path: Path,
    *,
    rules: Sequence[str] = DEFAULT_RULES,
    findings: Optional[Sequence[Any]] = None,
) -> List[FixAction]:
    """Re-read the file from disk and report any actions still pending.

    A non-empty result after a successful --apply means the change was not
    durable on disk. Common causes: corrupted source PPTX with duplicate zip
    entries (python-pptx writes both copies, only one fixed), or inherited
    bodyPr that the slide-level setter cannot override.
    """
    prs = Presentation(str(path))
    residual: List[FixAction] = []
    active_rules = _enabled_rules(rules)
    for shape, idx, sid, slide in _walk(prs):
        for rule in active_rules:
            det = DETECTORS.get(rule)
            if det is None:
                continue
            action = det(shape, idx, sid, slide)
            if action is not None:
                action = _apply_matching_finding_fixability(action, findings)
            if action is not None and action.status == "apply":
                residual.append(action)
    return residual


# ---- Output ----------------------------------------------------------------


def _format_loc(a: FixAction) -> str:
    if a.shape_name and a.shape_id is not None:
        return f"{a.shape_name} (id={a.shape_id})"
    return a.shape_name or "-"


def format_actions(actions: List[FixAction], applied: bool) -> str:
    if not actions:
        return "OK: no fixable issues found.\n"
    by_rule: dict = {}
    for a in actions:
        by_rule.setdefault(a.rule, []).append(a)
    head = "Applied" if applied else "Would apply"
    applyable = [a for a in actions if a.status == "apply"]
    manual = [a for a in actions if a.status == "manual_required"]
    counts = ", ".join(f"{r}: {len(by_rule[r])}" for r in sorted(by_rule))
    lines = [
        f"{head} {len(applyable)} fixes; manual_required: {len(manual)} ({counts})",
        "",
    ]
    for rule in sorted(by_rule):
        lines.append(f"--- {rule} ---")
        for a in by_rule[rule]:
            loc = _format_loc(a)
            status = "" if a.status == "apply" else "  manual_required: " + ", ".join(a.reasons)
            if rule == "autofit":
                lines.append(
                    f"slide {a.slide_index}: {loc}  {a.before['auto_size']} -> NONE{status}"
                )
            elif rule == "geometry":
                diff = ", ".join(
                    f"{k}: {a.before[k]:g}->{a.after[k]:g}pt" for k in sorted(a.after)
                )
                lines.append(f"slide {a.slide_index}: {loc}  {diff}{status}")
            elif rule == "font_size":
                examples = a.after.get("updates", [])[:2]
                diff = ", ".join(
                    (
                        f"{r['before_size_pt']:g}->{r['after_size_pt']:g}pt "
                        f"({r['normalized_size_pt']:g}->{r['target_normalized_size_pt']:g}pt normalized)"
                    )
                    for r in examples
                )
                if len(a.after.get("updates", [])) > 2:
                    diff += f", ... {len(a.after.get('updates', []))} run(s)"
                lines.append(f"slide {a.slide_index}: {loc}  {diff}{status}")
        lines.append("")
    return "\n".join(lines)


def format_residual(residual: List[FixAction]) -> str:
    by_rule: dict = {}
    for a in residual:
        by_rule.setdefault(a.rule, []).append(a)
    counts = ", ".join(f"{r}: {len(by_rule[r])}" for r in sorted(by_rule))
    lines = [
        f"WARNING: self-check found {len(residual)} residual actions after save ({counts}).",
        "         The change did not persist on disk (corrupted PPTX zip or inherited bodyPr).",
        "         First few:",
    ]
    for a in residual[:5]:
        lines.append(
            f"  slide {a.slide_index}: {_format_loc(a)}  rule={a.rule}  before={a.before}"
        )
    if len(residual) > 5:
        lines.append(f"  ... and {len(residual) - 5} more")
    lines.append("")
    return "\n".join(lines)


# ---- CLI -------------------------------------------------------------------


def main(argv: List[str]) -> int:
    ap = argparse.ArgumentParser(description="PPTX auto-fixer (v1 guideline)")
    ap.add_argument("pptx", type=Path, help="path to .pptx")
    ap.add_argument(
        "--apply",
        action="store_true",
        help="write changes back to the file (default: dry-run)",
    )
    ap.add_argument(
        "--backup",
        action="store_true",
        help="with --apply, copy the original to <path>.bak before saving if absent",
    )
    ap.add_argument(
        "--rules",
        default=",".join(DEFAULT_RULES),
        help=(
            f"comma-separated subset of {{{','.join(ALL_RULES)}}} "
            f"(default: {','.join(DEFAULT_RULES)})"
        ),
    )
    ap.add_argument("--json", action="store_true", help="emit actions as JSON")
    args = ap.parse_args(argv)

    if not args.pptx.exists():
        print(f"file not found: {args.pptx}", file=sys.stderr)
        return 1

    rules = [r.strip() for r in args.rules.split(",") if r.strip()]
    bad = [r for r in rules if r not in ALL_RULES]
    if bad:
        print(
            f"unknown rule(s): {bad}; valid: {list(ALL_RULES)}",
            file=sys.stderr,
        )
        return 1

    if args.backup and not args.apply:
        print("note: --backup has no effect without --apply", file=sys.stderr)

    actions = fix_pptx(args.pptx, apply=args.apply, backup=args.backup, rules=rules)

    residual: List[FixAction] = []
    if args.apply and actions:
        residual = verify_pptx(args.pptx, rules=rules)

    if args.json:
        payload = {
            "applied": args.apply,
            "actions": [asdict(a) for a in actions],
            "residual": [asdict(a) for a in residual],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(format_actions(actions, applied=args.apply), end="")
        if residual:
            print(format_residual(residual), end="", file=sys.stderr)

    return 2 if residual else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
