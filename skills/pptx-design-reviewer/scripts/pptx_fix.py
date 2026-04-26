#!/usr/bin/env python3
"""PPTX auto-fixer for safe mechanical violations of slide-guideline-v1.

Fixes (mechanical, no semantic decisions):
- autofit   text frame auto-size != NONE -> set to MSO_AUTO_SIZE.NONE
- geometry  shape left/top/width/height in EMU rounded to nearest 1pt,
            but only when the current value is within 0.1pt of an integer
            (catches float drift; preserves intentional sub-pt placements)

Out of fixer scope (require human judgment):
- font_family, font_size_scale, overflow, safe_text_area, animation_present,
  slide_size

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
from typing import Callable, Iterable, List, Optional, Sequence

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_AUTO_SIZE


EMU_PER_PT = 12700
GEOMETRY_ROUND_TOL_PT = 0.1  # only fix when drift is well below half-pt

ALL_RULES = ("autofit", "geometry")


@dataclass
class FixAction:
    rule: str
    slide_index: int  # 1-based
    slide_id: Optional[int]
    shape_id: Optional[int]
    shape_name: Optional[str]
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
            yield shape, idx, slide_id


# ---- Per-rule detectors (pure: no mutation) --------------------------------


def _detect_autofit(shape, slide_idx, slide_id) -> Optional[FixAction]:
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


def _detect_geometry(shape, slide_idx, slide_id) -> Optional[FixAction]:
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


DETECTORS: dict = {
    "autofit": _detect_autofit,
    "geometry": _detect_geometry,
}


# ---- Apply -----------------------------------------------------------------


def _backup_once(path: Path) -> None:
    backup_path = Path(str(path) + ".bak")
    if backup_path.exists():
        return
    shutil.copy2(path, backup_path)


def _apply_action(shape, action: FixAction) -> None:
    if action.rule == "autofit":
        shape.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    elif action.rule == "geometry":
        for c, pt in action.after.items():
            setattr(shape, c, int(round(pt * EMU_PER_PT)))


# ---- Driver ----------------------------------------------------------------


def fix_pptx(
    path: Path,
    *,
    apply: bool = False,
    backup: bool = False,
    rules: Sequence[str] = ALL_RULES,
) -> List[FixAction]:
    prs = Presentation(str(path))
    actions: List[FixAction] = []

    for shape, idx, sid in _walk(prs):
        for rule in rules:
            det: Optional[Callable] = DETECTORS.get(rule)
            if det is None:
                continue
            action = det(shape, idx, sid)
            if action is None:
                continue
            actions.append(action)
            _apply_action(shape, action)

    if apply and actions:
        if backup:
            _backup_once(path)
        prs.save(str(path))

    return actions


def verify_pptx(path: Path, *, rules: Sequence[str] = ALL_RULES) -> List[FixAction]:
    """Re-read the file from disk and report any actions still pending.

    A non-empty result after a successful --apply means the change was not
    durable on disk. Common causes: corrupted source PPTX with duplicate zip
    entries (python-pptx writes both copies, only one fixed), or inherited
    bodyPr that the slide-level setter cannot override.
    """
    prs = Presentation(str(path))
    residual: List[FixAction] = []
    for shape, idx, sid in _walk(prs):
        for rule in rules:
            det = DETECTORS.get(rule)
            if det is None:
                continue
            action = det(shape, idx, sid)
            if action is not None:
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
    counts = ", ".join(f"{r}: {len(by_rule[r])}" for r in sorted(by_rule))
    lines = [f"{head} {len(actions)} fixes ({counts})", ""]
    for rule in sorted(by_rule):
        lines.append(f"--- {rule} ---")
        for a in by_rule[rule]:
            loc = _format_loc(a)
            if rule == "autofit":
                lines.append(
                    f"slide {a.slide_index}: {loc}  {a.before['auto_size']} -> NONE"
                )
            else:
                diff = ", ".join(
                    f"{k}: {a.before[k]:g}->{a.after[k]:g}pt" for k in sorted(a.after)
                )
                lines.append(f"slide {a.slide_index}: {loc}  {diff}")
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
        default=",".join(ALL_RULES),
        help=f"comma-separated subset of {{{','.join(ALL_RULES)}}} (default: all)",
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
