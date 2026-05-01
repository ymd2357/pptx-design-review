#!/usr/bin/env python3
"""Smoke test for pptx_fix against generated fixtures.

Asserts:
- bad.pptx text_autofit_disabled findings drop to 0 after autofit fix
- non-targeted lint checks are not regressed (count unchanged)
- geometry rounding fixture: drifted (81.05pt) coord becomes exactly 81pt;
  intentional half-pt (81.5pt) is left untouched
- dry-run does not modify the file on disk
- existing .bak files are not overwritten by --backup
- self-check (verify_pptx) reports zero residual on a successfully fixed file

Exit code: 0 on success, 1 on any failed assertion.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.util import Emu, Pt

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import make_examples  # noqa: E402
import pptx_fix  # noqa: E402
import pptx_lint  # noqa: E402


DRIFT_LEFT_EMU = round(81.05 * 12700)  # 1029335 -> should round to 81pt
HALF_LEFT_EMU = round(81.5 * 12700)    # 1034605 -> drift 0.5pt, must stay


def _build_geometry_fixture(out: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(1440)
    prs.slide_height = Pt(810)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.shapes.add_textbox(Emu(DRIFT_LEFT_EMU), Pt(40), Pt(200), Pt(50))
    slide.shapes.add_textbox(Emu(HALF_LEFT_EMU), Pt(120), Pt(200), Pt(50))
    prs.save(str(out))


def _build_safe_font_size_fixture(out: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(720)
    prs.slide_height = Pt(405)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Pt(40), Pt(40), Pt(300), Pt(80))
    box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "Short label"
    run.font.name = "Noto Sans JP"
    run.font.size = Pt(14.75)
    prs.save(str(out))


def _build_manual_font_size_fixture(out: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(720)
    prs.slide_height = Pt(405)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Pt(40), Pt(40), Pt(300), Pt(27))
    box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "Cover label"
    run.font.name = "Noto Sans JP"
    run.font.size = Pt(22.5)
    prs.save(str(out))


def main() -> int:
    failures: list = []

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)

        # --- autofit fix on bad.pptx ---
        bad = tmp_dir / "bad.pptx"
        make_examples.make_bad(bad)

        pre = pptx_lint.lint_pptx(bad)
        pre_autofit = [f for f in pre if f.check == "text_autofit_disabled"]
        if not pre_autofit:
            failures.append("bad.pptx had no autofit findings to fix (fixture drift?)")

        actions = pptx_fix.fix_pptx(bad, apply=True, rules=("autofit",))
        if not any(a.rule == "autofit" for a in actions):
            failures.append("fixer reported no autofit actions on bad.pptx")

        post = pptx_lint.lint_pptx(bad)
        post_autofit = [f for f in post if f.check == "text_autofit_disabled"]
        if post_autofit:
            failures.append(
                f"autofit findings remain after fix: {len(post_autofit)}"
            )

        pre_other = sorted(f.check for f in pre if f.check != "text_autofit_disabled")
        post_other = sorted(f.check for f in post if f.check != "text_autofit_disabled")
        if pre_other != post_other:
            failures.append(
                f"non-targeted checks changed: pre={pre_other} post={post_other}"
            )

        # --- geometry fix on a drift+half fixture ---
        geo = tmp_dir / "geo.pptx"
        _build_geometry_fixture(geo)

        actions = pptx_fix.fix_pptx(geo, apply=True, rules=("geometry",))
        if not any(a.rule == "geometry" for a in actions):
            failures.append("geometry fixer reported no actions on drifted fixture")

        prs = Presentation(str(geo))
        shapes = list(prs.slides[0].shapes)
        drift_left_pt = shapes[0].left / 12700
        half_left_pt = shapes[1].left / 12700
        if abs(drift_left_pt - 81.0) > 1e-6:
            failures.append(
                f"drift_box.left expected 81pt; got {drift_left_pt}"
            )
        if abs(half_left_pt - 81.5) > 1e-6:
            failures.append(
                f"half_box.left should stay at 81.5pt; got {half_left_pt}"
            )

        # --- dry-run must not write ---
        bad2 = tmp_dir / "bad2.pptx"
        make_examples.make_bad(bad2)
        before_mtime = bad2.stat().st_mtime_ns
        actions = pptx_fix.fix_pptx(bad2, apply=False, rules=("autofit",))
        if not actions:
            failures.append("dry-run produced no actions on bad.pptx")
        if bad2.stat().st_mtime_ns != before_mtime:
            failures.append("dry-run modified the file on disk")

        # --- backup must preserve the oldest saved state ---
        backup_deck = tmp_dir / "backup.pptx"
        make_examples.make_bad(backup_deck)
        backup_path = Path(str(backup_deck) + ".bak")
        sentinel = b"existing original backup"
        backup_path.write_bytes(sentinel)
        pptx_fix.fix_pptx(backup_deck, apply=True, backup=True, rules=("autofit",))
        if backup_path.read_bytes() != sentinel:
            failures.append("--backup overwrote an existing .bak file")

        # --- font_size is behind an explicit feature flag ---
        safe_font = tmp_dir / "safe-font.pptx"
        _build_safe_font_size_fixture(safe_font)
        actions = pptx_fix.fix_pptx(safe_font, apply=True, rules=("font_size",))
        if actions:
            failures.append("font_size fixer should be disabled by default")
        prs = Presentation(str(safe_font))
        unchanged_size = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.size.pt
        if abs(unchanged_size - 14.75) > 1e-6:
            failures.append(f"disabled font_size fixer should leave 14.75pt; got {unchanged_size}")

        old_font_size_enabled = pptx_fix.RULE_ENABLED["font_size"]
        pptx_fix.RULE_ENABLED["font_size"] = True
        try:
            actions = pptx_fix.fix_pptx(safe_font, apply=True, rules=("font_size",))
            if not any(a.rule == "font_size" and a.status == "apply" for a in actions):
                failures.append("font_size fixer did not report an apply action for safe fixture")
            prs = Presentation(str(safe_font))
            fixed_size = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.size.pt
            if abs(fixed_size - 14.0) > 1e-6:
                failures.append(f"font_size safe fixture expected 14pt; got {fixed_size}")

            manual_font = tmp_dir / "manual-font.pptx"
            _build_manual_font_size_fixture(manual_font)
            actions = pptx_fix.fix_pptx(manual_font, apply=True, rules=("font_size",))
            if not any(a.rule == "font_size" and a.status == "manual_required" for a in actions):
                failures.append("font_size fixer did not report manual_required for unsafe fixture")
            prs = Presentation(str(manual_font))
            unchanged_size = prs.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.size.pt
            if abs(unchanged_size - 22.5) > 1e-6:
                failures.append(f"font_size manual fixture should remain 22.5pt; got {unchanged_size}")
        finally:
            pptx_fix.RULE_ENABLED["font_size"] = old_font_size_enabled

        # --- self-check (verify_pptx) reports no residual on fixed file ---
        bad3 = tmp_dir / "bad3.pptx"
        make_examples.make_bad(bad3)
        pptx_fix.fix_pptx(bad3, apply=True, rules=pptx_fix.ALL_RULES)
        residual = pptx_fix.verify_pptx(bad3, rules=pptx_fix.ALL_RULES)
        if residual:
            failures.append(
                f"verify_pptx reported residual on a clean fix: {len(residual)}"
            )

    if failures:
        print("FAIL:")
        for line in failures:
            print(f"- {line}")
        return 1

    print(
        "OK: pptx_fix removes autofit issues, rounds drifted geometry, "
        "and respects dry-run"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
