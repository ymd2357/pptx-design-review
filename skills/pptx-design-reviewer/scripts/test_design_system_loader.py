#!/usr/bin/env python3
"""Verify that doc/slide-guideline-v1.yml `rules.color.lint_palette` is the
single source of truth for the 6 lint color structures consumed by
pptx_lint.py (DS-COLOR-001).

Prior to 2026-05-20 this test compared a hardcoded Python copy against
the YAML-loaded version. The hardcoded copy has since been removed:
``pptx_lint.py`` always reads from YAML via ``design_system_loader``. The
test therefore now smoke-checks that the loader populates every expected
structure with sensible content (non-empty, basic shape, sentinel values
present).
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import design_system_loader  # noqa: E402
import pptx_lint  # noqa: E402


def main() -> int:
    failures: list[str] = []

    palette = design_system_loader.load_lint_palette()

    if not palette.allowed_text_colors_hex:
        failures.append("LintPalette.allowed_text_colors_hex is empty")
    if not palette.allowed_fill_colors_hex:
        failures.append("LintPalette.allowed_fill_colors_hex is empty")
    if not palette.text_color_token_by_hex:
        failures.append("LintPalette.text_color_token_by_hex is empty")
    if not palette.fill_color_token_by_hex:
        failures.append("LintPalette.fill_color_token_by_hex is empty")
    if not palette.contrast_repair_color_families:
        failures.append("LintPalette.contrast_repair_color_families is empty")
    if not palette.fill_repair_color_families:
        failures.append("LintPalette.fill_repair_color_families is empty")

    # Module-level constants must mirror the loader (proves pptx_lint
    # imports without a stale hardcoded copy).
    if set(palette.allowed_text_colors_hex) != set(pptx_lint.ALLOWED_TEXT_COLORS_HEX):
        failures.append(
            "pptx_lint.ALLOWED_TEXT_COLORS_HEX drifted from YAML loader"
        )
    if set(palette.allowed_fill_colors_hex) != set(pptx_lint.ALLOWED_FILL_COLORS_HEX):
        failures.append(
            "pptx_lint.ALLOWED_FILL_COLORS_HEX drifted from YAML loader"
        )
    if dict(palette.text_color_token_by_hex) != dict(pptx_lint.TEXT_COLOR_TOKEN_BY_HEX):
        failures.append(
            "pptx_lint.TEXT_COLOR_TOKEN_BY_HEX drifted from YAML loader"
        )
    if dict(palette.fill_color_token_by_hex) != dict(pptx_lint.FILL_COLOR_TOKEN_BY_HEX):
        failures.append(
            "pptx_lint.FILL_COLOR_TOKEN_BY_HEX drifted from YAML loader"
        )
    if tuple(palette.contrast_repair_color_families) != tuple(
        pptx_lint.CONTRAST_REPAIR_COLOR_FAMILIES
    ):
        failures.append(
            "pptx_lint.CONTRAST_REPAIR_COLOR_FAMILIES drifted from YAML loader"
        )
    if tuple(palette.fill_repair_color_families) != tuple(
        pptx_lint.FILL_REPAIR_COLOR_FAMILIES
    ):
        failures.append(
            "pptx_lint.FILL_REPAIR_COLOR_FAMILIES drifted from YAML loader"
        )

    # Sentinel: brand black + brand white + text primary must be allowed.
    for expected in ("#000000", "#FFFFFF"):
        if expected not in palette.allowed_text_colors_hex:
            failures.append(
                f"allowed_text_colors_hex should contain {expected}; got "
                f"{sorted(palette.allowed_text_colors_hex)}"
            )

    if failures:
        print("FAIL:")
        for f in failures:
            print(f"- {f}")
        return 1
    print(
        "OK: rules.color.lint_palette loads and matches pptx_lint module "
        f"constants ({len(palette.allowed_text_colors_hex)} text + "
        f"{len(palette.allowed_fill_colors_hex)} fill colors)"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
