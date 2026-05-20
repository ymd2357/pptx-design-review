#!/usr/bin/env python3
"""Assert that the YAML-loaded lint_palette matches the hardcoded Python
constants in pptx_lint.py. Both sources are intentionally maintained side
by side (DS-COLOR-001) until the YAML version becomes the single source of
truth. The test is the gate that prevents silent drift between them.
"""

from __future__ import annotations

import importlib
import os
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))


def _reimport_lint(palette_source: str):
    os.environ["PPTX_PALETTE_SOURCE"] = palette_source
    sys.modules.pop("pptx_lint", None)
    return importlib.import_module("pptx_lint")


def main() -> int:
    failures: list[str] = []

    hard = _reimport_lint("hardcoded")
    yml = _reimport_lint("yaml")

    # Restore default for any downstream imports
    os.environ.pop("PPTX_PALETTE_SOURCE", None)

    if set(hard.ALLOWED_TEXT_COLORS_HEX) != set(yml.ALLOWED_TEXT_COLORS_HEX):
        failures.append(
            "ALLOWED_TEXT_COLORS_HEX diverges: "
            f"hardcoded-only={set(hard.ALLOWED_TEXT_COLORS_HEX) - set(yml.ALLOWED_TEXT_COLORS_HEX)}, "
            f"yaml-only={set(yml.ALLOWED_TEXT_COLORS_HEX) - set(hard.ALLOWED_TEXT_COLORS_HEX)}"
        )

    if set(hard.ALLOWED_FILL_COLORS_HEX) != set(yml.ALLOWED_FILL_COLORS_HEX):
        failures.append(
            "ALLOWED_FILL_COLORS_HEX diverges: "
            f"hardcoded-only={set(hard.ALLOWED_FILL_COLORS_HEX) - set(yml.ALLOWED_FILL_COLORS_HEX)}, "
            f"yaml-only={set(yml.ALLOWED_FILL_COLORS_HEX) - set(hard.ALLOWED_FILL_COLORS_HEX)}"
        )

    if dict(hard.TEXT_COLOR_TOKEN_BY_HEX) != dict(yml.TEXT_COLOR_TOKEN_BY_HEX):
        failures.append(
            "TEXT_COLOR_TOKEN_BY_HEX diverges between hardcoded and yaml"
        )

    if dict(hard.FILL_COLOR_TOKEN_BY_HEX) != dict(yml.FILL_COLOR_TOKEN_BY_HEX):
        failures.append(
            "FILL_COLOR_TOKEN_BY_HEX diverges between hardcoded and yaml"
        )

    if tuple(hard.CONTRAST_REPAIR_COLOR_FAMILIES) != tuple(yml.CONTRAST_REPAIR_COLOR_FAMILIES):
        failures.append("CONTRAST_REPAIR_COLOR_FAMILIES diverges")

    if tuple(hard.FILL_REPAIR_COLOR_FAMILIES) != tuple(yml.FILL_REPAIR_COLOR_FAMILIES):
        failures.append("FILL_REPAIR_COLOR_FAMILIES diverges")

    if failures:
        print("FAIL:")
        for f in failures:
            print(f"- {f}")
        return 1

    print("OK: design_system_loader (yaml) matches pptx_lint hardcoded constants for all 6 structures")
    return 0


if __name__ == "__main__":
    sys.exit(main())
