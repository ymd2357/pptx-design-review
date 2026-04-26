#!/usr/bin/env python3
"""Smoke test for pptx_lint against the generated examples.

Asserts:
- good.pptx produces zero findings (errors and warnings)
- bad.pptx triggers each expected implemented check
- every check emitted by pptx_lint.py is defined in rules.lint.checks

Exit code: 0 on success, 1 on any failed assertion.
"""

from __future__ import annotations

import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import make_examples  # noqa: E402
import pptx_lint  # noqa: E402


EXPECTED_BAD_CHECKS = {
    "overflow_text",
    "safe_text_area_text",
    "text_autofit_disabled",
    "font_family",
    "font_size_scale",
    "safe_margins",
    "line_height",
    "alignment_left_top",
    "geometry_rounding",
    "text_color_allowlist",
    "background_color_palette",
    "animation_present",
}

KNOWN_EMITTED_CHECKS = EXPECTED_BAD_CHECKS | {
    "slide_size",
    "overflow_shapes",
    "overflow_images",
}


def _guideline_lint_check_keys() -> set[str]:
    guideline = HERE.parents[2] / "doc" / "slide-guideline-v1.yml"
    lines = guideline.read_text(encoding="utf-8").splitlines()
    for idx, line in enumerate(lines):
        if line == "  lint:":
            keys: set[str] = set()
            in_checks = False
            for child in lines[idx + 1:]:
                if child and not child.startswith("    "):
                    break
                if child == "    checks:":
                    in_checks = True
                    continue
                if in_checks:
                    if child.startswith("      ") and child.endswith(":"):
                        keys.add(child.strip()[:-1])
                    elif child.startswith("    ") and child.strip() and not child.startswith("      "):
                        break
            return keys
    raise AssertionError("rules.lint.checks was not found in doc/slide-guideline-v1.yml")


def main() -> int:
    failures: list = []

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)
        good = tmp_dir / "good.pptx"
        bad = tmp_dir / "bad.pptx"
        make_examples.make_good(good)
        make_examples.make_bad(bad)

        good_findings = pptx_lint.lint_pptx(good)
        if good_findings:
            failures.append(
                "good.pptx had {n} findings; expected 0:\n  {lines}".format(
                    n=len(good_findings),
                    lines="\n  ".join(f"[{f.severity}] {f.check}: {f.message}" for f in good_findings),
                )
            )

        bad_findings = pptx_lint.lint_pptx(bad)
        bad_check_set = {f.check for f in bad_findings}
        missing = EXPECTED_BAD_CHECKS - bad_check_set
        if missing:
            failures.append(
                f"bad.pptx did not trigger {sorted(missing)}; got {sorted(bad_check_set)}"
            )
        guideline_checks = _guideline_lint_check_keys()
        unknown = KNOWN_EMITTED_CHECKS - guideline_checks
        if unknown:
            failures.append(
                f"pptx_lint.py emitted checks missing from rules.lint.checks: {sorted(unknown)}"
            )

    if failures:
        print("FAIL:")
        for line in failures:
            print(f"- {line}")
        return 1

    print(f"OK: good.pptx clean, bad.pptx triggered all {len(EXPECTED_BAD_CHECKS)} expected checks")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
