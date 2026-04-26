#!/usr/bin/env python3
"""Smoke test for pptx_lint against the generated examples.

Asserts:
- good.pptx produces zero findings (errors and warnings)
- bad.pptx triggers each of: overflow, safe_text_area, text_autofit_disabled,
  font_family, font_size_scale, text_color_allowlist, background_color_palette,
  animation_present

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
    "overflow",
    "safe_text_area",
    "text_autofit_disabled",
    "font_family",
    "font_size_scale",
    "text_color_allowlist",
    "background_color_palette",
    "animation_present",
}


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

    if failures:
        print("FAIL:")
        for line in failures:
            print(f"- {line}")
        return 1

    print(f"OK: good.pptx clean, bad.pptx triggered all {len(EXPECTED_BAD_CHECKS)} expected checks")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
