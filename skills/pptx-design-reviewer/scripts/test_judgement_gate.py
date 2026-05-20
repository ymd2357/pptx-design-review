#!/usr/bin/env python3
"""POLICY-001 段階 3: judgement_fix gate の挙動を verify する。

- apply_mode=judgement_fix + fixability=manual_required + candidate_values
  あり: judgement_gate=True (default) で skip、judgement_gate=False で apply。
- apply_finding_judgements_overrides で promote した finding は gate=True
  でも apply される (= SPA judgement 経由 apply 経路)。
- apply_mode=auto_fix の finding は gate に影響されない。
"""

from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import pptx_fix  # noqa: E402


def _finding(check: str, detail: dict) -> SimpleNamespace:
    return SimpleNamespace(
        check=check,
        slide_index=1,
        slide_id=256,
        shape_id=42,
        shape_name="probe",
        detail=detail,
    )


def _action(rule: str) -> pptx_fix.FixAction:
    return pptx_fix.FixAction(
        rule=rule,
        slide_index=1,
        slide_id=256,
        shape_id=42,
        shape_name="probe",
        before={},
        after={},
    )


def main() -> int:
    failures: list[str] = []

    # 1. judgement_fix + manual_required + candidate → gate (default) skips.
    f1 = _finding(
        "text_color_allowlist",
        {
            "fixability": "manual_required",
            "candidate_values": {"color_hex": "#E6033D"},
        },
    )
    gated = pptx_fix._apply_matching_finding_fixability(_action("text_color"), [f1])
    if gated.status != "manual_required":
        failures.append(
            "judgement_fix + manual_required + candidate should be skipped under "
            f"strict gate; got {gated.status}"
        )
    if "judgement_fix_gate_requires_spa_judgement" not in gated.reasons:
        failures.append(f"gate skip reason missing: {gated.reasons}")

    # 2. Same finding, gate=False → applies (legacy behavior).
    legacy = pptx_fix._apply_matching_finding_fixability(
        _action("text_color"), [f1], judgement_gate=False
    )
    if legacy.status != "apply":
        failures.append(
            "judgement_fix + manual_required + candidate with judgement_gate=False "
            f"should apply; got {legacy.status}"
        )

    # 3. judgement_fix + auto_fix_candidate (e.g. promoted by SPA judgement)
    #    → applies even under strict gate.
    f3 = _finding(
        "text_color_allowlist",
        {
            "fixability": "auto_fix_candidate",
            "candidate_values": {"color_hex": "#E6033D"},
        },
    )
    promoted = pptx_fix._apply_matching_finding_fixability(_action("text_color"), [f3])
    if promoted.status != "apply":
        failures.append(
            "judgement_fix + auto_fix_candidate (promoted) should apply under "
            f"strict gate; got {promoted.status}"
        )

    # 4. auto_fix policy: gate must not change behavior.
    f4 = _finding(
        "badge_alignment",
        {
            "fixability": "auto_fix_candidate",
            "candidate_values": {"alignment": "CENTER", "vertical_anchor": "MIDDLE"},
        },
    )
    auto_gated = pptx_fix._apply_matching_finding_fixability(
        _action("badge_alignment"), [f4]
    )
    if auto_gated.status != "apply":
        failures.append(
            f"auto_fix policy under strict gate should apply; got {auto_gated.status}"
        )

    # 5. apply_mode lookup: declared policy returns the right mode.
    assert pptx_fix._apply_mode_for_check("text_color_allowlist") == "judgement_fix"
    assert pptx_fix._apply_mode_for_check("text_autofit_disabled") == "auto_fix"
    assert pptx_fix._apply_mode_for_check("alt_text_required") == "no_fix"
    assert pptx_fix._apply_mode_for_check("nonexistent_check") is None

    if failures:
        print("FAIL:")
        for f in failures:
            print(f"- {f}")
        return 1
    print(
        "OK: judgement_fix gate skips manual_required, lets auto_fix_candidate "
        "(promoted) through, and leaves auto_fix policy untouched"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
