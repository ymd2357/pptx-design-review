#!/usr/bin/env python3
"""Focused tests for pptx_fix evidence-schema fixability consumption."""

from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import pptx_fix  # noqa: E402


def _action() -> pptx_fix.FixAction:
    return pptx_fix.FixAction(
        rule="geometry",
        slide_index=2,
        slide_id=256,
        shape_id=42,
        shape_name="drift-box",
        before={"left": 81.05},
        after={"left": 81.0},
    )


def _finding(detail: dict) -> SimpleNamespace:
    return SimpleNamespace(
        check="geometry_rounding",
        slide_index=2,
        slide_id=256,
        shape_id=42,
        shape_name="drift-box",
        detail=detail,
    )


def main() -> int:
    failures: list[str] = []

    auto_fixable = pptx_fix._apply_matching_finding_fixability(
        _action(),
        [
            _finding(
                {
                    "fixability": "auto_fix_candidate",
                    "candidate_values": [{"field": "left", "value_pt": 81.0}],
                }
            )
        ],
    )
    if auto_fixable.status != "apply":
        failures.append(
            f"auto_fix_candidate with candidates should apply; got {auto_fixable.status}"
        )

    # geometry_rounding is a judgement_fix-policy check (POLICY-001 段階 3).
    # With the strict gate (default), manual_required findings are skipped:
    manual_gated = pptx_fix._apply_matching_finding_fixability(
        _action(),
        [
            _finding(
                {
                    "fixability": "manual_required",
                    "manual_required_reason": "intentional_half_pt_grid",
                }
            )
        ],
    )
    if manual_gated.status != "manual_required":
        failures.append(
            "manual_required under judgement_fix policy should be skipped by "
            f"the strict gate; got {manual_gated.status}"
        )
    if "judgement_fix_gate_requires_spa_judgement" not in manual_gated.reasons:
        failures.append(
            "gated skip should record judgement_fix_gate_requires_spa_judgement; "
            f"got {manual_gated.reasons}"
        )

    # With judgement_gate=False (legacy callers), mechanical apply remains:
    manual_legacy = pptx_fix._apply_matching_finding_fixability(
        _action(),
        [
            _finding(
                {
                    "fixability": "manual_required",
                    "manual_required_reason": "intentional_half_pt_grid",
                }
            )
        ],
        judgement_gate=False,
    )
    if manual_legacy.status != "apply":
        failures.append(
            "manual_required with judgement_gate=False should still apply; "
            f"got {manual_legacy.status}"
        )
    if "intentional_half_pt_grid" not in manual_legacy.reasons:
        failures.append(
            f"manual_required reason was not preserved: {manual_legacy.reasons}"
        )

    missing_candidates = pptx_fix._apply_matching_finding_fixability(
        _action(),
        [_finding({"fixability": "auto_fix_candidate"})],
    )
    if missing_candidates.status != "manual_required":
        failures.append(
            "auto_fix_candidate without candidate_values should require manual review"
        )
    if "missing_candidate_values" not in missing_candidates.reasons:
        failures.append(
            f"missing candidate_values reason was not preserved: {missing_candidates.reasons}"
        )

    legacy = pptx_fix._apply_matching_finding_fixability(
        _action(),
        [_finding({"measured_value": 81.05, "threshold": "integer_pt"})],
    )
    if legacy.status != "apply" or legacy.reasons:
        failures.append(
            f"legacy finding should fall back to detector decision; got {legacy.status} {legacy.reasons}"
        )

    auto_rules = pptx_fix.auto_rules_from_findings(
        [
            _finding(
                {
                    "fixability": "auto_fix_candidate",
                    "fixability_rule": "geometry",
                    "candidate_values": {"rounded_values_pt": {"x": 81}},
                }
            ),
            _finding(
                {
                    "fixability": "manual_required",
                    "fixability_rule": "contrast",
                    "manual_required_reason": "complex_background",
                }
            ),
        ]
    )
    if auto_rules != ("geometry", "contrast"):
        failures.append(f"auto_rules_from_findings selected wrong rules: {auto_rules}")

    if failures:
        print("FAIL:")
        for line in failures:
            print(f"- {line}")
        return 1

    print("OK: pptx_fix consumes evidence-schema fixability without treating manual_required as manual修正")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
