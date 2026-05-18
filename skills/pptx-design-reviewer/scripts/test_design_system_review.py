#!/usr/bin/env python3
"""Regression test for design_system_review (DS-003).

Locks in the current baseline of design-system audit findings against
`doc/slide-guideline-v1.yml` so that *new* drift is caught immediately
while known issues stay visible. Update the expected counts here when
the YAML is fixed or new check_ids are added.
"""

from __future__ import annotations

import sys
from collections import Counter
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import design_system_review as DSR  # noqa: E402


# Expected number of findings per check_id against the YAML at HEAD.
# Adjust this dict when the YAML is intentionally changed.
EXPECTED_COUNTS = {
    "palette_family_tier_gap": 0,
    "repair_candidate_family_overlap": 0,
    "repair_candidate_palette_drift": 1,  # data_series.series_6 (list-vs-dict)
}


def main() -> int:
    findings = DSR.review_design_system()
    counts = Counter(f.check for f in findings)
    failures: list[str] = []

    for check_id, expected in EXPECTED_COUNTS.items():
        actual = counts.get(check_id, 0)
        if actual != expected:
            failures.append(
                f"{check_id}: expected {expected}, got {actual} ("
                + "; ".join(
                    f.message for f in findings if f.check == check_id
                )
                + ")"
            )

    # Catch any newly emitted check that the test does not yet account for.
    unexpected = sorted(set(counts) - set(EXPECTED_COUNTS))
    if unexpected:
        failures.append(
            f"unexpected check_id(s) appeared: {unexpected}. "
            "Add them to EXPECTED_COUNTS or fix the underlying YAML."
        )

    # Validate that to_json() returns the evidence-schema keys for at least
    # one finding (smoke test for downstream consumers).
    if findings:
        schema = findings[0].to_json()
        required = {
            "check_id", "evidence", "fixability", "fixability_reason",
            "candidate_values", "manual_required_reason", "measurement_confidence",
        }
        missing = required - set(schema)
        if missing:
            failures.append(f"Finding.to_json missing schema keys: {sorted(missing)}")

    if failures:
        print("FAIL:")
        for f in failures:
            print(f"- {f}")
        return 1

    print(
        "OK: design_system_review baseline matches "
        f"({sum(EXPECTED_COUNTS.values())} expected finding(s))"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
