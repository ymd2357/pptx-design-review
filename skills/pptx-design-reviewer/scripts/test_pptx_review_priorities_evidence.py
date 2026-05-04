#!/usr/bin/env python3
"""Tests for schema-based priority evidence formatting."""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import pptx_lint  # noqa: E402
import pptx_review_priorities  # noqa: E402


SYNTHETIC_LINT_JSON = [
    {
        "severity": "error",
        "check": "low_contrast",
        "slide_index": 2,
        "slide_id": 258,
        "shape_id": 7,
        "shape_name": "Title 1",
        "message": "do-not-parse: rendered contrast message is legacy text",
        "detail": {
            "review_status": "needs_review",
            "shape_kind": "text",
            "text_excerpt": "受講後に到達できる状態",
            "measured_value": 2.42,
            "threshold": 3.0,
            "delta": -0.58,
            "unit": "ratio",
            "evidence_source": "rendered_image",
            "evidence_confidence": "high",
            "fixability": "manual_required",
            "manual_required_reason": "brand color/background pairing needs visual judgment",
            "candidate_values": [
                {
                    "text_color_hex": "#333333",
                    "background_color_hex": "#FFFFFF",
                    "verified_ratio": 12.63,
                }
            ],
            "artifact_refs": ["review/slide-02-contrast.png"],
        },
    },
    {
        "severity": "warning",
        "check": "contrast_ratio",
        "slide_index": 5,
        "slide_id": 261,
        "shape_id": 9,
        "shape_name": "Body 2",
        "message": "do-not-parse: legacy rendered contrast message",
        "detail": {
            "contrast_ratio": 4.21,
            "required_ratio": 4.5,
            "measurement": "rendered_image",
            "text_class": "normal_text",
            "font_size_pt": 24,
        },
    },
    {
        "severity": "warning",
        "check": "font_family",
        "slide_index": 3,
        "slide_id": 259,
        "shape_id": 4,
        "shape_name": "Body 1",
        "message": "do-not-parse: font family not in allowlist",
        "detail": {
            "shape_kind": "text",
            "text_excerpt": "Before / After",
            "measured_value": "Arial",
            "threshold": ["Noto Sans JP", "Nunito Sans"],
            "evidence_source": "pptx_xml",
            "evidence_confidence": "high",
            "fixability": "manual_required",
            "manual_required_reason": "mixed Latin/Japanese typography needs author choice",
            "recommended_value": "Noto Sans JP",
        },
    },
    {
        "severity": "warning",
        "check": "geometry_rounding",
        "slide_index": 4,
        "slide_id": 260,
        "shape_id": 2,
        "shape_name": "Image 1",
        "message": "do-not-parse: geometry is not rounded",
        "detail": {
            "shape_kind": "picture",
            "bbox_pt": [92.25, 120.5, 320.0, 180.0],
            "measured_value": 0.5,
            "threshold": 0.01,
            "unit": "pt",
            "evidence_source": "pptx_xml",
            "evidence_confidence": "high",
            "fixability": "auto_fix_candidate",
            "candidate_values": [{"x": 92, "y": 120, "w": 320, "h": 180}],
            "recommended_value": {"x": 92, "y": 120, "w": 320, "h": 180},
        },
    },
]


def _findings_from_json(items: list[dict]) -> list[pptx_lint.Finding]:
    return [
        pptx_lint.Finding(
            severity=item["severity"],
            check=item["check"],
            slide_index=item["slide_index"],
            slide_id=item["slide_id"],
            shape_id=item["shape_id"],
            shape_name=item["shape_name"],
            message=item["message"],
            detail=item["detail"],
        )
        for item in items
    ]


def _issue_for(issues: list[pptx_review_priorities.PriorityIssue], priority: str, check: str):
    for issue in issues:
        if issue.priority == priority and check in issue.checks:
            return issue
    return None


def main() -> int:
    failures: list[str] = []
    issues = pptx_review_priorities._summarize_findings(_findings_from_json(SYNTHETIC_LINT_JSON))

    p0 = _issue_for(issues, "P0", "low_contrast")
    if p0 is None:
        failures.append("schema low_contrast did not produce P0")
    else:
        evidence = "\n".join(p0.evidence)
        for expected in (
            'text="受講後に到達できる状態"',
            "measured=2.42ratio",
            "threshold=3ratio",
            "source=rendered_image/high",
            "fixability=manual_required",
            "brand color/background pairing needs visual judgment",
            "review/slide-02-contrast.png",
        ):
            if expected not in evidence:
                failures.append(f"P0 schema evidence missing {expected!r}: {evidence}")
        if "do-not-parse" in evidence:
            failures.append(f"P0 evidence used legacy message despite schema detail: {evidence}")

    p1 = _issue_for(issues, "P1", "contrast_ratio")
    if p1 is None:
        failures.append("legacy contrast_ratio detail did not produce P1")
    else:
        evidence = "\n".join(p1.evidence)
        for expected in ("measured=4.21", "threshold=4.5", "source=rendered_image"):
            if expected not in evidence:
                failures.append(f"P1 transition evidence missing {expected!r}: {evidence}")
        if "do-not-parse" in evidence:
            failures.append(f"P1 transition evidence used message instead of detail: {evidence}")

    p2 = _issue_for(issues, "P2", "font_family")
    if p2 is None:
        failures.append("schema font_family did not produce P2")
    else:
        evidence = "\n".join(p2.evidence)
        for expected in ("fixability=manual_required", "mixed Latin/Japanese typography", "recommended=Noto Sans JP"):
            if expected not in evidence:
                failures.append(f"P2 manual_required evidence missing {expected!r}: {evidence}")

    p3 = _issue_for(issues, "P3", "geometry_rounding")
    if p3 is None:
        failures.append("schema geometry_rounding did not produce P3")
    else:
        evidence = "\n".join(p3.evidence)
        for expected in ("fixability=auto_fix_candidate", "recommended=", "candidates="):
            if expected not in evidence:
                failures.append(f"P3 fixability evidence missing {expected!r}: {evidence}")

    if failures:
        print("FAIL:")
        for failure in failures:
            print(f"- {failure}")
        return 1

    print("OK: priority evidence prefers schema detail and preserves transition detail fields")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
