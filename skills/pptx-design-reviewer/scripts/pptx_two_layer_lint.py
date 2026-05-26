#!/usr/bin/env python3
"""Run raw design lint plus normalized measurement lint for a PPTX deck.

The raw deck is the design-system source of evidence. The normalized derived
deck exists only for checks whose result depends on viewer/text-width font
resolution.
"""

from __future__ import annotations

import argparse
import json
import shutil
import sys
from dataclasses import asdict
from pathlib import Path
from typing import Any

import pptx_fix
import pptx_lint
from pptx_normalize_fonts import normalize_pptx


MEASUREMENT_SENSITIVE_CHECKS = frozenset(
    {
        "text_box_overflow",
        "text_canvas_overflow",
        "wrap_break_changes_meaning",
        "text_vertical_balance",
    }
)


def _lint_json(
    path: Path,
    *,
    profile: str,
    rendered_image_dir: Path | None,
    consolidate: bool,
    min_recurring_slides: int,
    layer: str,
    measurement_source: str,
) -> list[dict[str, Any]]:
    findings = pptx_lint.lint_pptx(
        path,
        profile=profile,
        rendered_image_dir=rendered_image_dir,
    )
    if consolidate:
        findings = pptx_lint.consolidate_recurring(
            findings,
            min_slides=min_recurring_slides,
        )
    return [
        _with_layer_metadata(
            pptx_lint.finding_to_json_dict(finding),
            layer=layer,
            measurement_source=measurement_source,
        )
        for finding in findings
    ]


def _with_layer_metadata(
    finding: dict[str, Any],
    *,
    layer: str,
    measurement_source: str,
) -> dict[str, Any]:
    detail = dict(finding.get("detail") or {})
    detail["layer"] = layer
    detail["measurement_source"] = measurement_source
    finding = dict(finding)
    finding["layer"] = layer
    finding["measurement_source"] = measurement_source
    finding["detail"] = detail
    return finding


def _write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def build_artifacts(
    pptx_path: Path,
    out_dir: Path,
    *,
    profile: str = "default",
    rendered_image_dir: Path | None = None,
    consolidate: bool = True,
    min_recurring_slides: int = 3,
    verify_fix: bool = False,
    judgement_gate: bool = True,
) -> dict[str, Any]:
    if not pptx_path.is_file():
        raise FileNotFoundError(f"input not found: {pptx_path}")

    out_dir.mkdir(parents=True, exist_ok=True)
    normalized_pptx = out_dir / f"{pptx_path.stem}.normalized.pptx"
    fixed_pptx = out_dir / f"{pptx_path.stem}.normalized.fixed.pptx"
    raw_lint_json = out_dir / "raw-design-lint.json"
    normalized_lint_json = out_dir / "normalized-measurement-lint.json"
    fix_actions_json = out_dir / "normalized-fix-actions.json"
    fixed_lint_json = out_dir / "normalized-fixed-lint.json"
    combined_json = out_dir / "review-artifact.json"

    raw_findings = _lint_json(
        pptx_path,
        profile=profile,
        rendered_image_dir=None,
        consolidate=consolidate,
        min_recurring_slides=min_recurring_slides,
        layer="raw_design",
        measurement_source="raw_pptx",
    )
    normalization_report = normalize_pptx(pptx_path, normalized_pptx)
    normalized_findings = _lint_json(
        normalized_pptx,
        profile=profile,
        rendered_image_dir=rendered_image_dir,
        consolidate=consolidate,
        min_recurring_slides=min_recurring_slides,
        layer="normalized_measurement",
        measurement_source="normalized_pptx",
    )

    review_findings = [
        finding
        for finding in raw_findings
        if finding.get("check") not in MEASUREMENT_SENSITIVE_CHECKS
    ]
    review_findings.extend(
        finding
        for finding in normalized_findings
        if finding.get("check") in MEASUREMENT_SENSITIVE_CHECKS
    )

    _write_json(raw_lint_json, raw_findings)
    _write_json(normalized_lint_json, normalized_findings)

    fix_verification: dict[str, Any] = {
        "enabled": verify_fix,
        "fixed_pptx": None,
        "fix_actions_json": None,
        "fixed_measurement_lint_json": None,
        "actions": [],
        "fixed_findings": [],
    }
    if verify_fix:
        shutil.copy2(normalized_pptx, fixed_pptx)
        rules = list(pptx_fix.auto_rules_from_findings(review_findings))
        actions = pptx_fix.fix_pptx(
            fixed_pptx,
            apply=True,
            rules=rules,
            findings=review_findings,
            judgement_gate=judgement_gate,
        )
        fixed_findings = _lint_json(
            fixed_pptx,
            profile=profile,
            rendered_image_dir=rendered_image_dir,
            consolidate=consolidate,
            min_recurring_slides=min_recurring_slides,
            layer="normalized_fix_verification",
            measurement_source="normalized_fixed_pptx",
        )
        actions_payload = {
            "source_findings": "review_findings",
            "rules": rules,
            "applied": True,
            "actions": [asdict(action) for action in actions],
        }
        _write_json(fix_actions_json, actions_payload)
        _write_json(fixed_lint_json, fixed_findings)
        fix_verification = {
            "enabled": True,
            "fixed_pptx": str(fixed_pptx),
            "fix_actions_json": str(fix_actions_json),
            "fixed_measurement_lint_json": str(fixed_lint_json),
            "rules": rules,
            "actions": actions_payload["actions"],
            "fixed_findings": fixed_findings,
        }

    artifact = {
        "artifact_version": 1,
        "profile": profile,
        "sources": {
            "input_pptx": str(pptx_path),
            "normalized_pptx": str(normalized_pptx),
            "normalized_fixed_pptx": str(fixed_pptx) if verify_fix else None,
            "raw_design_lint_json": str(raw_lint_json),
            "normalized_measurement_lint_json": str(normalized_lint_json),
            "normalized_fix_actions_json": str(fix_actions_json) if verify_fix else None,
            "normalized_fixed_lint_json": str(fixed_lint_json) if verify_fix else None,
            "combined_json": str(combined_json),
            "rendered_image_dir": str(rendered_image_dir) if rendered_image_dir else None,
        },
        "normalization": {
            **asdict(normalization_report),
            "purpose": "derived deck for viewer/text-width measurement only",
        },
        "measurement_sensitive_checks": sorted(MEASUREMENT_SENSITIVE_CHECKS),
        "layers": {
            "raw_design": {
                "source_path": str(pptx_path),
                "lint_json_path": str(raw_lint_json),
                "findings": raw_findings,
            },
            "normalized_measurement": {
                "source_path": str(normalized_pptx),
                "lint_json_path": str(normalized_lint_json),
                "findings": normalized_findings,
            },
        },
        "raw_findings": raw_findings,
        "normalized_findings": normalized_findings,
        "review_findings": review_findings,
        "findings": review_findings,
        "fix_verification": fix_verification,
        "merge_policy": {
            "raw_design_checks": "all checks except measurement_sensitive_checks",
            "normalized_measurement_checks": "measurement_sensitive_checks only",
            "fix_verification": (
                "optional; applies review_findings to normalized_pptx and re-lints "
                "the normalized fixed deck"
            ),
        },
    }
    _write_json(combined_json, artifact)
    return artifact


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Run raw + normalized PPTX lint pipeline")
    parser.add_argument("pptx", type=Path, help="input .pptx")
    parser.add_argument("--out-dir", type=Path, required=True, help="artifact output directory")
    parser.add_argument(
        "--profile",
        choices=sorted(pptx_lint.LINT_PROFILES),
        default="default",
        help="lint policy profile (default: default)",
    )
    parser.add_argument(
        "--rendered-image-dir",
        type=Path,
        help="rendered slide PNG directory for normalized measurement lint",
    )
    parser.add_argument(
        "--no-consolidate",
        action="store_true",
        help="disable deck-level consolidation of recurring identical findings",
    )
    parser.add_argument(
        "--min-recurring-slides",
        type=int,
        default=3,
        help="threshold for consolidating recurring findings (default: 3)",
    )
    parser.add_argument(
        "--verify-fix",
        action="store_true",
        help=(
            "copy the normalized deck, apply pptx_fix using review_findings, "
            "and write normalized fix actions plus fixed-deck lint JSON"
        ),
    )
    parser.add_argument(
        "--no-judgement-gate",
        action="store_true",
        help="pass through to pptx_fix when --verify-fix is enabled",
    )
    args = parser.parse_args(argv)

    try:
        artifact = build_artifacts(
            args.pptx,
            args.out_dir,
            profile=args.profile,
            rendered_image_dir=args.rendered_image_dir,
            consolidate=not args.no_consolidate,
            min_recurring_slides=args.min_recurring_slides,
            verify_fix=args.verify_fix,
            judgement_gate=not args.no_judgement_gate,
        )
    except FileNotFoundError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    print(json.dumps(artifact["sources"], ensure_ascii=False, indent=2))
    return 1 if any(f.get("severity") == "error" for f in artifact["review_findings"]) else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
