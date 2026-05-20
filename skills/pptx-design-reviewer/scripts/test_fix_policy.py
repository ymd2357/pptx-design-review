#!/usr/bin/env python3
"""Verify that doc/slide-guideline-v1.yml `rules.lint.fix_policy` stays the
single source of truth for which lint check maps to which pptx_fix rule.

This is the policy-side counterpart of test_design_system_review.py: the
guideline YAML declares apply_mode (auto_fix / judgement_fix / no_fix) and
fix_rule, and we assert that the implementation (pptx_lint.py and
pptx_fix.py) does not silently drift from it.

段階 1 (本テスト): YAML と現実装の対応関係を assert する。コードを refactor
する段階 2 で「implementation reads YAML」に切り替わったときも、本テストは
そのままで意味を持つ (= 仕様文書としての YAML から逸脱しないことの保証)。
"""

from __future__ import annotations

import sys
from pathlib import Path

import yaml

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import pptx_fix  # noqa: E402

REPO_ROOT = HERE.parents[2]
GUIDELINE_PATH = REPO_ROOT / "doc" / "slide-guideline-v1.yml"


def load_lint_section() -> dict:
    data = yaml.safe_load(GUIDELINE_PATH.read_text(encoding="utf-8"))
    return data["rules"]["lint"]


def main() -> int:
    failures: list[str] = []
    lint = load_lint_section()
    fix_policy = lint.get("fix_policy") or {}
    lint_checks = lint.get("checks") or {}

    if not fix_policy:
        failures.append("rules.lint.fix_policy is missing from doc/slide-guideline-v1.yml")
        print("\n".join(failures))
        return 1

    # 1. Every check declared under rules.lint.checks must have a fix_policy entry.
    missing_policy = sorted(set(lint_checks) - set(fix_policy))
    extra_policy = sorted(set(fix_policy) - set(lint_checks))
    if missing_policy:
        failures.append(
            f"fix_policy is missing entries for rules.lint.checks: {missing_policy}"
        )
    if extra_policy:
        failures.append(
            f"fix_policy declares checks that are not in rules.lint.checks: {extra_policy}"
        )

    # 2. apply_mode must be one of the documented enum values.
    allowed_modes = {"auto_fix", "judgement_fix", "no_fix"}
    for check, entry in fix_policy.items():
        if not isinstance(entry, dict):
            failures.append(f"fix_policy.{check} is not a mapping")
            continue
        mode = entry.get("apply_mode")
        if mode not in allowed_modes:
            failures.append(
                f"fix_policy.{check}.apply_mode={mode!r} is not in {sorted(allowed_modes)}"
            )

    # 3. apply_mode=no_fix → fix_rule must be null, AND check must NOT appear
    #    in pptx_fix.CHECK_TO_RULE.
    # 4. apply_mode in {auto_fix, judgement_fix} → fix_rule must be a string,
    #    AND pptx_fix.CHECK_TO_RULE[check] must equal that fix_rule.
    check_to_rule = pptx_fix.CHECK_TO_RULE
    for check, entry in fix_policy.items():
        if not isinstance(entry, dict):
            continue
        mode = entry.get("apply_mode")
        fix_rule = entry.get("fix_rule")
        if mode == "no_fix":
            if fix_rule is not None:
                failures.append(
                    f"fix_policy.{check}.apply_mode=no_fix but fix_rule={fix_rule!r} (expected null)"
                )
            if check in check_to_rule:
                failures.append(
                    f"fix_policy.{check}.apply_mode=no_fix but pptx_fix.CHECK_TO_RULE[{check!r}]"
                    f"={check_to_rule[check]!r} (expected the check to be absent)"
                )
            continue
        if not isinstance(fix_rule, str):
            failures.append(
                f"fix_policy.{check}.apply_mode={mode!r} requires a string fix_rule; got {fix_rule!r}"
            )
            continue
        actual_rule = check_to_rule.get(check)
        if actual_rule != fix_rule:
            failures.append(
                f"fix_policy.{check}.fix_rule={fix_rule!r} but pptx_fix.CHECK_TO_RULE[{check!r}]={actual_rule!r}"
            )

    # 5. Every pptx_fix.CHECK_TO_RULE key must be covered by fix_policy with
    #    a non-`no_fix` apply_mode. Catches the reverse drift (rule registered
    #    in pptx_fix.py but no policy declared in YAML).
    for check in sorted(check_to_rule):
        entry = fix_policy.get(check)
        if entry is None:
            failures.append(
                f"pptx_fix.CHECK_TO_RULE has {check!r} but fix_policy has no entry"
            )
            continue
        if entry.get("apply_mode") == "no_fix":
            failures.append(
                f"pptx_fix.CHECK_TO_RULE has {check!r} but fix_policy.{check}.apply_mode=no_fix"
            )

    if failures:
        print("FAIL:")
        for f in failures:
            print(f"- {f}")
        return 1

    auto_fix = sum(1 for e in fix_policy.values() if e.get("apply_mode") == "auto_fix")
    judgement_fix = sum(1 for e in fix_policy.values() if e.get("apply_mode") == "judgement_fix")
    no_fix = sum(1 for e in fix_policy.values() if e.get("apply_mode") == "no_fix")
    print(
        f"OK: fix_policy sync ({len(fix_policy)} checks: "
        f"auto_fix={auto_fix} judgement_fix={judgement_fix} no_fix={no_fix})"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
