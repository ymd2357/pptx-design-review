#!/usr/bin/env python3
"""Lint the design-system guideline YAML itself (DS-003).

Detects structural drift between the palette and the rules that consume
it, before silent inconsistencies turn into bad lint output for end
users.

Implemented checks:
  - palette_family_tier_gap: brand.{primary, black, secondary} are
    documented as 10-tier scales (50, 100, ..., 900). Flag any missing
    tier so the palette stays gap-free.
  - repair_candidate_family_overlap: the same hex appearing in two or
    more contrast repair families would make hue-preserving selection
    ambiguous.
  - repair_candidate_palette_drift: every (token, hex) under
    `rules.color.repair_candidates.text.hue_preserving_low_contrast.families`
    must resolve against `rules.color.palette` (token exists, hex matches).

Usage:
    python3 design_system_review.py [--yaml doc/slide-guideline-v1.yml] [--json]
"""

from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

REPO_ROOT = Path(__file__).resolve().parents[3]
DEFAULT_GUIDELINE_PATH = REPO_ROOT / "doc" / "slide-guideline-v1.yml"

EXPECTED_BRAND_TIERS = (50, 100, 200, 300, 400, 500, 600, 700, 800, 900)
BRAND_FAMILY_PREFIX = {"primary": "p", "black": "b", "secondary": "s"}

# Token namespace prefix that wraps every allowlist reference.
TOKEN_NAMESPACE_PREFIX = "tokens.semantic.color."


@dataclass
class Finding:
    check: str
    severity: str  # "warn" | "error"
    message: str
    evidence: dict = field(default_factory=dict)
    manual_required_reason: str = ""

    def to_json(self) -> dict:
        return {
            "check_id": self.check,
            "check": self.check,
            "severity": self.severity,
            "message": self.message,
            "evidence": dict(self.evidence),
            "fixability": "manual_required",
            "fixability_reason": "design_system_yaml_edit",
            "candidate_values": {},
            "manual_required_reason": self.manual_required_reason or self.message,
            "measurement_confidence": "high",
        }


def _walk_palette(prefix: tuple[str, ...], node: Any, out: dict[str, str]) -> None:
    if isinstance(node, dict):
        for key, value in node.items():
            if key == "usage_policy":
                continue
            child_prefix = prefix + (key,) if prefix or key != "palette" else (key,)
            _walk_palette(child_prefix, value, out)
    elif isinstance(node, str) and node.startswith("#"):
        out[".".join(prefix)] = node.upper()


def _check_palette_family_tier_gap(palette: dict) -> list[Finding]:
    findings: list[Finding] = []
    brand = palette.get("brand", {})
    for family_name, prefix in BRAND_FAMILY_PREFIX.items():
        family = brand.get(family_name) or {}
        present: set[int] = set()
        for key in family.keys():
            if not isinstance(key, str) or not key.startswith(prefix):
                continue
            try:
                present.add(int(key[len(prefix):]))
            except ValueError:
                continue
        missing = [t for t in EXPECTED_BRAND_TIERS if t not in present]
        if missing:
            findings.append(
                Finding(
                    check="palette_family_tier_gap",
                    severity="warn",
                    message=(
                        f"brand.{family_name} is missing tier(s) "
                        f"{missing} (expected the full 10-step scale)"
                    ),
                    evidence={
                        "family": f"brand.{family_name}",
                        "expected_tiers": list(EXPECTED_BRAND_TIERS),
                        "missing_tiers": missing,
                        "present_tiers": sorted(present),
                    },
                    manual_required_reason=(
                        "Add the missing tier hex value(s) to the palette so "
                        "downstream lint constants do not silently fall back."
                    ),
                )
            )
    return findings


def _check_repair_candidate_family_overlap(repair_families: dict) -> list[Finding]:
    findings: list[Finding] = []
    hex_to_families: dict[str, set[str]] = {}
    for family_name, candidates in repair_families.items():
        if not isinstance(candidates, list):
            continue
        for entry in candidates:
            if not isinstance(entry, dict):
                continue
            hex_value = str(entry.get("hex", "")).upper()
            if hex_value:
                hex_to_families.setdefault(hex_value, set()).add(family_name)
    for hex_value, families in hex_to_families.items():
        if len(families) > 1:
            findings.append(
                Finding(
                    check="repair_candidate_family_overlap",
                    severity="error",
                    message=(
                        f"hex {hex_value} appears in multiple contrast repair "
                        f"families {sorted(families)}; family selection would "
                        f"be ambiguous"
                    ),
                    evidence={
                        "hex": hex_value,
                        "families": sorted(families),
                    },
                    manual_required_reason=(
                        "Keep each hex in exactly one repair family. If the "
                        "color genuinely belongs to two hue groups, split it "
                        "into separate hexes or remove from the secondary one."
                    ),
                )
            )
    return findings


def _check_repair_candidate_palette_drift(
    repair_families: dict, palette_tokens: dict[str, str]
) -> list[Finding]:
    findings: list[Finding] = []
    for family_name, candidates in repair_families.items():
        if not isinstance(candidates, list):
            continue
        for entry in candidates:
            if not isinstance(entry, dict):
                continue
            token = str(entry.get("token", ""))
            hex_value = str(entry.get("hex", "")).upper()
            if not token:
                continue
            palette_hex = palette_tokens.get(token)
            if palette_hex is None:
                findings.append(
                    Finding(
                        check="repair_candidate_palette_drift",
                        severity="error",
                        message=(
                            f"repair candidate token '{token}' "
                            f"(family={family_name}) does not exist in the palette"
                        ),
                        evidence={
                            "family": family_name,
                            "token": token,
                            "declared_hex": hex_value,
                            "issue": "token_not_in_palette",
                        },
                        manual_required_reason=(
                            "Add the token to rules.color.palette or correct the "
                            "repair candidate to reference an existing token."
                        ),
                    )
                )
                continue
            if hex_value and palette_hex != hex_value:
                findings.append(
                    Finding(
                        check="repair_candidate_palette_drift",
                        severity="error",
                        message=(
                            f"repair candidate token '{token}' hex "
                            f"{hex_value} does not match palette hex {palette_hex}"
                        ),
                        evidence={
                            "family": family_name,
                            "token": token,
                            "declared_hex": hex_value,
                            "palette_hex": palette_hex,
                            "issue": "hex_mismatch",
                        },
                        manual_required_reason=(
                            "Update either the palette entry or the repair "
                            "candidate so both reference the same hex."
                        ),
                    )
                )
    return findings


def _collect_allowlist_token_refs(node: Any, out: set[str]) -> None:
    """Walk the allowlist subtree and collect every `tokens.semantic.color.*`
    reference (both exact tokens and family prefixes), stripped of the
    namespace prefix so they line up with palette token paths.
    """
    if isinstance(node, dict):
        for value in node.values():
            _collect_allowlist_token_refs(value, out)
    elif isinstance(node, list):
        for item in node:
            _collect_allowlist_token_refs(item, out)
    elif isinstance(node, str) and node.startswith(TOKEN_NAMESPACE_PREFIX):
        out.add(node[len(TOKEN_NAMESPACE_PREFIX):])


def _token_is_covered(token: str, refs: set[str]) -> str | None:
    """Return the matching reference (exact or family prefix) or None."""
    parts = token.split(".")
    for i in range(len(parts), 0, -1):
        candidate = ".".join(parts[:i])
        if candidate in refs:
            return candidate
    return None


def _check_palette_token_usage_coverage(
    palette_tokens: dict[str, str], allowlist: dict
) -> list[Finding]:
    """Every palette token should be reachable via the allowlist either
    by exact match or by a family prefix (e.g. `brand.primary` covers
    `brand.primary.p500`). Bare `data_series.<index>` indices are skipped
    because they are list positions, not allowlist-style tokens.
    """
    findings: list[Finding] = []
    refs: set[str] = set()
    _collect_allowlist_token_refs(allowlist or {}, refs)
    for token, hex_value in sorted(palette_tokens.items()):
        if token.startswith("data_series."):
            continue
        if _token_is_covered(token, refs) is None:
            findings.append(
                Finding(
                    check="palette_token_usage_uncovered",
                    severity="warn",
                    message=(
                        f"palette token '{token}' ({hex_value}) has no "
                        f"allowlist coverage; downstream consumers cannot tell "
                        f"where this color is permitted"
                    ),
                    evidence={
                        "token": token,
                        "hex": hex_value,
                    },
                    manual_required_reason=(
                        "Either add an allowlist entry referencing this token "
                        "(or a covering family prefix), or remove the color from "
                        "the palette if it is unused."
                    ),
                )
            )
    return findings


def _check_contrast_pair_coverage(allowlist: dict) -> list[Finding]:
    """Every (text_hex, background_hex) pair built from allowed text
    grays and allowed background surfaces should be classified — either
    listed under `text_on_background` with a contrast ratio, or under
    `prohibited.text_on_background`. Pairs that appear in neither list
    are silent gaps where future drift could go undetected.
    """
    findings: list[Finding] = []
    grays_text = (allowlist.get("grays") or {}).get("text") or []
    backgrounds = (allowlist.get("backgrounds") or {})
    bg_entries: list[dict] = []
    for category in ("page", "surface"):
        for entry in backgrounds.get(category) or []:
            if isinstance(entry, dict):
                bg_entries.append({"category": category, **entry})

    classified_pairs: set[tuple[str, str]] = set()
    for entry in (allowlist.get("text_on_background") or []):
        if isinstance(entry, dict):
            t = str(entry.get("text_hex", "")).upper()
            b = str(entry.get("background_hex", "")).upper()
            if t and b:
                classified_pairs.add((t, b))
    prohibited_roles: set[str] = set()
    for entry in (
        (allowlist.get("prohibited") or {}).get("text_on_background") or []
    ):
        if isinstance(entry, dict) and entry.get("role"):
            prohibited_roles.add(str(entry["role"]))

    for text_entry in grays_text:
        text_hex = str(text_entry.get("hex", "")).upper()
        text_role = str(text_entry.get("role", ""))
        if not text_hex:
            continue
        for bg_entry in bg_entries:
            bg_hex = str(bg_entry.get("hex", "")).upper()
            bg_role = str(bg_entry.get("role", ""))
            if not bg_hex:
                continue
            if (text_hex, bg_hex) in classified_pairs:
                continue
            # Check prohibited pairs by role naming convention
            # (`<text_role>_on_<bg_role>` or `<text_role>_on_<bg_category>`).
            candidate_roles = {
                f"{text_role}_on_{bg_role}",
                f"{text_role}_on_{bg_entry.get('category', '')}",
            }
            if candidate_roles & prohibited_roles:
                continue
            findings.append(
                Finding(
                    check="contrast_pair_unclassified",
                    severity="warn",
                    message=(
                        f"text {text_role} ({text_hex}) on background "
                        f"{bg_role} ({bg_hex}) is in neither text_on_background "
                        f"nor prohibited.text_on_background"
                    ),
                    evidence={
                        "text_role": text_role,
                        "text_hex": text_hex,
                        "background_role": bg_role,
                        "background_hex": bg_hex,
                        "background_category": bg_entry.get("category"),
                    },
                    manual_required_reason=(
                        "Add this pair to allowlist.text_on_background with "
                        "its measured contrast ratio, or document the gap in "
                        "allowlist.prohibited.text_on_background."
                    ),
                )
            )
    return findings


def review_design_system(yaml_path: Path | str | None = None) -> list[Finding]:
    path = Path(yaml_path) if yaml_path is not None else DEFAULT_GUIDELINE_PATH
    with path.open("r", encoding="utf-8") as fh:
        data = yaml.safe_load(fh)

    color_rules = data.get("rules", {}).get("color", {}) or {}
    palette = color_rules.get("palette", {}) or {}
    allowlist = color_rules.get("allowlist", {}) or {}
    palette_tokens: dict[str, str] = {}
    _walk_palette((), palette, palette_tokens)

    repair_families = (
        color_rules.get("repair_candidates", {})
        .get("text", {})
        .get("hue_preserving_low_contrast", {})
        .get("families", {})
        or {}
    )

    findings: list[Finding] = []
    findings.extend(_check_palette_family_tier_gap(palette))
    findings.extend(_check_repair_candidate_family_overlap(repair_families))
    findings.extend(_check_repair_candidate_palette_drift(repair_families, palette_tokens))
    findings.extend(_check_palette_token_usage_coverage(palette_tokens, allowlist))
    findings.extend(_check_contrast_pair_coverage(allowlist))
    return findings


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--yaml", default=str(DEFAULT_GUIDELINE_PATH))
    parser.add_argument("--json", action="store_true", help="Emit findings as JSON")
    args = parser.parse_args()

    findings = review_design_system(args.yaml)

    if args.json:
        json.dump(
            {"findings": [f.to_json() for f in findings]},
            sys.stdout,
            ensure_ascii=False,
            indent=2,
        )
        sys.stdout.write("\n")
    else:
        if not findings:
            print(f"OK: design-system review found no issues in {args.yaml}")
        else:
            print(f"{len(findings)} issue(s) found in {args.yaml}:")
            for f in findings:
                print(f"  [{f.severity}] {f.check}: {f.message}")

    return 1 if any(f.severity == "error" for f in findings) else 0


if __name__ == "__main__":
    sys.exit(main())
