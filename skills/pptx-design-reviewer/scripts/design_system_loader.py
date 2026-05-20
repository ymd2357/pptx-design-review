"""Design-system palette loader for pptx_lint.

Reads `doc/slide-guideline-v1.yml` → `rules.color.lint_palette` and emits
the same 5 structures that pptx_lint.py defines as module-level
constants (ALLOWED_TEXT_COLORS_HEX, ALLOWED_FILL_COLORS_HEX,
TEXT_COLOR_TOKEN_BY_HEX, FILL_COLOR_TOKEN_BY_HEX,
CONTRAST_REPAIR_COLOR_FAMILIES, FILL_REPAIR_COLOR_FAMILIES).

Called unconditionally at pptx_lint.py import time (DS-COLOR-001 was
completed on 2026-05-20: the Python-side hardcoded duplicates have been
removed, so this YAML file is the single source of truth for the lint
color palette).
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import yaml

REPO_ROOT = Path(__file__).resolve().parents[3]
DEFAULT_GUIDELINE_PATH = REPO_ROOT / "doc" / "slide-guideline-v1.yml"


@dataclass(frozen=True)
class LintPalette:
    allowed_text_colors_hex: frozenset[str]
    allowed_fill_colors_hex: frozenset[str]
    text_color_token_by_hex: dict[str, str]
    fill_color_token_by_hex: dict[str, str]
    contrast_repair_color_families: tuple[tuple[str, tuple[str, ...]], ...]
    fill_repair_color_families: tuple[tuple[str, tuple[str, ...]], ...]


def _normalize_hex(value: str) -> str:
    text = str(value).strip().upper()
    if not text.startswith("#"):
        text = "#" + text
    return text


def _families_from_yaml(node: dict) -> tuple[tuple[str, tuple[str, ...]], ...]:
    items: list[tuple[str, tuple[str, ...]]] = []
    for name, colors in node.items():
        if not isinstance(colors, list):
            raise ValueError(f"lint_palette family '{name}' must be a list of hex strings")
        items.append((name, tuple(_normalize_hex(c) for c in colors)))
    return tuple(items)


def load_lint_palette(path: Path | str | None = None) -> LintPalette:
    yaml_path = Path(path) if path is not None else DEFAULT_GUIDELINE_PATH
    with yaml_path.open("r", encoding="utf-8") as fh:
        data = yaml.safe_load(fh)
    try:
        node = data["rules"]["color"]["lint_palette"]
    except (KeyError, TypeError) as exc:
        raise ValueError(
            f"{yaml_path}: rules.color.lint_palette is missing"
        ) from exc

    text_token = {_normalize_hex(k): str(v) for k, v in (node.get("text_color_token_by_hex") or {}).items()}
    fill_token = {_normalize_hex(k): str(v) for k, v in (node.get("fill_color_token_by_hex") or {}).items()}

    return LintPalette(
        allowed_text_colors_hex=frozenset(_normalize_hex(c) for c in (node.get("allowed_text_colors_hex") or [])),
        allowed_fill_colors_hex=frozenset(_normalize_hex(c) for c in (node.get("allowed_fill_colors_hex") or [])),
        text_color_token_by_hex=text_token,
        fill_color_token_by_hex=fill_token,
        contrast_repair_color_families=_families_from_yaml(node.get("contrast_repair_color_families") or {}),
        fill_repair_color_families=_families_from_yaml(node.get("fill_repair_color_families") or {}),
    )
