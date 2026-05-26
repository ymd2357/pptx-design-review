#!/usr/bin/env python3
"""Publish existing review artifacts into a git-trackable static snapshot.

Happy path:
    python3 scripts/publish_review_snapshot.py --deck 260329-seminar-curriculum-proposal --rev 017

With two-layer font pipeline output:
    python3 scripts/publish_review_snapshot.py --deck mcp-cource --rev 007 \
      --lint-json tmp/review/mcp-cource/font-pipeline-rev-007/review-artifact.json

With before/after comparison images:
    python3 scripts/publish_review_snapshot.py --deck 260329-seminar-curriculum-proposal --rev 034 \
      --before-images tmp/review/260329-seminar-curriculum-proposal/font-pipeline-rev-034/render-normalized \
      --after-images tmp/review/260329-seminar-curriculum-proposal/font-pipeline-rev-034/render-fixed \
      --lint-json tmp/review/260329-seminar-curriculum-proposal/font-pipeline-rev-034/review-artifact.json

The script does not render or lint a deck. It rebuilds
tmp/review-snapshot/<deck>/rev-<NNN>/ from files already present under
tmp/review/<deck>/.
"""

from __future__ import annotations

import argparse
import json
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable


REPO_ROOT = Path(__file__).resolve().parents[1]
REVIEW_ROOT = REPO_ROOT / "tmp" / "review"
SNAPSHOT_ROOT = REPO_ROOT / "tmp" / "review-snapshot"
SLIDE_RE = re.compile(r"slide[-_ ]?0*(\d+)", re.IGNORECASE)


@dataclass(frozen=True)
class Candidate:
    path: Path
    score: int
    mtime: float


def main() -> int:
    args = parse_args()
    rev = normalize_rev(args.rev)
    source_dir = REVIEW_ROOT / args.deck
    if not source_dir.is_dir():
        raise SystemExit(f"Source deck directory not found: {source_dir}")
    if (args.before_images is None) != (args.after_images is None):
        raise SystemExit("--before-images and --after-images must be provided together")

    image_dir = select_image_dir(source_dir, rev)
    lint_json = args.lint_json or select_json(source_dir, rev, kind="lint")
    priorities_json = args.priorities_json or select_json(source_dir, rev, kind="priorities")

    if args.before_images is None and image_dir is None:
        raise SystemExit(f"No slide PNG directory found for rev-{rev} under {source_dir}")
    if lint_json is None:
        raise SystemExit(f"No lint JSON found for rev-{rev} under {source_dir}")
    if not lint_json.is_file():
        raise SystemExit(f"Lint JSON not found: {lint_json}")
    if priorities_json is not None and not priorities_json.is_file():
        raise SystemExit(f"Priorities JSON not found: {priorities_json}")

    output_dir = SNAPSHOT_ROOT / args.deck / f"rev-{rev}"
    if output_dir.exists():
        shutil.rmtree(output_dir)

    if args.before_images is not None and args.after_images is not None:
        before_count = copy_named_slide_images(args.before_images, output_dir / "images" / "before")
        after_count = copy_named_slide_images(args.after_images, output_dir / "images" / "after")
        if before_count == 0:
            raise SystemExit(f"No slide PNG files found in before image directory: {args.before_images}")
        if after_count == 0:
            raise SystemExit(f"No slide PNG files found in after image directory: {args.after_images}")
        copied_message = (
            f"before={before_count} from {repo_display_path(args.before_images)}, "
            f"after={after_count} from {repo_display_path(args.after_images)}"
        )
    else:
        assert image_dir is not None
        copied = copy_named_slide_images(image_dir, output_dir / "images")
        if copied == 0:
            raise SystemExit(f"No slide PNG files found in selected image directory: {image_dir}")
        copied_message = f"{copied} from {repo_display_path(image_dir)}"

    lint_payload = rewrite_lint_json(lint_json, output_dir / "lint.json")
    if isinstance(lint_payload, dict):
        rewrite_json(lint_json, output_dir / "review-artifact.json")
    if priorities_json is not None:
        rewrite_json(priorities_json, output_dir / "priorities.json")

    print(f"snapshot: {repo_display_path(output_dir)}")
    print(f"images: {copied_message}")
    print(f"lint: {repo_display_path(lint_json)}")
    if priorities_json is not None:
        print(f"priorities: {repo_display_path(priorities_json)}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Copy existing tmp/review artifacts into tmp/review-snapshot for Pages.",
    )
    parser.add_argument("--deck", required=True, help="Deck id under tmp/review/<deck>.")
    parser.add_argument("--rev", required=True, help="Revision number, for example 017 or rev-017.")
    parser.add_argument(
        "--lint-json",
        type=Path,
        help=(
            "Explicit lint JSON path. If the file is a two-layer review-artifact, "
            "its review_findings/findings array is written as lint.json and the "
            "full artifact is copied to review-artifact.json."
        ),
    )
    parser.add_argument(
        "--priorities-json",
        type=Path,
        help="Explicit priorities JSON path.",
    )
    parser.add_argument(
        "--before-images",
        type=Path,
        help="Explicit before slide PNG directory for compare snapshots.",
    )
    parser.add_argument(
        "--after-images",
        type=Path,
        help="Explicit after slide PNG directory for compare snapshots.",
    )
    return parser.parse_args()


def normalize_rev(value: str) -> str:
    match = re.fullmatch(r"(?:rev-)?(\d+)", value.strip(), re.IGNORECASE)
    if not match:
        raise SystemExit("--rev must be a number such as 017 or rev-017")
    return match.group(1).zfill(3)


def select_image_dir(source_dir: Path, rev: str) -> Path | None:
    candidates: list[Candidate] = []
    for directory in source_dir.rglob("*"):
        if not directory.is_dir() or any(part.startswith("_") for part in directory.parts):
            continue
        slide_pngs = list(iter_slide_pngs(directory))
        if not slide_pngs:
            continue
        rel = directory.relative_to(source_dir).as_posix().lower()
        score = 0
        if f"rev-{rev}" in rel:
            score += 100
        if "render" in rel:
            score += 40
        if "review-images" in rel or "images" in rel:
            score += 25
        if directory.name == "after":
            score += 20
        if "annotated" in rel:
            score += 5
        if directory.name == "before" or "diff" in rel:
            score -= 50
        candidates.append(Candidate(directory, score, newest_mtime(slide_pngs)))
    return best(candidates)


def select_json(source_dir: Path, rev: str, kind: str) -> Path | None:
    candidates: list[Candidate] = []
    for path in source_dir.rglob(f"rev-{rev}-*.json"):
        if any(part.startswith("_") for part in path.parts):
            continue
        name = path.name.lower()
        if kind == "lint" and "lint" not in name:
            continue
        if kind == "lint" and "priorit" in name:
            continue
        if kind == "priorities" and "priorit" not in name:
            continue
        score = 100
        if "rendered" in name:
            score += 20
        if "after" in name:
            score += 10
        if "unconsolidated" in name:
            score -= 30
        candidates.append(Candidate(path, score, path.stat().st_mtime))
    return best(candidates)


def best(candidates: Iterable[Candidate]) -> Path | None:
    ordered = sorted(candidates, key=lambda item: (item.score, item.mtime, item.path.as_posix()))
    return ordered[-1].path if ordered else None


def iter_slide_pngs(directory: Path) -> Iterable[Path]:
    for path in directory.glob("*.png"):
        if SLIDE_RE.search(path.name):
            yield path


def newest_mtime(paths: Iterable[Path]) -> float:
    return max(path.stat().st_mtime for path in paths)


def copy_named_slide_images(source_dir: Path, target_dir: Path) -> int:
    if not source_dir.is_dir():
        raise SystemExit(f"Slide PNG directory not found: {source_dir}")
    target_dir.mkdir(parents=True, exist_ok=True)
    copied = 0
    seen: set[int] = set()
    for path in sorted(iter_slide_pngs(source_dir), key=slide_sort_key):
        match = SLIDE_RE.search(path.name)
        if not match:
            continue
        slide_no = int(match.group(1))
        if slide_no in seen:
            continue
        seen.add(slide_no)
        shutil.copy2(path, target_dir / f"slide-{slide_no:02d}.png")
        copied += 1
    return copied


def slide_sort_key(path: Path) -> tuple[int, str]:
    match = SLIDE_RE.search(path.name)
    return (int(match.group(1)) if match else 10_000, path.name)


def rewrite_json(source: Path, target: Path) -> None:
    with source.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    with target.open("w", encoding="utf-8") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.write("\n")


def rewrite_lint_json(source: Path, target: Path) -> object:
    with source.open("r", encoding="utf-8") as handle:
        data = json.load(handle)
    if isinstance(data, dict):
        findings = data.get("review_findings")
        if findings is None:
            findings = data.get("findings")
        if not isinstance(findings, list):
            raise SystemExit(
                f"Lint artifact object must contain review_findings[] or findings[]: {source}"
            )
        output = findings
    elif isinstance(data, list):
        output = data
    else:
        raise SystemExit(f"Lint JSON must be a findings array or artifact object: {source}")
    with target.open("w", encoding="utf-8") as handle:
        json.dump(output, handle, ensure_ascii=False, indent=2)
        handle.write("\n")
    return data


def repo_display_path(path: Path) -> Path:
    resolved = path.resolve()
    try:
        return resolved.relative_to(REPO_ROOT)
    except ValueError:
        return resolved


if __name__ == "__main__":
    raise SystemExit(main())
