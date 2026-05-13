#!/usr/bin/env python3
"""Publish existing review artifacts into a git-trackable static snapshot.

Happy path:
    python3 scripts/publish_review_snapshot.py --deck 260329-seminar-curriculum-proposal --rev 017

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

    image_dir = select_image_dir(source_dir, rev)
    lint_json = select_json(source_dir, rev, kind="lint")
    priorities_json = select_json(source_dir, rev, kind="priorities")

    if image_dir is None:
        raise SystemExit(f"No slide PNG directory found for rev-{rev} under {source_dir}")
    if lint_json is None:
        raise SystemExit(f"No lint JSON found for rev-{rev} under {source_dir}")

    output_dir = SNAPSHOT_ROOT / args.deck / f"rev-{rev}"
    if output_dir.exists():
        shutil.rmtree(output_dir)
    (output_dir / "images").mkdir(parents=True)

    copied = copy_slide_images(image_dir, output_dir / "images")
    if copied == 0:
        raise SystemExit(f"No slide PNG files found in selected image directory: {image_dir}")

    rewrite_json(lint_json, output_dir / "lint.json")
    if priorities_json is not None:
        rewrite_json(priorities_json, output_dir / "priorities.json")

    print(f"snapshot: {output_dir.relative_to(REPO_ROOT)}")
    print(f"images: {copied} from {image_dir.relative_to(REPO_ROOT)}")
    print(f"lint: {lint_json.relative_to(REPO_ROOT)}")
    if priorities_json is not None:
        print(f"priorities: {priorities_json.relative_to(REPO_ROOT)}")
    return 0


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Copy existing tmp/review artifacts into tmp/review-snapshot for Pages.",
    )
    parser.add_argument("--deck", required=True, help="Deck id under tmp/review/<deck>.")
    parser.add_argument("--rev", required=True, help="Revision number, for example 017 or rev-017.")
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


def copy_slide_images(source_dir: Path, target_dir: Path) -> int:
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


if __name__ == "__main__":
    raise SystemExit(main())
