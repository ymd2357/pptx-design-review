#!/usr/bin/env python3
"""Smoke test for pptx_repair against generated fixtures.

Asserts:
- clean good.pptx has zero orphan slide parts
- an injected slide99.xml orphan is detected and removed by --apply
- a repaired file is idempotent
- dry-run does not modify the file on disk
- existing .bak files are not overwritten by --backup

Exit code: 0 on success, 1 on any failed assertion.
"""

from __future__ import annotations

import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import make_examples  # noqa: E402
import pptx_repair  # noqa: E402


def inject_orphan(pptx_path: Path, name: str = "ppt/slides/slide99.xml") -> None:
    with zipfile.ZipFile(pptx_path, "a") as zf:
        zf.writestr(name, "<?xml version='1.0'?><sld xmlns='...'/>")


def _zip_names(path: Path) -> set[str]:
    with zipfile.ZipFile(path, "r") as zf:
        return set(zf.namelist())


def main() -> int:
    failures: list = []

    with tempfile.TemporaryDirectory() as tmp:
        tmp_dir = Path(tmp)

        # --- clean deck ---
        clean = tmp_dir / "good.pptx"
        make_examples.make_good(clean)
        clean_orphans = pptx_repair.repair_pptx(clean, apply=False)
        if clean_orphans:
            failures.append(f"good.pptx had orphans; expected 0, got {len(clean_orphans)}")

        # --- injected orphan fixture ---
        dirty = tmp_dir / "dirty.pptx"
        make_examples.make_good(dirty)
        inject_orphan(dirty)

        dry_orphans = pptx_repair.repair_pptx(dirty, apply=False)
        slide_orphans = [entry for entry in dry_orphans if entry.kind == "slide"]
        if len(slide_orphans) != 1:
            failures.append(
                f"dry-run expected 1 orphan slide; got {len(slide_orphans)} entries={dry_orphans}"
            )
        elif slide_orphans[0].path != "ppt/slides/slide99.xml":
            failures.append(f"unexpected orphan path: {slide_orphans[0].path}")

        applied = pptx_repair.repair_pptx(dirty, apply=True)
        if len([entry for entry in applied if entry.kind == "slide"]) != 1:
            failures.append(f"apply expected 1 orphan slide action; got {applied}")

        names = _zip_names(dirty)
        if "ppt/slides/slide99.xml" in names:
            failures.append("orphan slide99.xml still exists after repair")

        # --- idempotency ---
        again = pptx_repair.repair_pptx(dirty, apply=False)
        if again:
            failures.append(f"repaired file is not idempotent; got {len(again)} orphans")

        # --- dry-run must not write ---
        dry = tmp_dir / "dry.pptx"
        make_examples.make_good(dry)
        inject_orphan(dry)
        before_mtime = dry.stat().st_mtime_ns
        orphans = pptx_repair.repair_pptx(dry, apply=False)
        if not orphans:
            failures.append("dry-run produced no orphan actions on injected fixture")
        if dry.stat().st_mtime_ns != before_mtime:
            failures.append("dry-run modified the file on disk")

        # --- backup must preserve the oldest saved state ---
        backup_deck = tmp_dir / "backup.pptx"
        make_examples.make_good(backup_deck)
        inject_orphan(backup_deck)
        backup_path = Path(str(backup_deck) + ".bak")
        sentinel = b"existing original backup"
        backup_path.write_bytes(sentinel)
        pptx_repair.repair_pptx(backup_deck, apply=True, backup=True)
        if backup_path.read_bytes() != sentinel:
            failures.append("--backup overwrote an existing .bak file")

    if failures:
        print("FAIL:")
        for line in failures:
            print(f"- {line}")
        return 1

    print("OK: pptx_repair detects orphan slides, removes them, and respects dry-run")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
