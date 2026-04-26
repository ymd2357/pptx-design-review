#!/usr/bin/env python3
"""PPTX source repair tool for orphan slide parts.

Removes physical slide XML parts that are no longer reachable from
``ppt/_rels/presentation.xml.rels``. This works directly on the ZIP package and
does not depend on python-pptx, so it can run before higher-level PPTX tooling.

Usage
    python3 pptx_repair.py DECK.pptx                  # dry-run
    python3 pptx_repair.py DECK.pptx --apply          # write in-place
    python3 pptx_repair.py DECK.pptx --apply --backup # write DECK.pptx.bak if absent
    python3 pptx_repair.py DECK.pptx --json

Exit code
    0 = success (including no orphans)
    1 = invocation error (missing file, invalid package, rewrite failure)
"""

from __future__ import annotations

import argparse
import json
import os
import posixpath
import re
import shutil
import sys
import tempfile
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import List, Sequence
from urllib.parse import unquote, urlsplit
from xml.etree import ElementTree as ET


PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels"
SLIDE_RE = re.compile(r"^ppt/slides/slide\d+\.xml$")
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
SLIDE_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
)


@dataclass
class OrphanEntry:
    path: str
    bytes: int
    kind: str  # "slide" | "rels"
    reason: str = ""


class RepairError(Exception):
    pass


def _zip_info_by_name(infos: Sequence[zipfile.ZipInfo]) -> dict:
    by_name: dict = {}
    for info in infos:
        # Duplicate entries are rare but legal. The remover skips by path, so
        # report aggregate bytes to describe the full amount that will vanish.
        by_name.setdefault(info.filename, 0)
        by_name[info.filename] += info.file_size
    return by_name


def _normalize_target(target: str) -> str:
    """Normalize a presentation.xml.rels Target into a ZIP path."""
    split = urlsplit(target)
    path = unquote(split.path)
    if path.startswith("/"):
        normalized = posixpath.normpath(path.lstrip("/"))
    else:
        normalized = posixpath.normpath(posixpath.join("ppt", path))
    return normalized


def _slide_rels_path(slide_path: str) -> str:
    parent, name = posixpath.split(slide_path)
    return posixpath.join(parent, "_rels", name + ".rels")


def reachable_slide_paths(zf: zipfile.ZipFile) -> set[str]:
    try:
        data = zf.read(PRESENTATION_RELS)
    except KeyError as exc:
        raise RepairError(f"missing required relationship part: {PRESENTATION_RELS}") from exc

    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        raise RepairError(f"failed to parse {PRESENTATION_RELS}: {exc}") from exc

    reachable: set[str] = set()
    for rel in root.findall(f"{{{REL_NS}}}Relationship"):
        target = rel.get("Target")
        if not target:
            continue
        target_mode = rel.get("TargetMode")
        if target_mode and target_mode.lower() == "external":
            continue
        rel_type = rel.get("Type")
        normalized = _normalize_target(target)
        if rel_type == SLIDE_REL_TYPE or SLIDE_RE.match(normalized):
            if SLIDE_RE.match(normalized):
                reachable.add(normalized)
    return reachable


def find_orphans(path: Path) -> List[OrphanEntry]:
    try:
        with zipfile.ZipFile(path, "r") as zf:
            infos = zf.infolist()
            sizes = _zip_info_by_name(infos)
            names = set(sizes)
            reachable = reachable_slide_paths(zf)
    except zipfile.BadZipFile as exc:
        raise RepairError(f"not a valid ZIP/PPTX package: {path}") from exc

    orphans: List[OrphanEntry] = []
    for name in sorted(n for n in names if SLIDE_RE.match(n)):
        if name in reachable:
            continue
        orphans.append(
            OrphanEntry(
                path=name,
                bytes=sizes[name],
                kind="slide",
                reason="not in presentation.xml.rels",
            )
        )
        rels = _slide_rels_path(name)
        if rels in names:
            orphans.append(
                OrphanEntry(
                    path=rels,
                    bytes=sizes[rels],
                    kind="rels",
                    reason=f"relationship part for orphan slide {name}",
                )
            )
    return orphans


def _copy_without_entries(src: Path, dst: Path, remove_paths: set[str]) -> None:
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(
        dst, "w", compression=zipfile.ZIP_DEFLATED
    ) as zout:
        for info in zin.infolist():
            if info.filename in remove_paths:
                continue
            data = zin.read(info.filename)
            zout.writestr(info, data)


def _backup_once(path: Path) -> None:
    backup_path = Path(str(path) + ".bak")
    if backup_path.exists():
        return
    shutil.copy2(path, backup_path)


def repair_pptx(path: Path, *, apply: bool = False, backup: bool = False) -> List[OrphanEntry]:
    orphans = find_orphans(path)
    if not apply or not orphans:
        return orphans

    if backup:
        _backup_once(path)

    remove_paths = {entry.path for entry in orphans}
    tmp_name = ""
    try:
        fd, tmp_name = tempfile.mkstemp(
            prefix=path.name + ".", suffix=".tmp", dir=str(path.parent)
        )
        os.close(fd)
        tmp = Path(tmp_name)
        _copy_without_entries(path, tmp, remove_paths)
        shutil.copymode(path, tmp, follow_symlinks=True)
        os.replace(tmp, path)
    except OSError as exc:
        if tmp_name:
            try:
                Path(tmp_name).unlink(missing_ok=True)
            except OSError:
                pass
        raise RepairError(f"failed to rewrite {path}: {exc}") from exc
    return orphans


def format_orphans(path: Path, orphans: List[OrphanEntry], applied: bool) -> str:
    if not orphans:
        return "OK: no orphans\n"

    slide_n = sum(1 for entry in orphans if entry.kind == "slide")
    rels_n = sum(1 for entry in orphans if entry.kind == "rels")
    if applied:
        return (
            f"Removed {len(orphans)} entries ({slide_n} slides + {rels_n} rels) "
            f"from {path}\n"
        )

    lines = [
        f"Would remove {slide_n} orphan slide parts (and {rels_n} .rels):",
    ]
    for entry in orphans:
        reason = f", {entry.reason}" if entry.reason else ""
        lines.append(f"  {entry.path} ({entry.bytes} bytes{reason})")
    lines.append("")
    return "\n".join(lines)


def main(argv: List[str]) -> int:
    ap = argparse.ArgumentParser(description="PPTX source repair for orphan slide parts")
    ap.add_argument("pptx", type=Path, help="path to .pptx")
    ap.add_argument(
        "--apply",
        action="store_true",
        help="write repaired package back to the file (default: dry-run)",
    )
    ap.add_argument(
        "--backup",
        action="store_true",
        help="with --apply, copy the original to <path>.bak before rewriting if absent",
    )
    ap.add_argument("--json", action="store_true", help="emit repair plan as JSON")
    args = ap.parse_args(argv)

    if not args.pptx.exists():
        print(f"file not found: {args.pptx}", file=sys.stderr)
        return 1
    if not args.pptx.is_file():
        print(f"not a file: {args.pptx}", file=sys.stderr)
        return 1
    if args.backup and not args.apply:
        print("note: --backup has no effect without --apply", file=sys.stderr)

    try:
        orphans = repair_pptx(args.pptx, apply=args.apply, backup=args.backup)
    except RepairError as exc:
        print(str(exc), file=sys.stderr)
        return 1

    if args.json:
        payload = {
            "applied": args.apply,
            "orphans": [asdict(entry) for entry in orphans],
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
    else:
        print(format_orphans(args.pptx, orphans, applied=args.apply), end="")

    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
