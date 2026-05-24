#!/usr/bin/env python3
"""Normalize font typeface attributes in a .pptx so PowerPoint Mac-specific
font name conventions (e.g. ``Calibri (MS) Bold``, ``Noto Sans JP Bold``)
become OOXML-conformant family-only names (e.g. ``Calibri``, ``Noto Sans
JP``).

Background
==========
PowerPoint Mac saves the actual full/PostScript name of the font into the
OOXML ``<a:latin typeface="..."/>`` attribute (Mac-side font lookup uses
PostScript names). Windows PowerPoint instead writes the family name.
Browsers (e.g. Chromium used by ``vscode-pptx-viewer``) require the family
name in CSS ``font-family`` and silently fall back to ``sans-serif`` when
they receive a name like ``Calibri (MS) Bold``, which makes the viewer's
rendered text width diverge from the true PowerPoint rendering by
several pt per shape.

This normalizer rewrites typeface values to family-only names. Weight
information is preserved on the parent ``<a:rPr>`` (or equivalent) when a
matching ``b`` / ``font-weight`` attribute is missing, so the rendered
weight is unaffected.

Usage::

    python3 pptx_normalize_fonts.py INPUT.pptx OUTPUT.pptx

(``OUTPUT.pptx`` is overwritten.) Default normalization is conservative:
known weight suffixes are stripped; unknown suffixes are left intact.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

# Recognised weight suffixes (case-insensitive, hyphen/space tolerant).
# Order matters: longer/multi-word entries are matched first.
_WEIGHT_NAMES = [
    "Extra Bold", "ExtraBold", "Ultra Bold", "UltraBold",
    "Semi Bold", "SemiBold", "Semi-Bold",
    "Demi Bold", "DemiBold", "Demi-Bold",
    "Bold", "Heavy", "Black",
    "Light", "Thin", "Extra Light", "ExtraLight", "Ultra Light", "UltraLight",
    "Medium", "Regular",
    "Italic", "Oblique",
]
_WEIGHT_RE = re.compile(
    r"(?:\s*\(MS\)|\s*\(Body\)|\s*\(Headings\))?\s+(?:"
    + "|".join(re.escape(w) for w in _WEIGHT_NAMES)
    + r")\b\s*$",
    re.IGNORECASE,
)
_MS_SUFFIX_RE = re.compile(r"\s*\((?:MS|Body|Headings)\)", re.IGNORECASE)

_BOLD_WEIGHT_KEYWORDS = ("Bold", "Heavy", "Black", "Extra", "Ultra")


@dataclass
class NormalizeReport:
    typeface_changes: dict[str, str]  # original → normalized
    occurrences: int  # total replacements across all parts


def normalize_typeface_name(value: str) -> str:
    """Return the family-only form of ``value``.

    Strips a trailing weight/style suffix (e.g. ``" Bold"``, ``" Medium"``)
    and Mac-specific ``(MS)`` / ``(Body)`` / ``(Headings)`` annotations.
    Empty input is returned unchanged.
    """
    if not value:
        return value
    s = value.strip()
    # Repeatedly strip recognised suffixes (handles "Calibri (MS) Bold").
    while True:
        new = _WEIGHT_RE.sub("", s).strip()
        new = _MS_SUFFIX_RE.sub("", new).strip()
        if new == s:
            break
        s = new
    return s


def looks_bold(original: str) -> bool:
    """Heuristic: does the *original* typeface name imply a bold weight?"""
    lower = original.lower()
    return any(k.lower() in lower for k in _BOLD_WEIGHT_KEYWORDS)


_TYPEFACE_ATTR_RE = re.compile(r'(\btypeface=")([^"]*)(")')


def _normalize_xml_blob(blob: bytes, report: NormalizeReport) -> bytes:
    """Rewrite ``typeface="..."`` occurrences in a single XML byte string."""
    text = blob.decode("utf-8")

    def _sub(match: re.Match) -> str:
        prefix, value, suffix = match.group(1), match.group(2), match.group(3)
        normalized = normalize_typeface_name(value)
        if normalized == value:
            return match.group(0)
        report.typeface_changes.setdefault(value, normalized)
        report.occurrences += 1
        return prefix + normalized + suffix

    new_text = _TYPEFACE_ATTR_RE.sub(_sub, text)
    return new_text.encode("utf-8")


_XML_PATH_PATTERNS = (
    re.compile(r"^ppt/slides/slide\d+\.xml$"),
    re.compile(r"^ppt/slideLayouts/slideLayout\d+\.xml$"),
    re.compile(r"^ppt/slideMasters/slideMaster\d+\.xml$"),
    re.compile(r"^ppt/notesSlides/notesSlide\d+\.xml$"),
    re.compile(r"^ppt/notesMasters/notesMaster\d+\.xml$"),
    re.compile(r"^ppt/handoutMasters/handoutMaster\d+\.xml$"),
    re.compile(r"^ppt/theme/theme\d+\.xml$"),
    re.compile(r"^ppt/tableStyles\.xml$"),
)


def _should_rewrite(path: str) -> bool:
    return any(p.match(path) for p in _XML_PATH_PATTERNS)


def normalize_pptx(src: Path, dst: Path) -> NormalizeReport:
    report = NormalizeReport(typeface_changes={}, occurrences=0)
    if src.resolve() == dst.resolve():
        raise ValueError("src and dst must differ to avoid clobbering input mid-write")
    dst.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(src, "r") as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if _should_rewrite(item.filename):
                data = _normalize_xml_blob(data, report)
            zout.writestr(item, data)
    return report


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("src", type=Path, help="input .pptx")
    ap.add_argument("dst", type=Path, help="output .pptx (overwritten)")
    ap.add_argument("--verbose", action="store_true", help="print every change")
    args = ap.parse_args()
    if not args.src.is_file():
        print(f"error: input not found: {args.src}", file=sys.stderr)
        return 2
    report = normalize_pptx(args.src, args.dst)
    print(f"normalized {report.occurrences} typeface occurrence(s) → {args.dst}")
    if report.typeface_changes:
        print("changes (original → normalized):")
        for orig, new in sorted(report.typeface_changes.items()):
            print(f"  {orig!r:>40} → {new!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
