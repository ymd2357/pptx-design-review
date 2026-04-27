#!/usr/bin/env python3
"""Summarize PPTX review findings as P0-P3 action priorities.

This is intentionally not a finding counter. It groups mechanical lint results
into decision-level issues that a reviewer can act on.
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import defaultdict
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Iterable

from pptx import Presentation

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

import pptx_lint  # noqa: E402


PRIORITY_ORDER = {"P0": 0, "P1": 1, "P2": 2, "P3": 3}


@dataclass
class PriorityIssue:
    priority: str
    title: str
    slides: list[int] = field(default_factory=list)
    evidence: list[str] = field(default_factory=list)
    action: str = ""
    checks: list[str] = field(default_factory=list)


def _slide_range(slides: Iterable[int]) -> str:
    ordered = sorted(set(s for s in slides if s > 0))
    if not ordered:
        return "deck"
    runs: list[tuple[int, int]] = []
    start = prev = ordered[0]
    for slide in ordered[1:]:
        if slide == prev + 1:
            prev = slide
        else:
            runs.append((start, prev))
            start = prev = slide
    runs.append((start, prev))
    return ", ".join(str(a) if a == b else f"{a}-{b}" for a, b in runs)


def _clip(text: str, limit: int = 80) -> str:
    text = " ".join(text.split())
    if len(text) <= limit:
        return text
    return text[: limit - 1] + "..."


def _issue(
    priority: str,
    title: str,
    slides: Iterable[int],
    evidence: Iterable[str],
    action: str,
    checks: Iterable[str],
) -> PriorityIssue:
    return PriorityIssue(
        priority=priority,
        title=title,
        slides=sorted(set(slides)),
        evidence=list(dict.fromkeys(evidence)),
        action=action,
        checks=sorted(set(checks)),
    )


def _text_quality_issues(path: Path) -> list[PriorityIssue]:
    prs = Presentation(str(path))
    replacement_slides: dict[int, list[str]] = defaultdict(list)

    for slide_idx, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if not getattr(shape, "has_text_frame", False):
                continue
            text = shape.text_frame.text.strip()
            if not text:
                continue
            if "\ufffd" in text:
                replacement_slides[slide_idx].append(_clip(text))

    issues: list[PriorityIssue] = []
    if replacement_slides:
        evidence = [
            f"slide {slide}: {sample}"
            for slide, samples in replacement_slides.items()
            for sample in samples[:2]
        ]
        issues.append(
            _issue(
                "P0",
                "文字化けしており、内容を正しく読めない",
                replacement_slides.keys(),
                evidence,
                "該当テキストを元原稿または正しい日本語から差し替える。",
                ["text_encoding"],
            )
        )
    return issues


def _slides_for(findings: list[pptx_lint.Finding], checks: set[str]) -> list[int]:
    slides: list[int] = []
    for finding in findings:
        if finding.check not in checks:
            continue
        affected = finding.detail.get("affected_slides") if finding.detail else None
        if affected:
            slides.extend(int(s) for s in affected)
        else:
            slides.append(finding.slide_index)
    return sorted(set(s for s in slides if s > 0))


def _evidence_for(findings: list[pptx_lint.Finding], checks: set[str], limit: int = 5) -> list[str]:
    evidence: list[str] = []
    for finding in findings:
        if finding.check not in checks:
            continue
        slide = "deck" if finding.slide_index == 0 else f"slide {finding.slide_index}"
        evidence.append(f"{slide}: {finding.check}: {_clip(finding.message)}")
        if len(evidence) >= limit:
            break
    return evidence


def summarize_priorities(path: Path) -> list[PriorityIssue]:
    raw_findings = pptx_lint.lint_pptx(path)
    findings = pptx_lint.consolidate_recurring(raw_findings)
    issues = _text_quality_issues(path)

    p0_checks = {"overflow_text", "text_autofit_disabled"}
    slides = _slides_for(findings, p0_checks)
    if slides:
        issues.append(
            _issue(
                "P0",
                "テキストが切れる、または自動縮小で読めなくなるリスクがある",
                slides,
                _evidence_for(findings, p0_checks),
                "該当スライドを目視確認し、テキスト枠・文字量・改行を調整する。",
                p0_checks,
            )
        )

    p1_motion = {"animation_present"}
    slides = _slides_for(findings, p1_motion)
    if slides:
        issues.append(
            _issue(
                "P1",
                "静的配布で失われる可能性があるアニメーション/遷移が含まれる",
                slides,
                _evidence_for(findings, p1_motion),
                "PPTX/PDF 配布で必要な情報が静止状態でも読めるか確認し、不要な動きは削除する。",
                p1_motion,
            )
        )

    p2_checks = {
        "overflow_images",
        "overflow_shapes",
        "safe_margins",
        "safe_text_area_text",
        "alignment_left_top",
        "font_size_scale",
        "font_family",
        "line_height",
        "image_aspect_distortion",
        "image_upscale_ratio",
        "text_color_allowlist",
        "background_color_palette",
        "slide_size",
        "alt_text_required",
    }
    slides = _slides_for(findings, p2_checks)
    if slides:
        issues.append(
            _issue(
                "P2",
                "テンプレート/ブランドルールから外れているが、許容判断できる項目がある",
                slides,
                _evidence_for(findings, p2_checks),
                "SHIFT AI テンプレートとして意図した表現なら許容ルールへ寄せ、そうでなければマスター/スタイルを直す。",
                p2_checks,
            )
        )

    p3_checks = {"geometry_rounding"}
    slides = _slides_for(findings, p3_checks)
    if slides:
        issues.append(
            _issue(
                "P3",
                "座標の微小な丸めズレがある",
                slides,
                _evidence_for(findings, p3_checks),
                "見た目に問題がなければ後回し。テンプレート整備時に一括補正する。",
                p3_checks,
            )
        )

    issues.sort(key=lambda issue: (PRIORITY_ORDER[issue.priority], issue.title))
    return issues


def format_markdown(issues: list[PriorityIssue]) -> str:
    if not issues:
        return "OK: P0-P3 review issues were not found.\n"

    lines: list[str] = []
    for priority in ("P0", "P1", "P2", "P3"):
        bucket = [issue for issue in issues if issue.priority == priority]
        if not bucket:
            continue
        lines.append(f"## {priority}")
        for issue in bucket:
            lines.append(f"- {issue.title}")
            lines.append(f"  - slides: {_slide_range(issue.slides)}")
            if issue.evidence:
                lines.append(f"  - evidence: {issue.evidence[0]}")
                for extra in issue.evidence[1:3]:
                    lines.append(f"    / {extra}")
            lines.append(f"  - action: {issue.action}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="Summarize PPTX review as P0-P3 priorities")
    ap.add_argument("pptx", type=Path, help="path to .pptx")
    ap.add_argument("--json", action="store_true", help="emit priority issues as JSON")
    args = ap.parse_args(argv)

    if not args.pptx.exists():
        print(f"file not found: {args.pptx}", file=sys.stderr)
        return 2

    issues = summarize_priorities(args.pptx)
    if args.json:
        print(json.dumps([asdict(issue) for issue in issues], ensure_ascii=False, indent=2))
    else:
        print(format_markdown(issues), end="")
    return 1 if any(issue.priority == "P0" for issue in issues) else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
