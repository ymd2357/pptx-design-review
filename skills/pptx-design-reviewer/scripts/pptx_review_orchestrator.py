#!/usr/bin/env python3
"""Build a Pn-n review matrix from intentional fixtures and deck results.

This is not a replacement for visual review. It answers two separate questions:

1. Does each Pn-n have an intentionally bad fixture proving the detector can fire?
2. What does the current target deck still require after lint/fix/render gates?
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import re
import shutil
import subprocess
import sys
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Callable

from PIL import Image, ImageChops
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_AUTO_SIZE, MSO_VERTICAL_ANCHOR, PP_ALIGN
from pptx.util import Emu, Pt

HERE = Path(__file__).parent
ROOT = HERE.parents[2]
VSCODE_RENDER_SCRIPT = ROOT / "tmp/review/260329-seminar-curriculum-proposal/scripts/render_with_vscode_pptx_viewer.js"
VSCODE_CAPTURE_SCRIPT = ROOT / "tmp/review/260329-seminar-curriculum-proposal/scripts/capture_vscode_pptx_viewer.js"
sys.path.insert(0, str(HERE))

import make_examples  # noqa: E402
import pptx_fix  # noqa: E402
import pptx_lint  # noqa: E402
import pptx_review_priorities  # noqa: E402
import test_pptx_lint as lint_fixtures  # noqa: E402


TASK_ROW_RE = re.compile(
    r"^\|\s*(P[0-3]-\d+)\s*\|\s*`([^`]+)`\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|\s*([^|]+?)\s*\|"
)
PRIORITY_ORDER = {"P0": 0, "P1": 1, "P2": 2, "P3": 3}
RESULT_ONLY_IF_ZERO = {
    "alignment_left_top",
    "alt_text_required",
    "animation_present",
    "background_color_palette",
    "color_only_meaning",
    "font_family",
    "font_size_scale",
    "heading_hierarchy_broken",
    "image_aspect_distortion",
    "image_upscale_ratio",
    "key_area_cropped",
    "line_height",
    "object_gap_too_small",
    "object_overlap",
    "overflow_images",
    "overflow_shapes",
    "overflow_text",
    "slide_size",
    "text_autofit_disabled",
    "text_color_allowlist",
    "text_encoding",
    "text_overlap",
    "text_vertical_balance",
    "wrap_break_changes_meaning",
}


@dataclass
class CatalogItem:
    pn: str
    check: str
    priority: str
    declared_status: str
    detection: str
    fix_policy: str
    viewpoint: str


@dataclass
class FixtureSpec:
    name: str
    checks: tuple[str, ...]
    builder: Callable[[Path, Path], None]
    detector: str = "lint"


@dataclass
class FixtureResult:
    fixture: str
    detector: str
    path: str
    checks: list[str]
    found_counts: dict[str, int]
    status: str


@dataclass
class ReviewRow:
    pn: str
    priority: str
    check: str
    fixture_status: str
    fixture_count: int
    fixture_paths: str
    deck_decision: str
    before_count: int
    after_count: int
    fixed_actions: int
    slides_after: str
    reviewer_mode: str
    next_action: str


@dataclass
class FixEvidenceRow:
    pn: str
    priority: str
    check: str
    fixture: str
    fixture_status: str
    before_count: int
    after_count: int
    applied_actions: int
    manual_actions: int
    render_status: str
    diff_nonempty_slides: int
    work_dir: str
    before_pptx: str
    after_pptx: str
    report_href: str


def _set_text(shape, text: str, *, size: int = 24, color: str | None = None) -> None:
    shape.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    run = shape.text_frame.paragraphs[0].add_run()
    run.text = text
    run.font.name = "Noto Sans JP"
    run.font.size = Pt(size)
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def _make_text_encoding_bad(out: Path, _work: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(1440)
    prs.slide_height = Pt(810)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide.shapes.add_textbox(Pt(120), Pt(60), Pt(900), Pt(80))
    _set_text(title, "文字化け\ufffdサンプル", size=32)
    prs.save(str(out))


def _make_slide_size_bad(out: Path, _work: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(1000)
    prs.slide_height = Pt(500)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    title = slide.shapes.add_textbox(Pt(80), Pt(60), Pt(600), Pt(80))
    _set_text(title, "Non proportional slide", size=24)
    prs.save(str(out))


def _make_safe_font_size_bad(out: Path, _work: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(720)
    prs.slide_height = Pt(405)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Pt(40), Pt(40), Pt(300), Pt(80))
    box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    run = box.text_frame.paragraphs[0].add_run()
    run.text = "Short label"
    run.font.name = "Noto Sans JP"
    run.font.size = Pt(14.75)
    prs.save(str(out))


def _make_safe_line_height_bad(out: Path, _work: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(720)
    prs.slide_height = Pt(405)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Pt(40), Pt(40), Pt(420), Pt(90))
    box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    para = box.text_frame.paragraphs[0]
    para.line_spacing = Pt(16.4)
    run = para.add_run()
    run.text = "Line height fixture"
    run.font.name = "Noto Sans JP"
    run.font.size = Pt(12)
    prs.save(str(out))


def _make_safe_alignment_bad(out: Path, _work: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(720)
    prs.slide_height = Pt(405)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Pt(40), Pt(40), Pt(300), Pt(80))
    box.text_frame.auto_size = MSO_AUTO_SIZE.NONE
    box.text_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    para = box.text_frame.paragraphs[0]
    para.alignment = PP_ALIGN.CENTER
    run = para.add_run()
    run.text = "Alignment fixture"
    run.font.name = "Noto Sans JP"
    run.font.size = Pt(12)
    prs.save(str(out))


def _make_safe_geometry_rounding_bad(out: Path, _work: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(1440)
    prs.slide_height = Pt(810)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    box = slide.shapes.add_textbox(Emu(round(81.05 * 12700)), Pt(40), Pt(200), Pt(50))
    _set_text(box, "Geometry fixture", size=24)
    prs.save(str(out))


def _make_overflow_shape_bad(out: Path, _work: Path) -> None:
    prs = Presentation()
    prs.slide_width = Pt(1440)
    prs.slide_height = Pt(810)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(1380), Pt(200), Pt(120), Pt(80))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string("EEEEEE")
    prs.save(str(out))


def _builder(fn: Callable[[Path], None]) -> Callable[[Path, Path], None]:
    return lambda out, _work: fn(out)


def _rendered_low_contrast(out: Path, work: Path) -> None:
    lint_fixtures._make_rendered_low_contrast_case(out, work / "rendered-low-contrast-images")


def fixture_specs() -> list[FixtureSpec]:
    return [
        FixtureSpec("text-encoding-bad", ("text_encoding",), _make_text_encoding_bad, "priority"),
        FixtureSpec("safe-font-size-bad", ("font_size_scale",), _make_safe_font_size_bad),
        FixtureSpec("safe-line-height-bad", ("line_height",), _make_safe_line_height_bad),
        FixtureSpec("safe-alignment-bad", ("alignment_left_top",), _make_safe_alignment_bad),
        FixtureSpec("safe-geometry-rounding-bad", ("geometry_rounding",), _make_safe_geometry_rounding_bad),
        FixtureSpec(
            "bad-multi-check",
            (
                "animation_present",
                "alt_text_required",
                "alignment_left_top",
                "background_color_palette",
                "contrast_ratio",
                "font_family",
                "font_size_scale",
                "geometry_rounding",
                "image_upscale_ratio",
                "line_height",
                "low_contrast",
                "safe_margins",
                "safe_text_area_text",
                "text_autofit_disabled",
                "text_color_allowlist",
                "overflow_text",
            ),
            _builder(make_examples.make_bad),
        ),
        FixtureSpec(
            "object-relationships-bad",
            (
                "inner_padding_imbalance",
                "object_gap_too_small",
                "object_overlap",
                "text_overlap",
            ),
            _builder(lint_fixtures._make_object_relationships_bad),
        ),
        FixtureSpec("content-overflow-image", ("overflow_images",), _builder(lint_fixtures._make_content_overflow_image)),
        FixtureSpec("overflow-shape-bad", ("overflow_shapes",), _make_overflow_shape_bad),
        FixtureSpec("slide-size-bad", ("slide_size",), _make_slide_size_bad),
        FixtureSpec("aspect-distorted-image", ("image_aspect_distortion",), _builder(lint_fixtures._make_aspect_distorted_image)),
        FixtureSpec("missing-title-bad", ("missing_required_element",), _builder(lint_fixtures._make_missing_title_bad)),
        FixtureSpec("heading-hierarchy-bad", ("heading_hierarchy_broken",), _builder(lint_fixtures._make_heading_hierarchy_bad)),
        FixtureSpec("reading-order-bad", ("reading_order",), _builder(lint_fixtures._make_reading_order_bad)),
        FixtureSpec("wrap-break-bad", ("wrap_break_changes_meaning",), _builder(lint_fixtures._make_wrap_break_bad)),
        FixtureSpec("key-area-cropped-bad", ("key_area_cropped",), _builder(lint_fixtures._make_key_area_cropped_bad)),
        FixtureSpec("color-only-bad", ("color_only_meaning",), _builder(lint_fixtures._make_color_only_bad)),
        FixtureSpec("top-anchor-bottom-void-bad", ("text_vertical_balance",), _builder(lint_fixtures._make_top_anchor_bottom_void_bad)),
        FixtureSpec("card-grid-consistency-bad", ("card_grid_consistency",), _builder(lint_fixtures._make_card_grid_consistency_bad)),
        FixtureSpec("rendered-low-contrast", ("low_contrast",), _rendered_low_contrast),
    ]


def load_json(path: Path | None, default: Any) -> Any:
    if path is None or not path.exists():
        return default
    return json.loads(path.read_text(encoding="utf-8"))


def iter_catalog_lines(tasks_md: Path) -> list[str]:
    lines: list[str] = []
    in_catalog = False
    for line in tasks_md.read_text(encoding="utf-8").splitlines():
        if line.startswith("## チェック観点一覧"):
            in_catalog = True
            continue
        if in_catalog and line.startswith("## "):
            break
        if in_catalog:
            lines.append(line)
    if not lines:
        raise RuntimeError(f"Pn-n catalog section was not found in {tasks_md}")
    return lines


def load_catalog(tasks_md: Path) -> list[CatalogItem]:
    items: list[CatalogItem] = []
    seen: set[tuple[str, str]] = set()
    for line in iter_catalog_lines(tasks_md):
        match = TASK_ROW_RE.match(line)
        if not match:
            continue
        pn, check, status, detection, fix_policy, viewpoint = match.groups()
        key = (pn.strip(), check.strip())
        if key in seen:
            continue
        seen.add(key)
        items.append(
            CatalogItem(
                pn=pn.strip(),
                check=check.strip(),
                priority=pn.split("-", 1)[0],
                declared_status=status.strip(),
                detection=detection.strip(),
                fix_policy=fix_policy.strip(),
                viewpoint=viewpoint.strip(),
            )
        )
    if not items:
        raise RuntimeError(f"Pn-n catalog rows were not found in {tasks_md}")
    return sorted(items, key=lambda item: (PRIORITY_ORDER[item.priority], int(item.pn.split("-")[1])))


def run_fixture(spec: FixtureSpec, fixture_dir: Path) -> FixtureResult:
    work = fixture_dir / spec.name
    work.mkdir(parents=True, exist_ok=True)
    pptx_path = work / f"{spec.name}.pptx"
    spec.builder(pptx_path, work)

    if spec.detector == "priority":
        issues = pptx_review_priorities.summarize_priorities(pptx_path)
        found = Counter(check for issue in issues for check in issue.checks)
    else:
        rendered_dir = work / "rendered-low-contrast-images"
        rendered_arg = rendered_dir if rendered_dir.exists() else None
        findings = pptx_lint.lint_pptx(pptx_path, rendered_image_dir=rendered_arg)
        found = Counter(finding.check for finding in findings)

    expected = set(spec.checks)
    found_counts = {check: found.get(check, 0) for check in spec.checks}
    missing = [check for check, count in found_counts.items() if count <= 0]
    return FixtureResult(
        fixture=spec.name,
        detector=spec.detector,
        path=str(pptx_path),
        checks=sorted(expected),
        found_counts=found_counts,
        status="pass" if not missing else "fail",
    )


def findings_by_check(findings: list[dict]) -> dict[str, list[dict]]:
    by: dict[str, list[dict]] = defaultdict(list)
    for finding in findings:
        by[finding.get("check", "")].append(finding)
    return by


def slides_for(findings: list[dict]) -> list[int]:
    slides: set[int] = set()
    for finding in findings:
        detail = finding.get("detail") or {}
        affected = detail.get("affected_slides")
        if isinstance(affected, list):
            slides.update(int(s) for s in affected if int(s) > 0)
        else:
            slide = finding.get("slide_index")
            if isinstance(slide, int) and slide > 0:
                slides.add(slide)
    return sorted(slides)


def slide_range(slides: list[int]) -> str:
    if not slides:
        return ""
    runs: list[tuple[int, int]] = []
    start = prev = slides[0]
    for slide in slides[1:]:
        if slide == prev + 1:
            prev = slide
            continue
        runs.append((start, prev))
        start = prev = slide
    runs.append((start, prev))
    return ", ".join(str(a) if a == b else f"{a}-{b}" for a, b in runs)


def action_counts_by_check(actions_json: dict) -> Counter:
    counts: Counter = Counter()
    for action in actions_json.get("actions", []):
        if action.get("status") != "apply":
            continue
        for update in (action.get("after") or {}).get("updates", []):
            check = update.get("check")
            if check:
                counts[check] += 1
    return counts


def fixability_counts(findings: list[dict]) -> tuple[int, int]:
    auto = 0
    manual = 0
    for finding in findings:
        detail = finding.get("detail") or {}
        if detail.get("fixability") == "auto_fix_candidate":
            auto += 1
        elif detail.get("fixability") == "manual_required":
            manual += 1
    return auto, manual


def classify_deck(item: CatalogItem, before: list[dict], after: list[dict], fixed_actions: int) -> tuple[str, str, str]:
    if after and fixability_counts(after)[0] == len(after):
        return "auto_fix_ready", "result_then_auto_fix", "fixer を適用し、after lint=0 と diff 対象範囲を確認する。"
    if after:
        if item.priority in {"P0", "P1"}:
            return "manual_review_required", "finding_detail_and_visual", "finding detail とレンダ画像で修正/許容/対象外を判断する。"
        if item.priority == "P2":
            return "accept_or_fix_decision_required", "sampled_visual_or_acceptance", "代表スライドでテンプレート意図として許容するか修正方針を決める。"
        return "defer_or_batch_fix", "result_only_unless_touching_layout", "見た目に問題がなければ defer。レイアウト修正時だけ一括確認する。"
    if fixed_actions:
        return "fixed_verified", "result_and_diff", "fix actions、after lint=0、before/after diff の対象スライドだけ確認する。"
    mode = "result_only_with_fixture" if item.check in RESULT_ONLY_IF_ZERO else "result_only"
    return "pass_by_result", mode, "after lint=0。ただし intentionally bad fixture が pass していることを前提に結果だけ見る。"


def build_rows(
    catalog: list[CatalogItem],
    before_lint: list[dict],
    after_lint: list[dict],
    actions_json: dict,
    fixture_results: list[FixtureResult],
) -> list[ReviewRow]:
    before_by = findings_by_check(before_lint)
    after_by = findings_by_check(after_lint)
    actions_by = action_counts_by_check(actions_json)
    fixtures_by_check: dict[str, list[FixtureResult]] = defaultdict(list)
    for result in fixture_results:
        for check, count in result.found_counts.items():
            if count > 0:
                fixtures_by_check[check].append(result)

    rows: list[ReviewRow] = []
    for item in catalog:
        before = before_by.get(item.check, [])
        after = after_by.get(item.check, [])
        fixed = actions_by[item.check]
        decision, mode, next_action = classify_deck(item, before, after, fixed)
        fixtures = fixtures_by_check.get(item.check, [])
        rows.append(
            ReviewRow(
                pn=item.pn,
                priority=item.priority,
                check=item.check,
                fixture_status="pass" if fixtures else "missing",
                fixture_count=sum(result.found_counts.get(item.check, 0) for result in fixtures),
                fixture_paths="; ".join(result.path for result in fixtures),
                deck_decision=decision,
                before_count=len(before),
                after_count=len(after),
                fixed_actions=fixed,
                slides_after=slide_range(slides_for(after)),
                reviewer_mode=mode,
                next_action=next_action,
            )
        )
    return rows


def write_tsv(path: Path, rows: list[ReviewRow]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=list(asdict(rows[0]).keys()), delimiter="\t")
        writer.writeheader()
        for row in rows:
            writer.writerow(asdict(row))


def summary(rows: list[ReviewRow], fixtures: list[FixtureResult]) -> dict[str, Any]:
    return {
        "total_items": len(rows),
        "fixture_pass_items": [row.pn for row in rows if row.fixture_status == "pass"],
        "fixture_missing_items": [row.pn for row in rows if row.fixture_status == "missing"],
        "deck_decisions": dict(Counter(row.deck_decision for row in rows)),
        "result_only_pass_items": [row.pn for row in rows if row.deck_decision == "pass_by_result"],
        "manual_review_items": [
            row.pn
            for row in rows
            if row.deck_decision in {"manual_review_required", "accept_or_fix_decision_required"}
        ],
        "auto_fix_ready_items": [row.pn for row in rows if row.deck_decision == "auto_fix_ready"],
        "fixed_verified_items": [row.pn for row in rows if row.deck_decision == "fixed_verified"],
        "fixture_failures": [asdict(result) for result in fixtures if result.status != "pass"],
    }


def write_json(path: Path, rows: list[ReviewRow], fixtures: list[FixtureResult]) -> None:
    payload = {
        "summary": summary(rows, fixtures),
        "rows": [asdict(row) for row in rows],
        "fixtures": [asdict(result) for result in fixtures],
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def write_markdown(path: Path, rows: list[ReviewRow], fixtures: list[FixtureResult]) -> None:
    s = summary(rows, fixtures)
    result_only = [row for row in rows if row.deck_decision == "pass_by_result"]
    active = [row for row in rows if row.deck_decision != "pass_by_result"]
    lines = [
        "# Pn-n Review Orchestrator Report",
        "",
        "## Summary",
        "",
        f"- total Pn-n items: {s['total_items']}",
        f"- intentionally bad fixture pass: {len(s['fixture_pass_items'])}",
        f"- intentionally bad fixture missing: {len(s['fixture_missing_items'])}",
        f"- result-only deck pass: {len(result_only)}",
        f"- active deck queue: {len(active)}",
        "",
        "## Test Flow",
        "",
        "1. Catalog gate: `doc/tasks.md` の `チェック観点一覧` から Pn-n/check を固定する。",
        "2. Intentional fixture gate: 各 check の bad fixture を生成し、検出器が意図通り fire するか確認する。",
        "3. Deck measurement gate: 対象 deck の before/after lint と fix actions を check 別に突き合わせる。",
        "4. Review routing gate: fixture pass かつ after=0 の項目は result-only、残件は auto-fix/manual/acceptance/defer に振り分ける。",
        "",
    ]
    if s["fixture_missing_items"]:
        lines.extend(["## Fixture Missing", ""])
        for row in rows:
            if row.fixture_status != "pass":
                lines.append(f"- {row.pn} `{row.check}`")
        lines.append("")

    lines.extend(["## Result-Only Deck Pass With Fixture", ""])
    for row in result_only:
        lines.append(
            f"- {row.pn} `{row.check}`: deck after=0, fixture_count={row.fixture_count}, fixture={row.fixture_paths}"
        )
    lines.append("")

    lines.extend(["## Active Deck Queue", ""])
    for row in active:
        lines.append(
            f"- {row.pn} `{row.check}`: {row.deck_decision}, before={row.before_count}, after={row.after_count}, "
            f"slides={row.slides_after or '-'}, fixture_count={row.fixture_count}"
        )
        lines.append(f"  - 次: {row.next_action}")
    lines.append("")

    lines.extend(["## Full Matrix", ""])
    lines.append("| Pn | check | fixture | fixture_count | deck_decision | after | slides | next |")
    lines.append("| --- | --- | --- | ---: | --- | ---: | --- | --- |")
    for row in rows:
        lines.append(
            f"| {row.pn} | `{row.check}` | {row.fixture_status} | {row.fixture_count} | "
            f"{row.deck_decision} | {row.after_count} | {row.slides_after or '-'} | {row.next_action} |"
        )
    lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def write_html(path: Path, markdown_path: Path) -> None:
    text = markdown_path.read_text(encoding="utf-8")
    escaped = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    path.write_text(
        f"""<!doctype html>
<html lang="ja">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>Pn-n Review Orchestrator</title>
  <style>
    body{{font-family:system-ui,-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;margin:24px;color:#222;line-height:1.55}}
    pre{{white-space:pre-wrap;background:#f7f7f7;border:1px solid #ddd;padding:16px;overflow:auto}}
  </style>
</head>
<body>
<pre>{escaped}</pre>
</body>
</html>
""",
        encoding="utf-8",
    )


def _slug(value: str) -> str:
    return re.sub(r"[^A-Za-z0-9_.-]+", "-", value).strip("-")


def _fixture_spec_by_check(specs: list[FixtureSpec]) -> dict[str, FixtureSpec]:
    mapping: dict[str, FixtureSpec] = {}
    for spec in specs:
        for check in spec.checks:
            mapping.setdefault(check, spec)
    return mapping


def _lint_fixture(path: Path, work: Path) -> list:
    rendered_dir = work / "rendered-low-contrast-images"
    rendered_arg = rendered_dir if rendered_dir.exists() else None
    return pptx_lint.lint_pptx(path, rendered_image_dir=rendered_arg)


def _count_check(item: CatalogItem, pptx_path: Path, findings: list) -> int:
    if item.check == "text_encoding":
        return sum(
            1
            for issue in pptx_review_priorities.summarize_priorities(pptx_path)
            if item.check in issue.checks
        )
    return sum(
        1
        for finding in findings
        if (
            finding.get("check")
            if isinstance(finding, dict)
            else getattr(finding, "check", None)
        )
        == item.check
    )


def _fix_rules_for_check(check: str) -> tuple[str, ...]:
    rule = pptx_fix.CHECK_TO_RULE.get(check)
    return (rule,) if rule else ()


def _run(cmd: list[str]) -> tuple[bool, str]:
    try:
        completed = subprocess.run(
            cmd,
            check=False,
            cwd=ROOT,
            text=True,
            capture_output=True,
        )
    except Exception as exc:  # noqa: BLE001 - keep evidence generation non-fatal.
        return False, str(exc)
    output = "\n".join(part for part in (completed.stdout, completed.stderr) if part)
    return completed.returncode == 0, output


def _render_pptx(pptx_path: Path, work: Path, label: str) -> tuple[bool, str, Path]:
    viewer_dir = work / f"{label}-viewer"
    render_dir = work / label
    ok, output = _run(["node", str(VSCODE_RENDER_SCRIPT), str(pptx_path), str(viewer_dir)])
    if not ok:
        return False, output, render_dir
    ok, capture_output = _run(["node", str(VSCODE_CAPTURE_SCRIPT), str(viewer_dir), str(render_dir)])
    return ok, output + "\n" + capture_output, render_dir


def _diff_renders(before_dir: Path, after_dir: Path, diff_dir: Path) -> int:
    diff_dir.mkdir(parents=True, exist_ok=True)
    nonempty = 0
    for before_path in sorted(before_dir.glob("slide-*.png")):
        after_path = after_dir / before_path.name
        if not after_path.exists():
            continue
        before = Image.open(before_path).convert("RGB")
        after = Image.open(after_path).convert("RGB")
        if before.size != after.size:
            after = after.resize(before.size)
        diff = ImageChops.difference(before, after)
        if diff.getbbox():
            nonempty += 1
        diff.point(lambda v: min(255, v * 4)).save(diff_dir / before_path.name)
    return nonempty


def _write_fixture_evidence_html(path: Path, rows: list[FixEvidenceRow], outdir: Path) -> None:
    html_dir = path.parent

    def fix_outcome(row: FixEvidenceRow) -> str:
        if row.applied_actions:
            return "autofixed"
        if row.manual_actions:
            return "manual_required"
        if row.before_count and row.after_count == row.before_count:
            return "no_autofix_rule"
        if row.before_count == 0:
            return "fixture_not_detected"
        return "no_change"

    lines = [
        "<!doctype html>",
        '<html lang="ja">',
        "<head>",
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width,initial-scale=1">',
        "<title>Pn-n Fixture Fix Evidence</title>",
        "<style>",
        "body{font-family:system-ui,-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;margin:24px;color:#222;background:#fafafa}",
        "h1{font-size:24px;margin:0 0 16px} h2{font-size:18px;margin:28px 0 8px}",
        "table{border-collapse:collapse;width:100%;background:white}td,th{border:1px solid #ddd;padding:6px 8px;text-align:left;font-size:13px}",
        ".grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin:8px 0 24px}.panel{background:white;border:1px solid #ddd;padding:8px}",
        "img{width:100%;height:auto;display:block}.muted{color:#666}.fail{color:#b00020}.ok{color:#0b6b2b}",
        "</style>",
        "</head><body>",
        "<h1>Pn-n Fixture Fix Evidence</h1>",
        "<p class='muted'>各Pn-nの intentionally bad fixture を生成し、現在実装済みの pptx_fix auto-fix だけを check 単位で適用した前後を vscode-pptx-viewer でレンダリングしています。after は「完全修正後」ではなく「現時点の auto-fix 実行後」です。LibreOfficeは未使用です。</p>",
        "<table><thead><tr><th>Pn</th><th>check</th><th>fixture</th><th>before</th><th>after autofix</th><th>applied</th><th>manual</th><th>outcome</th><th>render</th><th>diff slides</th></tr></thead><tbody>",
    ]
    for row in rows:
        cls = "ok" if row.render_status == "ok" else "fail"
        outcome = fix_outcome(row)
        lines.append(
            "<tr>"
            f"<td>{html.escape(row.pn)}</td>"
            f"<td><code>{html.escape(row.check)}</code></td>"
            f"<td>{html.escape(row.fixture)}</td>"
            f"<td>{row.before_count}</td><td>{row.after_count}</td>"
            f"<td>{row.applied_actions}</td><td>{row.manual_actions}</td>"
            f"<td>{html.escape(outcome)}</td>"
            f"<td class='{cls}'>{html.escape(row.render_status)}</td>"
            f"<td>{row.diff_nonempty_slides}</td>"
            "</tr>"
        )
    lines.append("</tbody></table>")

    for row in rows:
        work = Path(row.work_dir)
        lines.append(f"<h2>{html.escape(row.pn)} <code>{html.escape(row.check)}</code></h2>")
        lines.append(
            f"<p class='muted'>fixture={html.escape(row.fixture)} / "
            f"before={row.before_count}, after={row.after_count}, "
            f"applied={row.applied_actions}, manual={row.manual_actions}, "
            f"outcome={html.escape(fix_outcome(row))}, render={html.escape(row.render_status)}</p>"
        )
        before_img = work / "before" / "slide-01.png"
        after_img = work / "after" / "slide-01.png"
        diff_img = work / "diff" / "slide-01.png"
        if before_img.exists() and after_img.exists() and diff_img.exists():
            lines.append("<div class='grid'>")
            for label, image_path in (("before", before_img), ("after autofix", after_img), ("diff x4", diff_img)):
                rel = image_path.relative_to(html_dir).as_posix()
                lines.append(
                    f"<div class='panel'><b>{html.escape(label)}</b><img src='{html.escape(rel)}'></div>"
                )
            lines.append("</div>")
        else:
            lines.append("<p class='fail'>render image missing</p>")
    lines.append("</body></html>")
    path.write_text("\n".join(lines), encoding="utf-8")


def write_fixture_fix_evidence(outdir: Path, catalog: list[CatalogItem]) -> list[FixEvidenceRow]:
    evidence_dir = outdir / "fixture-fix-evidence"
    evidence_dir.mkdir(parents=True, exist_ok=True)
    spec_by_check = _fixture_spec_by_check(fixture_specs())
    rows: list[FixEvidenceRow] = []
    for item in catalog:
        spec = spec_by_check.get(item.check)
        work = evidence_dir / f"{item.pn}-{_slug(item.check)}"
        work.mkdir(parents=True, exist_ok=True)
        before_pptx = work / "before.pptx"
        after_pptx = work / "after.pptx"
        if spec is None:
            rows.append(
                FixEvidenceRow(
                    pn=item.pn,
                    priority=item.priority,
                    check=item.check,
                    fixture="",
                    fixture_status="missing",
                    before_count=0,
                    after_count=0,
                    applied_actions=0,
                    manual_actions=0,
                    render_status="missing_fixture",
                    diff_nonempty_slides=0,
                    work_dir=str(work),
                    before_pptx=str(before_pptx),
                    after_pptx=str(after_pptx),
                    report_href="",
                )
            )
            continue

        spec.builder(before_pptx, work)
        shutil.copy2(before_pptx, after_pptx)
        before_findings = _lint_fixture(before_pptx, work)
        before_count = _count_check(item, before_pptx, before_findings)
        before_findings_json = [
            pptx_lint.finding_to_json_dict(finding)
            for finding in before_findings
            if finding.check == item.check
        ]
        rules = _fix_rules_for_check(item.check)
        actions = pptx_fix.fix_pptx(
            after_pptx,
            apply=True,
            rules=rules,
            findings=before_findings_json,
        )
        after_findings = _lint_fixture(after_pptx, work)
        after_count = _count_check(item, after_pptx, after_findings)
        applied = sum(1 for action in actions if action.status == "apply")
        manual = sum(1 for action in actions if action.status == "manual_required")
        if before_count > 0 and not rules:
            manual = before_count

        before_ok, before_render_output, before_dir = _render_pptx(before_pptx, work, "before")
        after_ok, after_render_output, after_dir = _render_pptx(after_pptx, work, "after")
        (work / "render.log").write_text(
            before_render_output + "\n--- after ---\n" + after_render_output,
            encoding="utf-8",
        )
        if before_ok and after_ok:
            diff_nonempty = _diff_renders(before_dir, after_dir, work / "diff")
            render_status = "ok"
        else:
            diff_nonempty = 0
            render_status = "render_failed"

        rows.append(
            FixEvidenceRow(
                pn=item.pn,
                priority=item.priority,
                check=item.check,
                fixture=spec.name,
                fixture_status="pass" if before_count > 0 else "missing",
                before_count=before_count,
                after_count=after_count,
                applied_actions=applied,
                manual_actions=manual,
                render_status=render_status,
                diff_nonempty_slides=diff_nonempty,
                work_dir=str(work),
                before_pptx=str(before_pptx),
                after_pptx=str(after_pptx),
                report_href="fixture-fix-evidence.html",
            )
        )

    (outdir / "fixture-fix-evidence.json").write_text(
        json.dumps([asdict(row) for row in rows], ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    _write_fixture_evidence_html(outdir / "fixture-fix-evidence.html", rows, evidence_dir)
    return rows


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--tasks", type=Path, default=ROOT / "doc/tasks.md")
    ap.add_argument("--before-lint", type=Path)
    ap.add_argument("--after-lint", type=Path)
    ap.add_argument("--actions-json", type=Path)
    ap.add_argument("--outdir", type=Path, required=True)
    ap.add_argument(
        "--fixture-fix-evidence",
        action="store_true",
        help="generate per-Pn fixture auto-fix before/after/diff evidence with vscode-pptx-viewer",
    )
    args = ap.parse_args()

    catalog = load_catalog(args.tasks)
    fixture_dir = args.outdir / "fixtures"
    fixture_results = [run_fixture(spec, fixture_dir) for spec in fixture_specs()]
    rows = build_rows(
        catalog,
        load_json(args.before_lint, []),
        load_json(args.after_lint, []),
        load_json(args.actions_json, {"actions": []}),
        fixture_results,
    )

    args.outdir.mkdir(parents=True, exist_ok=True)
    write_tsv(args.outdir / "pn-review-orchestrator-matrix.tsv", rows)
    write_json(args.outdir / "pn-review-orchestrator-report.json", rows, fixture_results)
    markdown_path = args.outdir / "pn-review-orchestrator-report.md"
    write_markdown(markdown_path, rows, fixture_results)
    write_html(args.outdir / "index.html", markdown_path)
    if args.fixture_fix_evidence:
        evidence_rows = write_fixture_fix_evidence(args.outdir, catalog)
        print(args.outdir / "fixture-fix-evidence.html")
        print(args.outdir / "fixture-fix-evidence.json")
        print(f"fixture_fix_evidence_rows={len(evidence_rows)}")
    print(markdown_path)
    print(args.outdir / "index.html")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
