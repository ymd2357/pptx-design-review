"""Thin orchestrator that binds the existing pptx-design-reviewer engine
(lint / fix / font-normalize / render) for the lint/fix studio app.

No lint/fix logic is duplicated here: every analytical decision is delegated
to ``skills/pptx-design-reviewer/scripts/``. This module only:
  * manages per-upload session directories,
  * normalizes fonts (mandatory pre-step for every lint/fix),
  * runs lint and groups findings by slide -> check type,
  * applies the user's per-slide, per-check toggle via ``fix_pptx(selection=)``,
  * renders slides on demand through the vscode-pptx-viewer pipeline.
"""

from __future__ import annotations

import sys
import json
import shutil
import uuid
from dataclasses import asdict
from pathlib import Path
from typing import Any, Optional

STUDIO_DIR = Path(__file__).resolve().parent
REPO_ROOT = STUDIO_DIR.parent
SCRIPTS_DIR = REPO_ROOT / "skills" / "pptx-design-reviewer" / "scripts"
SESSIONS_DIR = STUDIO_DIR / "sessions"

if str(SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPTS_DIR))

import pptx_lint  # noqa: E402
import pptx_fix  # noqa: E402
import pptx_normalize_fonts  # noqa: E402
import pptx_review_orchestrator as orchestrator  # noqa: E402
from pptx import Presentation  # noqa: E402


# --------------------------------------------------------------------------
# check-type metadata (derived from the guideline fix_policy, single source)
# --------------------------------------------------------------------------

def _apply_mode(check: str) -> Optional[str]:
    """'auto_fix' | 'judgement_fix' | 'no_fix' | None for a lint check."""
    return pptx_fix._apply_mode_for_check(check)


def _rule_for_check(check: str) -> str:
    return pptx_fix.CHECK_TO_RULE.get(check, check)


def _check_is_fixable(check: str) -> bool:
    """A check is mechanically fixable when its fix rule is implemented and
    its guideline apply_mode is not ``no_fix``."""
    mode = _apply_mode(check)
    if mode not in ("auto_fix", "judgement_fix"):
        return False
    rule = _rule_for_check(check)
    return rule in pptx_fix.ALL_RULES and pptx_fix.RULE_ENABLED.get(rule, True)


# --------------------------------------------------------------------------
# sessions
# --------------------------------------------------------------------------

def _session_dir(session_id: str) -> Path:
    # guard against path traversal: session ids are uu4 hex only
    if not session_id or any(c not in "0123456789abcdef" for c in session_id):
        raise ValueError("invalid session id")
    return SESSIONS_DIR / session_id


def _lint_to_grouped(findings_json: list[dict]) -> list[dict]:
    """Group lint findings by slide_index -> check type.

    Returns a list of slide entries, each with a list of check-type groups
    that the UI renders as toggle chips.
    """
    by_slide: dict[int, dict[str, dict]] = {}
    for f in findings_json:
        slide = f.get("slide_index")
        check = f.get("check")
        if slide is None or check is None:
            continue
        slide_groups = by_slide.setdefault(slide, {})
        g = slide_groups.get(check)
        if g is None:
            detail = f.get("detail") or {}
            mode = _apply_mode(check)
            fixable = _check_is_fixable(check)
            g = {
                "check": check,
                "rule": _rule_for_check(check),
                "apply_mode": mode,
                "fixable": fixable,
                # auto_fix on by default; judgement_fix on by default (the
                # toggle itself is the human judgement); manual/no_fix off.
                "default_on": fixable,
                "severity": f.get("severity"),
                "count": 0,
                "messages": [],
            }
            slide_groups[check] = g
        g["count"] += 1
        if len(g["messages"]) < 4 and f.get("message"):
            g["messages"].append(f["message"])
        # escalate to error if any finding in the group is an error
        if f.get("severity") == "error":
            g["severity"] = "error"

    slides = []
    for slide in sorted(by_slide):
        groups = sorted(
            by_slide[slide].values(),
            key=lambda g: (not g["fixable"], g["check"]),
        )
        slides.append({"slide_index": slide, "checks": groups})
    return slides


def create_session(filename: str, data: bytes) -> dict:
    """Save the upload, normalize fonts, lint, and return grouped findings."""
    session_id = uuid.uuid4().hex
    sdir = _session_dir(session_id)
    sdir.mkdir(parents=True, exist_ok=True)

    upload_path = sdir / "upload.pptx"
    upload_path.write_bytes(data)

    # font normalize is the top-level prerequisite for all lint/fix
    source_path = sdir / "source.pptx"
    report = pptx_normalize_fonts.normalize_pptx(upload_path, source_path)

    findings = pptx_lint.lint_pptx(source_path)
    findings_json = [pptx_lint.finding_to_json_dict(f) for f in findings]
    (sdir / "lint.json").write_text(
        json.dumps(findings_json, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    prs = Presentation(str(source_path))
    slide_count = len(prs.slides)

    return {
        "session_id": session_id,
        "filename": filename,
        "slide_count": slide_count,
        "finding_count": len(findings_json),
        "normalize": {
            "occurrences": report.occurrences,
            "typeface_changes": report.typeface_changes,
        },
        "slides": _lint_to_grouped(findings_json),
    }


def _load_findings(session_id: str) -> list[dict]:
    sdir = _session_dir(session_id)
    return json.loads((sdir / "lint.json").read_text(encoding="utf-8"))


def apply_session(session_id: str, selections: list[dict]) -> dict:
    """Apply only the toggled-on (slide_index, check) fixes.

    ``selections`` is ``[{"slide_index": int, "check": str}, ...]``.
    Produces ``fixed.pptx`` in the session directory and returns a summary of
    the applied (and skipped-but-detected) FixActions.
    """
    sdir = _session_dir(session_id)
    source_path = sdir / "source.pptx"
    if not source_path.exists():
        raise FileNotFoundError("session source missing")

    findings_json = _load_findings(session_id)

    # (slide_index, check) the user turned on
    selected_pairs = {
        (int(s["slide_index"]), s["check"])
        for s in selections
        if s.get("check") is not None and s.get("slide_index") is not None
    }
    # gate the fixer by (slide_index, rule); shape-driven rules are 1:1 with
    # their check so this honors check-level toggling.
    selection = {(slide, _rule_for_check(check)) for slide, check in selected_pairs}
    rules = tuple(sorted({rule for _, rule in selection}))

    # finding-driven rules only touch shapes referenced by the findings we
    # pass; restrict to the toggled-on (slide, check) findings for precision.
    filtered_findings = [
        f
        for f in findings_json
        if (f.get("slide_index"), f.get("check")) in selected_pairs
    ]

    fixed_path = sdir / "fixed.pptx"
    shutil.copyfile(source_path, fixed_path)

    if not selection:
        return {"applied": [], "skipped": [], "rules": [], "selected": 0}

    actions = pptx_fix.fix_pptx(
        fixed_path,
        apply=True,
        rules=rules,
        findings=filtered_findings,
        # the toggle is the human judgement: let judgement_fix findings through
        judgement_gate=False,
        selection=selection,
    )

    applied, skipped = [], []
    for a in actions:
        item = {
            "rule": a.rule,
            "slide_index": a.slide_index,
            "shape_id": a.shape_id,
            "shape_name": a.shape_name,
            "status": a.status,
            "reasons": list(a.reasons),
            "before": a.before,
            "after": a.after,
        }
        (applied if a.status == "apply" else skipped).append(item)

    return {
        "applied": applied,
        "skipped": skipped,
        "rules": list(rules),
        "selected": len(selected_pairs),
    }


def render_session(session_id: str, which: str = "source") -> dict:
    """Render slides on demand via the vscode-pptx-viewer pipeline.

    ``which`` is ``'source'`` or ``'fixed'``. Returns the relative PNG paths
    (served by the http server) for each slide. Cached: re-renders only when
    the target PPTX is newer than the existing render directory.
    """
    if which not in ("source", "fixed"):
        raise ValueError("which must be 'source' or 'fixed'")
    sdir = _session_dir(session_id)
    pptx_path = sdir / f"{which}.pptx"
    if not pptx_path.exists():
        raise FileNotFoundError(f"{which}.pptx not found; apply first")

    work = sdir / "render"
    work.mkdir(parents=True, exist_ok=True)
    render_dir = work / which

    needs_render = True
    if render_dir.exists():
        existing = sorted(render_dir.glob("slide-*.png"))
        if existing and render_dir.stat().st_mtime >= pptx_path.stat().st_mtime:
            needs_render = False

    output = ""
    if needs_render:
        ok, output, render_dir = orchestrator._render_pptx(pptx_path, work, which)
        if not ok:
            raise RuntimeError(f"render failed:\n{output[-2000:]}")

    pngs = sorted(render_dir.glob("slide-*.png"))
    return {
        "which": which,
        "slides": [
            {
                "name": p.name,
                "url": f"/api/render/{session_id}/{which}/{p.name}",
            }
            for p in pngs
        ],
        "log_tail": output[-500:] if output else "",
    }


def render_png_path(session_id: str, which: str, name: str) -> Path:
    if which not in ("source", "fixed"):
        raise ValueError("bad which")
    if "/" in name or "\\" in name or not name.endswith(".png"):
        raise ValueError("bad png name")
    return _session_dir(session_id) / "render" / which / name


def fixed_path(session_id: str) -> Path:
    return _session_dir(session_id) / "fixed.pptx"
