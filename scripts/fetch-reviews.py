#!/usr/bin/env python3
"""Fetch encrypted review feedback from the shared Cloudflare KV and apply it
to doc/reviews/<deck>/rev-<rev>-decisions.tsv and
rev-<rev>-finding-judgements.json.

The web UI POSTs an age-encrypted payload to
``https://pptx-visual-review.pages.dev/api/feedback`` with metadata
``{app, deck, rev, ciphertext}``. Only payloads whose ``app`` field equals
``pptx-design-review`` are processed.

The age private key lives at ``~/.config/pptx-design-review/age-key.txt`` and
is required for decryption. Decryption is delegated to the ``age`` CLI
(`brew install age`).

Usage:
    python3 scripts/fetch-reviews.py --list
    python3 scripts/fetch-reviews.py --apply
    python3 scripts/fetch-reviews.py --apply --key review-2026-05-16T05:30:00Z
"""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Iterable

FEEDBACK_ENDPOINT = "https://pptx-visual-review.pages.dev/api/feedback"
AGE_KEY_PATH = Path.home() / ".config" / "pptx-design-review" / "age-key.txt"
STATE_PATH = Path.home() / ".config" / "pptx-design-review" / "fetch-reviews-state.json"
APP_NAMESPACE = "pptx-design-review"
REPO_ROOT = Path(__file__).resolve().parent.parent


def fetch_json(url: str) -> dict:
    # Cloudflare's default bot mitigation rejects the Python-urllib UA with 403.
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": "pptx-design-review fetch-reviews.py",
            "Accept": "application/json",
        },
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read())


def list_keys() -> list[str]:
    body = fetch_json(f"{FEEDBACK_ENDPOINT}?list=1")
    if not body.get("ok"):
        raise RuntimeError(f"list failed: {body}")
    return body.get("keys", [])


def fetch_payload(key: str) -> dict | None:
    qs = urllib.parse.urlencode({"key": key})
    body = fetch_json(f"{FEEDBACK_ENDPOINT}?{qs}")
    if not body.get("ok"):
        return None
    return body.get("payload")


def filter_app_entries(keys: Iterable[str]) -> list[tuple[str, dict]]:
    rows: list[tuple[str, dict]] = []
    for key in keys:
        payload = fetch_payload(key)
        if not payload or payload.get("app") != APP_NAMESPACE:
            continue
        rows.append((key, payload))
    return rows


def decrypt(ciphertext: str) -> dict:
    if not AGE_KEY_PATH.is_file():
        raise SystemExit(f"missing age key: {AGE_KEY_PATH}")
    proc = subprocess.run(
        ["age", "-d", "-i", str(AGE_KEY_PATH)],
        input=ciphertext.encode("utf-8"),
        capture_output=True,
        check=False,
    )
    if proc.returncode != 0:
        raise SystemExit(f"age -d failed: {proc.stderr.decode('utf-8', 'replace')}")
    return json.loads(proc.stdout)


def render_decisions_tsv(rows: list[dict]) -> str:
    header = [
        "review_no",
        "check_id",
        "priority",
        "latest_lint_count",
        "observation_decision",
        "finding_dispositions",
        "rationale",
        "related_artifacts",
    ]
    lines = ["\t".join(header)]
    for row in rows:
        lines.append("\t".join(str(row.get(col, "")) for col in header))
    return "\n".join(lines) + "\n"


def read_existing_compare_memos(path: Path) -> dict[str, str]:
    """Read the memo column from an existing compare TSV.

    SPA payloads frequently arrive with an empty memo (the UI does not
    require a free-text comment for adopt/reject), but a human reviewer may
    have hand-edited memo on disk after a previous --apply run. Returning
    the on-disk memos lets the caller preserve them when the new payload
    is silent on memo. Returns {} when the file is missing or unreadable.
    """
    if not path.exists():
        return {}
    memos: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError:
        return {}
    for line in lines[1:]:  # skip header
        if not line:
            continue
        cols = line.split("\t")
        if len(cols) < 3:
            continue
        slide_no, _decision = cols[0], cols[1]
        memo = "\t".join(cols[2:])
        if memo:
            memos[slide_no] = memo
    return memos


def render_compare_tsv(slides: list[dict], existing_memos: dict[str, str] | None = None) -> str:
    header = ["slide_no", "decision", "memo"]
    lines = ["\t".join(header)]
    memos = existing_memos or {}
    for slide in slides:
        slide_no = str(slide.get("slideNo", ""))
        decision = slide.get("decision") or ""
        memo = (slide.get("memo") or "").replace("\t", " ").replace("\n", " / ")
        if not memo and slide_no in memos:
            memo = memos[slide_no]
        lines.append("\t".join([slide_no, decision, memo]))
    return "\n".join(lines) + "\n"


def write_outputs(payload: dict, key: str, dry_run: bool) -> None:
    deck = payload.get("deck")
    rev = payload.get("rev")
    if not deck or not rev:
        print(f"skip {key}: missing deck/rev", file=sys.stderr)
        return

    out_dir = REPO_ROOT / "doc" / "reviews" / deck
    out_dir.mkdir(parents=True, exist_ok=True)

    decisions = payload.get("decisions")

    # Compare-mode payload (REV-017 onwards): {mode: "compare", slides: [...]}.
    if isinstance(decisions, dict) and decisions.get("mode") == "compare":
        slides = decisions.get("slides") or []
        if not isinstance(slides, list):
            print(f"skip {key}: compare.slides not list", file=sys.stderr)
            return
        compare_path = out_dir / f"rev-{rev}-compare.tsv"
        existing_memos = read_existing_compare_memos(compare_path)
        tsv = render_compare_tsv(slides, existing_memos)
        adopted = sum(1 for s in slides if s.get("decision") == "adopt")
        rejected = sum(1 for s in slides if s.get("decision") == "reject")
        undecided = sum(1 for s in slides if not s.get("decision"))
        print(f"{key}: deck={deck} rev={rev} mode=compare")
        print(
            f"  -> {compare_path.relative_to(REPO_ROOT)}"
            f" ({len(slides)} slides / 採用 {adopted} / 不採用 {rejected} / 未判定 {undecided})"
        )
        if dry_run:
            return
        compare_path.write_text(tsv, encoding="utf-8")
        return

    # Legacy observation-list payload: decisions is a list of decision rows.
    if not isinstance(decisions, list):
        print(f"skip {key}: decisions has unsupported shape", file=sys.stderr)
        return

    decisions_path = out_dir / f"rev-{rev}-decisions.tsv"
    judgements_path = out_dir / f"rev-{rev}-finding-judgements.json"
    tsv = render_decisions_tsv(decisions)
    judgements = payload.get("findingJudgements") or {}
    judgements_text = json.dumps(judgements, ensure_ascii=False, indent=2) + "\n"

    print(f"{key}: deck={deck} rev={rev}")
    print(f"  -> {decisions_path.relative_to(REPO_ROOT)} ({len(decisions)} rows)")
    print(
        f"  -> {judgements_path.relative_to(REPO_ROOT)}"
        f" ({len(judgements.get('judgements', {}))} findings)"
    )
    if dry_run:
        return
    decisions_path.write_text(tsv, encoding="utf-8")
    judgements_path.write_text(judgements_text, encoding="utf-8")


def load_processed_keys() -> set[str]:
    if not STATE_PATH.exists():
        return set()
    try:
        state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return set()
    keys = state.get("processed_keys") or []
    return set(keys) if isinstance(keys, list) else set()


def save_processed_keys(processed: set[str]) -> None:
    STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {"processed_keys": sorted(processed)}
    STATE_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--list", action="store_true", help="list pptx-design-review keys only")
    parser.add_argument("--apply", action="store_true", help="decrypt and write decisions/judgements files")
    parser.add_argument("--key", help="apply only the specified KV key (default: all)")
    parser.add_argument("--dry-run", action="store_true", help="show outputs without writing")
    parser.add_argument(
        "--force",
        action="store_true",
        help="reprocess KV keys even if they appear in the local processed-keys state",
    )
    args = parser.parse_args()

    if not args.list and not args.apply:
        parser.error("specify --list or --apply")

    all_keys = list_keys()
    if args.key:
        target_keys = [args.key]
    else:
        target_keys = all_keys

    entries = filter_app_entries(target_keys)
    if not entries:
        print("no pptx-design-review entries found", file=sys.stderr)
        return

    if args.list:
        processed = load_processed_keys()
        for key, payload in entries:
            mark = " (processed)" if key in processed else ""
            print(f"{key}\tdeck={payload.get('deck')}\trev={payload.get('rev')}{mark}")
        return

    processed = load_processed_keys()
    if not processed and not args.force:
        existing_tsvs = list((REPO_ROOT / "doc" / "reviews").glob("*/rev-*-compare.tsv"))
        if existing_tsvs and not args.dry_run:
            bootstrap_keys = {key for key, _ in entries}
            save_processed_keys(bootstrap_keys)
            print(
                f"Bootstrap: found {len(existing_tsvs)} existing rev-*-compare.tsv "
                f"file(s) but no processed-keys state. Marked all {len(bootstrap_keys)} "
                f"KV key(s) as already processed to preserve local edits. "
                f"Use --force --key <KV-KEY> to re-apply a specific entry."
            )
            return
    newly_processed = set()
    for key, payload in entries:
        if not args.force and key in processed:
            print(f"skip {key}: already processed (use --force to re-apply)")
            continue
        ciphertext = payload.get("ciphertext")
        if not ciphertext:
            print(f"skip {key}: missing ciphertext", file=sys.stderr)
            continue
        decrypted = decrypt(ciphertext)
        write_outputs(decrypted, key, args.dry_run)
        if not args.dry_run:
            newly_processed.add(key)
    if newly_processed:
        save_processed_keys(processed | newly_processed)


if __name__ == "__main__":
    main()
