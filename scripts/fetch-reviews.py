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
APP_NAMESPACE = "pptx-design-review"
REPO_ROOT = Path(__file__).resolve().parent.parent


def fetch_json(url: str) -> dict:
    with urllib.request.urlopen(url, timeout=30) as resp:
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


def write_outputs(payload: dict, key: str, dry_run: bool) -> None:
    deck = payload.get("deck")
    rev = payload.get("rev")
    if not deck or not rev:
        print(f"skip {key}: missing deck/rev", file=sys.stderr)
        return

    out_dir = REPO_ROOT / "doc" / "reviews" / deck
    out_dir.mkdir(parents=True, exist_ok=True)

    decisions_path = out_dir / f"rev-{rev}-decisions.tsv"
    judgements_path = out_dir / f"rev-{rev}-finding-judgements.json"

    decisions = payload.get("decisions") or []
    if not isinstance(decisions, list):
        print(f"skip {key}: decisions not list", file=sys.stderr)
        return
    tsv = render_decisions_tsv(decisions)

    judgements = payload.get("findingJudgements") or {}
    judgements_text = json.dumps(judgements, ensure_ascii=False, indent=2) + "\n"

    print(f"{key}: deck={deck} rev={rev}")
    print(f"  -> {decisions_path.relative_to(REPO_ROOT)} ({len(decisions)} rows)")
    print(f"  -> {judgements_path.relative_to(REPO_ROOT)} ({len(judgements.get('judgements', {}))} findings)")
    if dry_run:
        return
    decisions_path.write_text(tsv, encoding="utf-8")
    judgements_path.write_text(judgements_text, encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--list", action="store_true", help="list pptx-design-review keys only")
    parser.add_argument("--apply", action="store_true", help="decrypt and write decisions/judgements files")
    parser.add_argument("--key", help="apply only the specified KV key (default: all)")
    parser.add_argument("--dry-run", action="store_true", help="show outputs without writing")
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
        for key, payload in entries:
            print(f"{key}\tdeck={payload.get('deck')}\trev={payload.get('rev')}")
        return

    for key, payload in entries:
        ciphertext = payload.get("ciphertext")
        if not ciphertext:
            print(f"skip {key}: missing ciphertext", file=sys.stderr)
            continue
        decrypted = decrypt(ciphertext)
        write_outputs(decrypted, key, args.dry_run)


if __name__ == "__main__":
    main()
