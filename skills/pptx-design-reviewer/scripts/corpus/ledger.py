"""コーパス走行の重複台帳 (doc/corpus-ledger.tsv) の読み書き。

sha256 を主キーに 取得済み/処理済み/skip/error を記録する。
- 「同じファイルを2度やらない」ゲート = 台帳に sha256 があれば skip。
- checks_fired 列の集合は「これまでのコーパスで観測済みの check」を表し、
  novelty (特殊オブジェクト) 判定の既知集合に使う。
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Optional

FIELDS = [
    "sha256",
    "status",        # fetched | processed | skipped | error
    "source",
    "source_url",
    "filename",
    "fetched_at",
    "slide_count",
    "lint_before",
    "fix_applied",
    "lint_after",
    "delta",         # lint_after - lint_before (負 = 改善)
    "checks_fired",  # カンマ区切りの check id 集合
    "harvest",       # 1 = 特殊オブジェクトとして workdir を残した
    "processed_at",
    "note",
]


def load(path: Path) -> dict[str, dict]:
    rows: dict[str, dict] = {}
    if not path.exists():
        return rows
    with path.open(newline="", encoding="utf-8") as fh:
        for row in csv.DictReader(fh, delimiter="\t"):
            sha = (row.get("sha256") or "").strip()
            if sha:
                rows[sha] = row
    return rows


def save(path: Path, rows: dict[str, dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(
            fh, fieldnames=FIELDS, delimiter="\t", extrasaction="ignore"
        )
        writer.writeheader()
        for sha in sorted(rows):
            writer.writerow(rows[sha])


def upsert(path: Path, row: dict) -> None:
    """既存行があれば非 None の値だけ上書きマージして保存する。"""
    rows = load(path)
    sha = row["sha256"]
    merged = rows.get(sha, {})
    merged.update({k: v for k, v in row.items() if v is not None})
    merged.setdefault("sha256", sha)
    rows[sha] = merged
    save(path, rows)


def seen(path: Path, sha: str) -> Optional[dict]:
    return load(path).get(sha)


def known_checks(ledger_path: Path, fixture_index_path: Path) -> set[str]:
    """これまでに観測済み / fixture 化済みの check id 集合。

    novelty 判定の既知集合。これに無い check が発火したデッキだけを
    「特殊オブジェクト候補」として残す (自己ブートストラップ式)。
    """
    known: set[str] = set()
    for row in load(ledger_path).values():
        for check in (row.get("checks_fired") or "").split(","):
            check = check.strip()
            if check:
                known.add(check)
    if fixture_index_path.exists():
        with fixture_index_path.open(newline="", encoding="utf-8") as fh:
            for row in csv.DictReader(fh, delimiter="\t"):
                check = (row.get("check_id") or "").strip()
                if check:
                    known.add(check)
    return known
