#!/usr/bin/env python3
"""コーパス走行オーケストレータ (1デッキ = 1ラウンド)。

  重複ゲート → font 正規化 → before 描画 → lint → fix → after 描画
            → 再 lint → メトリクス → 収穫判定 → ディスク回収

de-overfit の測定オラクル。未知の実 PPTX で lint/fix を回し、
lint_before / lint_after / checks_fired を台帳に貯めて汎用性を測る。

ディスク規律:
- 平凡なデッキ (既知の check しか出ない) は評価後に workdir を削除。台帳1行だけ残す。
- novelty (既知集合に無い check が発火) を検出したデッキは workdir を残し、
  harvest=1 を台帳に立てる → 人が最小1スライドの回帰 fixture へ昇格させる。
  既知集合はこれまでの台帳 checks_fired + doc/corpus-fixture-index.tsv から
  自己ブートストラップするので、序盤以降 novelty は自然に希少化し reclaim が効く。

視覚評価 (before/after 目視) は SPA 経由でのみ行う運用ルール。ここでは
lint レベルの delta までを機械計測し、人の目が要る回だけ Tier B (SPA) に載せる。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ledger  # noqa: E402
from corpus_paths import (  # noqa: E402
    CAPTURE_JS,
    FIX,
    FIXTURE_INDEX_TSV,
    INBOX_DIR,
    LEDGER_TSV,
    LINT,
    NORMALIZE,
    RENDER_JS,
    WORK_DIR,
    resolve_viewer_ext_dir,
)


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def run(cmd: list[str], log: Path, stdout_path: Path | None = None, env: dict | None = None) -> int:
    """サブプロセスを実行。stderr は log へ、任意で stdout をファイルへ。"""
    with log.open("ab") as errfh:
        errfh.write(f"\n$ {' '.join(cmd)}\n".encode())
        out = stdout_path.open("wb") if stdout_path else None
        try:
            proc = subprocess.run(
                cmd, stdout=out or subprocess.DEVNULL, stderr=errfh, check=False, env=env
            )
        finally:
            if out:
                out.close()
    return proc.returncode


def extract_findings(obj: object) -> list[dict]:
    """lint.json の形状差を吸収して findings 配列を取り出す。"""
    if isinstance(obj, list):
        return [x for x in obj if isinstance(x, dict)]
    if isinstance(obj, dict):
        for key in ("findings", "results", "issues"):
            val = obj.get(key)
            if isinstance(val, list):
                return [x for x in val if isinstance(x, dict)]
        # 入れ子の最初の list-of-dict を探す
        for val in obj.values():
            if isinstance(val, list) and val and isinstance(val[0], dict) and "check" in val[0]:
                return val
    return []


def count_applied(fix_json: object) -> int | None:
    if not isinstance(fix_json, dict):
        return None
    for key in ("applied", "changes", "fixes", "actions", "applied_rules"):
        val = fix_json.get(key)
        if isinstance(val, list):
            return len(val)
        if isinstance(val, int) and not isinstance(val, bool):
            return val
    return None


def load_json(path: Path) -> object:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return None


def render(pptx: Path, out_images: Path, log: Path) -> bool:
    viewer_out = out_images.parent / (out_images.name + "_viewer")
    viewer_out.mkdir(parents=True, exist_ok=True)
    out_images.mkdir(parents=True, exist_ok=True)
    env = dict(os.environ)
    ext_dir = resolve_viewer_ext_dir()
    if ext_dir:
        env["PPTX_VIEWER_EXT_DIR"] = str(ext_dir)
    rc = run(["node", str(RENDER_JS), str(pptx), str(viewer_out)], log, env=env)
    if rc != 0:
        return False
    rc = run(["node", str(CAPTURE_JS), str(viewer_out), str(out_images)], log)
    if rc != 0:
        return False
    return any(out_images.glob("slide-*.png"))


def process_one(pptx: Path, keep: bool, force: bool, rendered_contrast: bool = False) -> str:
    sha = sha256_file(pptx)
    prior = ledger.seen(LEDGER_TSV, sha)
    if prior and prior.get("status") in {"processed", "skipped"} and not force:
        print(f"= 既処理 skip {sha[:8]} ({pptx.name})")
        return "skip"

    work = WORK_DIR / sha[:8]
    if work.exists():
        shutil.rmtree(work)
    work.mkdir(parents=True, exist_ok=True)
    log = work / "run.log"
    base = {
        "sha256": sha,
        "source": (prior or {}).get("source", "local"),
        "source_url": (prior or {}).get("source_url", str(pptx)),
        "filename": pptx.name,
        "fetched_at": (prior or {}).get("fetched_at", now()),
    }

    def fail(note: str) -> str:
        ledger.upsert(LEDGER_TSV, {**base, "status": "error", "processed_at": now(), "note": note})
        print(f"! {note} ({sha[:8]}) — log: {log}")
        return "error"

    # 1) 正規化
    normalized = work / "normalized.pptx"
    if run([sys.executable, str(NORMALIZE), str(pptx), str(normalized)], log) != 0 or not normalized.exists():
        return fail("normalize 失敗")

    # 2) before 描画
    if not render(normalized, work / "before", log):
        return fail("before 描画失敗")

    # rendered-contrast 検査 (--rendered-image-dir) は既定 ON。--no-rendered-contrast
    # で無効化できる (描画自体は視覚 diff 用に常に行う)。
    def rendered_args(images_dir: Path) -> list[str]:
        return ["--rendered-image-dir", str(images_dir)] if rendered_contrast else []

    # 3) lint (before)
    lint_before = work / "lint.json"
    run(
        [sys.executable, str(LINT), str(normalized), "--json", *rendered_args(work / "before")],
        log,
        stdout_path=lint_before,
    )
    findings = extract_findings(load_json(lint_before))
    checks = sorted({f.get("check", "") for f in findings if f.get("check")})
    slide_count = len(list((work / "before").glob("slide-*.png")))

    # 4) fix (--apply は入力を書き換えるので複製に対して適用)
    fixed = work / "fixed.pptx"
    shutil.copy2(normalized, fixed)
    fix_json_path = work / "fix.json"
    run(
        [
            sys.executable, str(FIX), str(fixed),
            "--auto", "--apply", "--backup",
            "--findings-json", str(lint_before),
            *rendered_args(work / "before"),
            "--json",
        ],
        log,
        stdout_path=fix_json_path,
    )
    applied = count_applied(load_json(fix_json_path))

    # 5) after 描画 + 再 lint
    lint_after_count: int | None = None
    if render(fixed, work / "after", log):
        lint_after_path = work / "lint-after.json"
        run(
            [sys.executable, str(LINT), str(fixed), "--json", *rendered_args(work / "after")],
            log,
            stdout_path=lint_after_path,
        )
        lint_after_count = len(extract_findings(load_json(lint_after_path)))

    delta = (lint_after_count - len(findings)) if lint_after_count is not None else None

    # 6) 収穫判定 (novelty = 既知集合に無い check が発火)
    known = ledger.known_checks(LEDGER_TSV, FIXTURE_INDEX_TSV)
    novel = [c for c in checks if c not in known]
    harvest = bool(novel)

    ledger.upsert(
        LEDGER_TSV,
        {
            **base,
            "status": "processed",
            "slide_count": str(slide_count),
            "lint_before": str(len(findings)),
            "fix_applied": "" if applied is None else str(applied),
            "lint_after": "" if lint_after_count is None else str(lint_after_count),
            "delta": "" if delta is None else str(delta),
            "checks_fired": ",".join(checks),
            "harvest": "1" if harvest else "0",
            "processed_at": now(),
            "note": ("novel:" + ",".join(novel)) if novel else "",
        },
    )

    # 7) ディスク回収
    if harvest or keep:
        print(
            f"★ 保持 {sha[:8]} slides={slide_count} lint {len(findings)}→"
            f"{lint_after_count} novel={novel or '-'} → {work}"
        )
        return "harvest" if harvest else "keep"
    shutil.rmtree(work)
    print(f"✓ 回収 {sha[:8]} slides={slide_count} lint {len(findings)}→{lint_after_count} (workdir 削除)")
    return "ok"


def iter_inputs(args: argparse.Namespace) -> list[Path]:
    if args.pptx:
        return [Path(p) for p in args.pptx]
    if not INBOX_DIR.exists():
        return []
    return sorted(INBOX_DIR.glob("*.pptx"))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("pptx", nargs="*", help="処理する PPTX。省略時は inbox/ を全処理。")
    parser.add_argument("--keep", action="store_true", help="平凡なデッキも workdir を削除しない")
    parser.add_argument("--force", action="store_true", help="既処理でも再走行する")
    parser.add_argument("--limit", type=int, default=0, help="処理件数上限 (0=無制限)")
    parser.add_argument(
        "--rendered-contrast",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="lint/fix に --rendered-image-dir を渡し rendered コントラスト検査を有効化 (既定 ON)",
    )
    args = parser.parse_args()

    inputs = iter_inputs(args)
    if not inputs:
        print("処理対象なし。fetch_corpus.py で inbox を満たすか PPTX を指定してください。", file=sys.stderr)
        return 1

    tally: dict[str, int] = {}
    for i, pptx in enumerate(inputs):
        if args.limit and i >= args.limit:
            break
        if not pptx.exists():
            print(f"! 見つからない: {pptx}", file=sys.stderr)
            continue
        result = process_one(
            pptx, keep=args.keep, force=args.force, rendered_contrast=args.rendered_contrast
        )
        tally[result] = tally.get(result, 0) + 1
    print("集計:", ", ".join(f"{k}={v}" for k, v in sorted(tally.items())) or "なし")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
