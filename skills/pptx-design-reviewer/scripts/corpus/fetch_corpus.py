#!/usr/bin/env python3
"""公開テンプレ配布サイトから実 PPTX をコーパスへ自動取得する。

doc/corpus-sources.tsv に登録した供給源を巡回し、.pptx への直リンクを
発見してダウンロードする。sha256 が台帳既知なら skip (2度取得しない)。
取得物は tmp/corpus/inbox/ に置き、run_corpus.py がそこから消費する。

供給源 kind:
  listing  URL を GET し、HTML 中の *.pptx アンカーを抽出して全て取得。
           サイトのカテゴリ/検索ページを指すと自動でデッキを発見できる。
  direct   URL 自体が .pptx。1件だけ取得。
  adapter  JS でリンクを組み立てるサイト用の専用アダプタ (未実装スタブ)。

stdlib のみ (urllib)。直リンクを露出しない JS ゲート型サイトは listing では
拾えない → adapter を各サイト向けに書く前提 (骨組みではスタブ)。
"""
from __future__ import annotations

import argparse
import hashlib
import io
import re
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

sys.path.insert(0, str(Path(__file__).resolve().parent))
import ledger  # noqa: E402
from corpus_paths import INBOX_DIR, LEDGER_TSV, SOURCES_TSV  # noqa: E402

UA = "Mozilla/5.0 (pptx-design-review corpus fetcher)"
PPTX_HREF = re.compile(r'href=["\']([^"\']+?\.pptx)(?:["\'?#])', re.IGNORECASE)


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def read_sources(path: Path) -> list[dict]:
    if not path.exists():
        return []
    import csv

    with path.open(newline="", encoding="utf-8") as fh:
        return [
            row
            for row in csv.DictReader(fh, delimiter="\t")
            if (row.get("url") or "").strip() and not (row.get("source_id") or "").startswith("#")
        ]


def http_get(url: str, timeout: int = 30) -> bytes:
    req = Request(url, headers={"User-Agent": UA})
    with urlopen(req, timeout=timeout) as resp:  # noqa: S310 (信頼できる登録済み供給源のみ)
        return resp.read()


def extract_pptx_links(html: str, base_url: str) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for match in PPTX_HREF.finditer(html):
        absolute = urljoin(base_url, match.group(1))
        if absolute not in seen:
            seen.add(absolute)
            out.append(absolute)
    return out


def is_pptx(data: bytes) -> bool:
    if not data.startswith(b"PK\x03\x04"):
        return False
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as zf:
            return "[Content_Types].xml" in zf.namelist()
    except zipfile.BadZipFile:
        return False


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def candidate_urls(source: dict) -> list[str]:
    kind = (source.get("kind") or "listing").strip()
    url = source["url"].strip()
    if kind == "direct":
        return [url]
    if kind == "listing":
        try:
            html = http_get(url).decode("utf-8", errors="replace")
        except Exception as exc:  # noqa: BLE001
            print(f"  ! listing 取得失敗 {url}: {exc}", file=sys.stderr)
            return []
        return extract_pptx_links(html, url)
    if kind == "adapter":
        print(f"  · adapter '{source.get('source_id')}' は未実装 (スタブ)", file=sys.stderr)
        return []
    print(f"  ! 未知の kind '{kind}'", file=sys.stderr)
    return []


def safe_basename(url: str, sha: str) -> str:
    base = Path(urlparse(url).path).name or "deck.pptx"
    base = re.sub(r"[^A-Za-z0-9._-]", "_", base)
    if not base.lower().endswith(".pptx"):
        base += ".pptx"
    return f"{sha[:8]}__{base}"


def fetch(limit: int, only_source: str | None) -> int:
    INBOX_DIR.mkdir(parents=True, exist_ok=True)
    sources = read_sources(SOURCES_TSV)
    if only_source:
        sources = [s for s in sources if s.get("source_id") == only_source]
    if not sources:
        print("供給源がありません。doc/corpus-sources.tsv に登録してください。", file=sys.stderr)
        return 0

    fetched = 0
    for source in sources:
        sid = source.get("source_id", "?")
        print(f"[{sid}] {source.get('url')}")
        for url in candidate_urls(source):
            if fetched >= limit:
                print(f"limit {limit} 到達。停止。")
                return fetched
            try:
                data = http_get(url)
            except Exception as exc:  # noqa: BLE001
                print(f"  ! DL 失敗 {url}: {exc}", file=sys.stderr)
                continue
            if not is_pptx(data):
                print(f"  · PPTX でない/壊れ: {url}", file=sys.stderr)
                continue
            sha = sha256_bytes(data)
            if ledger.seen(LEDGER_TSV, sha):
                print(f"  = 既知 skip {sha[:8]} {url}")
                continue
            filename = safe_basename(url, sha)
            (INBOX_DIR / filename).write_bytes(data)
            ledger.upsert(
                LEDGER_TSV,
                {
                    "sha256": sha,
                    "status": "fetched",
                    "source": sid,
                    "source_url": url,
                    "filename": filename,
                    "fetched_at": now(),
                },
            )
            fetched += 1
            print(f"  + 取得 {sha[:8]} → inbox/{filename}")
    print(f"完了: {fetched} 件を取得。")
    return fetched


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--limit", type=int, default=10, help="この実行で取得する上限件数")
    parser.add_argument("--source", help="doc/corpus-sources.tsv の source_id を1つに絞る")
    parser.add_argument("--list", action="store_true", help="登録済み供給源を表示して終了")
    args = parser.parse_args()

    if args.list:
        for s in read_sources(SOURCES_TSV):
            print(f"{s.get('source_id')}\t{s.get('kind')}\t{s.get('url')}")
        return 0

    fetch(args.limit, args.source)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
