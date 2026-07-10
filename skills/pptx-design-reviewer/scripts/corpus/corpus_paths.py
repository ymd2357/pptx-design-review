"""コーパス走行ハーネスが参照する固定パス群。

tmp/corpus/ は .gitignore の `tmp/*` ルールで自動的に git 無視される
(= 使い捨て走行の Tier A scratch)。台帳・fixture index は doc/ 配下の
小さなテキストなので git 管理し、ディスク掃除で scratch を消しても
「2度やらない / どんな特殊物を残したか」の記憶が永続する。
"""
from __future__ import annotations

import os
from pathlib import Path

# .../skills/pptx-design-reviewer/scripts/corpus/corpus_paths.py
SCRIPTS_DIR = Path(__file__).resolve().parent.parent
REPO_ROOT = SCRIPTS_DIR.parents[2]

# --- 使い捨て scratch (git 無視) ---
CORPUS_DIR = REPO_ROOT / "tmp" / "corpus"
INBOX_DIR = CORPUS_DIR / "inbox"   # 取得直後の生 PPTX 置き場
WORK_DIR = CORPUS_DIR / "work"     # 1デッキ = 1サブディレクトリの作業物

# --- 永続台帳 (git 管理) ---
LEDGER_TSV = REPO_ROOT / "doc" / "corpus-ledger.tsv"
SOURCES_TSV = REPO_ROOT / "doc" / "corpus-sources.tsv"
FIXTURE_INDEX_TSV = REPO_ROOT / "doc" / "corpus-fixture-index.tsv"

# --- 既存 CLI ---
NORMALIZE = SCRIPTS_DIR / "pptx_normalize_fonts.py"
LINT = SCRIPTS_DIR / "pptx_lint.py"
FIX = SCRIPTS_DIR / "pptx_fix.py"
RENDER_JS = SCRIPTS_DIR / "render_with_vscode_pptx_viewer.js"
CAPTURE_JS = SCRIPTS_DIR / "capture_vscode_pptx_viewer.js"

# render は viewer のソース源 (兄弟 repo vscode-pptx-viewer) を追う。repo の
# extension.ts が openPptx/parseSingleSlide を __parsePptx/__loadSlide として
# 安定 re-export しており、ビルドし直しても壊れない (mangled 名に依存しない)。
# バージョン固定はしない ＝「合わせていく」。env 指定があれば最優先で尊重。
VIEWER_REPO_DIR = REPO_ROOT.parent / "vscode-pptx-viewer"


def resolve_viewer_ext_dir() -> Path | None:
    env = os.environ.get("PPTX_VIEWER_EXT_DIR")
    if env:
        return Path(env)
    if (VIEWER_REPO_DIR / "dist" / "extension.js").exists():
        return VIEWER_REPO_DIR
    # repo が無ければ wrapper 自身の locateExtensionDir (最新インストール版) に委ねる。
    return None

