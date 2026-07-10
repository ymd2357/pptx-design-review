# コーパス駆動 de-overfit ループ

少数デッキ (seminar / timeecare) で閾値・分岐を詰めた lint/fix を、
**未知の実 PPTX コーパスを回してオーバーフィットを炙り出し、汎用化する**枠組み。
コーパスは「汎用性のテストオラクル」であって成果物ではない。走行の作業物は
使い捨て、貯めるのは台帳・メトリクス・収穫した最小 fixture だけ。

## パイプライン (1デッキ = 1ラウンド)

```
取得 → 重複ゲート → font正規化 → before描画 → lint → fix → after描画
     → 再lint → メトリクス → 収穫判定 → ディスク回収
```

| 段 | 実体 |
|---|---|
| 取得 | `corpus/fetch_corpus.py` — 公開テンプレ配布サイトを巡回し .pptx を inbox へ |
| 重複ゲート | 生 PPTX の sha256 を `doc/corpus-ledger.tsv` 照合。既知なら skip |
| 正規化 | `pptx_normalize_fonts.py` (Mac typeface → family-only。全 lint/fix の最上位前提) |
| 描画 | `render_with_vscode_pptx_viewer.js` + `capture_vscode_pptx_viewer.js` (Chromium 直 render のみ。PDF/LibreOffice 禁止) |
| lint/fix | `pptx_lint.py --json` / `pptx_fix.py --auto --apply` |
| メトリクス | lint_before / lint_after / delta / checks_fired を台帳へ |
| 収穫判定 | novelty を検出したデッキだけ workdir を残す |
| 回収 | 平凡なデッキは workdir 削除、台帳1行だけ残す |

## 汎用化 (de-overfit) の測り方

デッキ名依存は無い。オーバーフィットは**閾値・分岐の暗黙チューニング**として出る。

- rule ごとの fire-rate をコーパス横断で記録 (`checks_fired` 列)。tuning デッキで
  高精度・未知デッキで崩壊 = overfit のシグネチャ。
- `pptx_lint.py` に散在するマジック定数を named 閾値表に集約し、緩める対象を可視化。
- デッキ固有の暗黙前提 (例「タイトルは最上段シェイプ」) を N デッキ検証済みの
  構造的・相対的ロジックへ置換。
- どの汎用化変更も既存回帰 fixture (`test_pptx_lint.py` / `test_pptx_fix.py`) を必ず通す。

## 特殊オブジェクトの保持 (novelty / 収穫)

残す価値 = **既知集合に無い check が発火した**デッキ = 未カバーの挙動を捕獲したもの。

- 既知集合 = これまでの台帳 `checks_fired` + `doc/corpus-fixture-index.tsv`。
  自己ブートストラップなので序盤以降 novelty は自然に希少化し、reclaim が効く。
- 収穫は該当**1スライドに間引いた最小再現** + 期待 lint/fix を fixture 化する
  (デッキ丸ごとでなく1枚 → ディスク極小)。昇格したら `corpus-fixture-index.tsv` に登録。
- `harvest=1` の台帳行が「人が昇格させるべき候補」。workdir は削除されず残る。

## ディスク運用 (2ティア)

- **Tier A (使い捨て走行)**: `tmp/corpus/` (`.gitignore` の `tmp/*` で自動無視)。
  評価後に workdir 削除。SPA に載せない。大半はこれ。
- **Tier B (人の目が要る回)**: `tmp/review-snapshot/` + SPA。少数だけ。
  before/after の目視評価は SPA 経由でのみ行う (ローカル PNG で済ませない)。
- 台帳・fixture index は `doc/` の小さなテキストなので git 管理。scratch を
  掃除しても「2度やらない / どんな特殊物を残したか」の記憶は永続する。

## 使い方

```bash
cd skills/pptx-design-reviewer/scripts/corpus

# 供給源を登録 (doc/corpus-sources.tsv) してから取得
python3 fetch_corpus.py --limit 10
python3 fetch_corpus.py --list          # 登録済み供給源を確認

# inbox を全処理 (平凡は自動削除・novelty は保持)
python3 run_corpus.py
python3 run_corpus.py path/to/deck.pptx # 単発
python3 run_corpus.py --keep            # 全 workdir 保持 (デバッグ)
```

台帳 `doc/corpus-ledger.tsv` を見て、`delta > 0` (fix で lint 悪化) や
`harvest=1` の行を優先的に掘る。

## エンジン破損の扱い (走行前提)

2026-07-10 のコーパス走行で判明。詳細は memory `reference-render-lint-pipeline-breakages`。

- **[解決済] viewer render の mangled 名依存**: wrapper が bundle の mangled 名
  (`Ad`/`Nd`) 決め打ちで、viewer をビルドし直すたび `Nd is not defined` で破綻して
  いた。バージョン固定でなく、viewer repo の `src/extension.ts` が openPptx/
  parseSingleSlide を `__parsePptx`/`__loadSlide` として**安定 re-export** し、
  wrapper がその安定名を読む形へ変更 (esbuild CJS は export 名を mangle しない)。
  ハーネスは `resolve_viewer_ext_dir()` で viewer repo を追う (固定しない=合わせていく)。
- **[解決済] `pptx_lint.py --rendered-image-dir` の `get_flattened_data` crash**:
  Pillow に無い `crop.get_flattened_data()` を `list(crop.getdata())` に修正
  (pptx_lint.py)。ハーネスは rendered コントラスト検査を**既定 ON** に
  (`--no-rendered-contrast` で無効化可)。timeecare デッキで `contrast_ratio`×7 /
  `low_contrast`×1 の陽性発火を確認済み。
