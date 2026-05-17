# pptx-design-review

PowerPoint (`.pptx`) のデザインレビュー支援ツール群。スライドガイドライン
(`doc/slide-guideline-v1.yml`) を機械的に検査する **lint** と、検出された問題を
自動修正する **fix**、レビュー結果を採否確定するためのモバイル対応 SPA を
合わせた一式。

## 構成

```
doc/
  slide-guideline-v1.yml    # 検査ルール / トークン / DS-001 color repair policy
  tasks.md                  # 観点 / REV / FIX タスク台帳 (= 進捗の単一情報源)
  reviews/<deck>/           # rev-NNN-decisions.tsv / rev-NNN-finding-judgements.json
skills/pptx-design-reviewer/scripts/
  pptx_lint.py              # PPTX 構造 + rendered-image を見る lint 本体
  pptx_fix.py               # findings JSON を消費して PPTX を機械修正する
  pptx_review_orchestrator.py
  test_pptx_*.py            # スモーク + 評価テスト
scripts/
  fetch-reviews.py          # Cloudflare KV から判定 payload を age 復号して取り込み
  build-compare-meta.py     # before/after PNG を ImageMagick で差分比較 → meta.json
  capture-pp-mac.sh         # PowerPoint Mac → PDF → PNG (1 ファイル/スライド)
  publish_review_snapshot.py
web/                        # GH Pages にホストされる SPA (Vite + TypeScript)
tmp/review-snapshot/        # SPA がフェッチする lint.json / 画像群 (gitignore 除外あり)
```

## 主要ワークフロー

### 1. lint → fix → 比較レビュー

1. `pptx_lint.py DECK.pptx --json --no-consolidate > lint.json` で finding を出力。
2. `pptx_fix.py DECK.pptx --findings-json lint.json --apply --rules contrast,card_grid,...`
   で機械修正可能なものを当てる。
3. PowerPoint で before / after を PDF→PNG キャプチャ (`scripts/capture-pp-mac.sh`)。
4. `scripts/build-compare-meta.py --deck <id> --rev NNN` で差分スライドを検出して
   `tmp/review-snapshot/.../rev-NNN/meta.json` を生成。
5. master に push すると `.github/workflows/deploy-review-web.yml` が SPA を
   `review-pages` ブランチに書き出し、GH Pages が
   `/compare/?deck=...&rev=NNN` で公開する。
6. レビューワーは SPA でスライド単位に採用 / 不採用 + メモを付け、age 暗号化して
   `pptx-visual-review.pages.dev/api/feedback` (Cloudflare KV) へ POST。
7. PC で `scripts/fetch-reviews.py --apply` を実行すると KV から復号して
   `doc/reviews/<deck>/rev-NNN-*.tsv` / `*-finding-judgements.json` が書かれる。

### 2. 採否を踏まえた再修正

- `pptx_fix.py --finding-judgements-json rev-NNN-finding-judgements.json` を渡すと、
  人間が `judgement_reason=auto_fixable` と判定した finding を lint の
  `fixability=manual_required` から `auto_fix_candidate` に昇格して再適用できる
  (`FIX-007` の経路)。

## lint / fix 検査の代表例

| Check | 機械修正 | 概要 |
|---|---|---|
| `text_autofit_disabled` | auto | `text_frame.auto_size = NONE` をセット |
| `geometry_rounding` | auto (0.1pt 以内) | drifted bbox を整数 pt に丸め |
| `low_contrast` / `contrast_ratio` | auto (dual strategy) | DS-001 同系色から WCAG 充足候補。fg / bg のうち輝度デルタが小さい側を採用 (`FIX-006`) |
| `card_grid_consistency` | auto (slot 平均) | 行内 card の children を `(top, left)` ソート後の slot index で横断 median。text は slot 幅に揃え、image / shape は cx/cy 保持で位置のみ調整 (`FIX-008`) |
| `text_color_allowlist` | manual + 昇格可 | candidate_values は出るが既定 manual。人間が `auto_fixable` と判定すれば fix で適用 (`FIX-007`) |

## 直近の主な変更

- **FIX-006** (P0): `_contrast_candidate` を foreground / background dual-strategy に
  再設計。`FILL_REPAIR_COLOR_FAMILIES` を新設し、白テキスト on 赤グラデのように
  fg を維持したい局面で bg 側 fill を darken する経路を追加。REV-017 deck の既存
  20 件は従来通り fg-mode で auto-fix、XML 完全一致で視覚的回帰なし。
- **FIX-007** (P1): `pptx_fix.py --finding-judgements-json` を追加。
  `judgement_reason=auto_fixable` の finding を
  `detail.fixability=manual_required → auto_fix_candidate` に昇格してから
  `auto_rules_from_findings` を走らせる。
- **FIX-008** (P1): `card_grid_consistency` の auto-fix を slot ベースに改修。
  v1 の「children を一律 width 拡大」による画像破綻を回避しつつ、v3 の保守的
  平行移動だけでは出なかった「カード間でテキスト幅が揃って見える」効果を実現。
  `pptx_lint` の `row_containers` にも children を埋めた。
- **REV-032**: FIX-008 の視覚確認用 compare レビュー。`p0-rev-017-low-contrast-fixed.pptx`
  を before、FIX-008 適用後を after にして PNG 計 40 枚 + meta.json を
  `tmp/review-snapshot/.../rev-032/` に同梱。changed_slides = [2, 13, 20]
  (slide 7 は XML 上タイトル幅 115.2pt → 168pt 拡張のみで render 差 AE=0)。

## ローカル開発

### lint / fix のテスト

```bash
python3 skills/pptx-design-reviewer/scripts/test_pptx_lint.py
python3 skills/pptx-design-reviewer/scripts/test_pptx_fix.py
python3 skills/pptx-design-reviewer/scripts/test_pptx_fix_evidence_schema.py
```

### SPA の dev

```bash
cd web
npm install
npm run dev
```

ローカルでは `../doc/reviews/`, `../tmp/review-snapshot/` を Vite 経由でフェッチする。
詳細は `web/README.md` を参照。

## デプロイと公開

- master に push されると `.github/workflows/deploy-review-web.yml` が走り、
  `web/dist/` + `doc/reviews/` + `tmp/review-snapshot/` を `review-pages` ブランチへ
  orphan push する。GitHub Pages は `review-pages` ブランチを source として配信。
- 画面ゲート用 PIN は SPA に SHA-256 のみを埋め込んでおり、平文はリポジトリには
  記載しない。Claude (アシスタント) が URL を案内する際に併記する運用。

## ガイドライン文書

- `doc/slide-guideline-v1.yml` — 検査ルール、許容トークン、色 / フォント / レイアウト
  ポリシー。`rules.color.palette` / `repair_candidates` は `pptx_lint` の候補生成と
  対応している。
- `doc/tasks.md` — 観点 (P0-* / P1-* / P2-* / P3-*) と REV / FIX / LINT / WEB タスクを
  単一表で管理。新規作業は必ずここに行を追加する。
- `doc/issue-lint-json-schema.md` — lint JSON の evidence schema 仕様。
- `AGENTS.md` — エージェント (Claude / Codex 等) 向けの作業規約。
