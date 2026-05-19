# Tasks

このファイルをローカルのタスク管理表として使う。
詳細な背景リストは `doc/design-system-review-v0.md`、
実装済みルールは `doc/slide-guideline-v1.yml` を参照する。

## 運用ルール

状態:

- `todo`: 未着手
- `doing`: 作業中
- `blocked`: 外部判断待ち
- `done`: 完了

優先度:

- `P0`: 次に着手する
- `P1`: P0 完了後に着手
- `P2`: 後続で着手

各タスクは、完了条件を満たしたら `done` に変更する。

## Pn 見直しメモ

現行の P0-P3 は、実装済み lint チェックを後追いで束ねた暫定分類であり、
レビュー観点全体から逆算した確定体系ではない。
そのため、今後の基本チェック追加に伴って Pn 項目は増える前提で扱う。

Pn は finding 数ではなく、納品物への影響度で決める。
特に、文字衝突、オブジェクト重なり、読めないコントラスト、欠落、
意味が変わる折返し、画像の重要部分の切れ、階層崩れ、
テキストボックス内の収まり・上下余白バランスは、
現行 lint に未実装または未整理でも P0-P2 判定対象になり得る。

次の lint 追加は、個別実装の前に Pn 定義を棚卸しし、
各チェックが P0/P1/P2/P3 のどこに落ちるか、
機械検出・手動レビュー・auto-fix 対象のどれかを明示してから進める。

## チェック観点一覧

正本は `doc/slide-guideline-v1.yml` の
`rules.lint.priority_catalog` とする。
この表は、進行中のレビューで Pn 判定を確認するための索引であり、
件数順の finding 一覧ではない。
観点Noは `priority_catalog` 内の Pn 別枝番であり、
下段のレビュー作業IDとは別に扱う。

| 観点No | ID | 状態 | 検出 | 修正 | 観点 |
| --- | --- | --- | --- | --- | --- |
| P0-1 | `text_encoding` | implemented_priority | automated | manual_review | 文字化けで読めない |
| P0-2 | `text_overlap` | implemented_lint | automated | manual_review | 文字衝突で読めない |
| P0-3 | `low_contrast` | implemented_lint | automated | manual_review | 読めないコントラスト |
| P0-4 | `overflow_text` | implemented_lint | automated | manual_review | テキストが切れて読めない |
| P1-1 | `animation_present` | implemented_lint | automated | manual_review | 静的配布で失われる情報 |
| P1-2 | `alt_text_required` | implemented_lint | automated | manual_review | 意味を持つ画像の代替テキスト欠落 |
| P1-3 | `color_only_meaning` | implemented_lint | automated | manual_review | 色だけで意味を伝えている |
| P1-4 | `contrast_ratio` | implemented_lint | automated | manual_review | コントラスト不足 |
| P1-5 | `heading_hierarchy_broken` | implemented_lint | automated | manual_review | 見出し・本文のデザインシステム階層不一致 |
| P1-6 | `image_aspect_distortion` | implemented_lint | automated | manual_review | 画像の縦横比崩れ |
| P1-7 | `key_area_cropped` | implemented_lint | automated | manual_review | 画像重要部の欠け |
| P1-8 | `line_height` | implemented_lint | automated | manual_review | 行間が読みやすさを損なう |
| P1-9 | `missing_required_element` | implemented_lint | automated | manual_review | 必須要素欠落 |
| P1-10 | `object_overlap` | implemented_lint | automated | manual_review | オブジェクト重なり |
| P1-11 | `overflow_images` | implemented_lint | automated | manual_review | 画像のスライド外はみ出し |
| P1-12 | `overflow_shapes` | implemented_lint | automated | manual_review | 図形のスライド外はみ出し |
| P1-13 | `reading_order` | implemented_lint | automated | manual_review | 読み順不整合 |
| P1-14 | `text_autofit_disabled` | implemented_lint | automated | auto_fix | PowerPoint自動縮小による可読性リスク |
| P1-15 | `wrap_break_changes_meaning` | implemented_lint | automated | manual_review | 折返しで意味が変わる |
| P2-1 | `object_gap_too_small` | implemented_lint | automated | manual_review | 隣接オブジェクト間隔が狭い |
| P2-2 | `background_color_palette` | implemented_lint | automated | manual_review | 背景・塗り色がパレット外 |
| P2-3 | `font_family` | implemented_lint | automated | manual_review | 書体がテンプレート外 |
| P2-4 | `font_size_scale` | implemented_lint | automated | manual_review | 文字サイズが定義スケール外 |
| P2-5 | `image_upscale_ratio` | implemented_lint | automated | manual_review | 画像解像度不足 |
| P2-6 | `inner_padding_imbalance` | implemented_lint | automated | manual_review | テキストボックス内余白バランス不自然 |
| P2-7 | `safe_margins` | implemented_lint | automated | manual_review | 非テキスト要素が安全余白外 |
| P2-8 | `safe_text_area_text` | implemented_lint | automated | manual_review | テキストが安全領域外 |
| P2-9 | `slide_size` | implemented_lint | automated | manual_review | スライドサイズが基準比率外 |
| P2-10 | `text_color_allowlist` | implemented_lint | automated | manual_review | 文字色が許可リスト外 |
| P2-11 | `alignment_left_top` | implemented_lint | automated | manual_review | 文字揃えがテンプレート基準外 |
| P2-12 | `card_grid_consistency` | implemented_lint | automated | manual_review | 同種カード群のサイズ・内側配置が不統一 |
| P2-13 | `text_vertical_balance` | implemented_lint | automated | manual_review | テキストボックス内の縦余白バランス不自然 |
| P3-1 | `geometry_rounding` | implemented_lint | automated | auto_fix | 座標の微小な丸めズレ |

## デザインシステム整備タスク

| ID | 状態 | 優先度 | 内容 | 完了条件 |
| --- | --- | --- | --- | --- |
| DS-001 | done | P0 | 色tokenを raw palette / semantic role / usage policy / pair policy / repair policy に分離する | `low_contrast` / `contrast_ratio` の修復候補が、元色に最も近いデザインシステム色の同系色からコントラストを満たす色を選ぶ |
| DS-002 | partial | P1 | design-system loader を作り、lint/fix のハードコード色定数を guideline YAML 参照へ移す | **両バージョン並走 + env switch を実装**。`doc/slide-guideline-v1.yml` の `rules.color.lint_palette` 節に 5 構造 (allowed_text/fill, *_color_token_by_hex, contrast_repair / fill_repair families) を明示。`design_system_loader.py` が `LintPalette` データクラスに読み込み、`pptx_lint.py` は `PPTX_PALETTE_SOURCE=yaml` 時に hardcoded 値を YAML 由来で置換 (default は hardcoded のまま)。`test_design_system_loader.py` が両 source の完全一致を assert、`test_pptx_lint` / `test_pptx_lint_evidence_schema` / `test_pptx_fix` / `test_pptx_fix_evidence_schema` も両 source で pass。**残り**: pptx_fix.py 側のハードコード定数 (まだ無いが今後参照するかも) の取り込み、hardcoded path 退役の判断 (= guideline 改訂時の運用フローが固まってから)。 |
| DS-003 | done | P1 | design-system 自体のレビューを追加する | `design_system_review.py` で 5 check を実装: (1) `palette_family_tier_gap`、(2) `repair_candidate_family_overlap`、(3) `repair_candidate_palette_drift`、(4) `palette_token_usage_uncovered` (allowlist token ref を exact + prefix match で照合、palette 各色が少なくとも 1 ヶ所で被覆されているかを検証)、(5) `contrast_pair_unclassified` (allowed text gray × allowed background のうち `text_on_background` / `prohibited` のどちらにも書かれていない pair)。`data_series` を flat list → series_1..series_8 dict 構造に直して `repair_candidates` 側を `series_7` (= #56B4E9) に揃え、drift 1 件は根治。現 YAML against で 22 件 (uncovered 14 + unclassified pair 8、全 warn) — これは hardcoded 経由の挙動には影響しない (lint output には繋がっていない監査結果)。`test_design_system_review.py` で per-check ベースラインを lock。 |

## 260329 seminar deck 現状メモ

対象:
`tmp/review/260329-seminar-curriculum-proposal/260329_seminar_curriculum_proposal.p0-p2-15-geometry-fixed.pptx`

### レビュー作業記録

レビュー作業IDは優先度を含まない作業通番とする。
`p2-N` は過去の artifact ファイル名に残る接頭辞であり、
チェック観点Noや優先度を表すものとして使わない。

- `REV-000` (`p0`): 初期 fix と目視。
  対応 check はなし。
  主な artifact は `260329_seminar_curriculum_proposal.p0-fixed.pptx`、
  `p0-review-images/`。
- `REV-002` (`p2-2`): 目視レビュー。
  対応 check は未整理。主な artifact は `p2-2-review-images/`。
- `REV-003` (`p2-3`): 目視レビュー。
  対応 check は未整理。主な artifact は `p2-3-review-images/`。
- `REV-004` (`p2-4`): 目視レビュー。
  対応 check は未整理。主な artifact は `p2-4-review-images/`。
- `REV-005` (`p2-5`): サンプル生成と目視。
  対応 check は未整理。
  主な artifact は `p0-p2-5-sample-fixed.pptx`、
  `p2-5-sample-images/`。
- `REV-006` (`p2-6`): `line_height` (`P1-8`) 修正。
  主な artifact は `p0-p2-6-line-height-only-fixed.pptx`、
  `p0-p2-6-fixed.pptx`、`p2-6-powerpoint-review-images/`。
- `REV-007` (`p2-7`): `font_family` (`P2-3`) 修正。
  主な artifact は `p0-p2-7-font-family-fixed.pptx`、
  `p2-7-powerpoint-review-images/`。
- `REV-008` (`p2-8`): `font_size_scale` (`P2-4`) 修正。
  主な artifact は `p0-p2-8-font-size-fixed.pptx`、
  `p2-8-powerpoint-review-images/`。
- `REV-009` (`p2-9`): `object_overlap` (`P1-10`) の構造メタ化。
  commit は `e2790e1`。
  主な artifact は `p2-9-evaluation-{lint,priorities,structure}.json`、
  `p2-9-evaluation-images/`。
- `REV-010` (`p2-10`): `overflow_images` (`P1-11`) のルール修正。
  commit は `e680e62`。
  主な artifact は `p2-10-evaluation-{lint,priorities}.json`。
- `REV-011` (`p2-11`): 評価のみ。
  対応 check はなし。
  主な artifact は `p2-11-after-{lint,priorities}.json`、
  `p2-11-evaluation-{lint,priorities}.json`。
- `REV-012` (`p2-12`): `object_gap_too_small` (`P2-1`) 修正。
  主な artifact は `p0-p2-12-object-gap-fixed.pptx`、
  `p2-12-review-images/`。
- `REV-013` (`p2-13`): `alignment_left_top` (`P2-11`) 修正。
  主な artifact は `p0-p2-13-left-align-fixed.pptx`、
  `p2-13-review-images/`、`p2-13-review-images-corrected/`。
- `REV-014` (`p2-14`): `text_vertical_balance` (`P2-13`) と
  `font_size_scale` (`P2-4`) のルール修正。
  commit は `0044959`、`b944b86`。
  主な artifact は `p2-14-after-{lint,priorities}.json`。
- `REV-015` (`p2-15`): `geometry_rounding` (`P3-1`) の一部修正。
  浮動小数点誤差由来のみ 54 件を auto-fix。
  残 129 件は `.25/.5/.75pt` 単位のため auto-fix 対象外。
  主な artifact は `p0-p2-15-geometry-fixed.pptx`、
  `p2-15-after-{lint,priorities}.json`。
- `REV-017` (`p0-3`): `low_contrast` (`P0-3`) の auto-fix を before/after
  比較 UI でスライド単位採否レビュー。
  入力 `p0-p2-15-geometry-fixed.pptx` に `pptx_lint --no-consolidate +
  pptx_fix --rules contrast` を当てて 20 件の灰色テキスト色値置換
  (`#999999→#707070` 等) を反映。
  20 スライドすべて「採用」で確定 (`rev-017-compare.tsv`)。
  主な artifact は `p0-rev-017-low-contrast-fixed.pptx`、
  `tmp/review-snapshot/.../rev-017/images/{before,after,diff}/`、
  `doc/reviews/.../rev-017-compare.tsv`。
  なお赤背景の白文字など視覚的に大胆な修正は最新 lint が検出して
  おらず未反映 (= `FIX-006` で継続)。
- `REV-019` (`p2-12`): `card_grid_consistency` (`P2-12`) の auto-fix を
  before/after 比較 UI でスライド単位採否レビュー。
  入力 `p0-rev-017-low-contrast-fixed.pptx` に対し
  `pptx_lint --no-consolidate + pptx_fix --rules card_grid` を当てて
  4 finding (slide 2, 7, 13, 20) を container 単位 10 action に展開し
  6 件 apply / 4 件 manual_required で反映。
  初版 (v1) は children の width を一律 `target_child_width` に拡大して
  image を壊し、レビューで全件 reject された。Codex デバッグの結果、
  pptx_fix.py を「children の width/height を保持し平行移動のみ」に
  修正 (v3)。視覚的変化は微小 (padding 微調整のみ) だが、画像破壊は
  解消し保守的修正として採用。
  主な artifact は `p0-rev-017-low-contrast-fixed.pptx` (= 採用後の
  正本に変更なし、card_grid 修正の effect が微小なため上書きせず)、
  `rev-019-card-grid-autofix-v3.pptx`、
  `tmp/review-snapshot/.../rev-019/images/{before,after,diff}/`、
  `doc/reviews/.../rev-019-compare.tsv`。
  賢い修正 (shape kind 別の処理、image は縦横比保持で再配置、text は
  inner_width に揃え) は `FIX-008` で継続。

種別:

- `fix`: PPTX 自体を修正して新しい正本候補を生成した回
- `rule`: lint ルール側を改修し、PPTX は更新していない回（評価 JSON のみ）
- `review`: 目視レビューと finding 整理のみ
- `sample`: サンプル生成と目視
- `eval`: 既存 PPTX に対する評価のみ

進行上の前提:

- レビュー対象の元ファイルは
  `/Users/yamadakenichi/Documents/GitHub/vscode-pptx-viewer/samples/260329_seminar_curriculum_proposal.pptx`。
- 直近の正本候補は
  `260329_seminar_curriculum_proposal.p0-p2-15-geometry-fixed.pptx`。
  それ以前の `p2-6` から `p2-14` の中間生成物は比較・確認用として扱う。
- 画像 DIFF は PowerPoint 書き出しで作る。LibreOffice は使わない。
- 直近の確認用 DIFF:
  `tmp/review/260329-seminar-curriculum-proposal/p2-15-powerpoint-review-images/diff-vertical/index.html`
- 既に `_invalid-do-not-use/` に移した旧生成物は、判断材料として使わない。
- 生成 PPTX、PNG、PDF、HTML などのレビュー成果物は git に入れない。

2026-04-29 時点の `pptx_lint.py --json` 残件
(`REV-008` / `p2-8` fix 直後の値、参考):

| check | count | severity |
| ------- | ------- | ---------- |
| `overflow_images` | 1 | error |
| `safe_margins` | 6 | warning |
| `safe_text_area_text` | 27 | warning |
| `alignment_left_top` | 2 | warning |
| `geometry_rounding` | 138 | warning |

`REV-015` (`p2-15`) 後 (2026-05-03 時点) の状況:

- `severity=error` は 0 件、`pptx_review_priorities.py` の P0/P1 該当も 0 件
- `overflow_images` 1 件 → `REV-010` (`p2-10`) の装飾ラスター除外ルールで消滅
- `alignment_left_top` 2 件 → `REV-013` (`p2-13`) の left-align fix で解消
- `geometry_rounding` 138 → 129
  (`REV-015` / `p2-15` で浮動小数点誤差由来 54 件を解消)
- 残 lint: `geometry_rounding` 129, `safe_text_area_text` 27,
  `inner_padding_imbalance` 17, `safe_margins` 1
- 残 `geometry_rounding` 129 件はすべて `.25/.5/.75pt` 単位で、
  グリッド計算上の意図的な値の可能性が高い。auto-fix で integer に丸めると
  レイアウトが壊れるため対象外。
- 配布ゲートとしては止まる要因なし。次の作業は `REV-016` とし、
  新規 artifact 接頭辞は `rev-016-*` を使う。
  候補は
  (a) PowerPoint 書き出しで目視 DIFF を取り finalize するか、
  (b) 同種カード群の内側配置を `card_grid_consistency` として確認するかの
  どちらかが候補。

### 既存 artifact / evidence 復元メモ

`tmp/review/260329-seminar-curriculum-proposal/` 配下の既存 artifact は、
過去の `p2-N` 接頭辞を含む。これは現在の Pn 観点Noとは一致しない。
新規 artifact ではこの接頭辞を増やさず、既存分は下表で採用状態を読む。

| REV | 既存 artifact / evidence | 採用状態 | メモ |
| --- | --- | --- | --- |
| `REV-000` | `p0-fixed.pptx`, `p0-review-images/` | superseded | 初期 fix。後続 REV の入力系列に吸収済み。 |
| `REV-002` | `p0-p2-2-fixed.pptx`, `p2-2-review-images/` | superseded | 目視レビュー回。対応 check は未整理。 |
| `REV-003` | `p0-p2-3-fixed.pptx`, `p2-3-review-images/` | superseded | 目視レビュー回。対応 check は未整理。 |
| `REV-004` | `p0-p2-4-fixed.pptx`, `p2-4-review-images/` | superseded | 目視レビュー回。対応 check は未整理。 |
| `REV-005` | `p0-p2-5-sample-fixed.pptx`, `p2-5-sample-images/` | sample | `p2-5-review-images/` も存在。採用判断は未整理。 |
| `REV-006` | `p0-p2-6-line-height-only-fixed.pptx`, `p0-p2-6-fixed.pptx` | superseded | `p2-6-powerpoint-review-images/` が主確認 evidence。sample 系画像も存在。 |
| `REV-007` | `p0-p2-7-font-family-fixed.pptx`, `p2-7-powerpoint-review-images/` | superseded | `font_family` (`P2-3`) 修正回。 |
| `REV-008` | `p0-p2-8-font-size-fixed.pptx`, `p2-8-powerpoint-review-images/` | superseded | `font_size_scale` (`P2-4`) 修正回。2026-04-29 残件表の基準。 |
| `REV-009` | `p2-9-evaluation-{lint,priorities,structure}.json`, `p2-9-evaluation-images/` | rule_evidence | PPTX 正本更新ではなく `object_overlap` (`P1-10`) ルール評価。 |
| `REV-010` | `p2-10-evaluation-{lint,priorities}.json` | rule_evidence | `overflow_images` (`P1-11`) ルール評価。PPTX 正本更新なし。 |
| `REV-011` | `p2-11-after-{lint,priorities}.json`, `p2-11-evaluation-{lint,priorities}.json` | eval_only | 評価のみ。採用 deck を進めた回ではない。 |
| `REV-012` | `p0-p2-12-object-gap-fixed.pptx`, `p2-12-review-images/` | superseded | `object_gap_too_small` (`P2-1`) 修正。後続 REV に吸収済み。 |
| `REV-013` | `p0-p2-13-left-align-fixed.pptx`, `p2-13-review-images/` | adopted_then_superseded | `alignment_left_top` (`P2-11`) 修正。`p2-13-review-images-corrected/` も存在。 |
| `REV-014` | `p2-14-after-{lint,priorities}.json` | rule_evidence | `text_vertical_balance` (`P2-13`) ルール評価。`p2-14-after-lint.json` は `alignment_left_top` 2 件を含むため、正本系列の採用状態としては使わない。 |
| `REV-015` | `p0-p2-15-geometry-fixed.pptx`, `p2-15-after-{lint,priorities}.json` | current_candidate | 現在の正本候補。`p0-p2-13-left-align-fixed.pptx` へ geometry fix を適用した系列。 |
| `REV-015` | `p2-15-powerpoint-review-images/`, `p2-15-diff-stats.json` | visual_evidence | PowerPoint 書き出し DIFF。最大差分は slide 10 の 0.0712%。 |
| n/a | `_invalid-do-not-use/` | rejected | 旧生成物。判断材料として使わない。 |
| n/a | `review-images/`, `before.txt`, `after.txt`, `priorities-after.md` | legacy_evidence | 初期確認ログ。観点別採用判断へ復元するときの補助材料。 |
| n/a | `powerpoint-test/`, `.pptx.bak`, `.DS_Store` | local_noise | 採用判断の根拠にしない。 |

### REV-016 観点別レビュー判定表

復元用の生成物は
`tmp/review/260329-seminar-curriculum-proposal/rev-016-lint-timeline.tsv`
と
`tmp/review/260329-seminar-curriculum-proposal/rev-016-artifact-map.tsv`
に置く。どちらもローカル evidence であり git には入れない。

判定:

- `done`: 修正またはルール評価が後続の正本候補に反映済み。
- `inferred_done`: 最新 lint / priority では検出 0 件だが、
  個別の目視 evidence は復元できていない。
- `remaining`: 最新 lint に残件があり、許容または修正判断が残る。
- `not_recorded`: 手動レビュー観点として定義済みだが、この deck での
  evidence が復元できていない。
- `not_applicable`: この deck の確認対象外。

最新 lint 件数は `REV-017` の
`rev-017-rendered-contrast-lint.json` を基準にする。
`REV-014` の `p2-14-after-lint.json` は `alignment_left_top` 2 件を含むため、
正本系列の採用状態としては使わない。

| 観点No | ID | 判定 | 最新 lint | 目視確認 | 対応 REV | 根拠 / メモ |
| --- | --- | --- | ---: | --- | --- | --- |
| P0-1 | `text_encoding` | inferred_done | n/a | automated_only | `REV-015` | `p2-15-after-priorities.json` で P0/P1 なし。 |
| P0-2 | `text_overlap` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P0-3 | `low_contrast` | done | 0 (auto-fix 後) | rendered_automated | `REV-017` | `p0-rev-017-low-contrast-fixed.pptx` で灰色テキスト 20 件を `pptx_fix --rules contrast` で色値差し替え (`#999999→#707070` 等)、スライド単位の比較 UI で全 20 件「採用」確定 (`rev-017-compare.tsv`)。`FIX-006` で dual-strategy (foreground / background) repair candidate を実装し、白テキスト on 赤グラデなど fg を維持したい局面で bg 側 fill を darken する経路を追加 (回帰なし: 既存 20 件は fg-mode で同色置換)。 |
| P0-4 | `overflow_text` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P1-1 | `animation_present` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P1-2 | `alt_text_required` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P1-3 | `color_only_meaning` | inferred_done | 0 | automated_only | `REV-017` | 機械 lint で未検出。 |
| P1-4 | `contrast_ratio` | done | 0 (auto-fix 後) | rendered_automated | `REV-017`, `REV-018` | REV-017 の `pptx_fix --rules contrast` が low_contrast と contrast_ratio を一括処理したため、REV-018 として独立した修正は不要。`p0-rev-017-low-contrast-fixed.pptx` 上で XML / rendered 双方 lint が 0 件で確認済。 |
| P1-5 | `heading_hierarchy_broken` | inferred_done | 0 | automated_only | `REV-017` | 機械 lint で未検出。 |
| P1-6 | `image_aspect_distortion` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P1-7 | `key_area_cropped` | inferred_done | 0 | automated_only | `REV-017` | 機械 lint で未検出。 |
| P1-8 | `line_height` | done | 0 | visual_done | `REV-006` | `p0-p2-6-*.pptx` と `p2-6-powerpoint-review-images/`。 |
| P1-9 | `missing_required_element` | inferred_done | 0 | automated_only | `REV-017` | `p0-rev-017-low-contrast-fixed.pptx` に対する最新 lint (XML / rendered 双方) で 0 件。旧 lint で 3 スライド検出されていた候補欠落は最新 lint アルゴリズムで消えた。 |
| P1-10 | `object_overlap` | done | 0 | automated_only | `REV-009` | `p2-9-evaluation-structure.json` で構造メタ化。最新 lint で未検出。 |
| P1-11 | `overflow_images` | done | 0 | automated_only | `REV-010` | 装飾ラスター除外ルールで解消。最新 lint で未検出。 |
| P1-12 | `overflow_shapes` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P1-13 | `reading_order` | inferred_done | 0 | automated_only | `REV-017` | `p0-rev-017-low-contrast-fixed.pptx` に対する最新 lint で 0 件。旧 lint で 5 スライド検出されていた source/visual order inversion は最新 lint で消えた。 |
| P1-14 | `text_autofit_disabled` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P1-15 | `wrap_break_changes_meaning` | inferred_done | 0 | automated_only | `REV-017` | 機械 lint で未検出。 |
| P2-1 | `object_gap_too_small` | done | 0 | visual_done | `REV-012` | `p0-p2-12-object-gap-fixed.pptx` と `p2-12-review-images/`。 |
| P2-2 | `background_color_palette` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P2-3 | `font_family` | done | 0 | visual_done | `REV-007` | `p0-p2-7-font-family-fixed.pptx` と PowerPoint review images。 |
| P2-4 | `font_size_scale` | done | 0 | visual_done | `REV-008` | `p0-p2-8-font-size-fixed.pptx` と PowerPoint review images。 |
| P2-5 | `image_upscale_ratio` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P2-6 | `inner_padding_imbalance` | inferred_done | 0 | automated_only | `REV-015`, `REV-017` | `p0-rev-017-low-contrast-fixed.pptx` に対する最新 lint で 0 件。旧 lint で残っていた 17 件は最新 lint で消えた。 |
| P2-7 | `safe_margins` | inferred_done | 0 | automated_only | `REV-015`, `REV-017` | `p0-rev-017-low-contrast-fixed.pptx` に対する最新 lint で 0 件。旧 lint で残っていた 1 件は最新 lint で消えた。 |
| P2-8 | `safe_text_area_text` | inferred_done | 0 | automated_only | `REV-015`, `REV-017` | `p0-rev-017-low-contrast-fixed.pptx` に対する最新 lint で 0 件。旧 lint で残っていた 27 件は最新 lint で消えた。 |
| P2-9 | `slide_size` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P2-10 | `text_color_allowlist` | inferred_done | 0 | automated_only | `REV-015` | 最新 lint で未検出。 |
| P2-11 | `alignment_left_top` | done | 0 | visual_done | `REV-013` | `p0-p2-13-left-align-fixed.pptx`。`REV-014` の 2 件残りは正本系列に採用しない。 |
| P2-12 | `card_grid_consistency` | done | 0 (auto-fix 後) | needs_visual_judgment | `REV-019` | `pptx_fix --rules card_grid` の保守的修正 (子要素の bbox 保持 + 平行移動のみ、container 高さは group_medians に揃え) を 6 container に適用。children が inner box を超える 4 container は manual_required で skip。視覚変化は微小だが画像破壊なしで採用。賢い修正 (shape kind 別の処理) は `FIX-008` で継続。 |
| P2-13 | `text_vertical_balance` | done | 0 | automated_only | `REV-013`, `REV-014` | `REV-013` 系列で 0 件。`REV-014` はルール evidence のみで正本採用には使わない。 |
| P3-1 | `geometry_rounding` | remaining | 129 | visual_done | `REV-015` | 残件は `.25/.5/.75pt` 単位。auto-fix 対象外として defer。 |

### REV-017 判定運用

REV-017 完了条件の「判断記録」は、以下の二箇所のいずれかで持つ。
新規枠は作らず、既存 evidence schema
(`rules.lint.finding_evidence_schema.enums.review_status` と
`judgement_reason`) と `REV-016 観点別判定表` を使う。

判定の置き場所:

- 観点単位の判定台帳 (本 REV の正本) は
  `doc/reviews/260329-seminar-curriculum-proposal/rev-017-decisions.tsv`
  に置く。スキーマと完了条件は `doc/reviews/README.md` を参照。
  finding 件数の内訳は同 TSV の `finding_dispositions` 列に
  `<review_status>:<judgement_reason> x<count>` 形式で記録する。
- finding 個別判定が必要な場合 (例: 同観点内で複数の judgement に
  分かれる) は lint JSON の各 finding に `review_status` / `judgement_reason`
  を埋める。`unreviewed` 以外を入れる際は対応する `judgement_reason` も
  同時に記入 (enum は `rules.lint.finding_evidence_schema.enums`)。
- 観点単位の集約判定は上の `REV-016 観点別判定表` の
  「判定」列にも反映する (`done` / `remaining` / `inferred_done`
  / `not_recorded` / `not_applicable`)。決定 TSV と表が二重で更新される
  形になるが、deck 跨ぎの索引としての価値があるので両方維持する。

`judgement_reason` の使い分け (`review_status` ごとの subtype):

- `accepted` (テンプレート意図として残す):
  - `intentional_template_design`: テンプレートが意図したレイアウト/色
  - `within_visual_tolerance`: 閾値を僅かに超えたが視覚的に問題なし
  - `decorative_only`: 装飾要素で情報伝達に影響なし
  - `brand_approved_exception`: ブランドガイドの例外として承認済み
- `fix_required` (配布前に必ず直す):
  - `auto_fixable`: `pptx_fix.py` が `candidate_values` で機械適用できる
  - `manual_layout_fix`: 枠サイズ・配置の手動調整が必要
  - `manual_content_fix`: 本文・テキストの編集が必要
  - `master_template_fix`: スライドマスター側で直す。deck 個別では直さない
  - `requires_design_decision`: 修正前に色・書体などの設計判断が必要
- `fixed` (この deck の後続成果物で既に修正済み):
  - `fixed_in_later_artifact`: 新しい PPTX リビジョンで反映済み
  - `fixed_by_rule_update`: lint ルール改修で検出されなくなった
- `false_positive` (lint 側の誤検出):
  - `lint_rule_too_strict`: 閾値・スコープを緩める rule 改修が必要
  - `measurement_error`: 計測方法の問題
  - `missing_context`: lint が読めないコンテキストで誤判定
- `out_of_scope` (この deck では意図的に扱わない):
  - `master_owns`: スライドマスター / ブランド管理者が責務
  - `different_distribution`: 別配布形態でのみ問題になる
  - `legacy_asset_frozen`: 凍結済みレガシーアセット
  - `partner_owned`: 第三者コンテンツで編集不可

`false_positive` を選んだ場合は、直後に lint ルール改修タスク
(`LINT-*` / `FIX-*`) を起票する前提。`fix_required` の subtype が
`master_template_fix` の場合は、deck 個別では `out_of_scope` 相当として
扱い、別途マスター修正タスクを起票する。

件数 0 の観点 (`P1-3`, `P1-5`, `P1-7`, `P1-15`):

- finding が無いため `review_status` / `judgement_reason` を埋める対象は無い。
- 観点単位では `inferred_done` を最終判定として採用する。
- REV-017 完了判定では「対象 12 観点のうち 4 観点が
  `inferred_done` で確定済み」と扱う。

検証:

- `doc/reviews/260329-seminar-curriculum-proposal/rev-017-decisions.tsv`
  の `observation_decision` 列に `not_recorded` または空欄が無いことを確認する。
- `remaining` 行で `finding_dispositions` の件数合計が
  `latest_lint_count` と一致することを確認する。
- finding 単位判定を併用した場合は、lint JSON の `review_status` が
  `unreviewed` 以外、かつ `judgement_reason` が enum 値であることを
  `jq` などで確認する (deck-level / consolidated 双方)。
- `REV-016 観点別判定表` の対象 12 行に `not_recorded` が残って
  いないことを確認する。

## Next

| ID | 状態 | 優先度 | タスク | 完了条件 |
| ---- | ------ | -------- | -------- | ---------- |
| WEB-001 | doing | P1 | レビュー UI を GitHub Pages 公開し、KV 経由でローカルに判定を取り込む | (1) `ymd2357/pptx-design-review` を public で push 済、(2) `deploy-review-web` ワークフローが `review-pages` ブランチに orphan push、(3) Settings → Pages の Source は `review-pages` ブランチ、(4) 画面ゲートは PIN (SHA-256 を SPA に埋め込み、平文は Claude メモリのみ) で通る、(5) Submit ボタンで判定 payload を age 公開鍵で暗号化し `https://pptx-visual-review.pages.dev/api/feedback` に POST、(6) `scripts/fetch-reviews.py --apply` を PC で実行し KV から復号して `doc/reviews/<deck>/rev-NNN-decisions.tsv` & `rev-NNN-finding-judgements.json` を書き出す、(7) `git commit` して反映。**前提**: vscode-pptx-viewer 側の `gallery/functions/api/feedback.js` に `?key=<id>` GET を追加するパッチが production deploy 済 (Claude が当て済、ユーザーが main マージ + push) |
| REV-017 | done | P1 | P0-3 `low_contrast` の auto-fix を before/after 比較でスライド単位に採否確定する | 完了。`p0-rev-017-low-contrast-fixed.pptx` を確定、`rev-017-compare.tsv` で 20 件全部「採用」記録済。残った大胆な修正候補は `FIX-006` (検出範囲拡張) で継続。 |
| REV-018 | done | P1 | P1-4 `contrast_ratio` の auto-fix を before/after 比較でスライド単位に採否確定する | REV-017 と同時消化済 (`pptx_fix --rules contrast` が両 check を一括処理)。`p0-rev-017-low-contrast-fixed.pptx` 上で XML / rendered lint 双方 0 件確認済。本 REV は独立作業を持たない。 |
| REV-019 | done | P1 | P2-12 `card_grid_consistency` の 4 件を auto-fix で適用 | 完了。`pptx_fix --rules card_grid` の保守的修正 (= 子要素の bbox 保持 + container 内平行移動のみ) を 6 container に適用、children が inner box を超える 4 container は manual_required skip。視覚変化は微小だが画像破壊なし。レビューで全採用 (`rev-019-compare.tsv`)。賢い修正候補は `FIX-008` で継続。 |
| REV-032 | done | P1 | FIX-008 (slot-aware card_grid v4) の視覚レビュー | レビュー結果は **全 3 件 reject** (`rev-032-compare.tsv`)。指摘内容は (a) アイコン / ヘッダー / 本文の同 role 整列が出来ていない、(b) text 幅変更で slide 13 / 20 に不要な改行が出た、(c) 行数差に応じた vertical alignment 切替が必要。本 REV を受けて FIX-008 のコード (`74069f6`) は `1292895` で revert し v3 に戻した。エビデンスとして `tmp/review-snapshot/.../rev-032/` の before/after PNG と `rev-032-compare.tsv` を保持。継続は `FIX-012`。 |
| REV-033 | done | P1 | FIX-012 (role 推定 + 上揃え + badge 下揃え) の視覚レビュー | 完了。`rev-033-compare.tsv` で slide 2 / 13 とも **adopt** 確定 (slide 20 は AE=1611 の render ノイズで XML 不変、unchanged-prune 後はそもそも比較対象外)。SPA URL は `/compare/?deck=260329-seminar-curriculum-proposal&rev=033`。FIX-012 の挙動はユーザー視覚レビューで承認された。 |
| FIX-001 | done | P1 | `wrap_break_changes_meaning` に widen-to-fit 候補と auto-fix を追加する | **lint 側**: `check_wrap_break_changes_meaning` で run 内インライン break (`\n`/`\v`) を保持する shape を対象に、`_estimate_text_width_pt` (CJK 1.0× / Latin 0.55× の混在推定) で joined 1 行幅を計算。要求幅 + margin が `SLIDE_W - SAFE_MARGIN_RIGHT - shape.left` 以下かつ現幅より大きい場合のみ `candidate_values=[{strategy:"widen_to_fit", bbox_pt:[x,y,target_w,h], width_pt, font_size_pt, estimated_text_width_pt, available_width_pt, replace_inline_breaks:true}]` を出力し `fixability=auto_fix_candidate` に昇格。条件未充足は `manual_required` のまま (=既存挙動)。`finding_to_json_dict` 側でも `wrap_break_changes_meaning` 専用 branch を `_fixability_for_json` / `_candidate_values_for_json` に追加し schema 経由でも widen 候補が伝搬する。<br>**fix 側**: `_detect_finding_action` の `text_wrap` 分岐で candidate_values の widen_to_fit を抽出し、`action.after.widen_to_fit = {width_pt_actual, ...}` に詰める (list / 単一 dict どちらも受ける)。`_apply_finding_action` で `_replace_text_breaks` 後に `shape.width = Pt(width_pt_actual)` を適用。<br>**回帰**: `_build_wrap_break_widen_fixture` (textbox 200pt × 60pt + "Auto\\nmation improves review speed" 24pt) を追加し、`candidate_values.widen_to_fit` 出力、fix 後の shape width が ~455pt に拡大、`text == "Automation improves review speed"` を assert。既存 `_build_wrap_break_fixture` (900pt 幅、widen 不要) は従来通り `_replace_text_breaks` のみで処理されることも回帰検証済。REV-017 deck では現在 wrap_break_changes_meaning が 0 件のため deck 上の効果は将来検証。検証 deck の根拠は `claude-manual-visual-fixes-2026-05-10.md` の slide 15 / 47 |
| FIX-002 | done | P1 | badge コンテナ内テキストの中央揃えを検出・修正する | **lint 側**: 新規 check `badge_alignment` (warning) を追加。`_is_badge_like_shape` で「solid fill ある + 単一行 ≤20 文字 + bbox 240×120pt 以下 + アスペクト比 0.5〜2.5」を満たす shape を badge と分類し、paragraph alignment が CENTER 以外 / vertical_anchor が MIDDLE 以外なら misaligned 軸を列挙して finding 出力。`fixability=auto_fix_candidate` + `candidate_values={alignment:"CENTER", vertical_anchor:"MIDDLE"}`。`check_alignment` は badge shape を skip して二重発火を防止。<br>**fix 側**: `badge_alignment` rule を `FINDING_DRIVEN_RULES` に追加し、`_detect_finding_action` で apply action を生成、`_apply_finding_action` で `text_frame.vertical_anchor = MIDDLE` + 各段落 `alignment = CENTER` を設定。<br>**回帰**: `_make_badge_alignment_bad` (ROUNDED_RECTANGLE 120×48 + 紫 fill + "Beta" + LEFT/TOP) と `_make_badge_alignment_good` (同 shape + CENTER/MIDDLE) を `test_pptx_lint.py` に、`_build_badge_alignment_fixture` を `test_pptx_fix.py` に追加。bad fixture では `misaligned={alignment, vertical_anchor}` + fix 後の residual 0 を assert。REV-017 deck では現在 badge_alignment が 0 件 (false positive 無し)。検証 deck の根拠は `claude-manual-visual-fixes-2026-05-10.md` の slide 3 / 40 |
| LINT-008 | done | P2 | 孤立装飾 line / connector / arrow の検出を追加する | (旧 ID: FIX-003 → 自動削除を行わない検出 only タスクなので命名規則違反として LINT-008 にリネーム) `decorative_isolated_lines` check を新設 (warning)。`MSO_SHAPE_TYPE.LINE` (connector) / `auto_shape_type` 名に `ARROW` または `LINE` を含む autoshape / 極細矩形 (thickness ≤ 4pt + length ≥ 20pt) を line-like と分類し、最近傍 shape の bbox gap が 24pt 以上なら finding を出力。`fixability=decorative_review` で `candidate_values=[]` を返し、自動削除は提案しない。`test_pptx_lint.py` で connector + thin bar の検出と「companion text 隣接時に発火しない」回帰を assert。REV-017 deck で 0 件 (既存 deck には isolated 装飾なし)。doc/slide-guideline-v1.yml に `rules.lint.checks.decorative_isolated_lines` / `priority_catalog.decorative_isolated_lines` / `finding_evidence_schema.decorative_isolated_lines` を追加。検証 deck の根拠は同レポートの slide 12 / 44 |
| FIX-004 | done | P2 | `object_gap_too_small` を semantic ペアに拡張する | **lint 側**: `_semantic_pair_kind` を新設。`title_subtitle` = (vertical 隣接 + 上の font ≥28pt + 下の font ≤24pt + 左 x 差 ≤24pt) → gap 8pt 未満で発火、`label_group` = (horizontal 隣接 + font 差 ≤1.5pt + 縦中心差 ≤8pt) → gap 16pt 未満で発火。`check_object_relationships` は ds:element 由来 threshold と semantic threshold の max を使う。finding detail に `semantic_pair_kind` / `semantic_threshold_pt` を残し、メッセージ末尾に `[title_subtitle]` 等を付与。<br>**fix 側**: 既存 `spacing` rule が `threshold_pt` を参照して shape_b を移動する経路を流用 (改修不要)。<br>**guideline**: `rules.slide.object_spacing.semantic_pair_gap_pt_min` (`title_subtitle: 8.0`, `label_group: 16.0`) と `semantic_pair_detection` (font/axis/tolerance 条件) を追加。`rules.lint.checks.object_gap_too_small` の detection / thresholds も更新。<br>**回帰**: `_make_semantic_title_subtitle_bad` (40pt title + 20pt subtitle, vertical gap 2pt) と `_make_semantic_title_subtitle_good` (gap 12pt) を `test_pptx_lint.py` に追加。bad 側で `semantic_pair_kind=title_subtitle` + `threshold_pt=8.0` を assert、good 側で発火しないことを確認。REV-017 deck では 0 件 (deck 上に該当 pair なし)。検証 deck の根拠は `claude-manual-visual-fixes-2026-05-10.md` の slide 8 / 13 |
| FIX-005 | done | P2 | `text_overlap` / `overflow_text` の修正候補を多段化する | **lint 側**: `_multi_step_overflow_candidates` (overflow_text) と `_multi_step_overlap_candidates` (text_overlap) を新設し、finding detail に `multi_step_candidates` (list of dict) を埋め込み。overflow_text のストラテジは shrink_height / move / shrink_font_size / compress_line_height、text_overlap は move_shape_b / shrink_font_size / shrink_shape_b_height。各 strategy は geometry または font_size_pt/line_height_pt を含み、自己完結。<br>**fix 側**: 既存の `bbox_fit` (geometry move + canvas shrink)、`font_size` (next allowed smaller)、`line_height` (next allowed smaller)、`spacing` (shape_b nudge) が組合せれば 4 ストラテジすべて適用可能。新 rule は追加せず、既存 rule 群を組合せて `--rules bbox_fit,font_size,line_height` で multi-step を実現する運用に統合。<br>**guideline**: `rules.lint.checks.overflow_text` / `text_overlap` に `multi_step_candidates.ordered_strategies` 節を追加し、優先順位とフィックス経路を明示。<br>**回帰**: `test_pptx_lint.py` で bad.pptx の overflow_text finding に `multi_step_candidates` strategies が含まれること、`object_relationships_bad` の text_overlap finding に `move_shape_b` strategy が含まれることを assert。検証 deck の根拠は `claude-manual-visual-fixes-2026-05-10.md` の slide 28 |
| WEB-002 | done | P1 | 視覚レビューで finding 判定後に自動で次の未判定 finding へ進む | `review_status` を選んで判定が確定 (= `unreviewed` 以外 + `judgement_reason` が決定) した時点で drawer を閉じて、現観点内の次の未判定 finding を開く。最後の未判定が判定された場合は drawer を閉じてトーストで完了を示す。連続判定 200 タップを最少操作で回せる |
| WEB-003 | done | P1 | finding drawer 内に「前/次の finding」ナビを追加 | drawer ヘッダに `< 前へ` / `次へ >` ボタンを置き、現観点の findings 配列内で前後の finding に切り替えられる。スライドギャラリー側のスライド/bbox 強調も追従。判定の確定有無に関わらず手動移動できる (WEB-002 とは独立) |
| WEB-004 | done | P1 | drawer を閉じた直後に観点進捗と bbox 色を即時反映 | (1) drawer を閉じると `判定済 X / N finding` が即座に更新される (現状でも `updateJudgement` 内で更新だが描画タイミングを確認)、(2) SVG overlay の bbox 色が「未判定 = 赤 / 判定済 = 緑」に switch される (`unreviewed` 以外 + `judgement_reason` 決定済を判定済と扱う)、(3) 観点カード側の `判定済 X / N` バッジも `renderObservationCard` で同様に追従 |
| WEB-005 | done | P2 | UI ラベルから schema 名の英語併記を排除する | 「観点判定 (observation_decision)」「レビュー状態 (review_status)」「判定理由 (judgement_reason)」「補足コメント (rationale)」など、現状日本語の後に括弧で原語を併記しているラベルから括弧部分を撤去し、日本語のみで統一する。データとして保存する value (TSV / JSON) は元の英語 enum のまま維持 |
| WEB-006 | done | P1 | 観点カードの判定内訳セクションを読み取り専用化する | `observation_decision = remaining` のときに表示される判定内訳 (review_status × judgement_reason × count) は finding 単位の判定を集計した結果であり、人手で `count` 数値を変えたり「判定内訳を追加 / 削除」できるべきではない。集計結果の表示のみに変え、編集経路は finding drawer 経由に一本化する |
| WEB-007 | done | P1 | 視覚レビューから戻った時の観点一覧スクロール位置を復元する | `視覚レビューへ →` で遷移する前のスクロール位置 (もしくは対象観点カードの位置) を sessionStorage に保存し、戻り時に復元する。観点単位の deep link (例: `review/?...#P0-3`) でも fragment anchor が効くようにし、観点カードに `id="<review_no>"` を付与する |
| WEB-008 | done | P1 | 視覚レビューに finding 一覧パネルを追加し、bbox 無し finding も含めて drawer に到達可能にする | スライドギャラリーの下 (or タブ切替) に「現観点の全 finding 一覧」をリスト表示する。各行は `slide N / shape 名 / check_id / 簡易 message / 判定状態バッジ` を持ち、タップで `openFinding` を呼んで drawer を開く。`bboxPt` を持たない finding (タップ可能 SVG が無いタイプ) でも一覧から drawer に飛べる。スライド上の「このスライドには finding がありません。」表示には「下の一覧から確認できます」の案内を添える |
| FIX-006 | done | P0 | `low_contrast` の auto-fix 修復候補を広げ、灰色置換以外のパターンも機械適用できるようにする | `_contrast_candidate` を **dual-strategy (foreground / background)** に再設計。`FILL_REPAIR_COLOR_FAMILIES` を新設し、背景色側の同系色 (red / magenta / blue / green / gold / neutral) でも WCAG 充足候補を提示する。各 finding に `foreground_option` / `background_option` / `preferred_strategy` を並べて出力し、輝度デルタが小さい側を自動採用 (= 白テキスト on 赤グラデは fg を維持して bg を `#E6033D` に寄せる、灰色 on 白は従来通り fg を `#707070` に寄せる)。`pptx_fix.py` 側は `mode=foreground_run` / `background_fill` を扱えるよう更新し、`background_source=shape_solid_fill` のときは text shape の fill を、`behind_solid_fill:<label>` のときは同スライドの enclosing shape を探して fill を差し替える。テスト (`test_pptx_fix.py`) に bg-mode fixture を追加 (白テキスト on `#EBEBEB` rect → 矩形 fill が変更され白 run は維持されることを assert)。REV-017 deck で再走 → 20 件 fg-mode auto-applied、residual 0 で後方互換性確認済。`behind_solid_fill` を実画像 fixture で検証する追加観点は FIX-010 として切り出し。 |
| FIX-007 | done | P1 | finding 単位判定 (judgement_reason=auto_fixable) を `pptx_fix.py` に反映する経路を作る | `pptx_fix.py` に `--finding-judgements-json` フラグと `apply_finding_judgements_overrides()` API を追加。`fetch-reviews.py --apply` 出力の `rev-NNN-finding-judgements.json` を読み、`judgement_reason=auto_fixable` の finding はキー (web SPA の `${check}:${slideIndex}:${shapeId}`) で照合し、`detail.fixability` を `manual_required → auto_fix_candidate` に昇格してから `auto_rules_from_findings` を走らせる。`text_color_allowlist` を用いた end-to-end smoke (lint manual_required + judgement で promote → fix が `#FF0000→#E6033D` を適用、residual 0) を `test_pptx_fix.py` に追加し OK。orchestrator / fetch-reviews 側の追加配線 (例: deck repair パイプラインから自動で呼ぶ) は後続 (FIX-011 候補)。 |
| FIX-008 | rejected | P1 | `card_grid_consistency` の auto-fix を「保守的平行移動のみ」から **child の role / 関係性** を見る賢い修正に進化させる | **v4 (slot-aware) は REV-032 で 3 件全 reject** → commit `74069f6` を `1292895` で revert し、v3 (保守的平行移動) に戻した。<br>**v4 の失敗内容**: children を `(top, left)` ソート後のローカル slot index でグルーピングし、kind 別 (text は slot median 幅に揃え、image / shape は cx/cy 保持) で適用した。XML 上は image cx/cy 不変 + slide 7 タイトル幅 115.2pt → 168pt 拡張など狙い通りだったが、REV-032 視覚レビューで以下の致命的問題が判明:<br>(a) **同 role 整列が局所的**: slot index で揃えただけで、カード間の icon center_x / title baseline / body 左 padding を **実際に整列していない** (各カードが独立に slot median offset へ平行移動するため見た目で揃わない)。レビュー指摘: 「アイコンはアイコン同士、ヘッダーはヘッダー同士、本文は本文同士で合っていることが求められる」。<br>(b) **text width 変更で wrap が再走**: slot median 幅を text shape に当てた結果、slide 13 / 20 で「不要な改行が作られた」 (1 行のはずの text が 2 行になり縦位置がずれた)。事前に行数増を検査せず適用したのが原因。<br>(c) **行数差を考慮していない**: レビュー指摘 「1 行 / 2 行 / 3 行が混在するなら上揃え / 中央揃え / 下揃えなど臨機応変な vertical alignment が必要」。現実装は vertical anchor を一切扱っていない。<br>**残骸**: `tmp/review-snapshot/.../rev-032/` の before/after PNG と meta.json は revert せず、失敗エビデンスとして保持。`rev-032-compare.tsv` (全件 reject + メモ) も同様。<br>**継続先**: 上記 3 点を踏まえた本格設計は **FIX-012** に集約。 |
| FIX-010 | done | P2 | contrast bg-mode の `behind_solid_fill` 経路を実画像 fixture で検証 | `test_pptx_fix.py` に `_build_behind_solid_fill_contrast_fixture` を追加 (= rect fill `#FF5757` の上に独立 textbox fg `#FFFFFF` を載せる構成)。lint が `background_source=behind_solid_fill:...` を出し、`preferred_strategy=background` + `fixability=auto_fix_candidate` であること、`pptx_fix --rules contrast` で **rect の fill だけが変わり、textbox の white run 色は維持される + textbox に新規 solid fill が付与されない** ことを XML レベルで assert。`_find_bg_repair_target` の `behind_solid_fill` 経路 (text shape の中心点を覆う enclosing fill shape を最小面積で選ぶ) が回帰検知可能になった。 |
| FIX-012 | done | P1 | `card_grid_consistency` の auto-fix を role 推定 + wrap 事前検査 + vertical alignment 切替で再設計 | **REV-033 で全 adopt → done 確定**。観測 → 原因 → 適用 → 評価のループをユーザーレビュー指摘ベースで実施。実装方針:<br>**lint 側**: `_shape_record_detail_for_card_role` を新設し、card_grid finding の `row_containers` / `inconsistent_containers` の children に `solid_fill_hex` / `vertical_anchor` / `has_text` / `text_length` を埋める。`row_containers` にも `children` と `padding_pt` を含めるよう拡張 (= 行全体を fix 側で見られるように)。<br>**fix 側**: `_card_classify_role` で各 child を `leading_icon` / `header` / `body` / `badge` (bg+text ペア) / `illustration` / `decorative` に zone ベースで分類。`_card_grid_alignment_moves` がカード横断で以下のルールに従い per-child の target_top を返す: (a) `leading_icon`.center_y = そのカードの header 1 行目 center (= Option B, intra-card pair 整合) (b) `body`.top = 行内 max(header.h) ベースの「上揃え」 (c) `badge`.bottom = card.bottom − min(observed padding_bottom) (= 締まっている card に合わせる)。**幅は一切変えない** (= wrap 副作用の根本回避)。`_card_grid_item_manual_reasons` の inner-box 制約は role-based 移動があるときは bypass (v3 用の container 縮小前提のため)。<br>**REV-017 deck 再走 (XML diff)**: slide 2 で cards 0/1 の icon が −10.13pt (上へ)、cards 2/3 の body が +20.25pt (下へ); slide 13 で card 0 の badge (bg+text) が +16.79pt (下へ); slide 7 / 20 は変化なし。`row_containers` を使う前は manual_required skip された case (= slide 2 の cards 0/1) も移動を反映。<br>**自己評価 (PNG 目視)**: slide 2 で icon 4 つ・本文 4 つが揃って見えるようになった (3+4 のヘッダー〜本文間に意図的な空きが出るが上揃えの必然)、slide 13 で badge 2 つの top が揃った。wrap 起因の不要な改行ゼロ。<br>**未着手の範囲**: lint 側に「役割不整列」を別 finding として出す拡張 (= 5 の方向) は本イテレーションでは未実施。auto-fix 側が role を推定して動かす経路は成立しているので継続価値は確認できた。 |

### LINT-007 evidence schema 方針

目的:

- lint finding を単なる警告文ではなく、修正判断に使える evidence として
  残す。
- `low_contrast` / `contrast_ratio` だけでなく、全チェックで
  「なぜ検出したか」「自動修正してよいか」「何をどう直す候補があるか」を
  追跡できるようにする。
- deck 固有のレビューと、汎用 lint / fix 実装を分離し、
  一回限りの置換スクリプトに依存しない。

共通 schema に入れる項目:

- `check`, `severity`, `priority`, `review_status`
- `slide_index`, `slide_id`, `shape_id`, `shape_name`, `shape_kind`
- `text_excerpt` または対象要素の説明
- `bbox_pt`, `actual_bbox_pt`, `rendered_bbox_px`
- `measured_value`, `threshold`, `delta`, `unit`
- `evidence_source`: `pptx_xml` / `rendered_image` / `structure_json` /
  `manual_review`
- `evidence_confidence`: `high` / `medium` / `low`
- `fixability`: `auto_fix_candidate` / `manual_required` /
  `not_applicable` / `decorative_review`
- `candidate_values`: guideline token、hex、pt、座標などの候補と
  各候補の検証値
- `recommended_value` と、その採用理由
- `group_key`: 同一テンプレート由来・同一原因の finding を束ねるキー
- `artifact_refs`: rendered PNG、annotation PNG、structure JSON などの
  ローカル evidence パス

作業:

- [x] 既存 `Finding.detail` の項目を棚卸しし、チェック別に不足項目を表にする。
- [x] `doc/slide-guideline-v1.yml` に evidence schema と
      check 別必須 evidence を定義する。
- [x] `pptx_lint.py` の finding 出力を共通 schema に寄せる。
- [x] `pptx_review_priorities.py` が共通 schema から priority evidence を
      作るようにする。
- [x] `pptx_fix.py` の auto-fix 候補判定が `fixability` と
      `candidate_values` を参照できるようにする。
- [x] `low_contrast` / `contrast_ratio` では、rendered image の
      foreground/background、元 run 色、候補 token、再計算 ratio を残す。
- [x] `object_overlap`, `inner_padding_imbalance`, `card_grid_consistency`,
      `safe_*`, `font_*`, `geometry_rounding` でも、修正候補または
      manual_required 理由を残す。
- [x] 回帰テストで、主要 check の finding が schema 必須項目を満たすことを
      検証する。

完了条件:

- `pptx_lint.py --json` の全 finding が共通 schema 必須項目を持つ。
- `--rendered-image-dir` を使う check は、画像 evidence と測定信頼度を
  JSON に残す。
- auto-fix 可能な finding と手動判断が必要な finding が JSON だけで
  区別できる。
- `REV-017` の判断に必要な contrast evidence が、annotation 画像だけでなく
  JSON からも追える。

REV-017 準備済み artifact:

- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/index.html`
- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/rev-017-manual-review-checklist.tsv`
- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/rev-017-lint-focus.tsv`
- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/rev-017-rendered-contrast-lint.json`
- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/rev-017-rendered-contrast-priorities.json`
- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/rev-017-rendered-contrast-focus.tsv`
- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/rev-017-rendered-contrast-annotated/index.html`
- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/rev-017-rendered-contrast-annotated/contrast-overview.png`

## Done

| ID | 状態 | 優先度 | タスク | 完了条件 |
| ---- | ------ | -------- | -------- | ---------- |
| LINT-007 | done | P0 | lint finding の evidence schema を全チェック横断で設計・実装する | 各 finding が、検出理由、対象要素、値、閾値、根拠画像、修正可否、候補値、レビュー判断状態を共通 schema で出力し、auto-fix や目視レビューに使える粒度で JSON に残る |
| REV-016 | done | P0 | 既存 artifact/evidence から観点別レビュー判定表を復元する | `P0-*` / `P1-*` / `P2-1`〜`P2-13` / `P3-1` ごとに、最新 lint 件数、目視確認、判断、根拠 artifact、対応 REV が一覧化され、採用済み・不採用・要再確認が区別されている |
| LINT-005 | done | P2 | オブジェクト間の位置関係チェックを設計・実装する | 単一PPTX内で text box / shape / image / table の重なり、近接不足、揃い崩れ、カード内余白の不均衡を検出する方針が YAML に定義され、チェックまたは明示的な手動レビュー手順と回帰テストが追加されている |
| LINT-006 | done | P2 | テキストボックス内の収まり・余白バランスチェックを設計・実装する | 単一PPTX内で font size / line height / text box height / internal margin / vertical anchor の組み合わせを評価し、フォントサイズ変更とテキストボックスサイズ変更を一方が他方に内包されない対等な変更対象としてセットで判断する方針が YAML に定義され、文字が収まっていても上下余白や視覚中心が不自然なケースを検出するチェックまたは明示的な手動レビュー手順と回帰テストが追加されている |
| PRI-001 | done | P0 | P0-P3 判定体系をレビュー観点から棚卸しする | `rules.lint.priorities` と `rules.lint.priority_catalog` に Pn 規範、既存 lint ID、未実装候補の分類を定義し、`pptx_review_priorities.py` が catalog を参照する |
| DIST-002 | done | P2 | 配布ゲートの実行手順を定義する | `distribution.release_gate.required_checks` を実行するコマンドまたは手動チェックリストの置き場所が決まり、生成物を git に入れない運用と接続されている |
| LINT-004 | done | P1 | lint の画像・アクセシビリティチェック方針を固定する | `image_upscale_ratio`、`contrast_ratio`、`color_only_meaning`、`alt_text_required`、`reading_order` を機械検出するか手動レビュー対象にするかが決まり、実装または除外理由が YAML とテストに反映されている |
| LINT-003 | done | P1 | lint のレイアウト・文字チェックを拡張する | `safe_margins`、`line_height`、`alignment_left_top`、`geometry_rounding` の検出が `pptx_lint.py` と `test_pptx_lint.py` に追加されている |
| LINT-002 | done | P0 | lint チェック ID をガイドラインと同期する | `pptx_lint.py` が出力する `check` が `rules.lint.checks` のキーと一致するか、互換エイリアスが YAML 上で明示され、回帰テストで同期が検証されている |
| DS-012 | done | P0 | 最小テンプレート一覧を上流定義と照合する | `doc/design-system-review-v0.md` の最小セット（タイトル、1メッセージ、2カラム、図解、表、注意、まとめ、確認問題）が `layout_templates.minimum_template_ids` と `layout_templates.templates` にテンプレートまたは明示的な対応表として反映されている |
| DS-011 | done | P0 | ガイドラインの版数メモを現行化する | `doc/slide-guideline-v1.yml` の `meta.notes` と `scope.non_goals` に残る旧版ラベルが現行版の説明に更新され、どの版で何を固定したかが矛盾なく読める |
| DIST-001 | done | P2 | 配布物の正本と派生物を整理する | PowerPoint マスター、アセット置き場、命名規則、書き出し設定の置き場所が決まっている |
| DS-010 | done | P2 | アクセシビリティをチェック可能にする | コントラスト、色以外の意味表現、読み順、代替テキスト運用がチェック可能な形で定義されている |
| DS-009 | done | P2 | アニメーションルールを定義する | 遷移、アニメーション時間、順序が定義されている |
| DS-008 | done | P2 | データ可視化ルールを追加する | グラフ配色、軸、ラベル、注釈、目盛線、数値表記、表スタイルが定義されている |
| DS-007 | done | P2 | 図形、線、影をトークン化する | radius、線幅、影、区切り線の許可値が定義されている |
| LINT-001 | done | P1 | 自動チェック対象を拡張する | 余白、フォント、色、行間、画像拡大率、コントラストのうち次の 1 つ以上が `pptx_lint.py` とテストに追加されている |
| DS-006 | done | P2 | 画像ルールを実務レベルにする | 写真トーン、文字載せ、角丸、キャプション、出典、画像形式方針が定義されている |
| DS-005 | done | P1 | カラー許可リストを拡張する | 背景、グレー、アクセント、状態色、文字背景の許可組み合わせ、アクセント面積比率が定義されている |
| DS-004 | done | P1 | タイポグラフィ運用ルールを固める | 行間、段落前後、箇条書き、強調、代替フォント時の許容範囲が定義されている |
| DS-003 | done | P1 | レイアウト骨格をテンプレート単位で定義する | 12 カラム、ガター、ベースライン、余白スケール、最小テンプレート一覧が YAML 上で機械参照できる |
| DS-002 | done | P0 | トークン体系の命名規則を固定する | value / semantic / component の命名ルール、単位ルール、参照方法が `doc/slide-guideline-v1.yml` に反映されている |
| DS-001 | done | P0 | 目的と適用範囲を固定する | `scope` と `scope_decision_notes` が PowerPoint 正本、Canva 対象外、LMS はPPTX/PDF閲覧文脈のみ、動画・SCORM・印刷・mobile-first 対象外を説明し、`status` が `scope_locked` になっている |
| TOOL-001 | done | P0 | PPTX lint smoke test を整備する | `test_pptx_lint.py` で good/bad fixtures の基本検証が通る |
| TOOL-002 | done | P0 | PPTX auto-fix を整備する | `pptx_fix.py` が autofit と geometry の安全な機械修正を行い、self-check を持つ |
| TOOL-003 | done | P0 | PPTX orphan repair を整備する | `pptx_repair.py` が orphan slide parts と対応 rels を削除し、`test_pptx_repair.py` が通る |
| TOOL-004 | done | P0 | `.bak` 上書きを防ぐ | `pptx_fix.py` と `pptx_repair.py` の `--backup` が既存 `.bak` を保持し、回帰テストがある |
