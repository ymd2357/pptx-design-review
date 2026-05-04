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
| P0-3 | `low_contrast` | planned | manual_review | manual_review | 読めないコントラスト |
| P0-4 | `overflow_text` | implemented_lint | automated | manual_review | テキストが切れて読めない |
| P1-1 | `animation_present` | implemented_lint | automated | manual_review | 静的配布で失われる情報 |
| P1-2 | `alt_text_required` | implemented_lint | automated | manual_review | 意味を持つ画像の代替テキスト欠落 |
| P1-3 | `color_only_meaning` | manual_defined | manual_review | manual_review | 色だけで意味を伝えている |
| P1-4 | `contrast_ratio` | manual_defined | manual_review | manual_review | コントラスト不足 |
| P1-5 | `heading_hierarchy_broken` | planned | manual_review | manual_review | 見出し階層崩れ |
| P1-6 | `image_aspect_distortion` | implemented_lint | automated | manual_review | 画像の縦横比崩れ |
| P1-7 | `key_area_cropped` | planned | manual_review | manual_review | 画像重要部の欠け |
| P1-8 | `line_height` | implemented_lint | automated | manual_review | 行間が読みやすさを損なう |
| P1-9 | `missing_required_element` | planned | automated | manual_review | 必須要素欠落 |
| P1-10 | `object_overlap` | implemented_lint | automated | manual_review | オブジェクト重なり |
| P1-11 | `overflow_images` | implemented_lint | automated | manual_review | 画像のスライド外はみ出し |
| P1-12 | `overflow_shapes` | implemented_lint | automated | manual_review | 図形のスライド外はみ出し |
| P1-13 | `reading_order` | manual_defined | manual_review | manual_review | 読み順不整合 |
| P1-14 | `text_autofit_disabled` | implemented_lint | automated | auto_fix | 自動縮小による可読性リスク |
| P1-15 | `wrap_break_changes_meaning` | planned | manual_review | manual_review | 折返しで意味が変わる |
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
| P2-12 | `alignment_drift` | implemented_lint | automated | manual_review | 近接要素の揃いズレ |
| P2-13 | `text_vertical_balance` | implemented_lint | automated | manual_review | テキストボックス内の縦余白バランス不自然 |
| P3-1 | `geometry_rounding` | implemented_lint | automated | auto_fix | 座標の微小な丸めズレ |

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
- 残 lint: `geometry_rounding` 129, `alignment_drift` 89,
  `safe_text_area_text` 27, `inner_padding_imbalance` 17, `safe_margins` 1
- 残 `geometry_rounding` 129 件はすべて `.25/.5/.75pt` 単位で、
  グリッド計算上の意図的な値の可能性が高い。auto-fix で integer に丸めると
  レイアウトが壊れるため対象外。
- 配布ゲートとしては止まる要因なし。次の作業は `REV-016` とし、
  新規 artifact 接頭辞は `rev-016-*` を使う。
  候補は
  (a) PowerPoint 書き出しで目視 DIFF を取り finalize するか、
  (b) `alignment_drift` 再発系の根本原因 (マスター/スタイル) を直すかの
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

## Next

| ID | 状態 | 優先度 | タスク | 完了条件 |
| ---- | ------ | -------- | -------- | ---------- |
| REV-016 | todo | P0 | 既存 artifact/evidence から観点別レビュー判定表を復元する | `P0-*` / `P1-*` / `P2-1`〜`P2-13` / `P3-1` ごとに、最新 lint 件数、目視確認、判断、根拠 artifact、対応 REV が一覧化され、採用済み・不採用・要再確認が区別されている |

## Done

| ID | 状態 | 優先度 | タスク | 完了条件 |
| ---- | ------ | -------- | -------- | ---------- |
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
