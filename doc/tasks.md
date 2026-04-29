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

## 260329 seminar deck 現状メモ

対象:
`tmp/review/260329-seminar-curriculum-proposal/260329_seminar_curriculum_proposal.p0-p2-8-font-size-fixed.pptx`

進行上の前提:

- レビュー対象の元ファイルは
  `/Users/yamadakenichi/Documents/GitHub/vscode-pptx-viewer/samples/260329_seminar_curriculum_proposal.pptx`。
- 直近の正本候補は上記 `p0-p2-8-font-size-fixed.pptx`。
  それ以前の P2-6/P2-7/P2-8 中間生成物は比較・確認用として扱う。
- 画像 DIFF は PowerPoint 書き出しで作る。LibreOffice は使わない。
- 直近の確認用 DIFF:
  `tmp/review/260329-seminar-curriculum-proposal/p2-8-powerpoint-review-images/diff-vertical/index.html`
- 既に `_invalid-do-not-use/` に移した旧生成物は、判断材料として使わない。
- 生成 PPTX、PNG、PDF、HTML などのレビュー成果物は git に入れない。

2026-04-29 時点の `pptx_lint.py --json` 残件:

| check | count | severity |
| ------- | ------- | ---------- |
| `overflow_images` | 1 | error |
| `safe_margins` | 6 | warning |
| `safe_text_area_text` | 27 | warning |
| `alignment_left_top` | 2 | warning |
| `geometry_rounding` | 138 | warning |

`pptx_review_priorities.py` 上は P2 と P3 が残っている。
ただし Pn 定義自体が `PRI-001` の見直し対象であり、
現行の P2 一括判定を最終分類として扱わない。

## Next

| ID | 状態 | 優先度 | タスク | 完了条件 |
| ---- | ------ | -------- | -------- | ---------- |

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
