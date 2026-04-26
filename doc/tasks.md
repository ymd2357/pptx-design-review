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

## Next

| ID | 状態 | 優先度 | タスク | 完了条件 |
| ---- | ------ | -------- | -------- | ---------- |
| LINT-003 | todo | P1 | lint のレイアウト・文字チェックを拡張する | `safe_margins`、`line_height`、`alignment_left_top`、`geometry_rounding` の検出が `pptx_lint.py` と `test_pptx_lint.py` に追加されている |
| LINT-004 | todo | P1 | lint の画像・アクセシビリティチェック方針を固定する | `image_upscale_ratio`、`contrast_ratio`、`color_only_meaning`、`alt_text_required`、`reading_order` を機械検出するか手動レビュー対象にするかが決まり、実装または除外理由が YAML とテストに反映されている |
| DIST-002 | todo | P2 | 配布ゲートの実行手順を定義する | `distribution.release_gate.required_checks` を実行するコマンドまたは手動チェックリストの置き場所が決まり、生成物を git に入れない運用と接続されている |

## Done

| ID | 状態 | 優先度 | タスク | 完了条件 |
| ---- | ------ | -------- | -------- | ---------- |
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
