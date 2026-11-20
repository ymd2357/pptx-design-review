# Per-deck Review Decisions

`tmp/review/` 配下は `.gitignore` 対象なので、deck ごとのレビュー
判定を残せる場所として `doc/reviews/<deck-id>/` を使う。

## ファイル種別

| ファイル | 内容 |
| --- | --- |
| `<deck-id>/rev-NNN-decisions.tsv` | 観点単位の判定台帳 (12 観点程度) |

deck artifact (PNG / PPTX / 中間 JSON) は引き続き
`tmp/review/<deck-id>/` 配下に置き、git 管理外として運用する。
ここに置くのは「人間が下した判定の結論」だけ。

## decisions.tsv のスキーマ

タブ区切り。1 行 = 1 観点。

| カラム | 制約 | 説明 |
| --- | --- | --- |
| `review_no` | 観点No (`P0-3` など) | `chec​k 観点一覧` の Pn 観点No |
| `check_id` | snake_case 文字列 | `rules.lint.checks.<id>` のキー |
| `priority` | `P0` / `P1` / `P2` / `P3` | `priority_catalog.<id>.priority` |
| `latest_lint_count` | 整数 | 当該 deck の最新 lint JSON 件数 |
| `observation_decision` | enum 1 | 下表参照 |
| `finding_dispositions` | 半角セミコロン区切り | 下表参照 |
| `rationale` | 自由文 | 判定根拠 |
| `related_artifacts` | 半角カンマ区切りパス | 参照 lint JSON / 画像など |

### `observation_decision` の取りうる値

| 値 | 意味 |
| --- | --- |
| `done` | 当該 deck でこの観点の判定が完了。修正済みまたは許容済み |
| `inferred_done` | 機械 lint で 0 件。視覚 evidence は無いが confidence あり |
| `remaining` | finding が残っている。`finding_dispositions` に内訳を記録 |
| `not_recorded` | レビュー対象だが判定が未記入 (REV-NNN 完了時には残してはならない) |
| `not_applicable` | この deck では対象外 |

### `finding_dispositions` の表記

`<review_status>:<judgement_reason> x<count>` をセミコロン区切りで列挙する。
review_status / judgement_reason の enum 値は
`rules.lint.finding_evidence_schema.enums` を参照。

例:

```
fix_required:auto_fixable x18; accepted:within_visual_tolerance x3
```

合計は `latest_lint_count` と一致させる
(`inferred_done` の場合は空欄)。

## 完了条件 (REV-017 を含む全 REV-NNN 共通)

- `observation_decision` 列に `not_recorded` または空欄が無い。
- `remaining` のすべての行で `finding_dispositions` が
  `latest_lint_count` と整合している。
- `false_positive:*` を含む行は、対応する lint ルール改修タスク
  (`LINT-*` / `FIX-*`) が `doc/tasks.md` の Next に起票されている。
