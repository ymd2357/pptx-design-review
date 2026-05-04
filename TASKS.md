# TASKS

The canonical task ledger is `doc/tasks.md`.

This file exists for repository initialization tooling and pre-commit discovery.

## Active

### REV-016: 観点別レビュー判定表を復元する

目的:

- 既存 artifact / evidence の実体名と採用状態を根拠に、
  `P0-*` / `P1-*` / `P2-1`〜`P2-13` / `P3-1` ごとのレビュー状況を
  再構成する。
- `p2-N` 接頭辞を新規 artifact 名に使わず、今後の新規成果物は
  `rev-016-*` 接頭辞に寄せる。

入力:

- `doc/tasks.md` の `既存 artifact / evidence 復元メモ`
- `tmp/review/260329-seminar-curriculum-proposal/p2-*-lint.json`
- `tmp/review/260329-seminar-curriculum-proposal/p2-*-priorities.json`
- `tmp/review/260329-seminar-curriculum-proposal/*review-images*/index.html`
- 現在の正本候補:
  `260329_seminar_curriculum_proposal.p0-p2-15-geometry-fixed.pptx`

作業:

- [ ] `rev-016-lint-timeline.tsv` をローカル生成し、各 lint JSON の
      check 件数推移を一覧化する。
- [ ] `rev-016-artifact-map.tsv` をローカル生成し、既存 PPTX / JSON /
      review images を `REV-*` と採用状態に対応付ける。
- [ ] `doc/tasks.md` に観点別レビュー判定表を追加する。
- [ ] 各観点に `done` / `inferred_done` / `remaining` /
      `not_recorded` / `not_applicable` のいずれかを付ける。
- [ ] 各観点に根拠 artifact、対応 REV、最新 lint 件数、目視確認状態を
      記録する。
- [ ] `REV-014` の `p2-14-after-lint.json` は正本系列の採用状態として
      使わないことを、観点別表にも反映する。
- [ ] `REV-015` 正本候補の残件を起点に、再目視が必要な観点だけを
      `Next` に残す。
- [ ] 生成した TSV / JSON / 画像 / HTML は `tmp/review/` に置き、
      git には入れない。

完了条件:

- `doc/tasks.md` だけを見て、各観点が採用済み・不採用・要再確認の
  どれか判断できる。
- `P2` が優先度、`P2-N` が観点No、`REV-NNN` が作業IDであることが
  表と本文で混ざっていない。
- `REV-017` 以降に進む前の再確認対象が明確になっている。
