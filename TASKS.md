# TASKS

The canonical task ledger is `doc/tasks.md`.

This file exists for repository initialization tooling and pre-commit discovery.

## Active

### LINT-007: lint finding の evidence schema を全チェック横断で設計・実装する

目的:

- lint finding を単なる警告文ではなく、修正判断に使える evidence として
  残す。
- 全チェックで「なぜ検出したか」「自動修正してよいか」
  「何をどう直す候補があるか」を追跡できるようにする。
- deck 固有のレビューと、汎用 lint / fix 実装を分離する。

入力:

- `doc/tasks.md` の `LINT-007 evidence schema 方針`
- `doc/slide-guideline-v1.yml`
- `skills/pptx-design-reviewer/scripts/pptx_lint.py`
- `skills/pptx-design-reviewer/scripts/pptx_review_priorities.py`
- `skills/pptx-design-reviewer/scripts/pptx_fix.py`
- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/rev-017-rendered-contrast-lint.json`
- `tmp/review/260329-seminar-curriculum-proposal/rev-017-manual-review/rev-017-rendered-contrast-annotated/index.html`

作業:

- [x] 既存 `Finding.detail` の項目を棚卸しし、チェック別に不足項目を表にする。
- [x] `doc/slide-guideline-v1.yml` に evidence schema と
      check 別必須 evidence を定義する。
- [ ] `pptx_lint.py` の finding 出力を共通 schema に寄せる。
- [ ] `pptx_review_priorities.py` が共通 schema から priority evidence を
      作るようにする。
- [ ] `pptx_fix.py` の auto-fix 候補判定が `fixability` と
      `candidate_values` を参照できるようにする。
- [ ] `low_contrast` / `contrast_ratio` では、rendered image の
      foreground/background、元 run 色、候補 token、再計算 ratio を残す。
- [ ] その他の主要 check でも、修正候補または manual_required 理由を残す。
- [ ] 回帰テストで、主要 check の finding が schema 必須項目を満たすことを
      検証する。

完了条件:

- `pptx_lint.py --json` の全 finding が共通 schema 必須項目を持つ。
- `--rendered-image-dir` を使う check は、画像 evidence と測定信頼度を
  JSON に残す。
- auto-fix 可能な finding と手動判断が必要な finding が JSON だけで
  区別できる。
- `REV-017` の判断に必要な contrast evidence が、annotation 画像だけでなく
  JSON からも追える。
