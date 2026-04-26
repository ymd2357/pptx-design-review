# Distribution Release Gate

このファイルは `distribution.release_gate.required_checks` の実行手順の正本。
配布前に、対象デッキごとに以下を確認する。

## Inputs

- `DECK_ID`: `distribution.naming.deck_id.pattern` に合う lower-kebab-case 名
- `PPTX`: `tmp/distribution/{DECK_ID}/{DECK_ID}.pptx`
- `PDF`: `tmp/distribution/{DECK_ID}/{DECK_ID}.pdf`
- レビュー出力先: `tmp/review/{DECK_ID}/`

`tmp/` 配下の PPTX、PDF、PNG、JSON、メモは生成物として扱い、明示指示が
ない限り git に追加しない。

## Automated Checks

### 1. Guideline YAML

```bash
node -e "const fs=require('fs'); const yaml=require('js-yaml'); \
yaml.load(fs.readFileSync('doc/slide-guideline-v1.yml','utf8')); \
console.log('YAML OK')"
```

Pass condition: `YAML OK` が出力される。

### 2. Markdown Lint

```bash
npx --no-install markdownlint-cli2 doc/tasks.md AGENTS.md TASKS.md
```

Pass condition: `0 error(s)` で終了する。

### 3. PPTX Lint

```bash
python3 skills/pptx-design-reviewer/scripts/pptx_lint.py \
  "tmp/distribution/${DECK_ID}/${DECK_ID}.pptx" \
  --severity error
```

Pass condition: error severity の finding が 0 件で終了する。
必要に応じて詳細ログを生成する場合は、git 管理外の `tmp/review/` に出力する。

```bash
mkdir -p "tmp/review/${DECK_ID}"
python3 skills/pptx-design-reviewer/scripts/pptx_lint.py \
  "tmp/distribution/${DECK_ID}/${DECK_ID}.pptx" \
  --json > "tmp/review/${DECK_ID}/lint.json"
```

## Manual Checks

### 4. PDF Visual Review

`tmp/distribution/{DECK_ID}/{DECK_ID}.pdf` を画面表示し、以下を確認する。

- すべてのスライドが読める
- テキスト、図形、画像が意図せず重なっていない
- 重要情報がスライド外、セーフエリア外、または切れた状態になっていない
- アニメーションやクリック順に依存せず、必要な内容が静的に読める
- 画像上の文字は最終レンダリング上で十分なコントラストがある
- 色だけで状態、正誤、重要度、系列、増減を表現していない
- 意味のある画像、スクリーンショット、図解、ロゴには代替テキストがある
- PowerPoint の選択ウィンドウ順と見た目の読順が矛盾していない

必要に応じてレビュー画像やメモを `tmp/review/{DECK_ID}/` に置く。
これらは生成物なので、明示指示がない限り git に追加しない。

## Git Gate

配布前の最終確認:

```bash
git status --short
```

Pass condition: 生成された PPTX、PDF、PNG、JSON、レビュー用メモが
staged / unstaged に出ていない。ユーザーが明示した場合だけ追跡対象にする。
