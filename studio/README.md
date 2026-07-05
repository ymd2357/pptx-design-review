# PPTX Lint/Fix Studio

任意の `.pptx` を投入し、**スライドごとに検出された LINT 項目を check 種別単位で
ON/OFF** して、選択した分だけ機械修正 (fix) を当て、修正済み PPTX を
ダウンロードするローカルアプリ。

`web/` の SPA は before/after PNG の採否確定（lint/fix ロジックの検証）ツール。
本アプリはそれとは別で、**実運用として PPTX に lint/fix を当てる**ためのもの。

## 関係

```
studio/
  server.py    stdlib http.server のみ。依存追加なし
  engine.py    skills/pptx-design-reviewer/scripts の lint/fix/normalize/render を import で束ねる薄い層
  static/      単一ページ UI (バニラ JS, ビルド不要)
  sessions/    アップロードと作業成果物 (.gitignore)
```

lint/fix のロジックは一切複製せず、すべて
`skills/pptx-design-reviewer/scripts/` に委譲する:

- `pptx_normalize_fonts.normalize_pptx` — 投入直後に必ず通す font 正規化（全 lint/fix の前提）
- `pptx_lint.lint_pptx` / `finding_to_json_dict` — finding 検出
- `pptx_fix.fix_pptx(..., selection=)` — `(slide_index, rule)` 単位の選択ゲート付きで適用
- `pptx_review_orchestrator._render_pptx` — vscode-pptx-viewer + Playwright による描画

## 起動

```bash
python3 studio/server.py            # http://127.0.0.1:8765
PORT=9000 python3 studio/server.py  # ポート変更
```

ブラウザで開き、PPTX をドラッグ&ドロップ。スライドごとの check 種別チップを
ON/OFF して「選択を適用」→「修正済みを DL」。

## トグルの粒度と仕組み

- 粒度は **check 種別 × スライド**（同一スライド内の同種 check は集約）。
- UI の (slide, check) 選択を `(slide_index, rule)` 集合に変換し、
  `fix_pptx(selection=...)` に渡す。shape-driven rule は check と 1:1 なので
  これで check 粒度のトグルが成立する。finding-driven rule は渡す findings を
  選択分に絞ることでさらに正確に効く。
- 既定: `auto_fix` / `judgement_fix` の check は ON、機械修正の実装が無い
  `no_fix` の check は表示のみ（トグル不可）。
- `judgement_fix` の check は「トグル ON = 人間の判断」とみなし、
  `judgement_gate=False` で適用する。

## 描画プレビュー

各スライドの「描画して確認」で `_render_pptx` を呼び、before(source) /
after(fixed) を並置表示する（適用後のみ after が出る）。
vscode-pptx-viewer 拡張 + Playwright が必要（`skills/pptx-design-reviewer/SKILL.md`
の Visual export 要件参照）。描画が使えない環境でも lint/fix/DL は動く。

## API

| Method | Path | 役割 |
|---|---|---|
| POST | `/api/upload` | raw octet-stream PPTX (+ `X-Filename`) → session + grouped findings |
| POST | `/api/apply` | `{session_id, selections:[{slide_index,check}]}` → 適用サマリ |
| GET  | `/api/download?session_id=` | fixed.pptx |
| POST | `/api/render` | `{session_id, which}` → スライド PNG URL |
| GET  | `/api/render/<sid>/<which>/<slide-NN.png>` | PNG |
