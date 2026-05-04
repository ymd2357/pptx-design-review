# Issue Draft: `pptx_lint.py --json` スキーマバージョンと bbox / 座標系メタを固定する

Status: draft (未起票)
Owner: 未割当
Labels (起票時): `tooling`, `breaking-change`, `integration`
Related: `vscode-pptx-viewer` 側 lint overlay 連携

## 背景

`pptx_lint.py --json` 出力を `vscode-pptx-viewer` 側で消費し、
検出された問題を slide 上に overlay 表示する連携を進めたい。
現状は `[asdict(f) for f in selected]` で `Finding` を素の配列として
吐いているため、消費側で以下が困る:

1. **スキーマバージョンがない**: lint 側のフィールド追加・削除があると viewer が静かに壊れる
2. **bbox がない**: `slide_index` と `shape_id` だけでは overlay 座標が
   決まらない。viewer 側でもう一度同じ shape を解析することになる
3. **座標系の宣言がない**: 単位 (pt / EMU)、スライドサイズ
   (1440×810pt 基準 vs 実 deck サイズ) が外から判断できない
4. **deck と lint 結果の対応が確認できない**: lint 後に PPTX を編集しても
   古い JSON を読んで overlay すると座標がずれる
5. **エンベロープがない**: 配列直渡しなのでメタ情報を後から足せない

## 提案するスキーマ (v1)

```json
{
  "schema_version": 1,
  "tool": {
    "name": "pptx-design-review/pptx_lint",
    "version": "1.0.0",
    "profile": "default"
  },
  "deck": {
    "path": "/abs/path/to/deck.pptx",
    "sha256": "abc123...",
    "slide_count": 12
  },
  "coordinate_system": {
    "unit": "pt",
    "base_size": { "w": 1440, "h": 810 },
    "actual_size": { "w": 960, "h": 540 },
    "scale_x": 1.5,
    "scale_y": 1.5
  },
  "generated_at": "2026-05-03T10:00:00Z",
  "findings": [
    {
      "severity": "error",
      "check": "overflow_text",
      "slide_index": 3,
      "slide_id": 256,
      "shape_id": 12,
      "shape_name": "TextBox 5",
      "message": "テキスト枠がスライド下端を超えています",
      "bbox": { "x": 100, "y": 700, "w": 400, "h": 200 },
      "detail": { "overflow_pt": 90 }
    }
  ]
}
```

主な変更点:

- **トップレベルをオブジェクト化** (配列 → `{ findings: [...] }`)
- **`schema_version: 1`** 必須
- **`coordinate_system`**: 単位と base/actual を明示。
  viewer は actual_size 基準に変換できる
- **`deck.sha256`**: viewer 側で「今開いている PPTX とこの lint 結果は一致するか」を検証
- **`bbox`**: `LintContext.bbox_pt` 由来の値を Finding に必ず含める
  (現状 `detail` 経由で時々入っている)。`coordinate_system.unit` と
  同じ単位で出す

## 互換性

破壊的変更。次の段階で導入する:

1. `--json-v0` フラグで旧出力 (素配列) を残す
2. デフォルトの `--json` を v1 に切り替え
3. `README` と `SKILL.md` の例を更新
4. v0 を 1 リリース後に削除

回帰テスト (`test_pptx_lint.py`) は v1 出力を assert する。

## viewer 側の連携 (参考)

viewer はこの JSON を以下から読む想定:

1. PPTX と同じディレクトリの `<deck>.lint.json`
2. VS Code コマンド `PPTX: Show Lint Findings` でユーザー指定パスから読み込み

`schema_version` が viewer の対応バージョンと違えば notification で警告して読み込みを止める。

## Definition of Done

- [ ] `Finding` dataclass に `bbox: Optional[BBox]` を追加
- [ ] `lint_pptx` 内の各 check が bbox を埋める
      (overflow / overlap / safe_area 系は既に bbox 計算済み)
- [ ] `LintReport` (新 dataclass) を追加し、`schema_version` / `tool` /
      `deck` / `coordinate_system` / `findings` を保持
- [ ] `--json` をデフォルトで v1 出力に変更し、`--json-v0` で旧出力を維持
- [ ] `test_pptx_lint.py` を v1 出力に追従
- [ ] `--structure-json` も同じエンベロープ規約に揃える (別 issue 化可)
- [ ] `doc/lint-json-schema.md` を新規作成して公式定義として置く

## 起票時の手順 (メモ)

- このリポジトリは現状 git remote 未設定。GitHub に push してから
  `gh issue create --title "..." --body-file doc/issue-lint-json-schema.md`
  で起票
- 起票後、本ファイルの Status を `filed (#NN)` に更新するか削除する
