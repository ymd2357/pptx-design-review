# Manual-fix Recipes (python-pptx one-shot)

`pptx_fix.py` のルール (`autofit` / `geometry` / 任意 opt-in の `font_size`) で
カバーできない `manual_review` 項目を、その場限りの python-pptx スクリプトで
適用するときの手順とレシピ。260329 deck の `REV-006`〜`REV-013`
(`p2-6`〜`p2-13`) で実運用してきた
やり方を、ロストしないように台帳化する。

## 適用対象

`fix_policy: manual_review` のうち、検出だけが automated で `pptx_fix.py` に
ルールが無いもの。具体的には:

- `line_height`
- `font_family`
- `font_size_scale` (fixer flag OFF 時)
- `object_gap_too_small`
- `alignment_left_top`
- `inner_padding_imbalance` (一部)
- `safe_text_area_text` (位置調整で済むケース)

`auto_fix` 指定の `geometry_rounding` / `text_autofit_disabled` は
`pptx_fix.py` を使う。混在させない。

## 運用ルール

1. **正本パスを固定**。元 deck は読み取り専用扱いで、出力先は
   `tmp/review/<deck-id>/<basename>.<topic>-fixed.pptx` の命名で別ファイルにする。
2. **スクリプトを捨てない**。書き捨てではなく
   `tmp/review/<deck-id>/scripts/<rev-id>-<topic>.py` に保存する。
   `rev-id` は `rev-006` のようなレビュー作業IDに対応させる。
   デザイン修正の根拠を後から追えるようにする
   (過去 `REV-006`〜`REV-013` / `p2-6`〜`p2-13` の反省点)。
3. **lint json を before/after で残す**。修正前 (= 元正本の lint json) と
   修正後 (= 出力 PPTX の lint json) を両方 `tmp/review/<deck-id>/` に書き、
   `tasks.md` のレビュー作業記録から参照できるようにする。
4. **PowerPoint UI で開かない**。python-pptx は `dcterms:modified` を
   保持するので、すべての iteration で同じ値が並ぶ。UI で開いて保存すると
   タイムスタンプが変わり、追跡が崩れる。
5. **画像 DIFF は PowerPoint 書き出しで作る**。LibreOffice は使わない
   (`AGENTS.md` / `tasks.md` の方針と同じ)。
6. **生成 PPTX / PNG / PDF は git に入れない**。スクリプトと lint json は
   `tmp/review/` のローカル作業証跡として残し、git には要約と運用文書だけを
   入れる。

## スクリプト雛形

```python
# tmp/review/<deck-id>/scripts/rev-NNN-<topic>.py
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt, Emu

SRC = Path("tmp/review/<deck-id>/<basename>.<previous-topic>-fixed.pptx")
DST = Path("tmp/review/<deck-id>/<basename>.<this-topic>-fixed.pptx")

prs = Presentation(str(SRC))
changed = 0
for slide_idx, slide in enumerate(prs.slides, start=1):
    for shape in slide.shapes:
        # 修正対象の特定条件をここに書く (slide_idx, shape.name, bbox 等で絞る)
        if not _is_target(shape, slide_idx):
            continue
        _apply(shape)
        changed += 1

prs.save(str(DST))
print(f"changed={changed} src={SRC.name} dst={DST.name}")
```

実行後は必ず lint を再実行して finding が想定どおり減ったか確認する:

```bash
python3 skills/pptx-design-reviewer/scripts/pptx_lint.py \
  "$DST" --json > tmp/review/<deck-id>/rev-NNN-after-lint.json
python3 skills/pptx-design-reviewer/scripts/pptx_review_priorities.py \
  "$DST" --json > tmp/review/<deck-id>/rev-NNN-after-priorities.json
```

## レシピ

### line_height (`REV-006` / `p2-6` で実施)

`paragraph.line_spacing` を `Pt(N)` で指定する。固定値が `{24, 30, 36, 42,
66, 90}` (lint の許容セット) のいずれかに収める。

```python
from pptx.util import Pt

for para in shape.text_frame.paragraphs:
    if para.line_spacing is None:
        continue
    current_pt = para.line_spacing.pt if hasattr(para.line_spacing, "pt") else None
    if current_pt is None:
        continue
    target = _nearest_allowed(current_pt, {24, 30, 36, 42, 66, 90})
    if abs(current_pt - target) > 0.1:
        para.line_spacing = Pt(target)
```

注意点:

- `line_spacing` が `float` (倍率) の paragraph は触らない。
- 段落単位で本文/見出しが混在することがあるので `para.runs[0].font.size`
  と組み合わせて判定する。

### font_family (`REV-007` / `p2-7` で実施)

`run.font.name` をテンプレート許容書体に寄せる。`font_family` lint の
許容リストは `doc/slide-guideline-v1.yml` の
`tokens.value.typography.font_family.allowed` を参照する。

```python
ALLOWED = {"Noto Sans JP", "Inter"}
TEMPLATE_DEFAULT = "Noto Sans JP"

for para in shape.text_frame.paragraphs:
    for run in para.runs:
        name = run.font.name
        if name in ALLOWED or name is None:
            continue
        run.font.name = TEMPLATE_DEFAULT
```

注意点:

- `run.font.name = None` で master 継承に戻したい場合がある。テーマ
  ベースの書体指定はそのままにする。
- 日本語と英数字で書体が分かれている deck では `<a:ea>` / `<a:latin>` を
  XML 直編集する必要がある。python-pptx は ASCII 側しか触らないので、
  必要なら `run._r.find("./a:rPr/...", NS)` で手当てする。

### font_size (`REV-008` / `p2-8` で実施 / 現在は flag OFF)

`pptx_fix.py --rules font_size` を `FONT_SIZE_FIXER_ENABLED=True` に
して 1 回限定で使うのが第一選択。flag OFF 既定なので、有効化する場合は
セット判断 (枠サイズ・行間・縦中心) を一緒にレビューする前提で使う。

直接スクリプトで書く場合の最小コード:

```python
from pptx.util import Pt

ALLOWED = {80, 64, 56, 48, 40, 36, 32, 28, 24, 22, 20}

for para in shape.text_frame.paragraphs:
    for run in para.runs:
        if run.font.size is None:
            continue
        current = run.font.size.pt
        target = _nearest_in(ALLOWED, current)
        if abs(current - target) <= 1.0:  # 1pt 以内のみ
            run.font.size = Pt(target)
```

`pptx_fix.py` 側の安全条件 (テキストが単一行、resize 後も枠内、bbox が
他図形と重ならない) は最低限再現する。手抜き実装で枠外にはみ出すと
P0 (`overflow_text`) 化するので必ず lint 再実行で確認する。

### object_gap_too_small (`REV-012` / `p2-12` で実施)

隣接オブジェクトの最小ギャップを 8pt 以上に揃える。lint は normalized
bbox で判定するので、修正側も `Emu` ↔ `pt` 変換を意識する。

```python
from pptx.util import Emu

GAP_MIN_EMU = Emu(int(8 * 12700))  # 8pt

# 同列 (top が近い) の左右隣接ペアに対して、左の right + GAP_MIN を超える
# 位置に右を移す。bbox を pt で正規化して並べ替えてから処理する。
```

注意点:

- 意図した密着 (アイコン+ラベル等) はテンプレ側で許容判定にする。
  闇雲に 8pt 開けるとカード内のグルーピングが崩れる。
- 既に重なっている (`object_overlap`) ペアは P1 として別 iteration で
  対応する。

### alignment_left_top (`REV-013` / `p2-13` で実施)

`paragraph.alignment` を `PP_ALIGN.LEFT` に揃え、`text_frame.vertical_anchor`
を `MSO_ANCHOR.TOP` に揃える。

```python
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

if shape.has_text_frame:
    shape.text_frame.vertical_anchor = MSO_ANCHOR.TOP
    for para in shape.text_frame.paragraphs:
        if para.alignment != PP_ALIGN.LEFT:
            para.alignment = PP_ALIGN.LEFT
```

注意点:

- 中央揃えが意図した表現の slide (タイトル中央寄せ等) は対象外。
  slide_idx か shape.name で除外する。
- `card_grid_consistency` (旧 `alignment_drift` の責務を吸収、`REV-012` / `p2-12` 系)
  はテンプレ起因が多く、中身を直すよりマスターを直すべきケースがある。
  再発するなら deck ではなくテンプレに上げる。

## 参考

- 反復記録: `doc/tasks.md` の `### レビュー作業記録`
- ルール定義: `doc/slide-guideline-v1.yml` の
  `rules.lint.priority_catalog` と `rules.lint.checks`
- 既存スクリプト: `skills/pptx-design-reviewer/scripts/pptx_fix.py`
  (autofit / geometry / font_size のみ)
