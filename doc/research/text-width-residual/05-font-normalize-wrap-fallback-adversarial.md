# Font normalize / wrap / fallback investigation

## Scope

対象は mcp-cource slide 8 shape 12:

```text
は、AI とシステムを繋ぐための「共通規格」である。
```

このメモは、`normalize 前 706.37pt -> normalize 後 699.28pt` という説明を
同一測定系で再検証し、どの値が何を測っていたかを切り分ける。

## Current Finding

`706.37pt` と full-deck before render の `660.25pt` は同じ量ではない。

- `706.37pt`: controlled fixture の nowrap / red-tail 測定による unwrapped visible width。
- `660.25pt`: full-deck before render で 2 行 wrap した 1 行目の ink right。

full-deck before の DOM range は次の 2 行に割れている:

```text
line 1: は、AI とシステムを繋ぐための「共通規格」であ
line 2: る。
```

そのため、full-deck before の screenshot で `706.37pt` を期待するのは誤り。
ただし、2 行の DOM advance を合算し、末尾 `。` の right side bearing を引くと
controlled fixture の `706.37pt` と 2px 以内で一致する。

normalized 後の `699.28pt` と estimator `701.90pt` の差は、CJK fallback を
`Noto Sans JP` と仮定したことが主因。Chrome DevTools Protocol の
`CSS.getPlatformFontsForNode` では、normalized 後の実描画 font は次の通り。

```text
Calibri-Bold: 3 glyphs
HiraginoSans-W6: 23 glyphs
```

つまり normalized 後は `Calibri Bold + Hiragino Sans W6` として評価すべき。
さらに PNG の右端は、黒 pixel の x 座標ではなく pixel の右境界 `x + 1` を
visible edge として扱う必要がある。

## Acceptance Target

いじわる試験では、次のすべてを満たすことを合格条件にする。

- wrap あり / nowrap の違いを説明できる。
- platform font が `Noto Sans JP` ではなく `Hiragino Sans W6` であることを確認できる。
- `Hiragino Sans W6` の末尾 glyph bbox による rsb 補正後、screenshot との差が 2px 以下。
- PNG edge convention (`x` vs `x + 1`) の影響を数値で示す。

## Adversarial Test Log

### Measurement Rule

line ごとの予測式:

```text
predicted_ink_width
  = DOM Range line width
  - invisible trailing whitespace advance, if the line ends with whitespace
  - right side bearing of the last glyph in the actual platform font
  - trailing letter-spacing, if CSS letter-spacing is non-normal
```

actual platform font は CDP `CSS.getPlatformFontsForNode` で取得する。

PNG の right edge は単一値ではなく、rightmost dark pixel が占める
pixel interval として扱う。

```text
pixel interval = [right_px, right_px + 1]
```

`x` だけ、または `x + 1` だけを正解に固定すると DPR / antialias /
threshold で 2px を少し超えることがある。interval 距離で評価する。

### Test Matrix

48 isolated captures:

- 12 scenarios
- 4 viewport / DPR variants (`1200x675 dsf1`, `1200x675 dsf2`,
  `2400x1350 dsf2`, `2400x1350 dsf4`)
- capture path: `/tmp/text-width-adversarial-isolated/`
- full result files:
  - `/tmp/text-width-adversarial-isolated/dom.json`
  - `/tmp/text-width-adversarial-isolated/analysis-adjusted.json`
  - `/tmp/text-width-adversarial-isolated/summary-adjusted.json`

Non-isolated full-slide scan was also run first at `/tmp/text-width-adversarial/`.
That run intentionally exposed a failure mode: wrapped second lines can overlap
nearby slide text in y-range, so naive black-pixel scans can pick a different
text object. Width verification must either isolate the target shape or use a
tighter pixel mask.

### Results

All isolated scenarios are within 2px when using actual platform fonts,
last-glyph rsb, trailing letter-spacing correction, and pixel interval scoring.

```text
scenario                    lines  max interval diff  note
after_font_size_24          1      1.00px             pass
after_font_size_40          2      1.35px             pass
after_force_hiragino        1      1.07px             W7 fallback
after_force_noto            1      1.88px             forced Noto
after_force_serif           1      1.22px             sans fallback
after_letter_spacing_0_5    2      1.21px             subtract trailing LS
after_normal                1      1.22px             pass
after_nowrap                1      1.22px             pass
after_width_650             2      1.27px             wrap line-by-line
after_width_800             1      1.22px             pass
before_normal               2      1.83px             wrapped before
before_nowrap               1      1.75px             control-like
```

Platform-font summaries:

- after normal family: `HiraginoSans-W6` for 23 glyphs,
  `Calibri-Bold` for 3 glyphs.
- before normal family: `HiraginoSans-W6` for 23 glyphs,
  `Helvetica-Bold` for 3 glyphs.
- forced Noto family: `NotoSansJP-Thin_Bold` for 23 glyphs,
  `Calibri-Bold` for 3 glyphs.

### Conclusions After Adversarial Tests

1. `706.37pt` is not the full-deck before visible width. It is the nowrap /
   unwrapped visible width from the controlled red-tail fixture.
2. Full-deck before is wrapped; its screenshot width `660.25pt` is line 1 only.
3. Normalized render uses `Calibri-Bold` for 3 Latin glyphs and
   `HiraginoSans-W6` for 23 CJK glyphs, not `Noto Sans JP`.
4. The estimator must use the actual platform font from CDP or an equivalent
   font-resolution model. Hard-coding `Noto Sans JP` is not accurate on macOS.
5. For PNG verification, use the pixel interval `[right_px, right_px + 1]`,
   not a single edge convention.
6. If `letter-spacing` is non-normal, subtract trailing letter spacing in
   addition to the last glyph rsb.

### Hardening Pass

The first adversarial pass was still too narrow: it only proved the specific
shape and a small variant set. A broader hardening pass was run at:

```text
/tmp/text-width-hardening/
```

Coverage:

- 168 synthetic cases.
- 12 actual mcp-cource slide 8 viewer runs.
- 180 total runs.
- 263 measured lines.
- varied text endings: `。`, kana, kanji, Latin, `.`, `）`, `」`.
- varied font sizes: 18, 24, 30, 40pt.
- varied wrapping widths: 260, 520, 760pt.
- varied font family forcing: normal fallback, Noto, Hiragino, Calibri (MS).
- letter-spacing case: `0.5px`.

The stricter pass initially failed for three useful reasons:

1. Synthetic harness bug: canvas was set to 1440 CSS px, which broke the
   viewer-equivalent mapping of `1920 CSS px == 1440pt`.
2. Measurement model bug: wrapped lines can end with invisible trailing spaces.
   DOM width includes those spaces, but PNG ink does not. The model must subtract
   trailing whitespace advances before applying last-glyph rsb.
3. Test harness bug: `Range.getClientRects()` can return duplicate or fragment
   rects for one visual line. The final harness groups character rects by `top`
   and uses the union as the line box.

After those fixes:

```text
total_runs        180
total_lines       263
failures          0
max_interval_px   0.447934179687536
mean_interval_px  0.08890020270895235
p95_interval_px   0.39249999999999924
```

The final artifacts are:

- `/tmp/text-width-hardening/dom.json`
- `/tmp/text-width-hardening/analysis.json`
- `/tmp/text-width-hardening/summary.json`

This raises the confirmed claim from "one shape can be explained" to:

> For the tested macOS Chromium/vscode-pptx-viewer environment, line-level ink
> right can be predicted under 1px for this mixed Latin/CJK class when using
> actual platform fonts, character-rect line grouping, trailing whitespace
> subtraction, last-glyph rsb, trailing letter-spacing correction, and pixel
> interval scoring.
