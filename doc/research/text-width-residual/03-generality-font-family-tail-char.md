完了しました。最終レポートは [codex-residual-generality.md](/tmp/codex-residual-generality.md) にあります。40 行なので 80 行制約内です。

結論は、`0.110 × font_size` の renderer 残差は今回の sweep では汎用ではありませんでした。全 16 ケースの `residual/font_size` は `min -0.029 / mean -0.010 / max 0.019` で、`0.110` 近傍のケースはありません。

生成・編集は `/tmp` のみです。repo 配下の git tree は触っていません。