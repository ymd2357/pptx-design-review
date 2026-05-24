完了しました。最終レポートは `/tmp/codex-residual-v2.md` にあります。

測定は `/tmp/measure_residual_v2.py` で実行し、中間成果物は `/tmp/pptx-residual-v2-work/` に置いています。既存 git tree には書き込んでいません。

結論は、residual は絶対値固定でも文字数単純比例でもなく、`text 内容 × font_size` の組合せ依存です。pixel→pt scale は 10.67px/pt、scale 誤差は最大 ±0.05pt、N=3 の std は 0.00pt でした。