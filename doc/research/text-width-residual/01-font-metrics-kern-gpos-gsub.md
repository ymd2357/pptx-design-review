`/tmp/codex-font-metrics.md` に指定フォーマットでまとめました。41行です。

結論として、font file 内の kern/GPOS/GSUB/rsb では +2-3pt は説明できません。該当する font-internal contribution は `kern 0.0pt`, `GPOS -1.2pt`, `GSUB 0.0pt` で、残りは renderer/clip 側の residual として切り分けています。