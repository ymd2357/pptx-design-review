完了しました。最終レポートは `/tmp/codex-residual-boundary.md` に作成済みです。60 行なので指定の 100 行以内です。

結論は、`residual` は boundary count に線形比例せず、順序依存・構成依存が強いです。今回の実測では `boundary_count` の係数は小さく、むしろ Latin glyph 数の寄与が支配的でした。

検証:
- `/tmp/codex-residual-boundary.md` 存在確認済み
- 行数 60 行
- git tree は触っていません。表示されている未コミット差分は作業前からある既存差分です。