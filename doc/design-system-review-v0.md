# やることの優先度順バックログ

タスク管理の正本は `doc/tasks.md`。
このファイルは背景と分解前の作業リストとして残す。

## 1. 目的と適用範囲を固定する

- 対象: PowerPoint編集, PPTX配布, PDF配布の優先順位を固定
- 対象外: Canva編集, LMS連携, SCORM化, 動画化, 印刷最適化を明示
- 閲覧環境: PC全画面, プロジェクタ, LMS内PPTX/PDF表示の優先順位を固定
- 例外条件: フルブリード背景など例外の許容範囲を定義

決定メモ:

- PowerPoint を正本にする。自動 lint/fix/repair は PPTX のパッケージ構造、
  図形座標、テキストフレーム、relationship を前提にしているため。
- Canva はスコープ外にする。過去の Canva export は値抽出の参考元として残すが、
  Canva 編集、Canva テンプレート互換、Canva export 検証は要求にしない。
- LMS は最終閲覧場所としてだけ扱う。LMS 内の PPTX/PDF ビューアで読めることは
  意識するが、アップロード、SCORM、quiz 連携、受講履歴、動画化は扱わない。
- PDF は静的配布・レンダリング確認の対象にする。PDF 編集や印刷最適化は扱わない。
- スマホは主設計対象にしない。16:9 スライドの screen delivery を優先し、
  モバイルは必要時の補助確認に留める。

## 2. トークン体系を3層に分離して命名規則を決める

- 値トークン: color.hex, space.pt, type.pt など
- 意味トークン: color.text.primary, space.md, type.body など
- 部品トークン: component.card.padding など
- 単位ルール: pt, px, 比率の変換方針を定義

決定メモ:

- `doc/slide-guideline-v1.yml` の `tokens` は value / semantic / component の
  3層で運用する。
- value token は `tokens.value.<category>.<name>` とし、pt, px, hex,
  ratio_0_1 などの raw value だけを持つ。別トークンへの参照は持たせない。
- semantic token は `tokens.semantic.<domain>.<role>` とし、用途や意味を表す。
  leaf は `tokens.value.*` への dot-path 参照にする。
- component token は `tokens.component.<component>.<property>` とし、部品固有の
  既定値を表す。原則 `tokens.semantic.*` を参照し、意味トークンがない低レベル値は
  `tokens.value.*` を参照する。
- 直接値を書ける場所は `tokens.value.*`、ポリシー定義、lint/rules/observations
  に限定する。semantic/component のデザイン値は参照必須にする。
- 例外として、component の boolean switch や非視覚 enum は、参照化しても意味が
  増えない場合に限って直接値を許容する。

## 3. レイアウトの骨格を追加する

- 12カラム: 1440pt幅, 左右81pt安全余白を前提にコンテンツ幅1278ptで定義
- ガター: 18pt固定
- ベースライン: 6pt刻みで定義
- 余白スケール: 6pt刻みの段階を定義
- テンプレ一覧: タイトル, 1メッセージ, 2カラム, 図解, 表, 注意, まとめ, 確認問題を最小セットとして定義

## 4. タイポグラフィを運用可能な形にする

- 行間: 各サイズに対してptで固定
- 段落前後: 見出し後, 箇条書き前後などを固定
- 箇条書き: インデント, ぶら下げ, 記号, 行間を固定
- 強調: 太字, 色, 囲みの許容ルールを固定
- 代替フォント時の崩れ: Calibri置換時の許容範囲と調整手順を定義

## 5. カラーを拡張し許可リストで運用する

- 背景色の基本値を定義
- グレースケール段階を定義
- アクセントの明度違いを段階化
- 状態色: 成功, 注意, エラー, 情報を定義
- 文字と背景の許可組み合わせリストを定義
- アクセント面積比率の上限を定義

## 6. 画像ルールを実務レベルにする

- 写真トーン: 彩度, コントラスト, 被写体距離感を定義
- 文字載せ: オーバーレイの濃度トークンを定義
- 角丸: 画像フレームのradiusを定義
- キャプション: 位置, サイズ, 余白を定義
- 出典表記: 必要条件と表記位置を定義
- 形式方針: 写真はJPEG, 透過はPNGなどの基準を定義

## 7. 図形, 線, 影をトークン化する

- radius段階を定義
- 線幅: 1pt, 2pt, 4ptなどに制限
- 影: 禁止か, 許容なら距離とぼかしを定義
- 区切り線: 色と余白を定義

## 8. データ可視化の規定を追加する

- グラフ配色セットを定義
- 軸, ラベル, 注釈のフォントサイズを定義
- 目盛線の線幅と色を定義
- 数値の小数点桁と単位表記を定義
- 表スタイル: 行高, 罫線, ヘッダー背景を定義

## 9. アニメーションルールを定義する

- 遷移: 許容一覧を固定
- アニメーション時間と順序を固定

## 10. アクセシビリティをチェック可能にする

- コントラスト許可リストを定義
- 色だけで意味を持たせないルールを定義
- 読み順の整備方針を定義
- 代替テキスト運用を定義

## 11. 自動チェック対象を拡張する

- 余白, はみ出し, フォント, 色, 行間, 画像拡大率, コントラストを検知
- 丸め方針: 0.5pt観測値を最終で1ptへ丸めるなど方式を固定
- 例外申請: 理由と期限を残す運用を定義

## 12. 配布物を揃える

- PowerPointスライドマスターを正本として実装
- アセット置き場, 命名規則, 書き出し設定を定義

## 推奨ファイル構成案

- design-system/tokens.yaml
- design-system/components.md
- design-system/templates.pptx
- design-system/checklist.md
- design-system/changelog.md
- design-system/examples/good.pptx
- design-system/examples/bad.pptx
