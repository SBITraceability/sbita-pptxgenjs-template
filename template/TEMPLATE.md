# SBI Traceability PPTX テンプレート仕様書

このドキュメントは SBI Traceability 向け PptxGenJS テンプレートの仕様を定義します。

## ファイル構成

```
template/
├── theme.js              # カラー・フォント・レイアウト定数
├── generate-template.js  # テンプレート生成スクリプト
├── sbita-template.pptx   # 生成されたテンプレートファイル
└── TEMPLATE.md           # この仕様書
```

## デザイン規格

### カラーパレット

| 用途 | カラー名 | HEX値 |
|------|----------|-------|
| ゴールド（アクセント） | BRAND_GOLD | `#CBC269` |
| レッド（強調） | BRAND_RED | `#EA4545` |
| サーモンレッド（Confidential） | BRAND_SALMON | `#ED6A5B` |
| メインテキスト | TEXT_PRIMARY | `#454545` |
| サブテキスト | TEXT_SECONDARY | `#535353` |
| ライトグレー | TEXT_LIGHT | `#7F7F7F` |
| 白背景 | BG_WHITE | `#FFFFFF` |
| ライトグレー背景 | BG_LIGHT | `#F5F5F5` |

### フォント設定

| 用途 | フォント |
|------|----------|
| 英語見出し | Century Gothic |
| 日本語見出し | メイリオ |
| 英語本文 | Century Gothic |
| 日本語本文 | メイリオ |
| UI（フッターなど） | Meiryo UI |
| コード | Consolas |

### フォントサイズ

| 用途 | サイズ (pt) |
|------|-------------|
| タイトル（下線付き） | 32 |
| サブタイトル | 24 |
| 大見出し | 24 |
| 見出し | 20 |
| 本文 | 16 |
| 小さめの本文 | 14 |
| キャプション | 10.5 |
| フッター | 8 |
| 統計数字 | 48 |

### レイアウト設定

- **スライドサイズ**: 16:9 ワイド (13.333" × 7.5")
- **マージン**: 0.8インチ
- **ヘッダーライン**: ゴールド (`#CBC269`)、幅 7.1インチ
- **フッター**: "Strictly Confidential" + 著作権表示

### スタイル特徴

1. **ゴールド下線**: タイトルや重要テキストに `#CBC269` の下線
2. **ヘッダーライン**: スライド上部中央にゴールドの水平線
3. **フッター**: 左に "Strictly Confidential"（サーモンレッド）、中央に著作権表示
4. **赤字下線**: 強調テキストに赤色 (`#EA4545`) + ゴールド下線

## スライドレイアウト一覧 (16種類)

### 1. タイトルスライド
- 用途: プレゼンテーション表紙
- 要素: タイトル（ゴールド下線付き）、サブタイトル、発表者名、日付

### 2. セクションディバイダー（Vision）
- 用途: セクション区切り、中扉
- 要素: ヘッダーテキスト、ヘッダーライン、大見出し、備考

### 3. 標準コンテンツスライド
- 用途: 箇条書きコンテンツ
- 要素: ヘッダー、ヘッダーライン、箇条書きリスト

### 4. 2カラムレイアウト
- 用途: 対比、左右分割コンテンツ

### 5. 3カラムレイアウト
- 用途: 複数項目の並列表示

### 6. KPIカードスライド
- 用途: 主要指標の表示
- ボーダー: ゴールド

### 7. タイムライン/プロセスフロー
- 用途: 手順、プロセスの説明
- アクセント: ゴールドの円と線

### 8. 比較スライド (Before/After)
- 用途: 改善前後の比較
- Before: 赤ボーダー、After: 緑ボーダー

### 9. 引用スライド
- 用途: 重要なメッセージ、引用
- アクセント: ゴールドの下線

### 10. アジェンダスライド
- 用途: 目次、進行状況
- アクティブ項目: ゴールドの番号

### 11. テーブルスライド
- 用途: 表データの表示

### 12. 画像中央配置スライド
- 用途: 図、スクリーンショット

### 13. テキスト + 図スライド
- 用途: 説明文と図の組み合わせ

### 14. アイコングリッドスライド
- 用途: 機能一覧、サービス紹介
- アイコン: ゴールドの円

### 15. 強調セクションスライド
- 用途: 右側にタイトル + 下線
- 元テンプレートのslide5スタイル

### 16. クロージングスライド
- 用途: 締めくくり
- タイトル: ゴールド下線付き

## 使用方法

### テンプレート生成

```bash
npm run generate-template
```

### スクリプトからの利用

```javascript
const pptxgen = require("pptxgenjs");
const { COLORS, FONTS, LAYOUT } = require("./template/theme");
const {
  addTitleSlide,
  addContentSlide,
  addHeaderLine,
  addFooter,
} = require("./template/generate-template");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

addTitleSlide(pres, "タイトル", "サブタイトル");
addContentSlide(pres, "見出し", ["ポイント1", "ポイント2"]);

pres.writeFile({ fileName: "output.pptx" });
```

## カスタマイズ

### 著作権年の変更

`theme.js` の `FOOTER.COPYRIGHT` を編集:

```javascript
const FOOTER = {
  COPYRIGHT: "Copyright © 2025 SBI Traceability co., Ltd. All Rights Reserved",
  // ...
};
```

### カラーの変更

`theme.js` の `COLORS` オブジェクトを編集。変更後は `npm run generate-template` で再生成してください。
