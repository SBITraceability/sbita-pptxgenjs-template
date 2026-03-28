# SBI Traceability PowerPoint テンプレート - AIエージェント向けガイド

## 概要

このリポジトリは SBI Traceability のブランドに沿った PowerPoint スライドを PptxGenJS で作成するためのテンプレートです。

## ファイル構成

```
sbita-pptxgenjs-template/
├── AGENTS.md                     # このファイル（AI向けガイド）
├── package.json                  # 依存関係
├── template/
│   ├── theme.js                  # カラー・フォント・レイアウト定数
│   ├── generate-template.js      # テンプレート生成スクリプト
│   ├── sbita-template.pptx       # 生成済みテンプレート
│   └── TEMPLATE.md               # 詳細仕様書
├── examples/                     # サンプルスクリプト & PPTX
└── output/                       # 生成したPPTXの出力先（.gitignore対象）
```

## 重要: 出力先について

**新しくスライドを作成する場合は、必ず `output/` フォルダに出力してください。**

```javascript
pres.writeFile({ fileName: "output/my-presentation.pptx" });
```

`output/` フォルダは `.gitignore` で除外されているため、リポジトリを汚しません。

## デザインルール

### 1. カラー基準

| 用途 | カラー | 使用場面 |
|------|--------|----------|
| ゴールド `#CBC269` | 下線、ヘッダーライン、KPIボーダー、タイムライン |
| メインテキスト `#454545` | 本文、見出し |
| サブテキスト `#535353` | 補足テキスト |
| 強調レッド `#EA4545` | 強調テキスト |
| サーモン `#ED6A5B` | "Strictly Confidential" |

### 2. フォント基準

- **英語**: Century Gothic
- **日本語**: メイリオ
- **UI/フッター**: Meiryo UI

### 3. 共通要素

すべてのスライドに以下を含める:

1. **ヘッダーライン**: スライド上部中央にゴールド (`#CBC269`) の水平線
2. **フッター**:
   - 左: "Strictly Confidential" (サーモンレッド)
   - 中央: "Copyright © 2024 SBI Traceability co., Ltd. All Rights Reserved"

### 4. 下線スタイル

重要なテキストにはゴールドの下線を使用:
```javascript
underline: { type: "sng", color: "CBC269" }
```

## レイアウト選択ガイド

| 内容 | 推奨レイアウト |
|------|----------------|
| 表紙 | タイトルスライド |
| 章の区切り | セクションディバイダー |
| 箇条書き | 標準コンテンツスライド |
| 対比 | 2カラムレイアウト / 比較スライド |
| 複数項目 | 3カラムレイアウト / アイコングリッド |
| 数値・KPI | KPIカードスライド |
| 手順・フロー | タイムラインスライド |
| 引用・強調 | 引用スライド |
| 目次 | アジェンダスライド |
| 表データ | テーブルスライド |
| 図・画像 | 画像中央配置 / テキスト+図 |
| 右タイトル | 強調セクションスライド |
| 締め | クロージングスライド |

## 作成方法

### 新規作成（PptxGenJS）

```javascript
const pptxgen = require("pptxgenjs");
const { COLORS, FONTS, LAYOUT } = require("./template/theme");
const { addHeaderLine, addFooter } = require("./template/generate-template");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

const slide = pres.addSlide();

// ヘッダーライン追加
addHeaderLine(slide);

// コンテンツ追加
slide.addText("タイトル", {
  x: 0.8,
  y: 2.0,
  w: 11.7,
  h: 1,
  fontFace: FONTS.HEADING_JP,
  fontSize: 24,
  color: COLORS.TEXT_PRIMARY,
  bold: true,
  underline: { type: "sng", color: COLORS.BRAND_GOLD },
});

// フッター追加
addFooter(slide);

pres.writeFile({ fileName: "output/presentation.pptx" });
```

### テンプレート関数を使用

```javascript
const { addTitleSlide, addContentSlide } = require("./template/generate-template");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

addTitleSlide(pres, "プレゼンタイトル", "サブタイトル", "発表者", "2024/01/01");
addContentSlide(pres, "見出し", ["ポイント1", "ポイント2", "ポイント3"]);

pres.writeFile({ fileName: "output/presentation.pptx" });
```

## 利用可能な関数

| 関数名 | 用途 |
|--------|------|
| `addTitleSlide(pres, title, subtitle, author, date)` | タイトルスライド |
| `addSectionSlide(pres, title, headerText)` | セクション区切り |
| `addContentSlide(pres, heading, bulletPoints[])` | 箇条書きスライド |
| `addTwoColumnSlide(pres, heading, left, right)` | 2カラム |
| `addThreeColumnSlide(pres, heading, columns[])` | 3カラム |
| `addKPISlide(pres, heading, kpis[])` | KPIカード |
| `addTimelineSlide(pres, heading, steps[])` | タイムライン |
| `addComparisonSlide(pres, heading, before, after)` | Before/After比較 |
| `addQuoteSlide(pres, quote, author)` | 引用 |
| `addAgendaSlide(pres, items[], currentIndex)` | アジェンダ |
| `addTableSlide(pres, heading, data[][])` | テーブル |
| `addImageCenterSlide(pres, heading, imagePath, caption)` | 画像中央 |
| `addTextImageSlide(pres, heading, text, imagePosition)` | テキスト+図 |
| `addIconGridSlide(pres, heading, items[])` | アイコングリッド |
| `addHighlightSlide(pres, title)` | 強調セクション |
| `addClosingSlide(pres, thankYou, contact)` | クロージング |
| `addHeaderLine(slide)` | ヘッダーライン追加 |
| `addHeaderText(slide, text)` | ヘッダーテキスト追加 |
| `addFooter(slide, showConfidential)` | フッター追加 |

## セットアップ

```bash
# 依存関係のインストール
npm install

# テンプレートの生成（任意）
npm run generate-template
```

## プロンプト例

AI に以下のようなプロンプトでスライド作成を依頼できます：

- 「トレーサビリティソリューションの提案書を作成してください」
- 「新サービスの紹介スライドを5枚作成してください」
- 「月次レポートのテンプレートを作成してください」
