# SBI Traceability PptxGenJS テンプレート

PptxGenJS を使用した PowerPoint スライド自動生成テンプレートです。SBI Traceability のブランドデザイン（カラー、フォント、レイアウト）に準拠したスライドを、AIエージェント（Claude Code）と組み合わせて簡単に作成できます。

## このリポジトリの良いところ

| メリット | 説明 |
|----------|------|
| **ブランド統一** | カラー・フォント・レイアウトがテーマファイルで一元管理。誰が作っても同じデザインに |
| **AI連携** | Claude Codeに「〇〇のスライドを作って」と指示するだけでPPTXを自動生成 |
| **16種類のレイアウト** | タイトル、箇条書き、KPI、タイムライン、比較表など、よく使うレイアウトを網羅 |
| **カスタマイズ可能** | theme.js を編集するだけでカラーやフォントを変更可能 |
| **コードで管理** | スライドがコードなのでGit管理・レビュー・再利用が容易 |

## セットアップ

```bash
# 1. リポジトリをクローン
git clone <repository-url>
cd sbita-pptxgenjs-template

# 2. 依存パッケージをインストール
npm install

# 3. サンプルテンプレートを生成（動作確認）
npm run generate-template
# → template/sbita-template.pptx が生成される
```

## 使い方

### 方法1: Claude Code に依頼する（推奨）

このディレクトリで Claude Code を起動し、以下のように依頼：

```
「クラウド移行の提案書を5枚で作成して」
「新サービス紹介のスライドを作って」
「月次レポート用のテンプレートを作って」
```

Claude Code がスクリプトを作成・実行し、PPTXファイルを生成します。

### 方法2: 自分でスクリプトを書く

```javascript
const pptxgen = require("pptxgenjs");
const {
  addTitleSlide,
  addContentSlide,
  addClosingSlide,
} = require("./template/generate-template");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";

// スライドを追加
addTitleSlide(pres, "プレゼンタイトル", "サブタイトル", "発表者", "2026年");
addContentSlide(pres, "概要", ["ポイント1", "ポイント2", "ポイント3"]);
addClosingSlide(pres, "ご清聴ありがとうございました");

// ファイル出力
pres.writeFile({ fileName: "output.pptx" });
```

```bash
node your-script.js
```

## 利用可能なレイアウト（16種類）

| 関数名 | 用途 |
|--------|------|
| `addTitleSlide` | タイトル（表紙） |
| `addSectionSlide` | セクション区切り |
| `addContentSlide` | 箇条書きコンテンツ |
| `addTwoColumnSlide` | 2カラム |
| `addThreeColumnSlide` | 3カラム |
| `addKPISlide` | KPIカード（数値表示） |
| `addTimelineSlide` | タイムライン/プロセスフロー |
| `addComparisonSlide` | Before/After 比較 |
| `addQuoteSlide` | 引用・メッセージ |
| `addAgendaSlide` | アジェンダ（目次） |
| `addTableSlide` | テーブル |
| `addImageCenterSlide` | 画像中央配置 |
| `addTextImageSlide` | テキスト + 図 |
| `addIconGridSlide` | アイコングリッド |
| `addHighlightSlide` | 強調セクション |
| `addClosingSlide` | クロージング |

## ファイル構成

```
sbita-pptxgenjs-template/
├── package.json              # 依存関係
├── AGENTS.md                 # AI向けガイド（Claude Codeが参照）
├── template/
│   ├── theme.js              # カラー・フォント・レイアウト定数
│   ├── generate-template.js  # テンプレート関数 & サンプル生成
│   ├── TEMPLATE.md           # 詳細仕様書
│   └── sbita-template.pptx   # 生成されたサンプル
└── examples/
    ├── how-to-use.js         # 使い方説明スライド生成スクリプト
    └── how-to-use.pptx       # 生成されたサンプル
```

## カスタマイズ

`template/theme.js` を編集：

```javascript
// カラーを変更
const COLORS = {
  BRAND_GOLD: "CBC269",    // ← ここを変更
  BRAND_RED: "EA4545",
  // ...
};

// フッターを変更
const FOOTER = {
  COPYRIGHT: "Copyright © 2026 Your Company All Rights Reserved",
  // ...
};
```

変更後は `npm run generate-template` で再生成。

## 依存パッケージ

- **pptxgenjs**: PowerPoint生成
- **react / react-dom**: アイコン変換用
- **react-icons**: アイコンライブラリ
- **sharp**: 画像処理
