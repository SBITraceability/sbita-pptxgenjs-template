/**
 * サンプル: テンプレートの使い方説明スライド
 */

const pptxgen = require("pptxgenjs");
const path = require("path");

// テンプレート関数をインポート
const {
  addTitleSlide,
  addSectionSlide,
  addContentSlide,
  addAgendaSlide,
  addTimelineSlide,
  addThreeColumnSlide,
  addKPISlide,
  addClosingSlide,
} = require("../template/generate-template");

const { COLORS, FONTS } = require("../template/theme");

async function createHowToUseSlides() {
  console.log("使い方説明スライドを生成中...\n");

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "SBI Traceability";
  pres.title = "PptxGenJS テンプレート 使い方ガイド";

  // 1. タイトルスライド
  console.log("1. タイトル");
  addTitleSlide(
    pres,
    "PptxGenJS テンプレート",
    "AIを活用した",
    "スライド自動生成システム - ",
    "SBI Traceability"
  );

  // 2. アジェンダ
  console.log("2. アジェンダ");
  addAgendaSlide(
    pres,
    [
      "このシステムについて",
      "セットアップ方法",
      "スライドの作成方法",
      "利用可能なレイアウト",
      "まとめ",
    ],
    -1
  );

  // 3. セクション1
  console.log("3. セクション1");
  addSectionSlide(pres, "このシステムについて", "Overview");

  // 4. システム概要
  console.log("4. システム概要");
  addContentSlide(pres, "PptxGenJS テンプレートとは", [
    "Node.js + PptxGenJS を使用したスライド自動生成システム",
    "SBI Traceability のブランドデザインに準拠",
    "AIエージェント（Claude Code等）と組み合わせて使用可能",
    "16種類のプリセットレイアウトを用意",
    "カスタマイズ可能なテーマ設定（カラー・フォント）",
  ]);

  // 5. メリット
  console.log("5. メリット");
  addThreeColumnSlide(pres, "導入のメリット", [
    "効率化\n\nスライド作成時間を大幅に短縮。定型フォーマットの再利用が可能",
    "統一感\n\nブランドガイドラインに沿った一貫性のあるデザイン",
    "自動化\n\nAIと連携してプロンプトからスライドを自動生成",
  ]);

  // 6. セクション2
  console.log("6. セクション2");
  addSectionSlide(pres, "セットアップ方法", "Setup");

  // 7. セットアップ手順
  console.log("7. セットアップ手順");
  addTimelineSlide(pres, "セットアップ手順", [
    { title: "Step 1", desc: "リポジトリをクローン" },
    { title: "Step 2", desc: "npm install" },
    { title: "Step 3", desc: "テーマをカスタマイズ" },
    { title: "Step 4", desc: "スライド生成" },
  ]);

  // 8. コマンド説明
  console.log("8. コマンド説明");
  addContentSlide(pres, "基本コマンド", [
    "npm install - 依存パッケージをインストール",
    "npm run generate-template - サンプルテンプレートを生成",
    "node examples/how-to-use.js - このスライドを生成",
  ]);

  // 9. セクション3
  console.log("9. セクション3");
  addSectionSlide(pres, "スライドの作成方法", "Usage");

  // 10. 基本的な使い方
  console.log("10. 基本的な使い方");
  addContentSlide(pres, "基本的な使い方", [
    "1. pptxgenjs と テンプレート関数をインポート",
    "2. new pptxgen() でプレゼンテーションを作成",
    "3. addXXXSlide() 関数でスライドを追加",
    "4. writeFile() でPPTXファイルを出力",
  ]);

  // 11. セクション4
  console.log("11. セクション4");
  addSectionSlide(pres, "利用可能なレイアウト", "Layouts");

  // 12. レイアウト一覧
  console.log("12. レイアウト一覧");
  addKPISlide(pres, "16種類のレイアウト", [
    { value: "6種", label: "基本レイアウト" },
    { value: "5種", label: "データ表示" },
    { value: "5種", label: "特殊用途" },
  ]);

  // 13. 基本レイアウト詳細
  console.log("13. 基本レイアウト");
  addContentSlide(pres, "基本レイアウト", [
    "addTitleSlide - タイトル（表紙）",
    "addSectionSlide - セクション区切り",
    "addContentSlide - 標準コンテンツ（箇条書き）",
    "addTwoColumnSlide - 2カラム",
    "addThreeColumnSlide - 3カラム",
    "addClosingSlide - クロージング",
  ]);

  // 14. データ表示レイアウト
  console.log("14. データ表示レイアウト");
  addContentSlide(pres, "データ表示レイアウト", [
    "addKPISlide - KPIカード（数値表示）",
    "addTimelineSlide - タイムライン/プロセスフロー",
    "addComparisonSlide - Before/After 比較",
    "addTableSlide - テーブル",
    "addAgendaSlide - アジェンダ（目次）",
  ]);

  // 15. 特殊用途レイアウト
  console.log("15. 特殊用途レイアウト");
  addContentSlide(pres, "特殊用途レイアウト", [
    "addQuoteSlide - 引用・メッセージ",
    "addImageCenterSlide - 画像中央配置",
    "addTextImageSlide - テキスト + 図",
    "addIconGridSlide - アイコングリッド（機能一覧）",
    "addHighlightSlide - 強調セクション",
  ]);

  // 16. まとめ
  console.log("16. まとめ");
  addSectionSlide(pres, "まとめ", "Summary");

  // 17. ポイント
  console.log("17. ポイント");
  addContentSlide(pres, "ポイント", [
    "テーマ設定は theme.js で一元管理",
    "スライド関数を組み合わせて自由にプレゼンを構成",
    "AIエージェントと連携して自動生成も可能",
    "カスタマイズは theme.js のカラー・フォントを変更",
  ]);

  // 18. クロージング
  console.log("18. クロージング");
  addClosingSlide(
    pres,
    "ご清聴ありがとうございました",
    "ご質問があればお気軽にどうぞ"
  );

  // ファイル出力
  const outputPath = path.join(__dirname, "how-to-use.pptx");
  await pres.writeFile({ fileName: outputPath });

  console.log(`\n✅ 生成完了: ${outputPath}`);
  console.log(`📊 合計 ${pres.slides.length} スライド`);
}

// 実行
createHowToUseSlides().catch(console.error);
