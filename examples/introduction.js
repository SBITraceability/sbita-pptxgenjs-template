/**
 * README紹介スライド - 新テンプレート版
 */

const pptxgen = require("pptxgenjs");
const path = require("path");

const {
  addTitleSlide,
  addSectionSlide,
  addContentSlide,
  addThreeColumnSlide,
  addKPISlide,
  addTimelineSlide,
  addComparisonSlide,
  addQuoteSlide,
  addClosingSlide,
  addHeaderLine,
  addFooter,
} = require("../template/generate-template");

const { COLORS, FONTS, FONT_SIZES, LAYOUT } = require("../template/theme");

async function createIntroSlides() {
  console.log("使い方紹介スライドを生成中...\n");

  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE";
  pres.author = "SBI Traceability";
  pres.title = "PptxGenJS テンプレート紹介";

  // 1. タイトルスライド（新フォーマット）
  console.log("1. タイトル");
  addTitleSlide(
    pres,
    "PptxGenJS テンプレート",
    "AIを活用した",
    "スライド自動生成システム - ",
    "使い方ガイド"
  );

  // 2. 問題提起
  console.log("2. 問題提起");
  const slide2 = pres.addSlide();
  addHeaderLine(slide2);
  slide2.addText("こんな経験ありませんか？", {
    x: 0,
    y: 0.1,
    w: LAYOUT.WIDTH,
    h: 0.4,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.HEADER,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  });

  const problems = [
    "毎回パワポのデザインがバラバラになる...",
    "フォントやカラーを揃えるのが面倒...",
    "同じようなスライドを何度も作っている...",
    "急ぎの資料作成で深夜残業...",
    "「もっとシンプルに作れないの？」と思う毎日",
  ];

  slide2.addText(problems.map(p => ({ text: p, options: { bullet: { code: "25CF" }, color: COLORS.BRAND_RED } })), {
    x: 1.5,
    y: 1.2,
    w: LAYOUT.WIDTH - 3,
    h: 4.5,
    fontFace: FONTS.BODY_JP,
    fontSize: 22,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
    lineSpacing: 40,
  });
  addFooter(slide2);

  // 3. 解決策ドーン
  console.log("3. 解決策");
  const slide3 = pres.addSlide();
  slide3.addShape("rect", {
    x: 0, y: 0, w: LAYOUT.WIDTH, h: LAYOUT.HEIGHT,
    fill: { color: COLORS.BRAND_GOLD },
  });
  slide3.addText("そこで", {
    x: 0.5,
    y: 1.5,
    w: LAYOUT.WIDTH - 1,
    h: 0.8,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: 28,
    color: COLORS.BG_WHITE,
  });
  slide3.addText("PptxGenJS テンプレート", {
    x: 0.5,
    y: 2.5,
    w: LAYOUT.WIDTH - 1,
    h: 1.2,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: 48,
    color: COLORS.BG_WHITE,
    bold: true,
  });
  slide3.addText("AIに話しかけるだけでスライドが完成", {
    x: 0.5,
    y: 4.0,
    w: LAYOUT.WIDTH - 1,
    h: 0.8,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: 24,
    color: COLORS.TEXT_PRIMARY,
  });

  // 4. デモ：こう言うだけ
  console.log("4. デモ");
  const slide4 = pres.addSlide();
  addHeaderLine(slide4);
  slide4.addText("使い方はカンタン", {
    x: 0,
    y: 0.1,
    w: LAYOUT.WIDTH,
    h: 0.4,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.HEADER,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  });

  slide4.addShape("rect", {
    x: 1.5,
    y: 1.2,
    w: LAYOUT.WIDTH - 3,
    h: 1.8,
    fill: { color: "F0F0F0" },
    line: { color: COLORS.TEXT_LIGHT, width: 1 },
  });

  slide4.addText("Claude Code に話しかける:", {
    x: 1.7,
    y: 1.3,
    w: 4,
    h: 0.4,
    fontFace: FONTS.BODY_JP,
    fontSize: 14,
    color: COLORS.TEXT_LIGHT,
  });

  slide4.addText("「クラウド移行の提案書を5枚で作成して」", {
    x: 1.7,
    y: 1.8,
    w: LAYOUT.WIDTH - 3.4,
    h: 0.8,
    fontFace: FONTS.HEADING_JP,
    fontSize: 28,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  });

  slide4.addText("これだけ。マジで。", {
    x: 0.5,
    y: 3.5,
    w: LAYOUT.WIDTH - 1,
    h: 0.8,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: 36,
    color: COLORS.BRAND_GOLD,
    bold: true,
  });

  slide4.addText("あとは AI がスクリプトを書いて実行、PPTXファイルが完成", {
    x: 0.5,
    y: 4.5,
    w: LAYOUT.WIDTH - 1,
    h: 0.6,
    align: "center",
    fontFace: FONTS.BODY_JP,
    fontSize: 18,
    color: COLORS.TEXT_SECONDARY,
  });
  addFooter(slide4);

  // 5. Before/After
  console.log("5. Before/After");
  addComparisonSlide(
    pres,
    "何が変わる？",
    "従来のパワポ作成\n\n- PowerPointを開く\n- テンプレ探す\n- コピペ繰り返し\n- デザイン調整...\n- 「フォント違う！」\n- 「色が微妙...」\n- 修正地獄\n\n所要時間: 2〜3時間",
    "これからのパワポ作成\n\n- Claude Code起動\n- 「〇〇のスライド作って」\n- 待つ（数秒）\n- 完成\n\n\n\n\n所要時間: 1分"
  );

  // 6. メリット3つ
  console.log("6. メリット");
  addThreeColumnSlide(pres, "ここがスゴイ", [
    "ブランド統一\n\nカラー・フォントが\n自動で統一\n\n誰が作っても\n同じクオリティ",
    "爆速生成\n\n話しかけるだけで\n数秒で完成\n\n深夜残業とは\nおさらば",
    "16種類の\nレイアウト\n\nよく使うパターンを\n網羅済み\n\n組み合わせ自由",
  ]);

  // 7. KPIスライド - 数字でインパクト
  console.log("7. 数字でインパクト");
  addKPISlide(pres, "どれくらい楽になる？", [
    { value: "1/60", label: "作成時間" },
    { value: "16種", label: "レイアウト" },
    { value: "0円", label: "追加コスト" },
  ]);

  // 8. セットアップ
  console.log("8. セットアップ");
  addTimelineSlide(pres, "セットアップも超カンタン", [
    { title: "Clone", desc: "リポジトリをクローン" },
    { title: "Install", desc: "npm install" },
    { title: "Generate", desc: "npm run generate-template" },
    { title: "Done!", desc: "すぐ使える" },
  ]);

  // 9. レイアウト一覧
  console.log("9. レイアウト一覧");
  addContentSlide(pres, "16種類のレイアウトを用意", [
    "タイトル / セクション区切り / クロージング",
    "箇条書き / 2カラム / 3カラム",
    "KPIカード / タイムライン / Before-After比較",
    "アジェンダ / テーブル / 引用",
    "画像中央 / テキスト+図 / アイコングリッド / 強調",
  ]);

  // 10. コード例
  console.log("10. コード例");
  const slide10 = pres.addSlide();
  addHeaderLine(slide10);
  slide10.addText("自分でスクリプトも書ける", {
    x: 0,
    y: 0.1,
    w: LAYOUT.WIDTH,
    h: 0.4,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.HEADER,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  });

  slide10.addShape("rect", {
    x: 0.8,
    y: 0.9,
    w: LAYOUT.WIDTH - 1.6,
    h: 4.8,
    fill: { color: "2D2D2D" },
  });

  const codeLines = [
    "const pptxgen = require('pptxgenjs');",
    "const { addTitleSlide, addContentSlide }",
    "  = require('./template/generate-template');",
    "",
    "const pres = new pptxgen();",
    "pres.layout = 'LAYOUT_WIDE';",
    "",
    "// タイトルスライドを追加",
    "addTitleSlide(pres, 'メインタイトル', 'サブ', '備考', 'ハイライト');",
    "",
    "pres.writeFile({ fileName: 'output/presentation.pptx' });",
  ];

  slide10.addText(codeLines.join("\n"), {
    x: 1.0,
    y: 1.1,
    w: LAYOUT.WIDTH - 2,
    h: 4.4,
    fontFace: "Consolas",
    fontSize: 14,
    color: "E0E0E0",
    valign: "top",
  });
  addFooter(slide10);

  // 11. 出力先について
  console.log("11. 出力先");
  addContentSlide(pres, "出力先フォルダについて", [
    "output/ フォルダに出力すればリポジトリを汚さない",
    "output/ は .gitignore で除外済み",
    "サンプルPPTXは template/ と examples/ に配置",
    "新規作成時は output/ に保存を推奨",
  ]);

  // 12. カスタマイズ
  console.log("12. カスタマイズ");
  addContentSlide(pres, "カスタマイズも自由自在", [
    "theme.js を編集するだけ",
    "カラーパレットを自社ブランドに変更可能",
    "フォントも自由に設定",
    "フッターの著作権表示もカンタン変更",
    "変更後は npm run generate-template で反映",
  ]);

  // 13. 引用（推しコメント風）
  console.log("13. 推しコメント");
  addQuoteSlide(
    pres,
    "パワポ作成が苦痛だった過去の自分に教えてあげたい。\nこのテンプレートがあれば、もっと早く帰れたのに。",
    "とある開発者"
  );

  // 14. まとめ
  console.log("14. まとめ");
  const slide14 = pres.addSlide();
  addHeaderLine(slide14);
  slide14.addText("まとめ", {
    x: 0,
    y: 0.1,
    w: LAYOUT.WIDTH,
    h: 0.4,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.HEADER,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  });

  const summaryPoints = [
    "Claude Code に話しかけるだけでスライド完成",
    "ブランドデザインが自動で統一される",
    "16種類のレイアウトですぐ使える",
    "セットアップは npm install だけ",
    "カスタマイズも theme.js を編集するだけ",
  ];

  slide14.addText(summaryPoints.map(p => ({
    text: p,
    options: { bullet: { code: "2713", color: COLORS.BRAND_GOLD } }
  })), {
    x: 1.5,
    y: 1.2,
    w: LAYOUT.WIDTH - 3,
    h: 4.5,
    fontFace: FONTS.BODY_JP,
    fontSize: 22,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
    lineSpacing: 36,
  });
  addFooter(slide14);

  // 15. CTA
  console.log("15. CTA");
  const slide15 = pres.addSlide();
  slide15.addShape("rect", {
    x: 0, y: 0, w: LAYOUT.WIDTH, h: LAYOUT.HEIGHT,
    fill: { color: COLORS.TEXT_PRIMARY },
  });
  slide15.addText("今すぐ試してみよう", {
    x: 0.5,
    y: 2.0,
    w: LAYOUT.WIDTH - 1,
    h: 1.0,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: 40,
    color: COLORS.BG_WHITE,
    bold: true,
  });
  slide15.addText("git clone → npm install → 完了", {
    x: 0.5,
    y: 3.3,
    w: LAYOUT.WIDTH - 1,
    h: 0.8,
    align: "center",
    fontFace: FONTS.HEADING_EN,
    fontSize: 28,
    color: COLORS.BRAND_GOLD,
    bold: true,
  });
  slide15.addText("あなたのパワポ作成ライフが変わります", {
    x: 0.5,
    y: 4.5,
    w: LAYOUT.WIDTH - 1,
    h: 0.6,
    align: "center",
    fontFace: FONTS.BODY_JP,
    fontSize: 20,
    color: COLORS.TEXT_LIGHT,
  });

  // 16. クロージング
  console.log("16. クロージング");
  addClosingSlide(
    pres,
    "Let's Automate!",
    "質問があればお気軽にどうぞ"
  );

  // ファイル出力
  const outputPath = path.join(__dirname, "introduction.pptx");
  await pres.writeFile({ fileName: outputPath });

  console.log(`\n✅ 生成完了: ${outputPath}`);
  console.log(`📊 合計 ${pres.slides.length} スライド`);
}

// 実行
createIntroSlides().catch(console.error);
