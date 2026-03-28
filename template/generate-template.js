/**
 * SBI Traceability PowerPoint Template Generator
 * PptxGenJSを使用してテンプレートPPTXを生成
 */

const pptxgen = require("pptxgenjs");
const path = require("path");

// テーマ設定を読み込み
const {
  COLORS,
  FONTS,
  FONT_SIZES,
  LAYOUT,
  STYLES,
  FOOTER,
  makeShadow,
  makeCardShadow,
} = require("./theme");

// ============================================
// 共通要素
// ============================================

/**
 * ヘッダーライン（ゴールド）を追加
 */
function addHeaderLine(slide) {
  const lineX = (LAYOUT.WIDTH - LAYOUT.HEADER_LINE_WIDTH) / 2;
  slide.addShape("line", {
    x: lineX,
    y: LAYOUT.HEADER_LINE_Y,
    w: LAYOUT.HEADER_LINE_WIDTH,
    h: 0,
    line: { color: COLORS.BRAND_GOLD, width: 1.25 },
  });
}

/**
 * ヘッダーテキストを追加
 */
function addHeaderText(slide, text) {
  slide.addText(text, {
    x: 0,
    y: LAYOUT.HEADER_Y,
    w: LAYOUT.WIDTH,
    h: 0.4,
    align: "center",
    ...STYLES.header,
  });
}

/**
 * フッターを追加（著作権表示 + Strictly Confidential）
 */
function addFooter(slide, showConfidential = true) {
  // Strictly Confidential (左側)
  if (showConfidential) {
    slide.addText(FOOTER.CONFIDENTIAL, {
      x: 0.17,
      y: LAYOUT.FOOTER_Y,
      w: 1.2,
      h: 0.2,
      ...STYLES.confidential,
    });
  }

  // Copyright (中央)
  slide.addText(FOOTER.COPYRIGHT, {
    x: 4.8,
    y: LAYOUT.FOOTER_Y,
    w: 4,
    h: 0.2,
    align: "right",
    ...STYLES.footer,
  });
}

// ============================================
// スライド生成関数
// ============================================

/**
 * 1. タイトルスライド
 */
function addTitleSlide(pres, title, subtitle, author = "", date = "") {
  const slide = pres.addSlide();

  // タイトル（下線付き）
  slide.addText(title || "プレゼンテーションタイトル", {
    x: LAYOUT.MARGIN,
    y: 2.8,
    w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
    h: 1.0,
    align: "center",
    lineSpacing: 36,
    ...STYLES.titleWithUnderline,
  });

  // サブタイトル
  if (subtitle) {
    slide.addText(subtitle, {
      x: LAYOUT.MARGIN,
      y: 4.0,
      w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
      h: 0.6,
      align: "center",
      fontFace: FONTS.HEADING_JP,
      fontSize: FONT_SIZES.HEADING2,
      color: COLORS.TEXT_SECONDARY,
    });
  }

  // 著者・日付
  if (author || date) {
    slide.addText(`${author}${author && date ? "  |  " : ""}${date}`, {
      x: LAYOUT.MARGIN,
      y: 5.5,
      w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
      h: 0.4,
      align: "center",
      fontFace: FONTS.BODY_JP,
      fontSize: FONT_SIZES.BODY_SMALL,
      color: COLORS.TEXT_LIGHT,
    });
  }

  addFooter(slide);
  return slide;
}

/**
 * 2. セクションディバイダー（Vision/中扉スライド）
 */
function addSectionSlide(pres, title, headerText = "") {
  const slide = pres.addSlide();

  // ヘッダー
  if (headerText) {
    addHeaderText(slide, headerText);
  }
  addHeaderLine(slide);

  // 大見出し（中央）
  slide.addText(title || "セクションタイトル", {
    x: LAYOUT.MARGIN,
    y: 2.2,
    w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
    h: 1.5,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.HEADING1,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  });

  // 備考（中央下）
  slide.addText("備考", {
    x: LAYOUT.MARGIN,
    y: 3.5,
    w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
    h: 0.4,
    align: "center",
    fontFace: FONTS.HEADING_JP,
    fontSize: 12,
    color: COLORS.TEXT_LIGHT,
    bold: true,
  });

  addFooter(slide);
  return slide;
}

/**
 * 3. 標準コンテンツスライド（ヘッダー + 箇条書き）
 */
function addContentSlide(pres, heading, bulletPoints = []) {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "ヘッダー");
  addHeaderLine(slide);

  // 箇条書き
  const items = bulletPoints.length > 0 ? bulletPoints : [
    "文字サイズ",
  ];

  const textItems = items.map((text) => ({
    text: text,
    options: {
      bullet: { type: "bullet", code: "l", color: COLORS.TEXT_PRIMARY },
      indentLevel: 0,
    },
  }));

  slide.addText(textItems, {
    x: LAYOUT.CONTENT_LEFT,
    y: LAYOUT.CONTENT_TOP + 0.2,
    w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
    h: LAYOUT.HEIGHT - LAYOUT.CONTENT_TOP - 1.5,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.BODY,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
    valign: "top",
    lineSpacing: 24,
  });

  addFooter(slide);
  return slide;
}

/**
 * 4. 2カラムレイアウト
 */
function addTwoColumnSlide(pres, heading, leftContent, rightContent) {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "2カラムレイアウト");
  addHeaderLine(slide);

  const columnWidth = (LAYOUT.WIDTH - LAYOUT.MARGIN * 2 - LAYOUT.COLUMN_GAP) / 2;

  // 左カラム
  slide.addText(leftContent || "左側コンテンツ", {
    x: LAYOUT.CONTENT_LEFT,
    y: LAYOUT.CONTENT_TOP,
    w: columnWidth,
    h: 5.0,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.BODY,
    color: COLORS.TEXT_PRIMARY,
    valign: "top",
  });

  // 右カラム
  slide.addText(rightContent || "右側コンテンツ", {
    x: LAYOUT.CONTENT_LEFT + columnWidth + LAYOUT.COLUMN_GAP,
    y: LAYOUT.CONTENT_TOP,
    w: columnWidth,
    h: 5.0,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.BODY,
    color: COLORS.TEXT_PRIMARY,
    valign: "top",
  });

  addFooter(slide);
  return slide;
}

/**
 * 5. 3カラムレイアウト
 */
function addThreeColumnSlide(pres, heading, columns = []) {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "3カラムレイアウト");
  addHeaderLine(slide);

  const columnWidth = (LAYOUT.WIDTH - LAYOUT.MARGIN * 2 - LAYOUT.COLUMN_GAP * 2) / 3;
  const defaultColumns = ["カラム1", "カラム2", "カラム3"];

  for (let i = 0; i < 3; i++) {
    const xPos = LAYOUT.CONTENT_LEFT + i * (columnWidth + LAYOUT.COLUMN_GAP);
    const content = columns[i] || defaultColumns[i];

    // カラムカード
    slide.addShape("rect", {
      x: xPos,
      y: LAYOUT.CONTENT_TOP,
      w: columnWidth,
      h: 4.5,
      fill: { color: COLORS.BG_LIGHT },
      line: { color: COLORS.BORDER, width: 1 },
      shadow: makeCardShadow(),
    });

    slide.addText(content, {
      x: xPos + 0.2,
      y: LAYOUT.CONTENT_TOP + 0.2,
      w: columnWidth - 0.4,
      h: 4.1,
      fontFace: FONTS.BODY_JP,
      fontSize: FONT_SIZES.BODY,
      color: COLORS.TEXT_PRIMARY,
      valign: "top",
    });
  }

  addFooter(slide);
  return slide;
}

/**
 * 6. KPIカードスライド
 */
function addKPISlide(pres, heading, kpis = []) {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "主要指標");
  addHeaderLine(slide);

  const defaultKPIs = [
    { value: "99.9%", label: "稼働率" },
    { value: "50%", label: "コスト削減" },
    { value: "2倍", label: "処理速度" },
  ];

  const kpiData = kpis.length > 0 ? kpis : defaultKPIs;
  const cardWidth = (LAYOUT.WIDTH - LAYOUT.MARGIN * 2 - LAYOUT.COLUMN_GAP * (kpiData.length - 1)) / kpiData.length;

  kpiData.forEach((kpi, i) => {
    const xPos = LAYOUT.CONTENT_LEFT + i * (cardWidth + LAYOUT.COLUMN_GAP);

    // カード背景
    slide.addShape("rect", {
      x: xPos,
      y: LAYOUT.CONTENT_TOP + 0.5,
      w: cardWidth,
      h: 3.0,
      fill: { color: COLORS.BG_WHITE },
      line: { color: COLORS.BRAND_GOLD, width: 2 },
      shadow: makeCardShadow(),
    });

    // 数値
    slide.addText(kpi.value, {
      x: xPos,
      y: LAYOUT.CONTENT_TOP + 1.0,
      w: cardWidth,
      h: 1.2,
      fontFace: FONTS.HEADING_EN,
      fontSize: FONT_SIZES.STAT,
      color: COLORS.TEXT_PRIMARY,
      bold: true,
      align: "center",
    });

    // ラベル
    slide.addText(kpi.label, {
      x: xPos,
      y: LAYOUT.CONTENT_TOP + 2.3,
      w: cardWidth,
      h: 0.5,
      fontFace: FONTS.BODY_JP,
      fontSize: FONT_SIZES.BODY,
      color: COLORS.TEXT_SECONDARY,
      align: "center",
    });
  });

  addFooter(slide);
  return slide;
}

/**
 * 7. タイムライン/プロセスフロー
 */
function addTimelineSlide(pres, heading, steps = []) {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "プロセス/手順");
  addHeaderLine(slide);

  const defaultSteps = [
    { title: "Step 1", desc: "計画" },
    { title: "Step 2", desc: "実行" },
    { title: "Step 3", desc: "評価" },
    { title: "Step 4", desc: "改善" },
  ];

  const stepData = steps.length > 0 ? steps : defaultSteps;
  const stepWidth = (LAYOUT.WIDTH - LAYOUT.MARGIN * 2) / stepData.length;

  // 接続線
  slide.addShape("line", {
    x: LAYOUT.CONTENT_LEFT + stepWidth / 2,
    y: 2.8,
    w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2 - stepWidth,
    h: 0,
    line: { color: COLORS.BRAND_GOLD, width: 2 },
  });

  stepData.forEach((step, i) => {
    const xPos = LAYOUT.CONTENT_LEFT + i * stepWidth + stepWidth / 2;

    // 円
    slide.addShape("ellipse", {
      x: xPos - 0.4,
      y: 2.4,
      w: 0.8,
      h: 0.8,
      fill: { color: COLORS.BRAND_GOLD },
    });

    // ステップ番号
    slide.addText((i + 1).toString(), {
      x: xPos - 0.4,
      y: 2.5,
      w: 0.8,
      h: 0.6,
      fontFace: FONTS.HEADING_EN,
      fontSize: FONT_SIZES.HEADING2,
      color: COLORS.BG_WHITE,
      bold: true,
      align: "center",
      valign: "middle",
    });

    // タイトル
    slide.addText(step.title, {
      x: xPos - stepWidth / 2 + 0.1,
      y: 3.4,
      w: stepWidth - 0.2,
      h: 0.5,
      fontFace: FONTS.HEADING_JP,
      fontSize: FONT_SIZES.BODY,
      color: COLORS.TEXT_PRIMARY,
      bold: true,
      align: "center",
    });

    // 説明
    slide.addText(step.desc, {
      x: xPos - stepWidth / 2 + 0.1,
      y: 3.9,
      w: stepWidth - 0.2,
      h: 1.0,
      fontFace: FONTS.BODY_JP,
      fontSize: FONT_SIZES.CAPTION,
      color: COLORS.TEXT_SECONDARY,
      align: "center",
    });
  });

  addFooter(slide);
  return slide;
}

/**
 * 8. 比較スライド (Before/After)
 */
function addComparisonSlide(pres, heading, beforeContent, afterContent) {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "Before / After 比較");
  addHeaderLine(slide);

  const columnWidth = (LAYOUT.WIDTH - LAYOUT.MARGIN * 2 - LAYOUT.COLUMN_GAP) / 2;

  // Before
  slide.addShape("rect", {
    x: LAYOUT.CONTENT_LEFT,
    y: LAYOUT.CONTENT_TOP,
    w: columnWidth,
    h: 4.5,
    fill: { color: COLORS.BG_LIGHT },
    line: { color: COLORS.BRAND_RED, width: 2 },
  });

  slide.addText("Before", {
    x: LAYOUT.CONTENT_LEFT,
    y: LAYOUT.CONTENT_TOP + 0.1,
    w: columnWidth,
    h: 0.5,
    fontFace: FONTS.HEADING_EN,
    fontSize: FONT_SIZES.HEADING2,
    color: COLORS.BRAND_RED,
    bold: true,
    align: "center",
  });

  slide.addText(beforeContent || "改善前の状態", {
    x: LAYOUT.CONTENT_LEFT + 0.2,
    y: LAYOUT.CONTENT_TOP + 0.7,
    w: columnWidth - 0.4,
    h: 3.6,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.BODY,
    color: COLORS.TEXT_PRIMARY,
    valign: "top",
  });

  // After
  slide.addShape("rect", {
    x: LAYOUT.CONTENT_LEFT + columnWidth + LAYOUT.COLUMN_GAP,
    y: LAYOUT.CONTENT_TOP,
    w: columnWidth,
    h: 4.5,
    fill: { color: COLORS.BG_WHITE },
    line: { color: COLORS.ACCENT6, width: 2 },
  });

  slide.addText("After", {
    x: LAYOUT.CONTENT_LEFT + columnWidth + LAYOUT.COLUMN_GAP,
    y: LAYOUT.CONTENT_TOP + 0.1,
    w: columnWidth,
    h: 0.5,
    fontFace: FONTS.HEADING_EN,
    fontSize: FONT_SIZES.HEADING2,
    color: COLORS.ACCENT6,
    bold: true,
    align: "center",
  });

  slide.addText(afterContent || "改善後の状態", {
    x: LAYOUT.CONTENT_LEFT + columnWidth + LAYOUT.COLUMN_GAP + 0.2,
    y: LAYOUT.CONTENT_TOP + 0.7,
    w: columnWidth - 0.4,
    h: 3.6,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.BODY,
    color: COLORS.TEXT_PRIMARY,
    valign: "top",
  });

  addFooter(slide);
  return slide;
}

/**
 * 9. 引用スライド
 */
function addQuoteSlide(pres, quote, author = "") {
  const slide = pres.addSlide();

  // 引用文
  slide.addText(quote || "重要なメッセージや引用文をここに記載します。", {
    x: LAYOUT.MARGIN + 0.5,
    y: 2.5,
    w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2 - 1,
    h: 2.0,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.HEADING1,
    color: COLORS.TEXT_PRIMARY,
    italic: true,
    align: "center",
    valign: "middle",
  });

  // 下線アクセント
  slide.addShape("line", {
    x: (LAYOUT.WIDTH - 3) / 2,
    y: 4.8,
    w: 3,
    h: 0,
    line: { color: COLORS.BRAND_GOLD, width: 2 },
  });

  // 著者
  if (author) {
    slide.addText(`— ${author}`, {
      x: LAYOUT.MARGIN,
      y: 5.0,
      w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
      h: 0.5,
      fontFace: FONTS.BODY_JP,
      fontSize: FONT_SIZES.BODY,
      color: COLORS.TEXT_SECONDARY,
      align: "center",
    });
  }

  addFooter(slide);
  return slide;
}

/**
 * 10. アジェンダスライド
 */
function addAgendaSlide(pres, items = [], currentIndex = -1) {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, "Agenda");
  addHeaderLine(slide);

  const defaultItems = ["はじめに", "現状分析", "提案内容", "スケジュール", "まとめ"];
  const agendaItems = items.length > 0 ? items : defaultItems;

  agendaItems.forEach((item, i) => {
    const yPos = LAYOUT.CONTENT_TOP + 0.2 + i * 0.8;
    const isActive = i === currentIndex;

    // 番号（円）
    slide.addShape("ellipse", {
      x: LAYOUT.CONTENT_LEFT,
      y: yPos,
      w: 0.5,
      h: 0.5,
      fill: { color: isActive ? COLORS.BRAND_GOLD : COLORS.TEXT_LIGHT },
    });

    slide.addText((i + 1).toString(), {
      x: LAYOUT.CONTENT_LEFT,
      y: yPos + 0.05,
      w: 0.5,
      h: 0.4,
      fontFace: FONTS.HEADING_EN,
      fontSize: FONT_SIZES.BODY_SMALL,
      color: COLORS.BG_WHITE,
      bold: true,
      align: "center",
      valign: "middle",
    });

    // 項目テキスト
    slide.addText(item, {
      x: LAYOUT.CONTENT_LEFT + 0.7,
      y: yPos,
      w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2 - 0.7,
      h: 0.5,
      fontFace: FONTS.BODY_JP,
      fontSize: FONT_SIZES.HEADING2,
      color: isActive ? COLORS.TEXT_PRIMARY : COLORS.TEXT_SECONDARY,
      bold: isActive,
      valign: "middle",
    });
  });

  addFooter(slide);
  return slide;
}

/**
 * 11. テーブルスライド
 */
function addTableSlide(pres, heading, tableData = []) {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "データテーブル");
  addHeaderLine(slide);

  const defaultData = [
    ["項目", "値", "備考"],
    ["項目A", "100", "サンプル"],
    ["項目B", "200", "サンプル"],
    ["項目C", "300", "サンプル"],
  ];

  const data = tableData.length > 0 ? tableData : defaultData;

  slide.addTable(data, {
    x: LAYOUT.CONTENT_LEFT,
    y: LAYOUT.CONTENT_TOP,
    w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.CAPTION,
    color: COLORS.TEXT_PRIMARY,
    border: { pt: 1, color: COLORS.BORDER },
    fill: { color: COLORS.BG_WHITE },
    colW: Array(data[0].length).fill((LAYOUT.WIDTH - LAYOUT.MARGIN * 2) / data[0].length),
    rowH: 0.5,
    align: "center",
    valign: "middle",
  });

  // ヘッダー行のスタイル適用（最初の行）
  // 注: PptxGenJSでは個別セルのスタイル指定が必要

  addFooter(slide);
  return slide;
}

/**
 * 12. 画像中央配置スライド
 */
function addImageCenterSlide(pres, heading, imagePath = null, caption = "") {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "画像スライド");
  addHeaderLine(slide);

  // 画像プレースホルダー
  slide.addShape("rect", {
    x: 3,
    y: LAYOUT.CONTENT_TOP + 0.3,
    w: 7.33,
    h: 4,
    fill: { color: COLORS.BG_LIGHT },
    line: { color: COLORS.BORDER, width: 1, dashType: "dash" },
  });

  slide.addText("画像をここに配置", {
    x: 3,
    y: LAYOUT.CONTENT_TOP + 1.8,
    w: 7.33,
    h: 0.5,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.BODY,
    color: COLORS.TEXT_LIGHT,
    align: "center",
  });

  // キャプション
  if (caption) {
    slide.addText(caption, {
      x: LAYOUT.MARGIN,
      y: 5.8,
      w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
      h: 0.4,
      fontFace: FONTS.BODY_JP,
      fontSize: FONT_SIZES.CAPTION,
      color: COLORS.TEXT_SECONDARY,
      align: "center",
    });
  }

  addFooter(slide);
  return slide;
}

/**
 * 13. テキスト + 図レイアウト
 */
function addTextImageSlide(pres, heading, text, imagePosition = "right") {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "テキスト + 図");
  addHeaderLine(slide);

  const columnWidth = (LAYOUT.WIDTH - LAYOUT.MARGIN * 2 - LAYOUT.COLUMN_GAP) / 2;
  const textX = imagePosition === "right" ? LAYOUT.CONTENT_LEFT : LAYOUT.CONTENT_LEFT + columnWidth + LAYOUT.COLUMN_GAP;
  const imageX = imagePosition === "right" ? LAYOUT.CONTENT_LEFT + columnWidth + LAYOUT.COLUMN_GAP : LAYOUT.CONTENT_LEFT;

  // テキストエリア
  slide.addText(text || "説明テキストをここに記載します。", {
    x: textX,
    y: LAYOUT.CONTENT_TOP,
    w: columnWidth,
    h: 4.5,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.BODY,
    color: COLORS.TEXT_PRIMARY,
    valign: "top",
  });

  // 画像プレースホルダー
  slide.addShape("rect", {
    x: imageX,
    y: LAYOUT.CONTENT_TOP,
    w: columnWidth,
    h: 4.5,
    fill: { color: COLORS.BG_LIGHT },
    line: { color: COLORS.BORDER, width: 1, dashType: "dash" },
  });

  slide.addText("図・画像", {
    x: imageX,
    y: LAYOUT.CONTENT_TOP + 2.0,
    w: columnWidth,
    h: 0.5,
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.BODY,
    color: COLORS.TEXT_LIGHT,
    align: "center",
  });

  addFooter(slide);
  return slide;
}

/**
 * 14. アイコングリッドスライド
 */
function addIconGridSlide(pres, heading, items = []) {
  const slide = pres.addSlide();

  // ヘッダー
  addHeaderText(slide, heading || "機能一覧");
  addHeaderLine(slide);

  const defaultItems = [
    { title: "機能1", desc: "説明文" },
    { title: "機能2", desc: "説明文" },
    { title: "機能3", desc: "説明文" },
    { title: "機能4", desc: "説明文" },
    { title: "機能5", desc: "説明文" },
    { title: "機能6", desc: "説明文" },
  ];

  const gridItems = items.length > 0 ? items : defaultItems;
  const cols = 3;
  const rows = Math.ceil(gridItems.length / cols);
  const itemWidth = (LAYOUT.WIDTH - LAYOUT.MARGIN * 2 - LAYOUT.COLUMN_GAP * (cols - 1)) / cols;
  const itemHeight = (LAYOUT.HEIGHT - LAYOUT.CONTENT_TOP - 1.5) / rows - 0.2;

  gridItems.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const xPos = LAYOUT.CONTENT_LEFT + col * (itemWidth + LAYOUT.COLUMN_GAP);
    const yPos = LAYOUT.CONTENT_TOP + row * (itemHeight + 0.2);

    // アイコンプレースホルダー（円）
    slide.addShape("ellipse", {
      x: xPos + itemWidth / 2 - 0.35,
      y: yPos + 0.1,
      w: 0.7,
      h: 0.7,
      fill: { color: COLORS.BRAND_GOLD },
    });

    // タイトル
    slide.addText(item.title, {
      x: xPos,
      y: yPos + 0.9,
      w: itemWidth,
      h: 0.5,
      fontFace: FONTS.HEADING_JP,
      fontSize: FONT_SIZES.BODY,
      color: COLORS.TEXT_PRIMARY,
      bold: true,
      align: "center",
    });

    // 説明
    slide.addText(item.desc, {
      x: xPos,
      y: yPos + 1.4,
      w: itemWidth,
      h: 0.6,
      fontFace: FONTS.BODY_JP,
      fontSize: FONT_SIZES.CAPTION,
      color: COLORS.TEXT_SECONDARY,
      align: "center",
    });
  });

  addFooter(slide);
  return slide;
}

/**
 * 15. 強調セクションスライド（右側タイトル + 下線）
 */
function addHighlightSlide(pres, title, lineColor = COLORS.TEXT_PRIMARY) {
  const slide = pres.addSlide();

  // 右側にタイトル
  slide.addText(title || "タイトル", {
    x: 8.5,
    y: 3.0,
    w: 4,
    h: 0.6,
    align: "right",
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.HEADING2,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  });

  // 下線
  slide.addShape("line", {
    x: 5.5,
    y: 3.65,
    w: 7.0,
    h: 0,
    line: { color: lineColor, width: 2 },
  });

  addFooter(slide);
  return slide;
}

/**
 * 16. クロージングスライド
 */
function addClosingSlide(pres, thankYouText, contactInfo = "") {
  const slide = pres.addSlide();

  // Thank you テキスト
  slide.addText(thankYouText || "ご清聴ありがとうございました", {
    x: LAYOUT.MARGIN,
    y: 2.5,
    w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
    h: 1.0,
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.TITLE,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
    align: "center",
    underline: { type: "sng", color: COLORS.BRAND_GOLD },
  });

  // コンタクト情報
  if (contactInfo) {
    slide.addText(contactInfo, {
      x: LAYOUT.MARGIN,
      y: 4.5,
      w: LAYOUT.WIDTH - LAYOUT.MARGIN * 2,
      h: 0.8,
      fontFace: FONTS.BODY_JP,
      fontSize: FONT_SIZES.BODY,
      color: COLORS.TEXT_SECONDARY,
      align: "center",
    });
  }

  addFooter(slide);
  return slide;
}

// ============================================
// メイン処理
// ============================================

async function generateTemplate() {
  console.log("SBI Traceability PowerPoint テンプレートを生成中...\n");

  const pres = new pptxgen();

  // プレゼンテーション設定
  pres.layout = "LAYOUT_WIDE";  // 13.33" x 7.5"
  pres.author = "SBI Traceability";
  pres.title = "SBI Traceability Presentation Template";
  pres.subject = "PptxGenJS Template";

  // 全スライドを生成
  console.log("1. タイトルスライド");
  addTitleSlide(pres, "プレゼンテーションタイトル", "サブタイトルをここに", "発表者名", "2024年1月1日");

  console.log("2. セクションディバイダー（Vision）");
  addSectionSlide(pres, "大見出し", "Vision");

  console.log("3. 標準コンテンツスライド");
  addContentSlide(pres, "ヘッダー", [
    "文字サイズ",
  ]);

  console.log("4. 2カラムレイアウト");
  addTwoColumnSlide(pres, "2カラムレイアウト", "左側のコンテンツ\n\n詳細な説明をここに記載します。", "右側のコンテンツ\n\n対比する情報を記載します。");

  console.log("5. 3カラムレイアウト");
  addThreeColumnSlide(pres, "3カラムレイアウト", ["カラム1\nの内容", "カラム2\nの内容", "カラム3\nの内容"]);

  console.log("6. KPIカードスライド");
  addKPISlide(pres, "主要指標", [
    { value: "99.9%", label: "稼働率" },
    { value: "50%", label: "コスト削減" },
    { value: "2x", label: "処理速度" },
  ]);

  console.log("7. タイムラインスライド");
  addTimelineSlide(pres, "導入プロセス", [
    { title: "Phase 1", desc: "要件定義" },
    { title: "Phase 2", desc: "設計・開発" },
    { title: "Phase 3", desc: "テスト" },
    { title: "Phase 4", desc: "リリース" },
  ]);

  console.log("8. 比較スライド");
  addComparisonSlide(pres, "Before / After", "従来の方法:\n- 手動作業が多い\n- 時間がかかる\n- ミスが発生しやすい", "新しい方法:\n- 自動化された処理\n- 大幅な時間短縮\n- 高い精度");

  console.log("9. 引用スライド");
  addQuoteSlide(pres, "イノベーションとは、変化を機会として利用することである。", "ピーター・ドラッカー");

  console.log("10. アジェンダスライド");
  addAgendaSlide(pres, ["はじめに", "課題の整理", "解決策の提案", "導入効果", "まとめ"], 0);

  console.log("11. テーブルスライド");
  addTableSlide(pres, "比較表", [
    ["機能", "プランA", "プランB", "プランC"],
    ["ストレージ", "10GB", "50GB", "無制限"],
    ["ユーザー数", "5", "20", "無制限"],
    ["サポート", "メール", "チャット", "24/7電話"],
    ["価格", "¥1,000/月", "¥3,000/月", "¥10,000/月"],
  ]);

  console.log("12. 画像中央配置スライド");
  addImageCenterSlide(pres, "システム構成図", null, "図1: 全体アーキテクチャ");

  console.log("13. テキスト + 図スライド");
  addTextImageSlide(pres, "機能概要", "この機能により、以下のことが実現できます：\n\n• 作業効率の向上\n• コストの削減\n• 品質の改善", "right");

  console.log("14. アイコングリッドスライド");
  addIconGridSlide(pres, "サービス機能一覧", [
    { title: "セキュリティ", desc: "高度な保護機能" },
    { title: "スピード", desc: "高速な処理" },
    { title: "スケーラビリティ", desc: "柔軟な拡張" },
    { title: "サポート", desc: "24時間対応" },
    { title: "統合", desc: "外部連携" },
    { title: "分析", desc: "詳細レポート" },
  ]);

  console.log("15. 強調セクションスライド");
  addHighlightSlide(pres, "タイトル");

  console.log("16. クロージングスライド");
  addClosingSlide(pres, "ご清聴ありがとうございました", "お問い合わせ: contact@sbitraceability.co.jp");

  // ファイル出力
  const outputPath = path.join(__dirname, "sbita-template.pptx");
  await pres.writeFile({ fileName: outputPath });

  console.log(`\n✅ テンプレート生成完了: ${outputPath}`);
  console.log(`📊 合計 ${pres.slides.length} スライド`);
}

// 実行
generateTemplate().catch(console.error);

// エクスポート（他のスクリプトから利用可能）
module.exports = {
  addTitleSlide,
  addSectionSlide,
  addContentSlide,
  addTwoColumnSlide,
  addThreeColumnSlide,
  addKPISlide,
  addTimelineSlide,
  addComparisonSlide,
  addQuoteSlide,
  addAgendaSlide,
  addTableSlide,
  addImageCenterSlide,
  addTextImageSlide,
  addIconGridSlide,
  addHighlightSlide,
  addClosingSlide,
  addHeaderLine,
  addHeaderText,
  addFooter,
};
