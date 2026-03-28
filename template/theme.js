/**
 * SBI Traceability PPTX Theme Configuration
 * ブランドカラー、フォント、レイアウト定数を定義
 */

// ============================================
// カラーパレット
// ============================================
const COLORS = {
  // ブランドカラー
  BRAND_GOLD: "CBC269",        // メインアクセント（ゴールド/黄土色）- 下線、ライン用
  BRAND_RED: "EA4545",         // 強調用レッド
  BRAND_SALMON: "ED6A5B",      // Strictly Confidential用サーモンレッド

  // テキストカラー
  TEXT_PRIMARY: "454545",      // メインテキスト（ダークグレー）
  TEXT_SECONDARY: "535353",    // サブテキスト（グレー）
  TEXT_LIGHT: "7F7F7F",        // 薄いテキスト（50%グレー）

  // 背景色
  BG_WHITE: "FFFFFF",          // 白背景
  BG_LIGHT: "F5F5F5",          // ライトグレー背景

  // ボーダー・区切り
  BORDER: "E0E0E0",            // ボーダー
  DIVIDER: "BDBDBD",           // 区切り線

  // Office テーマカラー（元テンプレートから）
  ACCENT1: "5B9BD5",           // ブルー
  ACCENT2: "ED7D31",           // オレンジ
  ACCENT3: "A5A5A5",           // グレー
  ACCENT4: "FFC000",           // イエロー
  ACCENT5: "4472C4",           // ダークブルー
  ACCENT6: "70AD47",           // グリーン
};

// ============================================
// フォント設定
// ============================================
const FONTS = {
  HEADING_EN: "Century Gothic",    // 英語見出し用
  HEADING_JP: "メイリオ",           // 日本語見出し用
  BODY_EN: "Century Gothic",       // 英語本文用
  BODY_JP: "メイリオ",              // 日本語本文用
  UI: "Meiryo UI",                 // UI用（フッターなど）
  CODE: "Consolas",                // コード用
};

// ============================================
// フォントサイズ (pt)
// ============================================
const FONT_SIZES = {
  TITLE: 32,          // タイトル（下線付き）
  SUBTITLE: 24,       // サブタイトル/タイトル（下線なし）
  HEADING1: 24,       // 大見出し
  HEADING2: 20,       // 見出し
  BODY: 16,           // 本文（箇条書き）
  BODY_SMALL: 14,     // 小さめの本文
  CAPTION: 10.5,      // キャプション/小文字
  FOOTER: 8,          // フッター
  HEADER: 14,         // ヘッダー
  STAT: 48,           // 統計数字
};

// ============================================
// スライドレイアウト (インチ)
// ============================================
const LAYOUT = {
  // スライドサイズ (16:9)
  WIDTH: 13.333,      // 12192000 EMU
  HEIGHT: 7.5,        // 6858000 EMU

  // マージン
  MARGIN: 0.8,
  MARGIN_NARROW: 0.4,

  // ヘッダー・フッター
  HEADER_Y: 0.1,
  HEADER_LINE_Y: 0.52,    // ヘッダーライン位置
  FOOTER_Y: 7.15,

  // コンテンツエリア
  CONTENT_TOP: 1.0,
  CONTENT_LEFT: 0.8,

  // カラムレイアウト
  COLUMN_GAP: 0.3,

  // ヘッダーライン幅
  HEADER_LINE_WIDTH: 7.1,  // 約6480000 EMU
};

// ============================================
// 下線スタイル（ゴールド）
// ============================================
const UNDERLINE_STYLE = {
  underline: {
    type: "sng",
    color: COLORS.BRAND_GOLD,
  },
};

// ============================================
// シャドウ効果
// ============================================
function makeShadow(blur = 4, offset = 2, opacity = 0.2) {
  return {
    type: "outer",
    blur: blur,
    offset: offset,
    angle: 45,
    color: "000000",
    opacity: opacity,
  };
}

function makeCardShadow() {
  return makeShadow(6, 3, 0.15);
}

// ============================================
// 共通スタイル
// ============================================
const STYLES = {
  // タイトルスライド（下線付き）
  titleWithUnderline: {
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.TITLE,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
    underline: { type: "sng", color: COLORS.BRAND_GOLD },
  },

  // 通常タイトル（下線なし）
  title: {
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.SUBTITLE,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  },

  // 備考タイトル
  noteTitle: {
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.BODY_SMALL,
    color: COLORS.TEXT_LIGHT,
    bold: true,
  },

  // 赤字下線（強調）
  emphasis: {
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.BODY_SMALL,
    color: COLORS.BRAND_RED,
    bold: true,
    underline: { type: "sng", color: COLORS.BRAND_GOLD },
  },

  // ヘッダーテキスト
  header: {
    fontFace: FONTS.HEADING_JP,
    fontSize: FONT_SIZES.HEADER,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  },

  // 本文
  body: {
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.BODY,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  },

  // 小文字
  caption: {
    fontFace: FONTS.BODY_JP,
    fontSize: FONT_SIZES.CAPTION,
    color: COLORS.TEXT_PRIMARY,
    bold: true,
  },

  // フッター
  footer: {
    fontFace: FONTS.UI,
    fontSize: FONT_SIZES.FOOTER,
    color: COLORS.TEXT_SECONDARY,
  },

  // Strictly Confidential
  confidential: {
    fontFace: FONTS.HEADING_EN,
    fontSize: FONT_SIZES.FOOTER,
    color: COLORS.BRAND_SALMON,
  },
};

// ============================================
// フッター情報
// ============================================
const FOOTER = {
  COPYRIGHT: "Copyright © 2026 SBI Traceability co., Ltd. All Rights Reserved",
  CONFIDENTIAL: "Strictly Confidential",
  LOGO_PATH: null,  // ロゴ画像パス（オプション）
};

// ============================================
// エクスポート
// ============================================
module.exports = {
  COLORS,
  FONTS,
  FONT_SIZES,
  LAYOUT,
  STYLES,
  FOOTER,
  makeShadow,
  makeCardShadow,
};
