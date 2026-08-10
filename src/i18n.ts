export const t = {
  appName: "Paneler",
  title: "作品タイトル",
  readingDirection: "読み方向",
  pages: "ページ一覧",
  addRow: "行追加",
  addPage: "ページ追加",
  addPageAfter: "次ページを挿入",
  duplicateRow: "複製",
  deleteRow: "削除",
  moveUp: "上へ",
  moveDown: "下へ",
  page: "ページ",
  emotion: "感情",
  size: "大きさ",
  role: "役割",
  shape: "形",
  camera: "カメラ",
  content: "内容",
  small: "小",
  medium: "中",
  large: "大",
  none: "なし",
  vertical: "縦",
  square: "正",
  horizontal: "横",
  importJson: "JSON読込",
  exportJson: "JSON保存",
  importSheet: "CSV/Excel読込",
  exportPng: "PNG保存",
  reset: "新規作品",
  saved: "自動保存済み",
  tooManyPanels: "このページはコマ数が多すぎる可能性があります。",
};

export const sizeLabel = {
  small: t.small,
  medium: t.medium,
  large: t.large,
} as const;

export const emotionLabel = {
  small: t.small,
  medium: t.medium,
  large: t.large,
  none: t.none,
} as const;

export const shapeLabel = {
  vertical: t.vertical,
  square: t.square,
  horizontal: t.horizontal,
} as const;
