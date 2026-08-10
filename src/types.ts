export type ReadingDirection = "rtl" | "ltr";
export type Size = "small" | "medium" | "large";
export type EmotionSize = Size | null;
export type Shape = "vertical" | "square" | "horizontal";

export type Panel = {
  id: string;
  emotionSize: EmotionSize;
  panelSize: Size;
  role: string;
  shape: Shape;
  camera: string;
  content: string;
  order: number;
};

export type Page = {
  id: string;
  pageNumber: number;
  panels: Panel[];
};

export type Project = {
  id: string;
  title: string;
  readingDirection: ReadingDirection;
  pages: Page[];
};

export type PanelRow = Panel & {
  pageNumber: number;
  warnings?: Partial<Record<"pageNumber" | "emotionSize" | "panelSize" | "shape", string>>;
};

export type LayoutPanel = Panel & {
  pageNumber: number;
  visualNumber: number;
  rowIndex: number;
  colStart: number;
  colSpan: number;
};

export type LayoutRow = {
  panels: LayoutPanel[];
  usedColumns: number;
};

export type PageLayout = {
  pageNumber: number;
  rows: LayoutRow[];
  warning?: string;
};
