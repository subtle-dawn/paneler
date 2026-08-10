export type ReadingDirection = "rtl" | "ltr";
export type EmotionSize = "small" | "medium" | "large" | null;
export type Size = "extraSmall" | "small" | "medium" | "large" | "extraLarge" | "fullPage";
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
  rowHeights?: Record<string, number[]>;
  rowWidths?: Record<string, number[][]>;
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
  stacks?: Array<{
    panels: LayoutPanel[];
    colStart: number;
    colSpan: number;
  }>;
  usedColumns: number;
};

export type PageLayout = {
  pageNumber: number;
  rows: LayoutRow[];
  warning?: string;
};
