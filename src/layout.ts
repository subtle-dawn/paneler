import { t } from "./i18n";
import type { LayoutPanel, LayoutRow, PageLayout, PanelRow, ReadingDirection, Size } from "./types";

const GRID_COLUMNS = 6;

const sizeToColumns: Record<Size, number> = {
  small: 2,
  medium: 3,
  large: 6,
};

export function rowsToLayouts(rows: PanelRow[], readingDirection: ReadingDirection): PageLayout[] {
  const pageNumbers = Array.from(new Set(rows.map((row) => row.pageNumber))).sort((a, b) => a - b);

  return pageNumbers.map((pageNumber) => {
    const pagePanels = rows
      .filter((row) => row.pageNumber === pageNumber)
      .sort((a, b) => a.order - b.order);

    const layoutRows: LayoutRow[] = [];
    let current: LayoutRow = { panels: [], usedColumns: 0 };

    pagePanels.forEach((panel, index) => {
      const colSpan = sizeToColumns[panel.panelSize];
      const shouldStartRow =
        current.panels.length > 0 && (panel.panelSize === "large" || current.usedColumns + colSpan > GRID_COLUMNS);

      if (shouldStartRow) {
        layoutRows.push(positionRow(current, readingDirection));
        current = { panels: [], usedColumns: 0 };
      }

      current.panels.push({
        ...panel,
        visualNumber: index + 1,
        rowIndex: layoutRows.length,
        colStart: 1,
        colSpan,
      });
      current.usedColumns += colSpan;

      if (panel.panelSize === "large") {
        layoutRows.push(positionRow(current, readingDirection));
        current = { panels: [], usedColumns: 0 };
      }
    });

    if (current.panels.length > 0) {
      layoutRows.push(positionRow(current, readingDirection));
    }

    return {
      pageNumber,
      rows: layoutRows,
      warning: layoutRows.length > 6 || pagePanels.length > 12 ? t.tooManyPanels : undefined,
    };
  });
}

function positionRow(row: LayoutRow, readingDirection: ReadingDirection): LayoutRow {
  let cursor = readingDirection === "rtl" ? GRID_COLUMNS + 1 : 1;

  const panels = row.panels.map((panel): LayoutPanel => {
    if (readingDirection === "rtl") {
      cursor -= panel.colSpan;
      return { ...panel, colStart: cursor };
    }

    const colStart = cursor;
    cursor += panel.colSpan;
    return { ...panel, colStart };
  });

  return { ...row, panels };
}
