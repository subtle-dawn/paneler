import { t } from "./i18n";
import type { LayoutPanel, LayoutRow, PageLayout, PanelRow, ReadingDirection, Size } from "./types";

const GRID_COLUMNS = 6;

const sizeToColumns: Record<Size, number> = {
  extraSmall: 2,
  small: 2,
  medium: 3,
  large: 4,
  extraLarge: 6,
  fullPage: 6,
};

export function rowsToLayouts(rows: PanelRow[], readingDirection: ReadingDirection): PageLayout[] {
  const pageNumbers = Array.from(new Set(rows.map((row) => row.pageNumber))).sort((a, b) => a - b);

  return pageNumbers.map((pageNumber) => {
    const pagePanels = rows
      .filter((row) => row.pageNumber === pageNumber)
      .sort((a, b) => a.order - b.order);
    const fullPagePanel = pagePanels.find((panel) => panel.panelSize === "fullPage");

    if (fullPagePanel) {
      return {
        pageNumber,
        rows: [
          positionRow(
            {
              panels: [
                {
                  ...fullPagePanel,
                  visualNumber: pagePanels.indexOf(fullPagePanel) + 1,
                  rowIndex: 0,
                  colStart: 1,
                  colSpan: GRID_COLUMNS,
                },
              ],
              usedColumns: GRID_COLUMNS,
            },
            readingDirection,
          ),
        ],
        warning: pagePanels.length > 1 ? t.tooManyPanels : undefined,
      };
    }

    const layoutRows: LayoutRow[] = [];
    let current: LayoutRow = { panels: [], usedColumns: 0 };

    for (let index = 0; index < pagePanels.length; index += 1) {
      const panel = pagePanels[index];
      const nextPanel = pagePanels[index + 1];
      const thirdPanel = pagePanels[index + 2];
      const fourthPanel = pagePanels[index + 3];
      const colSpan = sizeToColumns[panel.panelSize];
      const shouldStartRow =
        current.panels.length > 0 &&
        (panel.panelSize === "extraLarge" || current.usedColumns + colSpan > GRID_COLUMNS);

      if (shouldStartRow) {
        layoutRows.push(positionRow(expandHorizontalPanels(current), readingDirection));
        current = { panels: [], usedColumns: 0 };
      }

      if (
        current.panels.length === 0 &&
        isTallLeadStackPattern(panel, nextPanel, thirdPanel, fourthPanel)
      ) {
        const stackPanels =
          panel.shape === "vertical" && fourthPanel?.panelSize === "extraSmall"
            ? [nextPanel, thirdPanel, fourthPanel]
            : [nextPanel, thirdPanel];
        layoutRows.push(createTallLeadStackRow(panel, stackPanels, index, layoutRows.length, readingDirection, "lead-first"));
        index += stackPanels.length;
        continue;
      }

      if (
        current.panels.length === 0 &&
        isStackLeadPattern(panel, nextPanel, thirdPanel, fourthPanel)
      ) {
        const useThreeStack = fourthPanel?.panelSize === "large" && fourthPanel.shape === "vertical";
        const stackPanels = useThreeStack ? [panel, nextPanel, thirdPanel] : [panel, nextPanel];
        const leadPanel = useThreeStack ? fourthPanel : thirdPanel;
        layoutRows.push(createTallLeadStackRow(leadPanel, stackPanels, index, layoutRows.length, readingDirection, "stack-first"));
        index += stackPanels.length;
        continue;
      }

      if (current.panels.length === 0 && isExtraSmallPair(panel, nextPanel)) {
        layoutRows.push(createExtraSmallStackRow(panel, nextPanel, index, layoutRows.length, readingDirection));
        index += 1;
        continue;
      }

      current.panels.push({
        ...panel,
        visualNumber: index + 1,
        rowIndex: layoutRows.length,
        colStart: 1,
        colSpan,
      });
      current.usedColumns += colSpan;

      if (panel.panelSize === "extraLarge") {
        layoutRows.push(positionRow(current, readingDirection));
        current = { panels: [], usedColumns: 0 };
      }
    }

    if (current.panels.length > 0) {
      layoutRows.push(positionRow(expandHorizontalPanels(current), readingDirection));
    }

    return {
      pageNumber,
      rows: layoutRows,
      warning: layoutRows.length > 6 || pagePanels.length > 12 ? t.tooManyPanels : undefined,
    };
  });
}

function expandHorizontalPanels(row: LayoutRow): LayoutRow {
  const remainingColumns = GRID_COLUMNS - row.usedColumns;
  if (remainingColumns <= 0) return row;

  const horizontalPanels = row.panels.filter((panel) => panel.shape === "horizontal");
  if (horizontalPanels.length === 0) return row;

  const lastHorizontalPanel = horizontalPanels[horizontalPanels.length - 1];

  return {
    ...row,
    panels: row.panels.map((panel) =>
      panel.id === lastHorizontalPanel.id ? { ...panel, colSpan: panel.colSpan + remainingColumns } : panel,
    ),
    usedColumns: GRID_COLUMNS,
  };
}

function isTallLeadStackPattern(
  panel: PanelRow | undefined,
  nextPanel: PanelRow | undefined,
  thirdPanel: PanelRow | undefined,
  _fourthPanel: PanelRow | undefined,
) {
  return (
    panel?.panelSize === "large" &&
    Boolean(nextPanel) &&
    Boolean(thirdPanel) &&
    nextPanel?.panelSize === "extraSmall" &&
    thirdPanel?.panelSize === "extraSmall"
  );
}

function isExtraSmallPair(panel: PanelRow | undefined, nextPanel: PanelRow | undefined) {
  return panel?.panelSize === "extraSmall" && nextPanel?.panelSize === "extraSmall";
}

function isStackLeadPattern(
  panel: PanelRow | undefined,
  nextPanel: PanelRow | undefined,
  thirdPanel: PanelRow | undefined,
  fourthPanel: PanelRow | undefined,
) {
  return (
    panel?.panelSize === "extraSmall" &&
    nextPanel?.panelSize === "extraSmall" &&
    (thirdPanel?.panelSize === "large" || (thirdPanel?.panelSize === "extraSmall" && fourthPanel?.panelSize === "large"))
  );
}

function createTallLeadStackRow(
  leadPanel: PanelRow,
  stackPanels: PanelRow[],
  startIndex: number,
  rowIndex: number,
  readingDirection: ReadingDirection,
  order: "lead-first" | "stack-first",
): LayoutRow {
  const leadColSpan = 4;
  const stackColSpan = GRID_COLUMNS - leadColSpan;
  const shouldLeadStart = readingDirection === "ltr" ? order === "lead-first" : order === "stack-first";
  const leadColStart = shouldLeadStart ? 1 : stackColSpan + 1;
  const stackColStart = shouldLeadStart ? leadColSpan + 1 : 1;

  return {
    panels: [
      {
        ...leadPanel,
        visualNumber: startIndex + (order === "lead-first" ? 1 : stackPanels.length + 1),
        rowIndex,
        colStart: leadColStart,
        colSpan: leadColSpan,
      },
    ],
    stacks: [
      {
        colStart: stackColStart,
        colSpan: stackColSpan,
        panels: stackPanels.map((stackPanel, stackIndex) => ({
          ...stackPanel,
          visualNumber: startIndex + (order === "lead-first" ? stackIndex + 2 : stackIndex + 1),
          rowIndex,
          colStart: 1,
          colSpan: stackColSpan,
        })),
      },
    ],
    usedColumns: GRID_COLUMNS,
  };
}

function createExtraSmallStackRow(
  upperPanel: PanelRow,
  lowerPanel: PanelRow,
  startIndex: number,
  rowIndex: number,
  readingDirection: ReadingDirection,
): LayoutRow {
  const stackColSpan = 2;
  const stackColStart = readingDirection === "rtl" ? GRID_COLUMNS - stackColSpan + 1 : 1;

  return {
    panels: [],
    stacks: [
      {
        colStart: stackColStart,
        colSpan: stackColSpan,
        panels: [
          {
            ...upperPanel,
            visualNumber: startIndex + 1,
            rowIndex,
            colStart: 1,
            colSpan: stackColSpan,
          },
          {
            ...lowerPanel,
            visualNumber: startIndex + 2,
            rowIndex,
            colStart: 1,
            colSpan: stackColSpan,
          },
        ],
      },
    ],
    usedColumns: stackColSpan,
  };
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
