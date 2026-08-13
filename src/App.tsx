import { ChangeEvent, type CSSProperties, useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import {
  ArrowDown,
  ArrowLeft,
  ArrowRight,
  ArrowUp,
  ChevronDown,
  Copy,
  FileSpreadsheet,
  ImageDown,
  Plus,
  RotateCcw,
  Settings2,
  Trash2,
  X,
} from "lucide-react";
import { cameraOptionLabels, emotionLabel, faceSizeLabel, roleOptionLabels, shapeLabel, sizeLabel, t } from "./i18n";
import {
  downloadStoryboardXlsx,
  exportPagePng,
  exportPagesPngZip,
  exportPagesPdf,
  readSheetProject,
  readSheetRows,
} from "./io";
import { rowsToLayouts } from "./layout";
import {
  createDefaultProject,
  createEmptyProject,
  createPanelRow,
  loadProject,
  normalizeProject,
  projectToRows,
  rowsToProject,
  saveProject,
} from "./storage";
import type { BleedSide, EmotionSize, FaceSize, LayoutPanel, LayoutRow, PanelRow, Project, Shape, Size } from "./types";

const sizeOptions: Size[] = ["extraSmall", "small", "medium", "large", "extraLarge", "fullPage"];
const shapeOptions: Shape[] = ["square", "vertical", "horizontal"];
const emotionOptions = ["small", "medium", "large"] as const;
const faceSizeOptions: FaceSize[] = ["none", "extraSmall", "small", "medium", "large", "extraLarge"];
const roleOptions = roleOptionLabels;
const cameraOptions = cameraOptionLabels;
const minRowHeightWeight = 0.25;
const minColumnWidthWeight = 0.25;
const horizontalRowHeightWeight = 0.6;
const gridColumnCount = 6;
const pageBorderBleedAmount = 40;
const miniPageBorderBleedAmount = 4;
type DownloadFormat = "png" | "pdf";
type DownloadScope = "active" | "all";
type ExcelTone = "green" | "yellow" | "red";

const sizeTone: Partial<Record<Size, ExcelTone>> = {
  extraSmall: "green",
  small: "green",
  medium: "yellow",
  large: "red",
  extraLarge: "red",
  fullPage: "red",
};

const emotionTone: Partial<Record<NonNullable<EmotionSize>, ExcelTone>> = {
  small: "green",
  medium: "yellow",
  large: "red",
};

const shapeTone: Record<Shape, ExcelTone> = {
  vertical: "yellow",
  square: "green",
  horizontal: "yellow",
};

const faceSizeTone: Partial<Record<FaceSize, ExcelTone>> = {
  extraSmall: "green",
  small: "green",
  medium: "yellow",
  large: "red",
  extraLarge: "red",
};

const cameraToneByLabel: Partial<Record<(typeof cameraOptionLabels)[number], ExcelTone>> = {
  正: "green",
  横: "yellow",
  上: "yellow",
  下: "yellow",
  俯: "red",
  煽: "red",
};

export function App() {
  const [project, setProject] = useState<Project>(() => loadProject() ?? createDefaultProject());
  const [rows, setRows] = useState<PanelRow[]>(() =>
    projectToRows(loadProject() ?? createDefaultProject()).map(normalizeChoiceDefaults),
  );
  const [activePage, setActivePage] = useState(1);
  const [status, setStatus] = useState(t.saved);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isNewProjectConfirmOpen, setIsNewProjectConfirmOpen] = useState(false);
  const [isAppHelpOpen, setIsAppHelpOpen] = useState(false);
  const [isDownloadOpen, setIsDownloadOpen] = useState(false);
  const [downloadFormat, setDownloadFormat] = useState<DownloadFormat>("png");
  const [downloadScope, setDownloadScope] = useState<DownloadScope>("active");
  const [isDownloading, setIsDownloading] = useState(false);
  const [isPageHelpOpen, setIsPageHelpOpen] = useState(false);
  const [isStoryboardHelpOpen, setIsStoryboardHelpOpen] = useState(false);
  const [isPreviewHelpOpen, setIsPreviewHelpOpen] = useState(false);
  const [isPageListCollapsed, setIsPageListCollapsed] = useState(false);
  const [isStoryboardCollapsed, setIsStoryboardCollapsed] = useState(false);
  const [isProjectNoteCollapsed, setIsProjectNoteCollapsed] = useState(false);
  const [isPreviewCollapsed, setIsPreviewCollapsed] = useState(false);
  const [draggingPage, setDraggingPage] = useState<number | null>(null);
  const [dragOverPage, setDragOverPage] = useState<number | null>(null);
  const [draggingRowId, setDraggingRowId] = useState<string | null>(null);
  const [dragOverRowId, setDragOverRowId] = useState<string | null>(null);
  const [selectedRowId, setSelectedRowId] = useState<string | null>(null);
  const sheetInputRef = useRef<HTMLInputElement>(null);
  const pageRefs = useRef<Record<number, HTMLDivElement | null>>({});

  const syncedProject = useMemo(() => rowsToProject(project, rows), [project, rows]);
  const layouts = useMemo(() => {
    const generatedLayouts = rowsToLayouts(projectToRows(syncedProject), syncedProject.readingDirection);
    return syncedProject.pages
      .map(
        (page) =>
          generatedLayouts.find((layout) => layout.pageNumber === page.pageNumber) ?? {
            pageNumber: page.pageNumber,
            rows: [],
          },
      )
      .sort((a, b) => a.pageNumber - b.pageNumber);
  }, [syncedProject]);
  const activeLayout = layouts.find((layout) => layout.pageNumber === activePage) ?? layouts[0];
  const activeRows = useMemo(
    () => rows.filter((row) => row.pageNumber === activePage).sort((a, b) => a.order - b.order),
    [activePage, rows],
  );
  const pageSpreads = useMemo(() => {
    const spreads: typeof layouts[] = [];
    if (layouts[0]) spreads.push([layouts[0]]);
    for (let index = 1; index < layouts.length; index += 2) {
      spreads.push(layouts.slice(index, index + 2));
    }
    return spreads;
  }, [layouts]);
  const pageCountWarning = layouts.length % 4 === 0 ? undefined : t.pageCountWarning;
  const columnWarnings = useMemo(() => getColumnWarnings(activeRows), [activeRows]);

  useEffect(() => {
    saveProject(syncedProject);
    setStatus(t.saved);
  }, [syncedProject]);

  useEffect(() => {
    if (!activeLayout && layouts[0]) setActivePage(layouts[0].pageNumber);
  }, [activeLayout, layouts]);

  useEffect(() => {
    if (!isSettingsOpen) return;

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === "Escape" && !isNewProjectConfirmOpen) setIsSettingsOpen(false);
    }

    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [isNewProjectConfirmOpen, isSettingsOpen]);

  useEffect(() => {
    if (!isNewProjectConfirmOpen) return;

    function closeOnEscape(event: KeyboardEvent) {
      if (event.key === "Escape") setIsNewProjectConfirmOpen(false);
    }

    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [isNewProjectConfirmOpen]);

  function updateProject(next: Partial<Project>) {
    setStatus(t.edited);
    setProject((current) => ({ ...current, ...next }));
  }

  function createNewProject() {
    const nextProject = createEmptyProject();
    setStatus(t.edited);
    setProject(nextProject);
    setRows([]);
    setActivePage(1);
    setSelectedRowId(null);
    setDraggingPage(null);
    setDragOverPage(null);
    setDraggingRowId(null);
    setDragOverRowId(null);
    setIsNewProjectConfirmOpen(false);
    setIsSettingsOpen(false);
  }

  function updateRowHeights(pageNumber: number, nextHeights: number[]) {
    updateProject({
      rowHeights: {
        ...(project.rowHeights ?? {}),
        [pageNumber]: nextHeights,
      },
    });
  }

  function updateRowWidths(pageNumber: number, rowIndex: number, nextWidths: number[]) {
    const pageRowWidths = project.rowWidths?.[pageNumber] ?? [];
    const nextPageRowWidths = [...pageRowWidths];
    nextPageRowWidths[rowIndex] = nextWidths;

    updateProject({
      rowWidths: {
        ...(project.rowWidths ?? {}),
        [pageNumber]: nextPageRowWidths,
      },
    });
  }

  function clearActivePageLayoutAdjustments() {
    setStatus(t.edited);
    setProject((current) => {
      const nextRowHeights = { ...(current.rowHeights ?? {}) };
      const nextRowWidths = { ...(current.rowWidths ?? {}) };
      delete nextRowHeights[activePage];
      delete nextRowWidths[activePage];

      return {
        ...current,
        rowHeights: nextRowHeights,
        rowWidths: nextRowWidths,
      };
    });
    setRows((current) =>
      current.map((row) => (row.pageNumber === activePage ? { ...row, bleed: undefined } : row)).map(normalizeOrder),
    );
  }

  function updateRow(id: string, patch: Partial<PanelRow>) {
    setStatus(t.edited);
    setRows((current) => current.map((row) => (row.id === id ? { ...row, ...patch } : row)).map(normalizeOrder));
  }

  function selectRow(id: string, pageNumber: number) {
    const shouldSelect = selectedRowId !== id;
    setSelectedRowId(shouldSelect ? id : null);
    setActivePage(pageNumber);
  }

  function togglePanelFrame(id: string) {
    setStatus(t.edited);
    setRows((current) =>
      current.map((row) => (row.id === id ? { ...row, isFrameHidden: !row.isFrameHidden } : row)).map(normalizeOrder),
    );
  }

  function togglePanelBleed(id: string, side: BleedSide, blockedSide: "left" | "right") {
    if (side === blockedSide) return;

    setStatus(t.edited);
    setRows((current) =>
      current
        .map((row) => {
          if (row.id !== id) return row;

          const bleed = { ...(row.bleed ?? {}) };
          bleed[side] = !isPanelBleedSideActive(row, side);
          return { ...row, bleed };
        })
        .map(normalizeOrder),
    );
  }

  function duplicateRow(id: string) {
    setStatus(t.edited);
    setRows((current) => {
      const index = current.findIndex((row) => row.id === id);
      if (index < 0) return current;
      const next = [...current];
      next.splice(index + 1, 0, { ...current[index], id: `panel-${crypto.randomUUID()}`, order: index + 1 });
      return next.map(normalizeOrder);
    });
  }

  function deleteRow(id: string) {
    setStatus(t.edited);
    setRows((current) => {
      if (current.length === 1) return current;
      return current.filter((row) => row.id !== id).map(normalizeOrder);
    });
  }

  function appendRowToActivePage() {
    setStatus(t.edited);
    setRows((current) => {
      const lastIndex = current.reduce((foundIndex, row, index) => (row.pageNumber === activePage ? index : foundIndex), -1);
      const insertIndex = lastIndex >= 0 ? lastIndex + 1 : current.length;
      const next = [...current];
      next.splice(insertIndex, 0, createPanelRow(activePage, insertIndex));
      return next.map(normalizeOrder);
    });
  }

  function deletePage(pageNumber: number) {
    setStatus(t.edited);
    setRows((current) => {
      const remaining = current.filter((row) => row.pageNumber !== pageNumber);
      if (remaining.length === 0) {
        setActivePage(1);
        return [createPanelRow(1, 0)];
      }

      const renumbered = remaining.map((row) =>
        row.pageNumber > pageNumber ? { ...row, pageNumber: row.pageNumber - 1 } : row,
      );
      const nextActivePage = Math.min(pageNumber, Math.max(...renumbered.map((row) => row.pageNumber)));
      setActivePage(nextActivePage);
      return renumbered.map(normalizeOrder);
    });
  }

  function appendPage() {
    setStatus(t.edited);
    setRows((current) => {
      const maxPage = Math.max(...current.map((row) => row.pageNumber), 0);
      const nextPage = maxPage + 1;
      setActivePage(nextPage);
      return [...current, createPanelRow(nextPage, current.length)].map(normalizeOrder);
    });
  }

  function reorderPages(fromPage: number, toPage: number) {
    if (fromPage === toPage) return;

    setStatus(t.edited);
    setRows((current) => {
      const pageNumbers = Array.from(new Set(current.map((row) => row.pageNumber))).sort((a, b) => a - b);
      const fromIndex = pageNumbers.indexOf(fromPage);
      const toIndex = pageNumbers.indexOf(toPage);
      if (fromIndex < 0 || toIndex < 0) return current;

      const ordered = [...pageNumbers];
      const [moved] = ordered.splice(fromIndex, 1);
      ordered.splice(toIndex, 0, moved);
      const pageMap = new Map(ordered.map((oldPageNumber, index) => [oldPageNumber, index + 1]));
      const movedPageNumber = pageMap.get(fromPage) ?? toPage;

      setActivePage(movedPageNumber);
      return current.map((row) => ({ ...row, pageNumber: pageMap.get(row.pageNumber) ?? row.pageNumber })).map(normalizeOrder);
    });
  }

  function reorderRows(fromId: string, toId: string) {
    if (fromId === toId) return;

    setStatus(t.edited);
    setRows((current) => {
      const pageRows = current
        .filter((row) => row.pageNumber === activePage)
        .sort((a, b) => a.order - b.order);
      const fromIndex = pageRows.findIndex((row) => row.id === fromId);
      const toIndex = pageRows.findIndex((row) => row.id === toId);
      if (fromIndex < 0 || toIndex < 0) return current;

      const reorderedPageRows = [...pageRows];
      const [moved] = reorderedPageRows.splice(fromIndex, 1);
      reorderedPageRows.splice(toIndex, 0, moved);

      const reorderedIds = new Map(reorderedPageRows.map((row, index) => [row.id, index]));
      return current
        .map((row) =>
          row.pageNumber === activePage ? { ...row, order: reorderedIds.get(row.id) ?? row.order } : row,
        )
        .sort((a, b) => a.pageNumber - b.pageNumber || a.order - b.order)
        .map(normalizeOrder);
    });
  }

  async function importSheet(file?: File) {
    if (!file) return;
    const importedProject = await readSheetProject(file);
    if (importedProject) {
      const importedRows = await readSheetRows(file);
      const mergedRows = mergeRowsWithProjectRows(importedRows, importedProject);
      const mergedProject = normalizeProject(rowsToProject(importedProject, mergedRows));
      setProject(mergedProject);
      setRows(mergedRows.map(normalizeChoiceDefaults).map(normalizeOrder));
      setActivePage(mergedRows[0]?.pageNumber ?? mergedProject.pages[0]?.pageNumber ?? 1);
      setStatus(t.sheetLoaded);
      return;
    }

    const importedRows = await readSheetRows(file);
    const importedTitle = titleFromExcelFileName(file.name);
    if (importedTitle) {
      setProject((current) => ({ ...current, title: importedTitle }));
    }
    setRows(importedRows.map(normalizeChoiceDefaults).map(normalizeOrder));
    setActivePage(importedRows[0]?.pageNumber ?? 1);
    setStatus(t.sheetLoaded);
  }

  function formatPageNumber(pageNumber: number) {
    return String(pageNumber).padStart(3, "0");
  }

  async function downloadSelectedPages() {
    if (isDownloading) return;

    setIsDownloading(true);
    try {
      await new Promise<void>((resolve) => window.requestAnimationFrame(() => resolve()));

      if (downloadScope === "active") {
        const pageEl = pageRefs.current[activePage];
        if (!pageEl) return;

        if (downloadFormat === "png") {
          await exportPagePng(pageEl, `${project.title}_${t.panelingFileStem}_${formatPageNumber(activePage)}.png`);
        } else {
          await exportPagesPdf([pageEl], `${project.title}_${t.panelingFileStem}_${formatPageNumber(activePage)}.pdf`);
        }
        setIsDownloadOpen(false);
        return;
      }

      if (downloadFormat === "png") {
        const pages = layouts
          .map((layout) => {
            const pageEl = pageRefs.current[layout.pageNumber];
            return pageEl
              ? { pageEl, fileName: `${project.title}_${t.panelingFileStem}_${formatPageNumber(layout.pageNumber)}.png` }
              : undefined;
          })
          .filter((page): page is { pageEl: HTMLDivElement; fileName: string } => Boolean(page));
        await exportPagesPngZip(pages, `${project.title}_${t.panelingFileStem}.zip`);
      } else {
        const pageEls = layouts
          .map((layout) => pageRefs.current[layout.pageNumber])
          .filter((pageEl): pageEl is HTMLDivElement => Boolean(pageEl));
        await exportPagesPdf(pageEls, `${project.title}_${t.panelingFileStem}.pdf`);
      }

      setIsDownloadOpen(false);
    } catch (error) {
      console.error(error);
      window.alert(t.downloadFailed);
    } finally {
      setIsDownloading(false);
    }
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-area">
          <h1>{t.appName}</h1>
        </div>
        <div className="topbar-actions">
          <button
            type="button"
            className="header-button"
            aria-label={t.appHelpAria}
            onClick={() => setIsAppHelpOpen(true)}
          >
            {t.appHelpTitle}
          </button>
          <button type="button" className="header-button" onClick={() => setIsSettingsOpen(true)}>
            {t.settings}
          </button>
          <button type="button" className="header-button" onClick={() => sheetInputRef.current?.click()}>
            {t.import}
          </button>
          <button type="button" className="header-button" onClick={() => downloadStoryboardXlsx(rows, project.title, syncedProject)}>
            {t.export}
          </button>
          <button type="button" className="header-button" onClick={() => setIsDownloadOpen(true)}>
            {t.downloads}
          </button>
        </div>
        <input
          ref={sheetInputRef}
          type="file"
          accept=".xlsx,.xls"
          hidden
          onChange={(event) => {
            importSheet(event.target.files?.[0]);
            event.currentTarget.value = "";
          }}
        />
      </header>

      {isAppHelpOpen && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setIsAppHelpOpen(false)}>
          <section
            className="settings-modal page-help-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="app-help-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <h2 id="app-help-title">{t.appHelpTitle}</h2>
              <button type="button" className="icon-button" onClick={() => setIsAppHelpOpen(false)} title={t.close}>
                <X size={18} />
              </button>
            </div>
            <p className="page-help-text">
              {t.appHelpLine1}
              <br />
              {t.appHelpLine2}
              <br />
              {t.appHelpLine3}
              <br />
              {t.appHelpLine4}
            </p>
          </section>
        </div>
      )}

      {isDownloadOpen && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setIsDownloadOpen(false)}>
          <section
            className="settings-modal download-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="download-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <h2 id="download-title">{t.downloads}</h2>
              <button type="button" className="icon-button" onClick={() => setIsDownloadOpen(false)} title={t.close}>
                <X size={18} />
              </button>
            </div>
            <fieldset className="download-options">
              <legend>{t.downloadFormat}</legend>
              <label>
                <input
                  type="radio"
                  name="download-format"
                  checked={downloadFormat === "png"}
                  onChange={() => setDownloadFormat("png")}
                />
                <span>PNG</span>
              </label>
              <label>
                <input
                  type="radio"
                  name="download-format"
                  checked={downloadFormat === "pdf"}
                  onChange={() => setDownloadFormat("pdf")}
                />
                <span>PDF</span>
              </label>
            </fieldset>
            <fieldset className="download-options">
              <legend>{t.downloadScope}</legend>
              <label>
                <input
                  type="radio"
                  name="download-scope"
                  checked={downloadScope === "active"}
                  onChange={() => setDownloadScope("active")}
                />
                <span>{t.downloadActivePageOnly}</span>
              </label>
              <label>
                <input
                  type="radio"
                  name="download-scope"
                  checked={downloadScope === "all"}
                  onChange={() => setDownloadScope("all")}
                />
                <span>{t.downloadAllPages}</span>
              </label>
            </fieldset>
            <button
              type="button"
              className="button primary download-submit"
              onClick={downloadSelectedPages}
              disabled={isDownloading}
            >
              {isDownloading && <span className="download-spinner" aria-hidden="true" />}
              {isDownloading ? t.downloading : t.download}
            </button>
          </section>
        </div>
      )}

      {isSettingsOpen && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setIsSettingsOpen(false)}>
          <section
            className="settings-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="settings-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <h2 id="settings-title">{t.settings}</h2>
              <button type="button" className="icon-button" onClick={() => setIsSettingsOpen(false)} title={t.close}>
                <X size={18} />
              </button>
            </div>
            <label className="settings-field">
              <span>{t.title}</span>
              <input value={project.title} onChange={(event) => updateProject({ title: event.target.value })} />
            </label>
            <div className="settings-field">
              <span>{t.readingDirection}</span>
              <div className="segmented-control" aria-label={t.readingDirection}>
                <button
                  type="button"
                  className={project.readingDirection === "rtl" ? "active" : ""}
                  onClick={() => updateProject({ readingDirection: "rtl" })}
                >
                  {t.readingRightToLeft}
                </button>
                <button
                  type="button"
                  className={project.readingDirection === "ltr" ? "active" : ""}
                  onClick={() => updateProject({ readingDirection: "ltr" })}
                >
                  {t.readingLeftToRight}
                </button>
              </div>
            </div>
            <div className="settings-actions">
              <button
                type="button"
                className="button danger"
                onClick={() => setIsNewProjectConfirmOpen(true)}
              >
                {t.createNewProject}
              </button>
            </div>
          </section>
        </div>
      )}

      {isNewProjectConfirmOpen && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setIsNewProjectConfirmOpen(false)}>
          <section
            className="settings-modal confirm-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="new-project-confirm-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <h2 id="new-project-confirm-title">{t.createNewProject}</h2>
              <button
                type="button"
                className="icon-button"
                onClick={() => setIsNewProjectConfirmOpen(false)}
                title={t.close}
              >
                <X size={18} />
              </button>
            </div>
            <p className="confirm-text">
              {t.newProjectConfirmLine1}
              <br />
              {t.newProjectConfirmLine2}
              <br />
              {t.newProjectConfirmLine3}
            </p>
            <div className="confirm-actions">
              <button type="button" className="button" onClick={() => setIsNewProjectConfirmOpen(false)}>
                {t.no}
              </button>
              <button type="button" className="button primary" onClick={createNewProject}>
                {t.yes}
              </button>
            </div>
          </section>
        </div>
      )}

      {isPageHelpOpen && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setIsPageHelpOpen(false)}>
          <section
            className="settings-modal page-help-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="page-help-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <h2 id="page-help-title">{t.pages}</h2>
              <button type="button" className="icon-button" onClick={() => setIsPageHelpOpen(false)} title={t.close}>
                <X size={18} />
              </button>
            </div>
            <p className="page-help-text">
              {t.pageHelpLine1}
              <br />
              {t.pageHelpLine2}
              <br />
              {t.pageHelpLine3}
            </p>
          </section>
        </div>
      )}

      {isStoryboardHelpOpen && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setIsStoryboardHelpOpen(false)}>
          <section
            className="settings-modal page-help-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="storyboard-help-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <h2 id="storyboard-help-title">{t.storyboard}</h2>
              <button
                type="button"
                className="icon-button"
                onClick={() => setIsStoryboardHelpOpen(false)}
                title={t.close}
              >
                <X size={18} />
              </button>
            </div>
            <p className="page-help-text">
              {t.storyboardHelpLine1}
              <br />
              {t.storyboardHelpLine2}
              <br />
              {t.storyboardHelpLine3}
              <br />
              {t.storyboardHelpLine4}
              <br />
              {t.storyboardHelpLine5}
            </p>
          </section>
        </div>
      )}

      {isPreviewHelpOpen && (
        <div className="modal-backdrop" role="presentation" onMouseDown={() => setIsPreviewHelpOpen(false)}>
          <section
            className="settings-modal page-help-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="preview-help-title"
            onMouseDown={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <h2 id="preview-help-title">{t.preview}</h2>
              <button type="button" className="icon-button" onClick={() => setIsPreviewHelpOpen(false)} title={t.close}>
                <X size={18} />
              </button>
            </div>
            <p className="page-help-text">
              {t.previewHelpLine1}
              <br />
              {t.previewHelpLine2}
              <br />
              {t.previewHelpLine3}
              <br />
              {t.previewHelpLine4}
              <br />
              {t.previewHelpLine5}
            </p>
          </section>
        </div>
      )}

      <section className={`page-list-bar reading-${syncedProject.readingDirection}`} aria-label={t.pages}>
        <div className="page-list-header collapsible-header" onClick={() => setIsPageListCollapsed((current) => !current)}>
          <div className="page-list-title">
            <button
              type="button"
              className="collapse-toggle"
              aria-expanded={!isPageListCollapsed}
              onClick={(event) => {
                event.stopPropagation();
                setIsPageListCollapsed((current) => !current);
              }}
            >
              <span className="collapse-triangle" aria-hidden="true" />
              <h2>{t.pages}</h2>
            </button>
            <button
              type="button"
              className="help-button"
              aria-label={t.pageHelpAria}
              onClick={(event) => {
                event.stopPropagation();
                setIsPageHelpOpen(true);
              }}
            >
              ?
            </button>
            {pageCountWarning && <WarningBadge warning={pageCountWarning} />}
          </div>
        </div>
        {!isPageListCollapsed && <div className={`page-strip thumbnail-strip reading-${syncedProject.readingDirection}`}>
          {pageSpreads.map((spread) => (
            <div
              className={`page-spread reading-${syncedProject.readingDirection}`}
              key={spread.map((layout) => layout.pageNumber).join("-")}
            >
              {spread.map((layout) => (
                <button
                  key={layout.pageNumber}
                  type="button"
                  draggable
                  className={`page-thumb ${layout.pageNumber === activePage ? "active" : ""} ${
                    layout.pageNumber === draggingPage ? "dragging" : ""
                  } ${layout.pageNumber === dragOverPage ? "drag-over" : ""}`}
                  onClick={() => setActivePage(layout.pageNumber)}
                  onDragStart={(event) => {
                    event.dataTransfer.effectAllowed = "move";
                    event.dataTransfer.setData("text/plain", String(layout.pageNumber));
                    setDraggingPage(layout.pageNumber);
                  }}
                  onDragOver={(event) => {
                    event.preventDefault();
                    event.dataTransfer.dropEffect = "move";
                    setDragOverPage(layout.pageNumber);
                  }}
                  onDragLeave={() => {
                    setDragOverPage((current) => (current === layout.pageNumber ? null : current));
                  }}
                  onDrop={(event) => {
                    event.preventDefault();
                    const source = Number(event.dataTransfer.getData("text/plain")) || draggingPage;
                    if (source) reorderPages(source, layout.pageNumber);
                    setDraggingPage(null);
                    setDragOverPage(null);
                  }}
                  onDragEnd={() => {
                    setDraggingPage(null);
                    setDragOverPage(null);
                  }}
                >
                  <span className="thumb-heading">
                    {layout.pageNumber}
                    <span
                      role="button"
                      tabIndex={0}
                      className="thumb-delete"
                      title={t.deletePage}
                      onClick={(event) => {
                        event.stopPropagation();
                        deletePage(layout.pageNumber);
                      }}
                      onKeyDown={(event) => {
                        if (event.key !== "Enter" && event.key !== " ") return;
                        event.preventDefault();
                        event.stopPropagation();
                        deletePage(layout.pageNumber);
                      }}
                    >
                      <Trash2 size={14} />
                    </span>
                  </span>
                  <MiniPage
                    layout={layout}
                    readingDirection={syncedProject.readingDirection}
                    rowHeights={project.rowHeights?.[layout.pageNumber]}
                    rowWidths={project.rowWidths?.[layout.pageNumber]}
                  />
                </button>
              ))}
            </div>
          ))}
          <button type="button" className="page-thumb add-page-thumb" onClick={appendPage} title={t.addPage}>
            <Plus size={28} />
          </button>
        </div>}
      </section>

      <section className="workspace">
        <div className="editor-pane">
          <div className="editor-toolbar collapsible-header" onClick={() => setIsStoryboardCollapsed((current) => !current)}>
            <div className="page-list-title">
              <button
                type="button"
                className="collapse-toggle"
                aria-expanded={!isStoryboardCollapsed}
                onClick={(event) => {
                  event.stopPropagation();
                  setIsStoryboardCollapsed((current) => !current);
                }}
              >
                <span className="collapse-triangle" aria-hidden="true" />
                <h2>{t.storyboard}</h2>
              </button>
              <button
                type="button"
                className="help-button"
                aria-label={t.storyboardHelpAria}
                onClick={(event) => {
                  event.stopPropagation();
                  setIsStoryboardHelpOpen(true);
                }}
              >
                ?
              </button>
            </div>
          </div>
          {!isStoryboardCollapsed && <div className="storyboard-table-wrap">
            <table className="storyboard-table">
              <thead>
                <tr>
                  <th>No.</th>
                  <th>{t.emotion}</th>
                  <th><HeaderLabel label={t.size} warning={columnWarnings.panelSize} /></th>
                  <th><HeaderLabel label={t.role} warning={columnWarnings.role} /></th>
                  <th>{t.shape}</th>
                  <th><HeaderLabel label={t.camera} warning={columnWarnings.camera} /></th>
                  <th><HeaderLabel label={t.faceSize} warning={columnWarnings.faceSize} /></th>
                  <th>{t.content}</th>
                  <th aria-label="actions" />
                </tr>
              </thead>
              <tbody>
                    {activeRows.map((row, rowIndex) => (
                      <tr
                        key={row.id}
                        draggable
                        className={`is-active ${row.id === selectedRowId ? "selected-row" : ""} ${row.id === draggingRowId ? "dragging-row" : ""} ${
                          row.id === dragOverRowId ? "drag-over-row" : ""
                        }`}
                        onDragStart={(event) => {
                          event.dataTransfer.effectAllowed = "move";
                          event.dataTransfer.setData("text/plain", row.id);
                          setDraggingRowId(row.id);
                        }}
                        onDragOver={(event) => {
                          event.preventDefault();
                          event.dataTransfer.dropEffect = "move";
                          setDragOverRowId(row.id);
                        }}
                        onDragLeave={() => {
                          setDragOverRowId((current) => (current === row.id ? null : current));
                        }}
                        onDrop={(event) => {
                          event.preventDefault();
                          const source = event.dataTransfer.getData("text/plain") || draggingRowId;
                          if (source) reorderRows(source, row.id);
                          setDraggingRowId(null);
                          setDragOverRowId(null);
                        }}
                        onDragEnd={() => {
                          setDraggingRowId(null);
                          setDragOverRowId(null);
                        }}
                        onClick={() => selectRow(row.id, row.pageNumber)}
                      >
                        <td className="panel-index-cell">{rowIndex + 1}</td>
                        <td>
                          <SelectCell
                            value={row.emotionSize ?? "small"}
                            warning={row.warnings?.emotionSize}
                            tone={row.emotionSize ? emotionTone[row.emotionSize] : undefined}
                            onChange={(value) => updateRow(row.id, { emotionSize: value as NonNullable<EmotionSize> })}
                          >
                            {emotionOptions.map((value) => (
                              <option key={value} value={value} className={excelToneClass(emotionTone[value])}>
                                {emotionLabel[value]}
                              </option>
                            ))}
                          </SelectCell>
                        </td>
                        <td>
                          <SelectCell
                            value={row.panelSize}
                            warning={row.warnings?.panelSize}
                            tone={sizeTone[row.panelSize]}
                            onChange={(value) => updateRow(row.id, { panelSize: value as Size })}
                          >
                            {sizeOptions.map((value) => (
                              <option key={value} value={value} className={excelToneClass(sizeTone[value])}>
                                {sizeLabel[value]}
                              </option>
                            ))}
                          </SelectCell>
                        </td>
                        <td>
                          <RoleDropdownCell value={row.role} onChange={(value) => updateRow(row.id, { role: value })} />
                        </td>
                        <td>
                          <SelectCell
                            value={row.shape}
                            warning={row.warnings?.shape}
                            tone={shapeTone[row.shape]}
                            onChange={(value) => updateRow(row.id, { shape: value as Shape })}
                          >
                            {shapeOptions.map((value) => (
                              <option key={value} value={value} className={excelToneClass(shapeTone[value])}>
                                {shapeLabel[value]}
                              </option>
                            ))}
                          </SelectCell>
                        </td>
                        <td>
                          <SelectCell
                            value={cameraOptions.includes(row.camera as (typeof cameraOptions)[number]) ? row.camera : cameraOptions[0]}
                            tone={cameraTone(row.camera)}
                            onChange={(value) => updateRow(row.id, { camera: value })}
                          >
                            {cameraOptions.map((value) => (
                              <option key={value} value={value} className={excelToneClass(cameraTone(value))}>
                                {value}
                              </option>
                            ))}
                          </SelectCell>
                        </td>
                        <td>
                          <SelectCell
                            value={row.faceSize ?? "medium"}
                            warning={row.warnings?.faceSize}
                            tone={faceSizeTone[row.faceSize ?? "medium"]}
                            onChange={(value) => updateRow(row.id, { faceSize: value as FaceSize })}
                          >
                            {faceSizeOptions.map((value) => (
                              <option key={value} value={value} className={excelToneClass(faceSizeTone[value])}>
                                {faceSizeLabel[value]}
                              </option>
                            ))}
                          </SelectCell>
                        </td>
                        <td>
                          <textarea
                            value={row.content}
                            rows={1}
                            onChange={(event) => updateRow(row.id, { content: event.target.value })}
                          />
                        </td>
                        <td>
                          <div className="row-actions" onClick={(event) => event.stopPropagation()}>
                            <button type="button" title={t.duplicateRow} onClick={() => duplicateRow(row.id)}>
                              <Copy size={15} />
                            </button>
                            <button type="button" title={t.deleteRow} onClick={() => deleteRow(row.id)}>
                              <Trash2 size={15} />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                    <tr className="add-panel-row">
                      <td colSpan={8}>
                        <button type="button" onClick={appendRowToActivePage} title={t.addPanel}>
                          <Plus size={20} />
                        </button>
                      </td>
                    </tr>
              </tbody>
            </table>
          </div>}
          <div className="project-note-field">
            <div className="project-note-header collapsible-header" onClick={() => setIsProjectNoteCollapsed((current) => !current)}>
              <div className="page-list-title">
                <button
                  type="button"
                  className="collapse-toggle"
                  aria-expanded={!isProjectNoteCollapsed}
                  onClick={(event) => {
                    event.stopPropagation();
                    setIsProjectNoteCollapsed((current) => !current);
                  }}
                >
                  <span className="collapse-triangle" aria-hidden="true" />
                  <h2>{t.projectNote}</h2>
                </button>
              </div>
            </div>
            {!isProjectNoteCollapsed && <textarea
              value={project.note}
              placeholder={t.projectNotePlaceholder}
              rows={5}
              aria-label={t.projectNote}
              onChange={(event) => updateProject({ note: event.target.value })}
            />}
          </div>
        </div>

        <aside className="preview-pane">
          <div className="preview-toolbar collapsible-header" onClick={() => setIsPreviewCollapsed((current) => !current)}>
            <div className="page-list-title">
              <button
                type="button"
                className="collapse-toggle"
                aria-expanded={!isPreviewCollapsed}
                onClick={(event) => {
                  event.stopPropagation();
                  setIsPreviewCollapsed((current) => !current);
                }}
              >
                <span className="collapse-triangle" aria-hidden="true" />
                <h2>{t.preview}</h2>
              </button>
              <button
                type="button"
                className="help-button"
                aria-label={t.previewHelpAria}
                onClick={(event) => {
                  event.stopPropagation();
                  setIsPreviewHelpOpen(true);
                }}
              >
                ?
              </button>
            </div>
            {!isPreviewCollapsed && <button
              type="button"
              className="button preview-clear-button"
              onClick={(event) => {
                event.stopPropagation();
                clearActivePageLayoutAdjustments();
              }}
            >
              {t.clear}
            </button>}
          </div>
          {!isPreviewCollapsed && activeLayout && (
            <>
              {activeLayout.warning && <p className="warning">{activeLayout.warning}</p>}
              <PagePreview
                layout={activeLayout}
                readingDirection={syncedProject.readingDirection}
                rowHeights={project.rowHeights?.[activeLayout.pageNumber]}
                rowWidths={project.rowWidths?.[activeLayout.pageNumber]}
                selectedPanelId={selectedRowId}
                onPanelSelect={selectRow}
                onPanelFrameToggle={togglePanelFrame}
                onPanelBleedToggle={togglePanelBleed}
                onRowResize={(nextHeights) => updateRowHeights(activeLayout.pageNumber, nextHeights)}
                onColumnResize={(rowIndex, nextWidths) => updateRowWidths(activeLayout.pageNumber, rowIndex, nextWidths)}
                register={(node) => {
                  pageRefs.current[activeLayout.pageNumber] = node;
                }}
              />
            </>
          )}
        </aside>
      </section>
      <div className="export-pages" aria-hidden="true">
        {layouts.map((layout) => (
          <PageCanvas
            key={layout.pageNumber}
            layout={layout}
            readingDirection={syncedProject.readingDirection}
            rowHeights={project.rowHeights?.[layout.pageNumber]}
            rowWidths={project.rowWidths?.[layout.pageNumber]}
            register={(node) => {
              if (node) pageRefs.current[layout.pageNumber] = node;
            }}
          />
        ))}
      </div>
    </main>
  );
}

function PagePreview({
  layout,
  readingDirection,
  rowHeights,
  rowWidths,
  selectedPanelId,
  onPanelSelect,
  onPanelFrameToggle,
  onPanelBleedToggle,
  onRowResize,
  onColumnResize,
  register,
}: {
  layout: ReturnType<typeof rowsToLayouts>[number];
  readingDirection: Project["readingDirection"];
  rowHeights?: number[];
  rowWidths?: number[][];
  selectedPanelId: string | null;
  onPanelSelect: (id: string, pageNumber: number) => void;
  onPanelFrameToggle: (id: string) => void;
  onPanelBleedToggle: (id: string, side: BleedSide, blockedSide: "left" | "right") => void;
  onRowResize: (nextHeights: number[]) => void;
  onColumnResize: (rowIndex: number, nextWidths: number[]) => void;
  register: (node: HTMLDivElement | null) => void;
}) {
  return (
    <div className="page-stage">
      <PageCanvas
        layout={layout}
        readingDirection={readingDirection}
        rowHeights={rowHeights}
        rowWidths={rowWidths}
        selectedPanelId={selectedPanelId}
        onPanelSelect={onPanelSelect}
        onPanelFrameToggle={onPanelFrameToggle}
        onPanelBleedToggle={onPanelBleedToggle}
        onRowResize={onRowResize}
        onColumnResize={onColumnResize}
        register={register}
      />
    </div>
  );
}

function PageCanvas({
  layout,
  readingDirection,
  rowHeights,
  rowWidths,
  selectedPanelId,
  onPanelSelect,
  onPanelFrameToggle,
  onPanelBleedToggle,
  onRowResize,
  onColumnResize,
  register,
}: {
  layout: ReturnType<typeof rowsToLayouts>[number];
  readingDirection: Project["readingDirection"];
  rowHeights?: number[];
  rowWidths?: number[][];
  selectedPanelId?: string | null;
  onPanelSelect?: (id: string, pageNumber: number) => void;
  onPanelFrameToggle?: (id: string) => void;
  onPanelBleedToggle?: (id: string, side: BleedSide, blockedSide: "left" | "right") => void;
  onRowResize?: (nextHeights: number[]) => void;
  onColumnResize?: (rowIndex: number, nextWidths: number[]) => void;
  register: (node: HTMLDivElement | null) => void;
}) {
  const pageRef = useRef<HTMLDivElement | null>(null);
  const normalizedRowHeights = normalizeRowHeights(rowHeights, layout.rows);
  const normalizedRowWidths = normalizeRowWidths(rowWidths, layout.rows.length);
  const sideMarkPosition = getSideMarkPosition(layout.pageNumber, readingDirection);
  const canEditPanelBleed = Boolean(onPanelBleedToggle);

  function setPageNode(node: HTMLDivElement | null) {
    pageRef.current = node;
    register(node);
  }

  function startRowResize(event: React.PointerEvent<HTMLButtonElement>, rowIndex: number) {
    if (!onRowResize || !pageRef.current) return;
    event.preventDefault();

    const commitRowResize = onRowResize;
    const startY = event.clientY;
    const startHeights = normalizedRowHeights;
    const rowEls = Array.from(pageRef.current.querySelectorAll<HTMLElement>(".layout-row"));
    const availableHeight = rowEls.reduce((total, rowEl) => total + rowEl.getBoundingClientRect().height, 0);
    const totalWeight = startHeights.reduce((total, weight) => total + weight, 0);
    event.currentTarget.setPointerCapture(event.pointerId);

    function resize(moveEvent: PointerEvent) {
      const deltaWeight = ((moveEvent.clientY - startY) / Math.max(availableHeight, 1)) * totalWeight;
      const previous = startHeights[rowIndex] + deltaWeight;
      const next = startHeights[rowIndex + 1] - deltaWeight;

      if (previous < minRowHeightWeight || next < minRowHeightWeight) return;

      const resized = [...startHeights];
      resized[rowIndex] = previous;
      resized[rowIndex + 1] = next;
      commitRowResize(resized);
    }

    function stopResize() {
      window.removeEventListener("pointermove", resize);
      window.removeEventListener("pointerup", stopResize);
      window.removeEventListener("pointercancel", stopResize);
    }

    window.addEventListener("pointermove", resize);
    window.addEventListener("pointerup", stopResize);
    window.addEventListener("pointercancel", stopResize);
  }

  function startColumnResize(event: React.PointerEvent<HTMLButtonElement>, rowIndex: number, boundary: number) {
    if (!onColumnResize || !pageRef.current) return;
    event.preventDefault();

    const rowEl = pageRef.current.querySelectorAll<HTMLElement>(".layout-row")[rowIndex];
    if (!rowEl) return;

    const commitColumnResize = onColumnResize;
    const startX = event.clientX;
    const startWidths = normalizedRowWidths[rowIndex];
    const rowWidth = rowEl.getBoundingClientRect().width;
    const totalWeight = startWidths.reduce((total, weight) => total + weight, 0);
    const leftIndexes = Array.from({ length: boundary }, (_, index) => index);
    const rightIndexes = Array.from({ length: gridColumnCount - boundary }, (_, index) => boundary + index);
    const leftStart = sumIndexes(startWidths, leftIndexes);
    const rightStart = sumIndexes(startWidths, rightIndexes);
    event.currentTarget.setPointerCapture(event.pointerId);

    function resize(moveEvent: PointerEvent) {
      const deltaWeight = ((moveEvent.clientX - startX) / Math.max(rowWidth, 1)) * totalWeight;
      const nextLeft = leftStart + deltaWeight;
      const nextRight = rightStart - deltaWeight;
      const minLeft = leftIndexes.length * minColumnWidthWeight;
      const minRight = rightIndexes.length * minColumnWidthWeight;

      if (nextLeft < minLeft || nextRight < minRight) return;

      const resized = [...startWidths];
      applyScaledWeights(resized, leftIndexes, nextLeft / leftStart);
      applyScaledWeights(resized, rightIndexes, nextRight / rightStart);
      commitColumnResize(rowIndex, resized);
    }

    function stopResize() {
      window.removeEventListener("pointermove", resize);
      window.removeEventListener("pointerup", stopResize);
      window.removeEventListener("pointercancel", stopResize);
    }

    window.addEventListener("pointermove", resize);
    window.addEventListener("pointerup", stopResize);
    window.addEventListener("pointercancel", stopResize);
  }

  return (
    <div className={`manga-page side-mark-${sideMarkPosition}`} data-page-number={layout.pageNumber} ref={setPageNode}>
      <div className="page-side-mark" aria-hidden="true">
        <span />
        <span />
      </div>
      <div className="hidden-page-number" aria-hidden="true">
        {layout.pageNumber}
      </div>
      {layout.rows.map((row, rowIndex) => (
        <div className="layout-row-group" key={rowIndex}>
          <div
            className="layout-row"
            style={{
              flex: `${normalizedRowHeights[rowIndex]} 1 0`,
              gridTemplateColumns: normalizedRowWidths[rowIndex].map((weight) => `${weight}fr`).join(" "),
            }}
          >
          {row.panels.map((panel) => (
            <PanelArticle
              key={panel.id}
              panel={panel}
              selectedPanelId={selectedPanelId}
              sideMarkPosition={sideMarkPosition}
              allowedBleedSides={getPanelAllowedBleedSides(panel, rowIndex, layout.rows.length, sideMarkPosition)}
              canEditPanelBleed={canEditPanelBleed}
              style={{ gridColumn: `${panel.colStart} / span ${panel.colSpan}`, gridRow: 1 }}
              onPanelSelect={onPanelSelect}
              onPanelFrameToggle={onPanelFrameToggle}
              onPanelBleedToggle={onPanelBleedToggle}
            />
          ))}
            {row.stacks?.map((stack, stackIndex) => (
              <div
                className="panel-stack"
                key={stackIndex}
                style={{ gridColumn: `${stack.colStart} / span ${stack.colSpan}`, gridRow: 1 }}
              >
                {stack.panels.map((panel, panelIndex) => (
                  <PanelArticle
                    key={panel.id}
                    panel={panel}
                    selectedPanelId={selectedPanelId}
                    sideMarkPosition={sideMarkPosition}
                    allowedBleedSides={getStackPanelAllowedBleedSides(
                      stack,
                      panelIndex,
                      rowIndex,
                      layout.rows.length,
                      sideMarkPosition,
                    )}
                    canEditPanelBleed={canEditPanelBleed}
                    onPanelSelect={onPanelSelect}
                    onPanelFrameToggle={onPanelFrameToggle}
                    onPanelBleedToggle={onPanelBleedToggle}
                  />
                ))}
              </div>
            ))}
            {onColumnResize &&
              getColumnBoundaries(row).map((boundary) => (
                <button
                  type="button"
                  className="column-resizer"
                  aria-label={t.resizeColumn}
                  key={boundary}
                  style={{ gridColumn: `${boundary + 1} / span 1`, gridRow: 1 }}
                  onPointerDown={(event) => startColumnResize(event, rowIndex, boundary)}
                />
              ))}
          </div>
          {rowIndex < layout.rows.length - 1 &&
            (onRowResize ? (
              <button
                type="button"
                className="row-resizer"
                aria-label={t.resizeRow}
                onPointerDown={(event) => startRowResize(event, rowIndex)}
              />
            ) : (
              <div className="row-resizer row-resizer-static" />
            ))}
        </div>
      ))}
    </div>
  );
}

function PanelArticle({
  panel,
  selectedPanelId,
  sideMarkPosition,
  allowedBleedSides,
  canEditPanelBleed,
  style,
  onPanelSelect,
  onPanelFrameToggle,
  onPanelBleedToggle,
}: {
  panel: LayoutPanel;
  selectedPanelId?: string | null;
  sideMarkPosition: "left" | "right";
  allowedBleedSides: BleedSide[];
  canEditPanelBleed: boolean;
  style?: CSSProperties;
  onPanelSelect?: (id: string, pageNumber: number) => void;
  onPanelFrameToggle?: (id: string) => void;
  onPanelBleedToggle?: (id: string, side: BleedSide, blockedSide: "left" | "right") => void;
}) {
  const isSelected = panel.id === selectedPanelId;
  const bleedStyle = getPanelBleedStyle(panel, sideMarkPosition, pageBorderBleedAmount, allowedBleedSides);
  const bleedEdgeClasses = getBleedEdgeClassName(panel, allowedBleedSides);

  return (
    <article
      className={`panel-frame ${panel.shape} ${bleedEdgeClasses} ${isSelected ? "selected-panel" : ""} ${
        panel.isFrameHidden ? "frame-hidden" : ""
      }`}
      style={{ ...style, ...bleedStyle }}
      onClick={() => onPanelSelect?.(panel.id, panel.pageNumber)}
      onDoubleClick={() => onPanelFrameToggle?.(panel.id)}
    >
      {panel.faceSize !== "none" && (
        <span
          className={`face-size-marker ${faceSizeMarkerClass(panel.faceSize)}`}
          data-face-size={panel.faceSize}
          aria-hidden="true"
        />
      )}
      <div className="panel-fields" data-export-text>
        <span className="panel-field-number" title="No.">
          {panel.visualNumber}
        </span>
        <span title={t.role}>{panel.role || "-"}</span>
        <span title={t.camera}>{panel.camera || "-"}</span>
      </div>
      <p data-export-text>{panel.content}</p>
      {canEditPanelBleed && isSelected && (
        <div className="bleed-controls" onClick={(event) => event.stopPropagation()}>
          {allowedBleedSides.map((side) => {
            const control = getBleedControl(side);
            const Icon = control.icon;

            return (
              <button
                key={side}
                type="button"
                className={`bleed-button bleed-${side} ${isPanelBleedSideActive(panel, side) ? "active" : ""}`}
                title={control.label}
                aria-label={control.label}
                onClick={() => onPanelBleedToggle?.(panel.id, side, sideMarkPosition)}
              >
                <Icon size={13} />
              </button>
            );
          })}
        </div>
      )}
    </article>
  );
}

function MiniPage({
  layout,
  readingDirection,
  rowHeights,
  rowWidths,
}: {
  layout: ReturnType<typeof rowsToLayouts>[number];
  readingDirection: Project["readingDirection"];
  rowHeights?: number[];
  rowWidths?: number[][];
}) {
  const normalizedRowHeights = normalizeRowHeights(rowHeights, layout.rows);
  const normalizedRowWidths = normalizeRowWidths(rowWidths, layout.rows.length);
  const sideMarkPosition = getSideMarkPosition(layout.pageNumber, readingDirection);

  return (
    <div className={`mini-page side-mark-${sideMarkPosition}`}>
      <div className="mini-side-mark" aria-hidden="true">
        ×
      </div>
      {layout.rows.map((row, rowIndex) => (
        <div className="mini-row-group" key={rowIndex} style={{ flex: `${normalizedRowHeights[rowIndex]} 1 0` }}>
          <div
            className="mini-row"
            style={{ gridTemplateColumns: normalizedRowWidths[rowIndex].map((weight) => `${weight}fr`).join(" ") }}
          >
            {row.panels.map((panel) => {
              const allowedBleedSides = getPanelAllowedBleedSides(panel, rowIndex, layout.rows.length, sideMarkPosition);

              return (
                <span
                  key={panel.id}
                  className={`${panel.isFrameHidden ? "frame-hidden" : ""} ${getBleedEdgeClassName(panel, allowedBleedSides)}`}
                  style={{
                    gridColumn: `${panel.colStart} / span ${panel.colSpan}`,
                    gridRow: 1,
                    ...getPanelBleedStyle(
                      panel,
                      sideMarkPosition,
                      miniPageBorderBleedAmount,
                      allowedBleedSides,
                    ),
                  }}
                >
                  {panel.faceSize !== "none" && (
                    <b className={`mini-face-size-marker ${faceSizeMarkerClass(panel.faceSize)}`} aria-hidden="true" />
                  )}
                </span>
              );
            })}
            {row.stacks?.map((stack, stackIndex) => (
              <span
                className="mini-stack"
                key={`stack-${stackIndex}`}
                style={{ gridColumn: `${stack.colStart} / span ${stack.colSpan}`, gridRow: 1 }}
              >
                {stack.panels.map((panel, panelIndex) => {
                  const allowedBleedSides = getStackPanelAllowedBleedSides(
                    stack,
                    panelIndex,
                    rowIndex,
                    layout.rows.length,
                    sideMarkPosition,
                  );

                  return (
                    <i
                      key={panel.id}
                      className={`${panel.isFrameHidden ? "frame-hidden" : ""} ${getBleedEdgeClassName(panel, allowedBleedSides)}`}
                      style={getPanelBleedStyle(
                        panel,
                        sideMarkPosition,
                        miniPageBorderBleedAmount,
                        allowedBleedSides,
                      )}
                    >
                      {panel.faceSize !== "none" && (
                        <b className={`mini-face-size-marker ${faceSizeMarkerClass(panel.faceSize)}`} aria-hidden="true" />
                      )}
                    </i>
                  );
                })}
              </span>
            ))}
          </div>
          {rowIndex < layout.rows.length - 1 && <div className="mini-row-gap" />}
        </div>
      ))}
    </div>
  );
}

function SelectCell({
  value,
  warning,
  tone,
  onChange,
  children,
}: {
  value: string;
  warning?: string;
  tone?: ExcelTone;
  onChange: (value: string) => void;
  children: React.ReactNode;
}) {
  return (
    <label className="cell-control">
      <select className={[warning ? "invalid" : "", excelToneClass(tone)].filter(Boolean).join(" ")} value={value} onChange={(event: ChangeEvent<HTMLSelectElement>) => onChange(event.target.value)}>
        {children}
      </select>
      {warning && <span>{warning}</span>}
    </label>
  );
}

function RoleDropdownCell({ value, onChange }: { value: string; onChange: (value: string) => void }) {
  const [isOpen, setIsOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!isOpen) return;

    function closeOnOutsidePointer(event: MouseEvent) {
      if (!rootRef.current?.contains(event.target as Node)) setIsOpen(false);
    }

    window.addEventListener("mousedown", closeOnOutsidePointer);
    return () => window.removeEventListener("mousedown", closeOnOutsidePointer);
  }, [isOpen]);

  return (
    <div className="role-dropdown" ref={rootRef}>
      <input
        className={excelToneClass(roleTone(value))}
        value={value}
        onChange={(event) => onChange(event.target.value)}
        onFocus={() => setIsOpen(true)}
      />
      <button
        type="button"
        className="role-dropdown-button"
        aria-label={t.role}
        aria-expanded={isOpen}
        onClick={() => setIsOpen((current) => !current)}
      >
        <ChevronDown size={14} />
      </button>
      {isOpen && (
        <div className="role-dropdown-menu" role="listbox">
          {roleOptions.map((option) => (
            <button
              type="button"
              className={`role-dropdown-option ${excelToneClass(roleTone(option))}`}
              key={option}
              role="option"
              aria-selected={value === option}
              onClick={() => {
                onChange(option);
                setIsOpen(false);
              }}
            >
              {option}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

function excelToneClass(tone?: ExcelTone) {
  return tone ? `excel-tone-${tone}` : "";
}

function cameraTone(value: string) {
  return cameraOptions.includes(value as (typeof cameraOptionLabels)[number])
    ? cameraToneByLabel[value as (typeof cameraOptionLabels)[number]]
    : undefined;
}

function roleTone(value: string): ExcelTone | undefined {
  if (!value) return undefined;
  if (value.includes(roleOptionLabels[3])) return "red";
  if (value.includes(roleOptionLabels[4])) return "yellow";
  if (value.includes(roleOptionLabels[1]) || value.includes(roleOptionLabels[2])) return "green";
  return undefined;
}

function HeaderLabel({ label, warning }: { label: string; warning?: string }) {
  return (
    <span className="column-heading">
      {label}
      {warning && <WarningBadge warning={warning} />}
    </span>
  );
}

function WarningBadge({ warning }: { warning: string }) {
  const badgeRef = useRef<HTMLSpanElement>(null);
  const [tooltipPosition, setTooltipPosition] = useState<{ left: number; top: number } | null>(null);

  function showTooltip() {
    const rect = badgeRef.current?.getBoundingClientRect();
    if (!rect) return;
    setTooltipPosition({
      left: rect.left + rect.width / 2,
      top: rect.top - 8,
    });
  }

  return (
    <>
      <span
        ref={badgeRef}
        className="column-warning"
        aria-label={warning}
        onBlur={() => setTooltipPosition(null)}
        onFocus={showTooltip}
        onMouseEnter={showTooltip}
        onMouseLeave={() => setTooltipPosition(null)}
        tabIndex={0}
      >
        !
      </span>
      {tooltipPosition &&
        createPortal(
        <span
          className="column-warning-tooltip"
          role="tooltip"
          style={{ left: tooltipPosition.left, top: tooltipPosition.top }}
        >
          {warning}
        </span>,
        document.body,
      )}
    </>
  );
}

function getColumnWarnings(rows: PanelRow[]) {
  return {
    panelSize: needsSizeVariation(rows.map((row) => row.panelSize)) ? t.monotonousWarning : undefined,
    role: roleWarnings(rows),
    camera: rows.length > 0 && rows.every((row) => row.camera === cameraOptionLabels[0]) ? t.monotonousWarning : undefined,
    faceSize: needsFaceSizeVariation(rows.map((row) => row.faceSize)) ? t.monotonousWarning : undefined,
  };
}

function needsSizeVariation(values: Size[]) {
  if (values.length === 0) return false;
  const hasSmall = values.some((value) => value === "extraSmall" || value === "small");
  const hasLarge = values.some((value) => value === "large" || value === "extraLarge");
  return !hasSmall || !hasLarge;
}

function needsFaceSizeVariation(values: FaceSize[]) {
  const visibleValues = values.filter((value) => value !== "none");
  if (visibleValues.length === 0) return false;
  const hasSmall = visibleValues.some((value) => value === "extraSmall" || value === "small");
  const hasLarge = visibleValues.some((value) => value === "large" || value === "extraLarge");
  return !hasSmall || !hasLarge;
}

function roleWarnings(rows: PanelRow[]) {
  const warnings: string[] = [];
  if (rows.length === 0) return undefined;
  if (!rows.some((row) => row.role.includes(roleOptionLabels[1]))) warnings.push(t.missingLocationWarning);
  if (!rows.some((row) => row.role.includes(roleOptionLabels[3]))) warnings.push(t.missingHighlightPanelWarning);
  return warnings.length ? warnings.join("\n") : undefined;
}

function normalizeOrder<T extends PanelRow>(row: T, order: number): T {
  return { ...row, order };
}

function mergeRowsWithProjectRows(importedRows: PanelRow[], sourceProject: Project) {
  const sourceRowsByPage = new Map<number, PanelRow[]>();
  projectToRows(sourceProject).forEach((row) => {
    const pageRows = sourceRowsByPage.get(row.pageNumber) ?? [];
    pageRows.push(row);
    sourceRowsByPage.set(row.pageNumber, pageRows);
  });

  const importedRowIndexByPage = new Map<number, number>();
  return importedRows.map((row, order) => {
    const pageIndex = importedRowIndexByPage.get(row.pageNumber) ?? 0;
    importedRowIndexByPage.set(row.pageNumber, pageIndex + 1);
    const sourceRow = sourceRowsByPage.get(row.pageNumber)?.[pageIndex];

    return {
      ...sourceRow,
      ...row,
      id: sourceRow?.id ?? row.id,
      isFrameHidden: sourceRow?.isFrameHidden,
      bleed: sourceRow?.bleed,
      order,
    };
  });
}

function normalizeChoiceDefaults<T extends PanelRow>(row: T): T {
  return {
    ...row,
    emotionSize: row.emotionSize ?? "small",
    camera: cameraOptions.includes(row.camera as (typeof cameraOptions)[number]) ? row.camera : cameraOptions[0],
    faceSize: row.faceSize ?? "medium",
  };
}

function faceSizeMarkerClass(faceSize: FaceSize) {
  return `face-size-${faceSize.toLowerCase()}`;
}

function normalizeRowHeights(rowHeights: number[] | undefined, rows: LayoutRow[]) {
  return rows.map((row, index) => {
    const weight = rowHeights?.[index];
    if (typeof weight === "number" && Number.isFinite(weight) && weight > 0) return weight;
    if (row.stacks?.length) return 2;
    return row.panels.some((panel) => panel.shape === "horizontal") ? horizontalRowHeightWeight : 1;
  });
}

function normalizeRowWidths(rowWidths: number[][] | undefined, rowCount: number) {
  return Array.from({ length: rowCount }, (_, rowIndex) =>
    Array.from({ length: gridColumnCount }, (_, columnIndex) => {
      const weight = rowWidths?.[rowIndex]?.[columnIndex];
      return typeof weight === "number" && Number.isFinite(weight) && weight > 0 ? weight : 1;
    }),
  );
}

function getSideMarkPosition(pageNumber: number, readingDirection: Project["readingDirection"]) {
  const isLeftPage = readingDirection === "rtl" ? pageNumber % 2 === 1 : pageNumber % 2 === 0;
  return isLeftPage ? "right" : "left";
}

function getPanelBleedStyle(
  panel: Pick<LayoutPanel, "bleed" | "shape">,
  blockedSide: "left" | "right",
  amount: number,
  allowedSides: BleedSide[],
): CSSProperties {
  if (allowedSides.length === 0) return {};
  const outerSide = getOuterBleedSide(blockedSide);

  return {
    marginTop: allowedSides.includes("top") && isPanelBleedSideActive(panel, "top") ? -amount : undefined,
    marginRight:
      outerSide === "right" && allowedSides.includes("right") && isPanelBleedSideActive(panel, "right")
        ? -amount
        : undefined,
    marginBottom: allowedSides.includes("bottom") && isPanelBleedSideActive(panel, "bottom") ? -amount : undefined,
    marginLeft:
      outerSide === "left" && allowedSides.includes("left") && isPanelBleedSideActive(panel, "left")
        ? -amount
        : undefined,
    zIndex: allowedSides.some((side) => isPanelBleedSideActive(panel, side)) ? 1 : undefined,
  };
}

function getActiveBleedSides(panel: Pick<LayoutPanel, "bleed" | "shape">, allowedSides: BleedSide[]) {
  return allowedSides.filter((side) => isPanelBleedSideActive(panel, side));
}

function getBleedEdgeClassName(panel: Pick<LayoutPanel, "bleed" | "shape">, allowedSides: BleedSide[]) {
  return getActiveBleedSides(panel, allowedSides)
    .map((side) => `bleed-edge-${side}`)
    .join(" ");
}

function isPanelBleedSideActive(panel: Pick<PanelRow, "bleed" | "shape">, side: BleedSide) {
  return panel.bleed?.[side] ?? isDefaultBleedSideActive(panel.shape, side);
}

function isDefaultBleedSideActive(shape: Shape, side: BleedSide) {
  if (shape === "vertical") return side === "top" || side === "bottom";
  if (shape === "horizontal") return side === "left" || side === "right";
  return false;
}

function getOuterBleedSide(blockedSide: "left" | "right"): "left" | "right" {
  return blockedSide === "left" ? "right" : "left";
}

function getBleedControl(side: BleedSide) {
  if (side === "top") return { label: t.extendToPageTop, icon: ArrowUp };
  if (side === "bottom") return { label: t.extendToPageBottom, icon: ArrowDown };
  return side === "left"
    ? { label: t.extendToPageOuter, icon: ArrowLeft }
    : { label: t.extendToPageOuter, icon: ArrowRight };
}

function getPanelAllowedBleedSides(
  panel: LayoutPanel,
  rowIndex: number,
  rowCount: number,
  blockedSide: "left" | "right",
) {
  const sides: BleedSide[] = [];
  const outerSide = getOuterBleedSide(blockedSide);

  if (rowIndex === 0) sides.push("top");
  if (rowIndex === rowCount - 1) sides.push("bottom");
  if (outerSide === "left" && panel.colStart === 1) sides.push("left");
  if (outerSide === "right" && panel.colStart + panel.colSpan - 1 === gridColumnCount) sides.push("right");

  return sides;
}

function getStackPanelAllowedBleedSides(
  stack: NonNullable<LayoutRow["stacks"]>[number],
  panelIndex: number,
  rowIndex: number,
  rowCount: number,
  blockedSide: "left" | "right",
) {
  const sides: BleedSide[] = [];
  const outerSide = getOuterBleedSide(blockedSide);

  if (rowIndex === 0 && panelIndex === 0) sides.push("top");
  if (rowIndex === rowCount - 1 && panelIndex === stack.panels.length - 1) sides.push("bottom");
  if (outerSide === "left" && stack.colStart === 1) sides.push("left");
  if (outerSide === "right" && stack.colStart + stack.colSpan - 1 === gridColumnCount) sides.push("right");

  return sides;
}

function getColumnBoundaries(row: LayoutRow) {
  const items = [
    ...row.panels.map((panel) => ({ colStart: panel.colStart, colSpan: panel.colSpan })),
    ...(row.stacks ?? []).map((stack) => ({ colStart: stack.colStart, colSpan: stack.colSpan })),
  ].sort((a, b) => a.colStart - b.colStart);

  const boundaries: number[] = [];
  items.forEach((item, index) => {
    const next = items[index + 1];
    if (!next) return;

    const boundary = item.colStart + item.colSpan - 1;
    if (next.colStart === boundary + 1) boundaries.push(boundary);
  });

  return boundaries;
}

function sumIndexes(values: number[], indexes: number[]) {
  return indexes.reduce((total, index) => total + values[index], 0);
}

function applyScaledWeights(values: number[], indexes: number[], scale: number) {
  indexes.forEach((index) => {
    values[index] *= scale;
  });
}

function titleFromExcelFileName(fileName: string) {
  const baseName = fileName.replace(/\.[^.]+$/, "");
  const title = baseName.replace(new RegExp(`(?:[_-]?${t.textStoryboardFileStem}|[_-]?text[-_ ]?storyboard)$`, "i"), "").trim();
  return title || "";
}

