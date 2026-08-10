import { ChangeEvent, useEffect, useMemo, useRef, useState } from "react";
import {
  Copy,
  ChevronDown,
  FileSpreadsheet,
  ImageDown,
  Plus,
  Settings2,
  Trash2,
  X,
} from "lucide-react";
import { emotionLabel, shapeLabel, sizeLabel, t } from "./i18n";
import {
  downloadJson,
  downloadStoryboardXlsx,
  exportPagePng,
  exportPagesPdf,
  readProjectJson,
  readSheetRows,
} from "./io";
import { rowsToLayouts } from "./layout";
import { createDefaultProject, createPanelRow, loadProject, projectToRows, rowsToProject, saveProject } from "./storage";
import type { EmotionSize, LayoutRow, PanelRow, Project, Shape, Size } from "./types";

const sizeOptions: Size[] = ["extraSmall", "small", "medium", "large", "extraLarge", "fullPage"];
const shapeOptions: Shape[] = ["vertical", "square", "horizontal"];
const emotionOptions = ["none", "small", "medium", "large"] as const;
const minRowHeightWeight = 0.25;
const minColumnWidthWeight = 0.25;
const horizontalRowHeightWeight = 0.6;
const gridColumnCount = 6;

export function App() {
  const [project, setProject] = useState<Project>(() => loadProject() ?? createDefaultProject());
  const [rows, setRows] = useState<PanelRow[]>(() => projectToRows(loadProject() ?? createDefaultProject()));
  const [activePage, setActivePage] = useState(1);
  const [status, setStatus] = useState(t.saved);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [draggingPage, setDraggingPage] = useState<number | null>(null);
  const [dragOverPage, setDragOverPage] = useState<number | null>(null);
  const [draggingRowId, setDraggingRowId] = useState<string | null>(null);
  const [dragOverRowId, setDragOverRowId] = useState<string | null>(null);
  const jsonInputRef = useRef<HTMLInputElement>(null);
  const sheetInputRef = useRef<HTMLInputElement>(null);
  const pageRefs = useRef<Record<number, HTMLDivElement | null>>({});

  const syncedProject = useMemo(() => rowsToProject(project, rows), [project, rows]);
  const layouts = useMemo(
    () => rowsToLayouts(projectToRows(syncedProject), syncedProject.readingDirection),
    [syncedProject],
  );
  const activeLayout = layouts.find((layout) => layout.pageNumber === activePage) ?? layouts[0];
  const activeRows = useMemo(
    () => rows.filter((row) => row.pageNumber === activePage).sort((a, b) => a.order - b.order),
    [activePage, rows],
  );

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
      if (event.key === "Escape") setIsSettingsOpen(false);
    }

    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [isSettingsOpen]);

  function updateProject(next: Partial<Project>) {
    setStatus(t.edited);
    setProject((current) => ({ ...current, ...next }));
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

  function updateRow(id: string, patch: Partial<PanelRow>) {
    setStatus(t.edited);
    setRows((current) => current.map((row) => (row.id === id ? { ...row, ...patch } : row)).map(normalizeOrder));
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
    const importedRows = await readSheetRows(file);
    const importedTitle = titleFromExcelFileName(file.name);
    if (importedTitle) {
      setProject((current) => ({ ...current, title: importedTitle }));
    }
    setRows(importedRows.map(normalizeOrder));
    setActivePage(importedRows[0]?.pageNumber ?? 1);
    setStatus(t.sheetLoaded);
  }

  async function importProjectJson(file?: File) {
    if (!file) return;
    const importedProject = await readProjectJson(file);
    const importedRows = projectToRows(importedProject);
    setProject(importedProject);
    setRows(importedRows.map(normalizeOrder));
    setActivePage(importedRows[0]?.pageNumber ?? importedProject.pages[0]?.pageNumber ?? 1);
    setStatus(t.jsonLoaded);
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-area">
          <h1>{t.appName}</h1>
        </div>
        <div className="topbar-actions">
          <button type="button" className="header-button" onClick={() => setIsSettingsOpen(true)}>
            {t.settings}
          </button>
          <details className="header-menu">
            <summary>
              {t.import}
              <ChevronDown size={15} />
            </summary>
            <div className="dropdown-panel">
              <button type="button" className="button" onClick={() => sheetInputRef.current?.click()}>
                {t.importSheet}
              </button>
              <button type="button" className="button" onClick={() => jsonInputRef.current?.click()}>
                {t.importJson}
              </button>
            </div>
          </details>
          <details className="header-menu">
            <summary>
              {t.export}
              <ChevronDown size={15} />
            </summary>
            <div className="dropdown-panel">
              <button type="button" className="button" onClick={() => downloadStoryboardXlsx(rows, project.title)}>
                {t.exportSheet}
              </button>
              <button type="button" className="button" onClick={() => downloadJson(syncedProject)}>
                {t.exportJson}
              </button>
            </div>
          </details>
          <details className="header-menu">
            <summary>
              {t.downloads}
              <ChevronDown size={15} />
            </summary>
            <div className="dropdown-panel">
              <button
                type="button"
                className="button"
                onClick={() => {
                  const pageEl = pageRefs.current[activePage];
                  if (pageEl) exportPagePng(pageEl, `${project.title}-page-${activePage}.png`);
                }}
              >
                {t.exportPngActivePage}
              </button>
              <button
                type="button"
                className="button"
                onClick={() => {
                  const pageEls = layouts
                    .map((layout) => pageRefs.current[layout.pageNumber])
                    .filter((pageEl): pageEl is HTMLDivElement => Boolean(pageEl));
                  exportPagesPdf(pageEls, `${project.title}_コマ割り.pdf`);
                }}
              >
                {t.exportPdfAllPages}
              </button>
            </div>
          </details>
        </div>
        <input
          ref={jsonInputRef}
          type="file"
          accept=".json,application/json"
          hidden
          onChange={(event) => {
            importProjectJson(event.target.files?.[0]);
            event.currentTarget.value = "";
          }}
        />
        <input
          ref={sheetInputRef}
          type="file"
          accept=".csv,.tsv,.txt,.xlsx,.xls"
          hidden
          onChange={(event) => {
            importSheet(event.target.files?.[0]);
            event.currentTarget.value = "";
          }}
        />
      </header>

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
          </section>
        </div>
      )}

      <section className="page-list-bar" aria-label={t.pages}>
        <div className="page-strip thumbnail-strip">
          {layouts.map((layout) => (
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
                rowHeights={project.rowHeights?.[layout.pageNumber]}
                rowWidths={project.rowWidths?.[layout.pageNumber]}
              />
            </button>
          ))}
          <button type="button" className="page-thumb add-page-thumb" onClick={appendPage} title={t.addPage}>
            <Plus size={28} />
          </button>
        </div>
      </section>

      <section className="workspace">
        <div className="editor-pane">
          <div className="storyboard-table-wrap">
            <table className="storyboard-table">
              <thead>
                <tr>
                  <th>No.</th>
                  <th>{t.emotion}</th>
                  <th>{t.size}</th>
                  <th>{t.role}</th>
                  <th>{t.shape}</th>
                  <th>{t.camera}</th>
                  <th>{t.content}</th>
                  <th aria-label="actions" />
                </tr>
              </thead>
              <tbody>
                    {activeRows.map((row, rowIndex) => (
                      <tr
                        key={row.id}
                        draggable
                        className={`is-active ${row.id === draggingRowId ? "dragging-row" : ""} ${
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
                      >
                        <td className="panel-index-cell">{rowIndex + 1}</td>
                        <td>
                          <SelectCell
                            value={row.emotionSize ?? "none"}
                            warning={row.warnings?.emotionSize}
                            onChange={(value) =>
                              updateRow(row.id, { emotionSize: value === "none" ? null : (value as NonNullable<EmotionSize>) })
                            }
                          >
                            {emotionOptions.map((value) => (
                              <option key={value} value={value}>
                                {value === "none" ? emotionLabel.none : emotionLabel[value]}
                              </option>
                            ))}
                          </SelectCell>
                        </td>
                        <td>
                          <SelectCell
                            value={row.panelSize}
                            warning={row.warnings?.panelSize}
                            onChange={(value) => updateRow(row.id, { panelSize: value as Size })}
                          >
                            {sizeOptions.map((value) => (
                              <option key={value} value={value}>
                                {sizeLabel[value]}
                              </option>
                            ))}
                          </SelectCell>
                        </td>
                        <td>
                          <input value={row.role} onChange={(event) => updateRow(row.id, { role: event.target.value })} />
                        </td>
                        <td>
                          <SelectCell
                            value={row.shape}
                            warning={row.warnings?.shape}
                            onChange={(value) => updateRow(row.id, { shape: value as Shape })}
                          >
                            {shapeOptions.map((value) => (
                              <option key={value} value={value}>
                                {shapeLabel[value]}
                              </option>
                            ))}
                          </SelectCell>
                        </td>
                        <td>
                          <input
                            value={row.camera}
                            onChange={(event) => updateRow(row.id, { camera: event.target.value })}
                          />
                        </td>
                        <td>
                          <textarea
                            value={row.content}
                            rows={1}
                            onChange={(event) => updateRow(row.id, { content: event.target.value })}
                          />
                        </td>
                        <td>
                          <div className="row-actions">
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
          </div>
        </div>

        <aside className="preview-pane">
          {activeLayout && (
            <>
              {activeLayout.warning && <p className="warning">{activeLayout.warning}</p>}
              <PagePreview
                layout={activeLayout}
                rowHeights={project.rowHeights?.[activeLayout.pageNumber]}
                rowWidths={project.rowWidths?.[activeLayout.pageNumber]}
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
  rowHeights,
  rowWidths,
  onRowResize,
  onColumnResize,
  register,
}: {
  layout: ReturnType<typeof rowsToLayouts>[number];
  rowHeights?: number[];
  rowWidths?: number[][];
  onRowResize: (nextHeights: number[]) => void;
  onColumnResize: (rowIndex: number, nextWidths: number[]) => void;
  register: (node: HTMLDivElement | null) => void;
}) {
  return (
    <div className="page-stage">
      <PageCanvas
        layout={layout}
        rowHeights={rowHeights}
        rowWidths={rowWidths}
        onRowResize={onRowResize}
        onColumnResize={onColumnResize}
        register={register}
      />
    </div>
  );
}

function PageCanvas({
  layout,
  rowHeights,
  rowWidths,
  onRowResize,
  onColumnResize,
  register,
}: {
  layout: ReturnType<typeof rowsToLayouts>[number];
  rowHeights?: number[];
  rowWidths?: number[][];
  onRowResize?: (nextHeights: number[]) => void;
  onColumnResize?: (rowIndex: number, nextWidths: number[]) => void;
  register: (node: HTMLDivElement | null) => void;
}) {
  const pageRef = useRef<HTMLDivElement | null>(null);
  const normalizedRowHeights = normalizeRowHeights(rowHeights, layout.rows);
  const normalizedRowWidths = normalizeRowWidths(rowWidths, layout.rows.length);

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
    <div className="manga-page" ref={setPageNode}>
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
            <article
              key={panel.id}
              className={`panel-frame ${panel.shape}`}
              style={{ gridColumn: `${panel.colStart} / span ${panel.colSpan}`, gridRow: 1 }}
            >
              <div className="panel-number" data-export-text>
                {panel.visualNumber}
              </div>
              <div className="panel-meta" data-export-text>
                {sizeLabel[panel.panelSize]}{t.panelSeparator}{shapeLabel[panel.shape]}
                {panel.camera ? `${t.panelSeparator}${panel.camera}` : ""}
              </div>
              {panel.role && (
                <div className="panel-role" data-export-text>
                  {panel.role}
                </div>
              )}
              <p data-export-text>{panel.content}</p>
            </article>
          ))}
            {row.stacks?.map((stack, stackIndex) => (
              <div
                className="panel-stack"
                key={stackIndex}
                style={{ gridColumn: `${stack.colStart} / span ${stack.colSpan}`, gridRow: 1 }}
              >
                {stack.panels.map((panel) => (
                  <article key={panel.id} className={`panel-frame ${panel.shape}`}>
                    <div className="panel-number" data-export-text>
                      {panel.visualNumber}
                    </div>
                    <div className="panel-meta" data-export-text>
                      {sizeLabel[panel.panelSize]}{t.panelSeparator}{shapeLabel[panel.shape]}
                      {panel.camera ? `${t.panelSeparator}${panel.camera}` : ""}
                    </div>
                    {panel.role && (
                      <div className="panel-role" data-export-text>
                        {panel.role}
                      </div>
                    )}
                    <p data-export-text>{panel.content}</p>
                  </article>
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

function MiniPage({
  layout,
  rowHeights,
  rowWidths,
}: {
  layout: ReturnType<typeof rowsToLayouts>[number];
  rowHeights?: number[];
  rowWidths?: number[][];
}) {
  const normalizedRowHeights = normalizeRowHeights(rowHeights, layout.rows);
  const normalizedRowWidths = normalizeRowWidths(rowWidths, layout.rows.length);

  return (
    <div className="mini-page">
      {layout.rows.map((row, rowIndex) => (
        <div className="mini-row-group" key={rowIndex} style={{ flex: `${normalizedRowHeights[rowIndex]} 1 0` }}>
          <div
            className="mini-row"
            style={{ gridTemplateColumns: normalizedRowWidths[rowIndex].map((weight) => `${weight}fr`).join(" ") }}
          >
            {row.panels.map((panel) => (
              <span key={panel.id} style={{ gridColumn: `${panel.colStart} / span ${panel.colSpan}`, gridRow: 1 }} />
            ))}
            {row.stacks?.map((stack, stackIndex) => (
              <span
                className="mini-stack"
                key={`stack-${stackIndex}`}
                style={{ gridColumn: `${stack.colStart} / span ${stack.colSpan}`, gridRow: 1 }}
              >
                {stack.panels.map((panel) => (
                  <i key={panel.id} />
                ))}
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
  onChange,
  children,
}: {
  value: string;
  warning?: string;
  onChange: (value: string) => void;
  children: React.ReactNode;
}) {
  return (
    <label className="cell-control">
      <select className={warning ? "invalid" : ""} value={value} onChange={(event: ChangeEvent<HTMLSelectElement>) => onChange(event.target.value)}>
        {children}
      </select>
      {warning && <span>{warning}</span>}
    </label>
  );
}

function normalizeOrder<T extends PanelRow>(row: T, order: number): T {
  return { ...row, order };
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
  const title = baseName.replace(/(?:[_-]?文字ネーム|[_-]?text[-_ ]?storyboard)$/i, "").trim();
  return title || "";
}

