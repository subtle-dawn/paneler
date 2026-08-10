import { ChangeEvent, type CSSProperties, useEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import {
  ArrowDown,
  ArrowLeft,
  ArrowRight,
  ArrowUp,
  Copy,
  ChevronDown,
  FileSpreadsheet,
  ImageDown,
  Plus,
  RotateCcw,
  Settings2,
  Trash2,
  X,
} from "lucide-react";
import { emotionLabel, faceSizeLabel, shapeLabel, sizeLabel, t } from "./i18n";
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
import type { BleedSide, EmotionSize, FaceSize, LayoutPanel, LayoutRow, PanelRow, Project, Shape, Size } from "./types";

const sizeOptions: Size[] = ["extraSmall", "small", "medium", "large", "extraLarge", "fullPage"];
const shapeOptions: Shape[] = ["square", "vertical", "horizontal"];
const emotionOptions = ["small", "medium", "large"] as const;
const faceSizeOptions: FaceSize[] = ["none", "extraSmall", "small", "medium", "large", "extraLarge"];
const roleOptions = ["ー", "場", "時", "魅", "予"] as const;
const cameraOptions = ["正", "俯", "煽", "横", "上", "下"] as const;
const minRowHeightWeight = 0.25;
const minColumnWidthWeight = 0.25;
const horizontalRowHeightWeight = 0.6;
const gridColumnCount = 6;
const pageBorderBleedAmount = 40;
const miniPageBorderBleedAmount = 4;

export function App() {
  const [project, setProject] = useState<Project>(() => loadProject() ?? createDefaultProject());
  const [rows, setRows] = useState<PanelRow[]>(() =>
    projectToRows(loadProject() ?? createDefaultProject()).map(normalizeChoiceDefaults),
  );
  const [activePage, setActivePage] = useState(1);
  const [status, setStatus] = useState(t.saved);
  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [isPageHelpOpen, setIsPageHelpOpen] = useState(false);
  const [isStoryboardHelpOpen, setIsStoryboardHelpOpen] = useState(false);
  const [isPreviewHelpOpen, setIsPreviewHelpOpen] = useState(false);
  const [draggingPage, setDraggingPage] = useState<number | null>(null);
  const [dragOverPage, setDragOverPage] = useState<number | null>(null);
  const [draggingRowId, setDraggingRowId] = useState<string | null>(null);
  const [dragOverRowId, setDragOverRowId] = useState<string | null>(null);
  const [selectedRowId, setSelectedRowId] = useState<string | null>(null);
  const jsonInputRef = useRef<HTMLInputElement>(null);
  const sheetInputRef = useRef<HTMLInputElement>(null);
  const pageRefs = useRef<Record<number, HTMLDivElement | null>>({});
  const rowRefs = useRef<Record<string, HTMLTableRowElement | null>>({});

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
  const pageSpreads = useMemo(() => {
    const spreads: typeof layouts[] = [];
    if (layouts[0]) spreads.push([layouts[0]]);
    for (let index = 1; index < layouts.length; index += 2) {
      spreads.push(layouts.slice(index, index + 2));
    }
    return spreads;
  }, [layouts]);
  const pageCountWarning = layouts.length % 4 === 0 ? undefined : "ページ数が４の倍数でないです";
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
      if (event.key === "Escape") setIsSettingsOpen(false);
    }

    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [isSettingsOpen]);

  useEffect(() => {
    function closeMenusOnOutsideClick(event: PointerEvent) {
      if (!(event.target instanceof Node)) return;
      const target = event.target;

      document.querySelectorAll<HTMLDetailsElement>(".header-menu[open]").forEach((menu) => {
        if (!menu.contains(target)) menu.open = false;
      });
    }

    document.addEventListener("pointerdown", closeMenusOnOutsideClick);
    return () => document.removeEventListener("pointerdown", closeMenusOnOutsideClick);
  }, []);

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
    if (shouldSelect) {
      window.requestAnimationFrame(() => {
        rowRefs.current[id]?.scrollIntoView({ behavior: "smooth", block: "center" });
      });
    }
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
    const importedRows = await readSheetRows(file);
    const importedTitle = titleFromExcelFileName(file.name);
    if (importedTitle) {
      setProject((current) => ({ ...current, title: importedTitle }));
    }
    setRows(importedRows.map(normalizeChoiceDefaults).map(normalizeOrder));
    setActivePage(importedRows[0]?.pageNumber ?? 1);
    setStatus(t.sheetLoaded);
  }

  async function importProjectJson(file?: File) {
    if (!file) return;
    const importedProject = await readProjectJson(file);
    const importedRows = projectToRows(importedProject);
    setProject(importedProject);
    setRows(importedRows.map(normalizeChoiceDefaults).map(normalizeOrder));
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
              <h2 id="page-help-title">ページ一覧</h2>
              <button type="button" className="icon-button" onClick={() => setIsPageHelpOpen(false)} title={t.close}>
                <X size={18} />
              </button>
            </div>
            <p className="page-help-text">
              各ページを選択して切り替えられます。
              <br />
              ゴミ箱ボタンを押すとページを削除できます。
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
              <h2 id="storyboard-help-title">文字ネーム</h2>
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
              １行＝１コマです。
              <br />
              各列を入力することで、自動的にネームにコマ割りが生成されます。
              <br />
              行＝コマの複製と削除もできます。
              <br />
              最下行の＋ボタンを押すと、行＝コマを追加できます。
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
              <h2 id="preview-help-title">ネーム</h2>
              <button type="button" className="icon-button" onClick={() => setIsPreviewHelpOpen(false)} title={t.close}>
                <X size={18} />
              </button>
            </div>
            <p className="page-help-text">
              コマの幅と高さをドラッグで調節できます。
              <br />
              コマをクリックすると、文字ネームの該当行がハイライトされます。
              <br />
              コマをダブルクリックすると、コマ枠無しにできます。
              <br />
              外側のコマは、矢印をクリックすることで伸ばしたり縮めたりできます。
              <br />
              クリアボタンを押すと、自動生成されたコマ割りに戻ります。
            </p>
          </section>
        </div>
      )}

      <section className={`page-list-bar reading-${syncedProject.readingDirection}`} aria-label={t.pages}>
        <div className="page-list-header">
          <div className="page-list-title">
            <h2>ページ一覧</h2>
            <button
              type="button"
              className="help-button"
              aria-label="ページ一覧の説明"
              onClick={() => setIsPageHelpOpen(true)}
            >
              ?
            </button>
            {pageCountWarning && <WarningBadge warning={pageCountWarning} />}
          </div>
        </div>
        <div className={`page-strip thumbnail-strip reading-${syncedProject.readingDirection}`}>
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
        </div>
      </section>

      <section className="workspace">
        <div className="editor-pane">
          <div className="editor-toolbar">
            <div className="page-list-title">
              <h2>文字ネーム</h2>
              <button
                type="button"
                className="help-button"
                aria-label="文字ネームの説明"
                onClick={() => setIsStoryboardHelpOpen(true)}
              >
                ?
              </button>
            </div>
          </div>
          <div className="storyboard-table-wrap">
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
                        ref={(node) => {
                          rowRefs.current[row.id] = node;
                        }}
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
                      >
                        <td className="panel-index-cell">{rowIndex + 1}</td>
                        <td>
                          <SelectCell
                            value={row.emotionSize ?? "small"}
                            warning={row.warnings?.emotionSize}
                            onChange={(value) => updateRow(row.id, { emotionSize: value as NonNullable<EmotionSize> })}
                          >
                            {emotionOptions.map((value) => (
                              <option key={value} value={value}>
                                {emotionLabel[value]}
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
                          <input
                            list="role-options"
                            value={row.role}
                            onChange={(event) => updateRow(row.id, { role: event.target.value })}
                          />
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
                          <SelectCell
                            value={cameraOptions.includes(row.camera as (typeof cameraOptions)[number]) ? row.camera : cameraOptions[0]}
                            onChange={(value) => updateRow(row.id, { camera: value })}
                          >
                            {cameraOptions.map((value) => (
                              <option key={value} value={value}>
                                {value}
                              </option>
                            ))}
                          </SelectCell>
                        </td>
                        <td>
                          <SelectCell
                            value={row.faceSize ?? "medium"}
                            warning={row.warnings?.faceSize}
                            onChange={(value) => updateRow(row.id, { faceSize: value as FaceSize })}
                          >
                            {faceSizeOptions.map((value) => (
                              <option key={value} value={value}>
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
          <div className="preview-toolbar">
            <div className="page-list-title">
              <h2>ネーム</h2>
              <button
                type="button"
                className="help-button"
                aria-label="ネームの説明"
                onClick={() => setIsPreviewHelpOpen(true)}
              >
                ?
              </button>
            </div>
            <button type="button" className="button preview-clear-button" onClick={clearActivePageLayoutAdjustments}>
              クリア
            </button>
          </div>
          {activeLayout && (
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
      <datalist id="role-options">
        {roleOptions.map((value) => (
          <option key={value} value={value} />
        ))}
      </datalist>
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
    <div className={`manga-page side-mark-${sideMarkPosition}`} ref={setPageNode}>
      <div className="page-side-mark" aria-hidden="true">
        <span />
        <span />
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
    panelSize: needsSizeVariation(rows.map((row) => row.panelSize)) ? "単調です" : undefined,
    role: roleWarnings(rows),
    camera: rows.length > 0 && rows.every((row) => row.camera === "正") ? "単調です" : undefined,
    faceSize: needsFaceSizeVariation(rows.map((row) => row.faceSize)) ? "単調です" : undefined,
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
  if (!rows.some((row) => row.role.includes("場"))) warnings.push("場所がわかりません");
  if (!rows.some((row) => row.role.includes("魅"))) warnings.push("魅せゴマがありません");
  return warnings.length ? warnings.join("\n") : undefined;
}

function normalizeOrder<T extends PanelRow>(row: T, order: number): T {
  return { ...row, order };
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
  if (side === "top") return { label: "ページ上端まで伸ばす", icon: ArrowUp };
  if (side === "bottom") return { label: "ページ下端まで伸ばす", icon: ArrowDown };
  return side === "left"
    ? { label: "ページ外側まで伸ばす", icon: ArrowLeft }
    : { label: "ページ外側まで伸ばす", icon: ArrowRight };
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
  const title = baseName.replace(/(?:[_-]?文字ネーム|[_-]?text[-_ ]?storyboard)$/i, "").trim();
  return title || "";
}

