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
import { downloadStoryboardXlsx, exportPagePdf, exportPagePng, readSheetRows } from "./io";
import { rowsToLayouts } from "./layout";
import { createDefaultProject, createPanelRow, loadProject, projectToRows, rowsToProject, saveProject } from "./storage";
import type { EmotionSize, PanelRow, Project, Shape, Size } from "./types";

const sizeOptions: Size[] = ["small", "medium", "large"];
const shapeOptions: Shape[] = ["vertical", "square", "horizontal"];
const emotionOptions = ["none", "small", "medium", "large"] as const;

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
    setStatus("編集中");
    setProject((current) => ({ ...current, ...next }));
  }

  function updateRow(id: string, patch: Partial<PanelRow>) {
    setStatus("編集中");
    setRows((current) => current.map((row) => (row.id === id ? { ...row, ...patch } : row)).map(normalizeOrder));
  }

  function duplicateRow(id: string) {
    setStatus("編集中");
    setRows((current) => {
      const index = current.findIndex((row) => row.id === id);
      if (index < 0) return current;
      const next = [...current];
      next.splice(index + 1, 0, { ...current[index], id: `panel-${crypto.randomUUID()}`, order: index + 1 });
      return next.map(normalizeOrder);
    });
  }

  function deleteRow(id: string) {
    setStatus("編集中");
    setRows((current) => {
      if (current.length === 1) return current;
      return current.filter((row) => row.id !== id).map(normalizeOrder);
    });
  }

  function appendRowToActivePage() {
    setStatus("編集中");
    setRows((current) => {
      const lastIndex = current.reduce((foundIndex, row, index) => (row.pageNumber === activePage ? index : foundIndex), -1);
      const insertIndex = lastIndex >= 0 ? lastIndex + 1 : current.length;
      const next = [...current];
      next.splice(insertIndex, 0, createPanelRow(activePage, insertIndex));
      return next.map(normalizeOrder);
    });
  }

  function deletePage(pageNumber: number) {
    setStatus("編集中");
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
    setStatus("編集中");
    setRows((current) => {
      const maxPage = Math.max(...current.map((row) => row.pageNumber), 0);
      const nextPage = maxPage + 1;
      setActivePage(nextPage);
      return [...current, createPanelRow(nextPage, current.length)].map(normalizeOrder);
    });
  }

  function reorderPages(fromPage: number, toPage: number) {
    if (fromPage === toPage) return;

    setStatus("編集中");
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

    setStatus("編集中");
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
    setStatus("表データを読み込みました");
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div className="brand-area">
          <h1>{t.appName}</h1>
        </div>
        <div className="topbar-actions">
          <button type="button" className="header-button" onClick={() => setIsSettingsOpen(true)}>
            漫画の設定
          </button>
          <button type="button" className="header-button" onClick={() => sheetInputRef.current?.click()}>
            インポート
          </button>
          <button type="button" className="header-button" onClick={() => downloadStoryboardXlsx(rows, project.title)}>
            エクスポート
          </button>
          <details className="header-menu">
            <summary>
              ダウンロード
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
                PNG
              </button>
              <button
                type="button"
                className="button primary"
                onClick={() => {
                  const pageEl = pageRefs.current[activePage];
                  if (pageEl) exportPagePdf(pageEl, `${project.title}-page-${activePage}.pdf`);
                }}
              >
                PDF
              </button>
            </div>
          </details>
        </div>
        <input
          ref={sheetInputRef}
          type="file"
          accept=".csv,.tsv,.txt,.xlsx,.xls"
          hidden
          onChange={(event) => importSheet(event.target.files?.[0])}
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
              <h2 id="settings-title">漫画の設定</h2>
              <button type="button" className="icon-button" onClick={() => setIsSettingsOpen(false)} title="閉じる">
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
                  右から左
                </button>
                <button
                  type="button"
                  className={project.readingDirection === "ltr" ? "active" : ""}
                  onClick={() => updateProject({ readingDirection: "ltr" })}
                >
                  左から右
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
                  title="ページ削除"
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
              <MiniPage layout={layout} />
            </button>
          ))}
          <button type="button" className="page-thumb add-page-thumb" onClick={appendPage} title="ページ追加">
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
                    {activeRows.map((row) => (
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
                        <td>
                          <SelectCell
                            value={row.emotionSize ?? "none"}
                            warning={row.warnings?.emotionSize}
                            onChange={(value) =>
                              updateRow(row.id, { emotionSize: value === "none" ? null : (value as Size) })
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
                      <td colSpan={7}>
                        <button type="button" onClick={appendRowToActivePage} title="コマ追加">
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
                register={(node) => {
                  pageRefs.current[activeLayout.pageNumber] = node;
                }}
              />
            </>
          )}
        </aside>
      </section>
    </main>
  );
}

function PagePreview({
  layout,
  register,
}: {
  layout: ReturnType<typeof rowsToLayouts>[number];
  register: (node: HTMLDivElement | null) => void;
}) {
  return (
    <div className="page-stage">
      <div className="manga-page" ref={register}>
        {layout.rows.map((row, rowIndex) => (
          <div className="layout-row" key={rowIndex}>
            {row.panels.map((panel) => (
              <article
                key={panel.id}
                className={`panel-frame ${panel.shape}`}
                style={{ gridColumn: `${panel.colStart} / span ${panel.colSpan}` }}
              >
                <div className="panel-number" data-export-text>
                  {panel.visualNumber}
                </div>
                <div className="panel-meta" data-export-text>
                  {sizeLabel[panel.panelSize]}・{shapeLabel[panel.shape]}
                  {panel.camera ? `・${panel.camera}` : ""}
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
      </div>
    </div>
  );
}

function MiniPage({ layout }: { layout: ReturnType<typeof rowsToLayouts>[number] }) {
  return (
    <div className="mini-page">
      {layout.rows.map((row, rowIndex) => (
        <div className="mini-row" key={rowIndex}>
          {row.panels.map((panel) => (
            <span key={panel.id} style={{ gridColumn: `${panel.colStart} / span ${panel.colSpan}` }} />
          ))}
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

function titleFromExcelFileName(fileName: string) {
  const baseName = fileName.replace(/\.[^.]+$/, "");
  const title = baseName.replace(/(?:[_-]?文字ネーム|[_-]?text[-_ ]?storyboard)$/i, "").trim();
  return title || "";
}
