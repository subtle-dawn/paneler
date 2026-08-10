import { cameraOptionLabels, roleOptionLabels, t } from "./i18n";
import type { PanelRow, Project } from "./types";

const STORAGE_KEY = "paneler.project.v1";

export function saveProject(project: Project) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(project));
}

export function loadProject(): Project | null {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return null;

  try {
    return JSON.parse(raw) as Project;
  } catch {
    return null;
  }
}

export function createId(prefix: string) {
  if (crypto.randomUUID) return `${prefix}-${crypto.randomUUID()}`;
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

export function createPanelRow(pageNumber = 1, order = 0): PanelRow {
  return {
    id: createId("panel"),
    pageNumber,
    emotionSize: "medium",
    panelSize: "medium",
    role: "",
    shape: "square",
    camera: cameraOptionLabels[0],
    faceSize: "medium",
    content: "",
    order,
  };
}

export function createDefaultProject(): Project {
  return {
    id: createId("project"),
    title: t.defaultProjectTitle,
    readingDirection: "rtl",
    pages: [
      {
        id: createId("page"),
        pageNumber: 1,
        panels: [
          {
            id: createId("panel"),
            emotionSize: "small",
            panelSize: "small",
            role: `${roleOptionLabels[1]}・${roleOptionLabels[2]}`,
            shape: "vertical",
            camera: cameraOptionLabels[0],
            faceSize: "medium",
            content: t.defaultPanelContentExterior,
            order: 0,
          },
          {
            id: createId("panel"),
            emotionSize: "small",
            panelSize: "large",
            role: roleOptionLabels[1],
            shape: "square",
            camera: cameraOptionLabels[1],
            faceSize: "medium",
            content: t.defaultPanelContentInterior,
            order: 1,
          },
          {
            id: createId("panel"),
            emotionSize: "medium",
            panelSize: "medium",
            role: "",
            shape: "square",
            camera: cameraOptionLabels[1],
            faceSize: "medium",
            content: t.defaultPanelContentMonologue,
            order: 2,
          },
        ],
      },
    ],
  };
}

export function projectToRows(project: Project): PanelRow[] {
  return project.pages
    .flatMap((page) =>
      page.panels.map((panel) => ({
        ...panel,
        pageNumber: page.pageNumber,
      })),
    )
    .sort((a, b) => a.pageNumber - b.pageNumber || a.order - b.order)
    .map((row, order) => ({ ...row, order }));
}

export function rowsToProject(project: Project, rows: PanelRow[]): Project {
  const pageNumbers = Array.from(new Set(rows.map((row) => row.pageNumber))).sort((a, b) => a - b);

  return {
    ...project,
    pages: pageNumbers.map((pageNumber) => ({
      id: project.pages.find((page) => page.pageNumber === pageNumber)?.id ?? createId("page"),
      pageNumber,
      panels: rows
        .filter((row) => row.pageNumber === pageNumber)
        .sort((a, b) => a.order - b.order)
        .map(({ pageNumber: _pageNumber, warnings: _warnings, ...panel }, order) => ({ ...panel, order })),
    })),
  };
}
