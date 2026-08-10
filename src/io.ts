import { sizeLabel, shapeLabel } from "./i18n";
import { createId, createPanelRow } from "./storage";
import type { EmotionSize, PanelRow, Project, Shape, Size } from "./types";

type RawSheetRow = Record<string, unknown>;

const headerMap: Record<string, keyof PanelRow | "content"> = {
  page: "pageNumber",
  "page number": "pageNumber",
  ページ: "pageNumber",
  pagenumber: "pageNumber",
  emotion: "emotionSize",
  感情: "emotionSize",
  size: "panelSize",
  panelsize: "panelSize",
  大きさ: "panelSize",
  役割: "role",
  role: "role",
  shape: "shape",
  形: "shape",
  camera: "camera",
  カメラ: "camera",
  content: "content",
  内容: "content",
  本文: "content",
};

export function downloadJson(project: Project) {
  downloadBlob(`${safeFileName(project.title)}.json`, JSON.stringify(project, null, 2), "application/json");
}

export async function downloadStoryboardXlsx(rows: PanelRow[], title: string) {
  const xlsx = await import("xlsx");
  const sorted = rows.slice().sort((a, b) => a.pageNumber - b.pageNumber || a.order - b.order);
  const table: string[][] = [["感情", "大きさ", "役割", "形", "カメラ", "内容"]];
  let currentPage = 0;

  sorted.forEach((row) => {
    if (row.pageNumber !== currentPage) {
      currentPage = row.pageNumber;
      table.push([`${currentPage}ページ目`, pageSideLabel(currentPage), "", "", "", ""]);
    }

    table.push([
      row.emotionSize ? sizeLabel[row.emotionSize] : "なし",
      sizeLabel[row.panelSize],
      row.role,
      shapeLabel[row.shape],
      row.camera,
      row.content,
    ]);
  });

  const worksheet = xlsx.utils.aoa_to_sheet(table);
  worksheet["!cols"] = [{ wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 10 }, { wch: 12 }, { wch: 64 }];
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "文字ネーム");
  xlsx.writeFile(workbook, `${safeFileName(title)}_文字ネーム.xlsx`);
}

export function downloadCsv(rows: PanelRow[], title: string) {
  const csvRows = [
    ["ページ", "感情", "大きさ", "役割", "形", "カメラ", "内容"],
    ...rows.map((row) => [
      String(row.pageNumber),
      row.emotionSize ? sizeLabel[row.emotionSize] : "なし",
      sizeLabel[row.panelSize],
      row.role,
      shapeLabel[row.shape],
      row.camera,
      row.content,
    ]),
  ];
  const csv = csvRows.map((row) => row.map(escapeCsvValue).join(",")).join("\n");
  downloadBlob(`${safeFileName(title)}.csv`, `\uFEFF${csv}`, "text/csv;charset=utf-8");
}

export async function readProjectJson(file: File): Promise<Project> {
  const text = await file.text();
  return JSON.parse(text) as Project;
}

export async function readSheetRows(file: File): Promise<PanelRow[]> {
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (ext === "csv" || ext === "tsv" || ext === "txt") {
    return rowsFromTable(parseCsv((await file.text()).replace(/^\uFEFF/, ""), ext === "tsv" ? "\t" : ","));
  }

  const xlsx = await import("xlsx");
  const data = await file.arrayBuffer();
  const workbook = xlsx.read(data);
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const table = xlsx.utils.sheet_to_json<string[]>(firstSheet, { header: 1, defval: "", raw: false });
  return rowsFromTable(table);
}

export async function exportPagePng(pageEl: HTMLElement, fileName: string) {
  const canvas = renderPageToCanvas(pageEl);
  const blob = await canvasToBlob(canvas, "image/png");
  if (blob) downloadBlob(fileName, blob, "image/png");
}

export async function exportPagesPdf(pageEls: HTMLElement[], fileName: string) {
  if (pageEls.length === 0) return;

  const { jsPDF } = await import("jspdf");
  const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });

  pageEls.forEach((pageEl, index) => {
    if (index > 0) pdf.addPage();
    const canvas = renderPageToCanvas(pageEl);
    const image = canvas.toDataURL("image/png");
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const margin = 10;
    const maxWidth = pageWidth - margin * 2;
    const maxHeight = pageHeight - margin * 2;
    const imageRatio = canvas.width / canvas.height;
    let width = maxWidth;
    let height = width / imageRatio;

    if (height > maxHeight) {
      height = maxHeight;
      width = height * imageRatio;
    }

    pdf.addImage(image, "PNG", (pageWidth - width) / 2, (pageHeight - height) / 2, width, height);
  });

  pdf.save(fileName);
}

function renderPageToCanvas(pageEl: HTMLElement) {
  const rect = pageEl.getBoundingClientRect();
  const scale = 2;
  const canvas = document.createElement("canvas");
  canvas.width = Math.round(rect.width * scale);
  canvas.height = Math.round(rect.height * scale);
  const ctx = canvas.getContext("2d");
  if (!ctx) return canvas;

  ctx.scale(scale, scale);
  ctx.fillStyle = "#f7f1e4";
  ctx.fillRect(0, 0, rect.width, rect.height);

  pageEl.querySelectorAll<HTMLElement>(".panel-frame").forEach((panelEl) => {
    const bounds = panelEl.getBoundingClientRect();
    const x = bounds.left - rect.left;
    const y = bounds.top - rect.top;
    const width = bounds.width;
    const height = bounds.height;

    ctx.fillStyle = "#fffdfa";
    ctx.strokeStyle = "#181611";
    ctx.lineWidth = 2;
    ctx.fillRect(x, y, width, height);
    ctx.strokeRect(x, y, width, height);

    const lines = Array.from(panelEl.querySelectorAll<HTMLElement>("[data-export-text]")).map((node) =>
      node.innerText.trim(),
    );
    ctx.fillStyle = "#181611";
    ctx.font = "12px sans-serif";
    let lineY = y + 18;
    lines.forEach((line) => {
      wrapText(ctx, line, x + 10, lineY, width - 20, 16);
      lineY += Math.max(16, Math.ceil(ctx.measureText(line).width / Math.max(width - 20, 1)) * 16);
    });
  });

  return canvas;
}

function rowsFromTable(table: string[][]): PanelRow[] {
  const headerIndex = table.findIndex((row) => row.some((cell) => normalizeHeader(String(cell)) in headerMap));
  if (headerIndex < 0) return [];

  const headers = table[headerIndex].map((cell) => normalizeHeader(String(cell)));
  const hasPageColumn = headers.some((header) => headerMap[header] === "pageNumber");
  const dataRows = table.slice(headerIndex + 1);

  if (!hasPageColumn) {
    return rowsFromSectionedTable(headers, dataRows);
  }

  const objects = dataRows.map((row) =>
    headers.reduce<RawSheetRow>((object, header, index) => {
      object[header] = row[index] ?? "";
      return object;
    }, {}),
  );

  return rowsFromObjects(objects);
}

function rowsFromSectionedTable(headers: string[], table: string[][]): PanelRow[] {
  const rows: PanelRow[] = [];
  let currentPage = 1;

  table.forEach((cells) => {
    const first = String(cells[0] ?? "").trim();
    const pageMatch = first.match(/^(\d+)\s*ページ目?$/);
    if (pageMatch) {
      currentPage = Number(pageMatch[1]);
      return;
    }

    const isBlank = cells.every((cell) => String(cell ?? "").trim() === "");
    if (isBlank) return;

    const object = headers.reduce<RawSheetRow>((raw, header, index) => {
      const key = headerMap[header];
      if (key) raw[key] = cells[index] ?? "";
      return raw;
    }, {});

    rows.push(objectToPanelRow(object, rows.length, currentPage));
  });

  return rows;
}

function rowsFromObjects(objects: RawSheetRow[]): PanelRow[] {
  return objects
    .filter((raw) => Object.values(raw).some((value) => String(value ?? "").trim() !== ""))
    .map((raw, order) => objectToPanelRow(raw, order));
}

function objectToPanelRow(raw: RawSheetRow, order: number, fallbackPageNumber = 1): PanelRow {
    const row = createPanelRow(1, order);
    const normalized: Partial<PanelRow> = {};

    Object.entries(raw).forEach(([header, value]) => {
      const key = headerMap[normalizeHeader(header)];
      if (!key) return;
      normalized[key] = String(value ?? "") as never;
    });

    const pageNumber = parsePageNumber(String(normalized.pageNumber ?? fallbackPageNumber));
    const panelSize = parseSize(String(normalized.panelSize ?? row.panelSize));
    const emotionSize = parseEmotion(String(normalized.emotionSize ?? ""));
    const shape = parseShape(String(normalized.shape ?? row.shape));

    return {
      ...row,
      id: createId("panel"),
      pageNumber: Number.isFinite(pageNumber) && pageNumber > 0 ? pageNumber : 1,
      emotionSize: emotionSize.value,
      panelSize: panelSize.value,
      role: String(normalized.role ?? ""),
      shape: shape.value,
      camera: String(normalized.camera ?? ""),
      content: String(normalized.content ?? ""),
      order,
      warnings: {
        pageNumber: Number.isFinite(pageNumber) && pageNumber > 0 ? undefined : "ページ番号を確認してください",
        emotionSize: emotionSize.warning,
        panelSize: panelSize.warning,
        shape: shape.warning,
      },
    };
}

function parsePageNumber(value: string) {
  const match = value.trim().match(/\d+/);
  return match ? Number(match[0]) : NaN;
}

function parseSize(value: string): { value: Size; warning?: string } {
  const normalized = value.trim().toLowerCase();
  if (["小", "small", "s"].includes(normalized)) return { value: "small" };
  if (["中", "medium", "m"].includes(normalized)) return { value: "medium" };
  if (["大", "large", "l"].includes(normalized)) return { value: "large" };
  return { value: "small", warning: "小・中・大から選んでください" };
}

function parseEmotion(value: string): { value: EmotionSize; warning?: string } {
  const normalized = value.trim().toLowerCase();
  if (["", "なし", "none", "null", "-"].includes(normalized)) return { value: null };
  return parseSize(normalized);
}

function parseShape(value: string): { value: Shape; warning?: string } {
  const normalized = value.trim().toLowerCase();
  if (["縦", "vertical", "v"].includes(normalized)) return { value: "vertical" };
  if (["正", "square", "s"].includes(normalized)) return { value: "square" };
  if (["横", "horizontal", "h"].includes(normalized)) return { value: "horizontal" };
  return { value: "vertical", warning: "縦・正・横から選んでください" };
}

function parseCsv(text: string, delimiter: string) {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = "";
  let quoted = false;

  for (let index = 0; index < text.length; index += 1) {
    const char = text[index];
    const next = text[index + 1];

    if (char === '"' && quoted && next === '"') {
      cell += '"';
      index += 1;
    } else if (char === '"') {
      quoted = !quoted;
    } else if (char === delimiter && !quoted) {
      row.push(cell);
      cell = "";
    } else if ((char === "\n" || char === "\r") && !quoted) {
      if (char === "\r" && next === "\n") index += 1;
      row.push(cell);
      if (row.some((value) => value.trim() !== "")) rows.push(row);
      row = [];
      cell = "";
    } else {
      cell += char;
    }
  }

  row.push(cell);
  if (row.some((value) => value.trim() !== "")) rows.push(row);
  return rows;
}

function normalizeHeader(value: string) {
  return value.trim().toLowerCase();
}

function escapeCsvValue(value: string) {
  const escaped = value.replace(/"/g, '""');
  return /[",\n\r]/.test(escaped) ? `"${escaped}"` : escaped;
}

function downloadBlob(fileName: string, data: BlobPart, type: string) {
  const blob = data instanceof Blob ? data : new Blob([data], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function safeFileName(value: string) {
  return value.trim().replace(/[\\/:*?"<>|]/g, "_") || "paneler";
}

function pageSideLabel(pageNumber: number) {
  return pageNumber % 2 === 1 ? "左ページ" : "右ページ";
}

function canvasToBlob(canvas: HTMLCanvasElement, type: string) {
  return new Promise<Blob | null>((resolve) => {
    canvas.toBlob((blob) => resolve(blob), type);
  });
}

function wrapText(ctx: CanvasRenderingContext2D, text: string, x: number, y: number, maxWidth: number, lineHeight: number) {
  let line = "";
  Array.from(text).forEach((char) => {
    const testLine = line + char;
    if (ctx.measureText(testLine).width > maxWidth && line) {
      ctx.fillText(line, x, y);
      line = char;
      y += lineHeight;
    } else {
      line = testLine;
    }
  });
  if (line) ctx.fillText(line, x, y);
}
