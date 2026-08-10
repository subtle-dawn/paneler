import { faceSizeLabel, sizeLabel, shapeLabel } from "./i18n";
import { createId, createPanelRow } from "./storage";
import type { EmotionSize, FaceSize, PanelRow, Project, Shape, Size } from "./types";

type RawSheetRow = Record<string, unknown>;
const STORYBOARD_TEMPLATE_URL = "/templates/storyboard-template.xlsx";

const headerMap: Record<string, keyof PanelRow | "content"> = {
  page: "pageNumber",
  "page number": "pageNumber",
  ページ: "pageNumber",
  pagenumber: "pageNumber",
  emotion: "emotionSize",
  emotionsize: "emotionSize",
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
  facesize: "faceSize",
  "face size": "faceSize",
  顔サイズ: "faceSize",
  顔の大きさ: "faceSize",
  content: "content",
  内容: "content",
  本文: "content",
};

export function downloadJson(project: Project) {
  downloadBlob(`${safeFileName(project.title)}.json`, JSON.stringify(project, null, 2), "application/json");
}

export async function downloadStoryboardXlsx(rows: PanelRow[], title: string) {
  const table = storyboardRowsForExport(rows);
  const data = await buildStoryboardTemplateXlsx(table);
  downloadBlob(
    `${safeFileName(title)}_文字ネーム.xlsx`,
    data,
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  );
}

async function downloadStoryboardXlsxLegacy(rows: PanelRow[], title: string) {
  const xlsx = await import("xlsx");
  const sorted = rows.slice().sort((a, b) => a.pageNumber - b.pageNumber || a.order - b.order);
  const table: string[][] = [["感情", "大きさ", "役割", "形", "カメラ", "顔サイズ", "内容"]];
  let currentPage = 0;

  sorted.forEach((row) => {
    if (row.pageNumber !== currentPage) {
      currentPage = row.pageNumber;
      table.push([`${currentPage}ページ目`, pageSideLabel(currentPage), "", "", "", "", ""]);
    }

    table.push([
      row.emotionSize ? sizeLabel[row.emotionSize] : "なし",
      sizeLabel[row.panelSize],
      row.role,
      shapeLabel[row.shape],
      normalizeCameraForExport(row.camera),
      faceSizeLabel[row.faceSize],
      row.content,
    ]);
  });

  const worksheet = xlsx.utils.aoa_to_sheet(table);
  worksheet["!cols"] = [{ wch: 12 }, { wch: 12 }, { wch: 16 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 64 }];
  const workbook = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(workbook, worksheet, "文字ネーム");
  xlsx.writeFile(workbook, `${safeFileName(title)}_文字ネーム.xlsx`);
}

function storyboardRowsForExport(rows: PanelRow[]) {
  const sorted = rows.slice().sort((a, b) => a.pageNumber - b.pageNumber || a.order - b.order);
  const table: string[][] = [["感情", "大きさ", "役割", "形", "カメラ", "顔サイズ", "内容"]];
  let currentPage = 0;

  sorted.forEach((row) => {
    if (row.pageNumber !== currentPage) {
      currentPage = row.pageNumber;
      table.push([`${currentPage}ページ目`, pageSideLabel(currentPage), "", "", "", "", ""]);
    }

    table.push([
      row.emotionSize ? sizeLabel[row.emotionSize] : "なし",
      sizeLabel[row.panelSize],
      row.role,
      shapeLabel[row.shape],
      row.camera,
      faceSizeLabel[row.faceSize],
      row.content,
    ]);
  });

  return table;
}

async function loadStoryboardTemplateWorkbook(xlsx: typeof import("xlsx")) {
  const response = await fetch(STORYBOARD_TEMPLATE_URL);
  if (!response.ok) throw new Error("文字ネームテンプレートを読み込めませんでした");

  return xlsx.read(await response.arrayBuffer(), { cellStyles: true });
}

async function buildStoryboardTemplateXlsx(table: string[][]) {
  const fflate = await import("fflate");
  const response = await fetch(STORYBOARD_TEMPLATE_URL);
  if (!response.ok) throw new Error("文字ネームテンプレートを読み込めませんでした");

  const archive = fflate.unzipSync(new Uint8Array(await response.arrayBuffer()));
  const sheetPath = firstWorksheetPath(archive);
  const sheetXml = fflate.strFromU8(archive[sheetPath]);
  const nextSheetXml = replaceSheetData(sheetXml, table);

  archive[sheetPath] = fflate.strToU8(nextSheetXml);
  return fflate.zipSync(archive, { level: 6 });
}

function firstWorksheetPath(archive: Record<string, Uint8Array>) {
  const path = Object.keys(archive)
    .filter((key) => /^xl\/worksheets\/sheet\d+\.xml$/.test(key))
    .sort()[0];

  if (!path) throw new Error("文字ネームテンプレートにワークシートが見つかりませんでした");
  return path;
}

function replaceSheetData(sheetXml: string, table: string[][]) {
  const sheetData = `<sheetData>${table.map((cells, rowIndex) => storyboardXmlRow(cells, rowIndex)).join("")}</sheetData>`;
  const maxRow = Math.max(table.length, 1);
  const contentColumnWidth = calculateContentColumnWidth(table);

  const nextSheetXml = sheetXml
    .replace(/<dimension ref="[^"]*"\s*\/>/, `<dimension ref="A1:G${maxRow}"/>`)
    .replace(/<col min="7" max="7"[^>]*\/>/, `<col min="7" max="7" width="${contentColumnWidth}" style="3" customWidth="1"/>`)
    .replace(/<sheetData>[\s\S]*?<\/sheetData>/, sheetData);

  return nextSheetXml;
}

function storyboardConditionalFormattingXml() {
  return [
    `<conditionalFormatting sqref="B1:B1048576 F1:F1048576"><cfRule type="cellIs" dxfId="16" priority="1" operator="equal"><formula>"極大"</formula></cfRule><cfRule type="cellIs" dxfId="20" priority="2" operator="equal"><formula>"極小"</formula></cfRule></conditionalFormatting>`,
    `<conditionalFormatting sqref="E1:F1048576"><cfRule type="cellIs" dxfId="23" priority="3" operator="equal"><formula>"下"</formula></cfRule><cfRule type="cellIs" dxfId="23" priority="4" operator="equal"><formula>"上"</formula></cfRule><cfRule type="cellIs" dxfId="22" priority="12" operator="equal"><formula>"煽"</formula></cfRule><cfRule type="cellIs" dxfId="21" priority="13" operator="equal"><formula>"俯"</formula></cfRule></conditionalFormatting>`,
    `<conditionalFormatting sqref="C1:C1048576"><cfRule type="cellIs" dxfId="32" priority="5" operator="equal"><formula>"場、時"</formula></cfRule><cfRule type="cellIs" dxfId="31" priority="6" operator="equal"><formula>"時"</formula></cfRule><cfRule type="cellIs" dxfId="30" priority="14" operator="equal"><formula>"予"</formula></cfRule><cfRule type="containsText" dxfId="29" priority="15" operator="containsText" text="魅"><formula>NOT(ISERROR(SEARCH("魅",C1)))</formula></cfRule><cfRule type="cellIs" dxfId="28" priority="16" operator="equal"><formula>"場"</formula></cfRule></conditionalFormatting>`,
    `<conditionalFormatting sqref="D1:F1048576"><cfRule type="cellIs" dxfId="26" priority="8" operator="equal"><formula>"正"</formula></cfRule><cfRule type="cellIs" dxfId="25" priority="9" operator="equal"><formula>"横"</formula></cfRule></conditionalFormatting>`,
    `<conditionalFormatting sqref="D1:D1048576"><cfRule type="cellIs" dxfId="27" priority="10" operator="equal"><formula>"縦"</formula></cfRule></conditionalFormatting>`,
    `<conditionalFormatting sqref="A1:B1048576 F1:F1048576"><cfRule type="cellIs" dxfId="19" priority="17" operator="equal"><formula>"小"</formula></cfRule><cfRule type="cellIs" dxfId="18" priority="18" operator="equal"><formula>"中"</formula></cfRule><cfRule type="cellIs" dxfId="17" priority="19" operator="equal"><formula>"大"</formula></cfRule></conditionalFormatting>`,
    `<conditionalFormatting sqref="A1:G1048576"><cfRule type="expression" dxfId="33" priority="20"><formula>IF($G1="",TRUE)</formula></cfRule></conditionalFormatting>`,
  ].join("");
}

function calculateContentColumnWidth(table: string[][]) {
  const maxLength = Math.max(
    8,
    ...table.slice(1).map((cells) => stringDisplayWidth(cells[6] ?? "")),
  );

  return Math.min(Math.ceil(maxLength * 1.15 + 2), 120);
}

function stringDisplayWidth(value: string) {
  return Array.from(value).reduce((total, char) => total + (char.charCodeAt(0) > 0xff ? 2 : 1), 0);
}

function storyboardXmlRow(cells: string[], rowIndex: number) {
  const rowNumber = rowIndex + 1;
  const isHeader = rowIndex === 0;
  const isPageRow = rowIndex > 0 && (cells[0] ?? "").endsWith("ページ目");
  const rowAttributes = isPageRow ? ` r="${rowNumber}" spans="1:7" s="5" customFormat="1"` : ` r="${rowNumber}" spans="1:7"`;

  return `<row${rowAttributes}>${Array.from({ length: 7 }, (_, colIndex) =>
    storyboardXmlCell(cells[colIndex] ?? "", rowNumber, colIndex, isHeader, isPageRow),
  ).join("")}</row>`;
}

function storyboardXmlCell(value: string, rowNumber: number, colIndex: number, isHeader: boolean, isPageRow: boolean) {
  const address = `${String.fromCharCode(65 + colIndex)}${rowNumber}`;
  const style = isHeader ? 1 : isPageRow ? 4 : 2;
  if (!value) return `<c r="${address}" s="${style}"/>`;

  return `<c r="${address}" s="${style}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
}

function escapeXml(value: string) {
  return value
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function normalizeCameraForExport(value: string) {
  return value;
}

function writeStoryboardRowsToTemplate(
  xlsx: typeof import("xlsx"),
  worksheet: import("xlsx").WorkSheet,
  table: string[][],
) {
  const pageRowStyles = getRowStyles(worksheet, 2);
  const dataRowStyles = getRowStyles(worksheet, 3);
  const maxRows = Math.max(table.length, rangeEndRow(worksheet));

  for (let rowIndex = 1; rowIndex <= maxRows; rowIndex += 1) {
    for (let colIndex = 0; colIndex < 7; colIndex += 1) {
      const address = xlsx.utils.encode_cell({ r: rowIndex, c: colIndex });
      delete worksheet[address];
    }
  }

  table.forEach((cells, rowIndex) => {
    const isHeader = rowIndex === 0;
    const isPageRow = rowIndex > 0 && cells[0].endsWith("ページ目");

    cells.forEach((value, colIndex) => {
      const address = xlsx.utils.encode_cell({ r: rowIndex, c: colIndex });
      const existing = worksheet[address];
      const style = isHeader ? existing?.s : isPageRow ? pageRowStyles[colIndex] : dataRowStyles[colIndex];
      worksheet[address] = value ? { t: "s", v: value, s: style } : { t: "z", s: style };
    });
  });

  worksheet["!ref"] = xlsx.utils.encode_range({
    s: { r: 0, c: 0 },
    e: { r: Math.max(table.length - 1, 0), c: 6 },
  });
}

function getRowStyles(worksheet: import("xlsx").WorkSheet, rowNumber: number) {
  return Array.from({ length: 7 }, (_, colIndex) => {
    const address = String.fromCharCode(65 + colIndex) + rowNumber;
    return worksheet[address]?.s;
  });
}

function rangeEndRow(worksheet: import("xlsx").WorkSheet) {
  const range = worksheet["!ref"];
  if (!range) return 1;

  return Number(range.match(/\d+$/)?.[0] ?? 1);
}

export function downloadCsv(rows: PanelRow[], title: string) {
  const csvRows = [
    ["ページ", "感情", "大きさ", "役割", "形", "カメラ", "顔サイズ", "内容"],
    ...rows.map((row) => [
      String(row.pageNumber),
      row.emotionSize ? sizeLabel[row.emotionSize] : "なし",
      sizeLabel[row.panelSize],
      row.role,
      shapeLabel[row.shape],
      row.camera,
      faceSizeLabel[row.faceSize],
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
    const isFrameHidden = panelEl.classList.contains("frame-hidden");

    ctx.fillStyle = "#fffdfa";
    ctx.strokeStyle = "#181611";
    ctx.lineWidth = 2;
    ctx.fillRect(x, y, width, height);
    if (!isFrameHidden) strokePanelFrame(ctx, panelEl, x, y, width, height);

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

    drawFaceSizeMarker(ctx, panelEl, x, y, width, height);
  });

  return canvas;
}

function drawFaceSizeMarker(
  ctx: CanvasRenderingContext2D,
  panelEl: HTMLElement,
  x: number,
  y: number,
  width: number,
  height: number,
) {
  const faceSize = panelEl.querySelector<HTMLElement>("[data-face-size]")?.dataset.faceSize;
  const diameter = Math.min(width, height) * faceSizeMarkerRatio(faceSize);
  if (!diameter) return;

  const radius = diameter / 2;
  const centerX = x + width / 2;
  const centerY = y + height / 2;

  ctx.beginPath();
  ctx.arc(centerX, centerY, radius, 0, Math.PI * 2);
  ctx.fillStyle = "rgba(31, 62, 85, 0.12)";
  ctx.fill();
  ctx.strokeStyle = "#1f3e55";
  ctx.lineWidth = 2;
  ctx.stroke();
}

function faceSizeMarkerRatio(faceSize: string | undefined) {
  if (faceSize === "extraSmall") return 0.2;
  if (faceSize === "small") return 0.4;
  if (faceSize === "medium") return 0.6;
  if (faceSize === "large") return 0.8;
  if (faceSize === "extraLarge") return 1.2;
  return 0;
}

function strokePanelFrame(
  ctx: CanvasRenderingContext2D,
  panelEl: HTMLElement,
  x: number,
  y: number,
  width: number,
  height: number,
) {
  ctx.beginPath();
  if (!panelEl.classList.contains("bleed-edge-top")) {
    ctx.moveTo(x, y);
    ctx.lineTo(x + width, y);
  }
  if (!panelEl.classList.contains("bleed-edge-right")) {
    ctx.moveTo(x + width, y);
    ctx.lineTo(x + width, y + height);
  }
  if (!panelEl.classList.contains("bleed-edge-bottom")) {
    ctx.moveTo(x + width, y + height);
    ctx.lineTo(x, y + height);
  }
  if (!panelEl.classList.contains("bleed-edge-left")) {
    ctx.moveTo(x, y + height);
    ctx.lineTo(x, y);
  }
  ctx.stroke();
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
    const faceSize = parseFaceSize(String(normalized.faceSize ?? row.faceSize));

    return {
      ...row,
      id: createId("panel"),
      pageNumber: Number.isFinite(pageNumber) && pageNumber > 0 ? pageNumber : 1,
      emotionSize: emotionSize.value,
      panelSize: panelSize.value,
      role: String(normalized.role ?? ""),
      shape: shape.value,
      camera: String(normalized.camera ?? ""),
      faceSize: faceSize.value,
      content: String(normalized.content ?? ""),
      order,
      warnings: {
        pageNumber: Number.isFinite(pageNumber) && pageNumber > 0 ? undefined : "ページ番号を確認してください",
        emotionSize: emotionSize.warning,
        panelSize: panelSize.warning,
        shape: shape.warning,
        faceSize: faceSize.warning,
      },
    };
}

function parsePageNumber(value: string) {
  const match = value.trim().match(/\d+/);
  return match ? Number(match[0]) : NaN;
}

function parseSize(value: string): { value: Size; warning?: string } {
  const normalized = value.trim().toLowerCase();
  if (["極小", "xs", "extra small", "extrasmall"].includes(normalized)) return { value: "extraSmall" };
  if (["小", "small", "s"].includes(normalized)) return { value: "small" };
  if (["中", "medium", "m"].includes(normalized)) return { value: "medium" };
  if (["大", "large", "l"].includes(normalized)) return { value: "large" };
  if (["極大", "特大", "extra large", "extralarge", "xl"].includes(normalized)) return { value: "extraLarge" };
  if (["1ページ", "１ページ", "1page", "full page", "fullpage"].includes(normalized)) return { value: "fullPage" };
  return { value: "small", warning: "極小・小・中・大・特大・1ページから選んでください" };
}

function parseEmotion(value: string): { value: EmotionSize; warning?: string } {
  const normalized = value.trim().toLowerCase();
  if (["", "なし", "none", "null", "-"].includes(normalized)) return { value: null };
  if (["小", "small", "s"].includes(normalized)) return { value: "small" };
  if (["中", "medium", "m"].includes(normalized)) return { value: "medium" };
  if (["大", "large", "l"].includes(normalized)) return { value: "large" };
  return { value: null, warning: "小・中・大から選んでください" };
}

function parseShape(value: string): { value: Shape; warning?: string } {
  const normalized = value.trim().toLowerCase();
  if (["縦", "vertical", "v"].includes(normalized)) return { value: "vertical" };
  if (["正", "正方形", "square", "s"].includes(normalized)) return { value: "square" };
  if (["横", "horizontal", "h"].includes(normalized)) return { value: "horizontal" };
  return { value: "vertical", warning: "縦・正・横から選んでください" };
}

function parseFaceSize(value: string): { value: FaceSize; warning?: string } {
  const normalized = value.trim().toLowerCase();
  if (["", "ー", "-", "none", "null"].includes(normalized)) return { value: "none" };
  if (["極小", "xs", "extra small", "extrasmall"].includes(normalized)) return { value: "extraSmall" };
  if (["小", "small", "s"].includes(normalized)) return { value: "small" };
  if (["中", "medium", "m"].includes(normalized)) return { value: "medium" };
  if (["大", "large", "l"].includes(normalized)) return { value: "large" };
  if (["特大", "extra large", "extralarge", "xl"].includes(normalized)) return { value: "extraLarge" };
  return { value: "medium", warning: "ー・極小・小・中・大・特大から選んでください" };
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
