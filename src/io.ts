import { faceSizeLabel, sizeLabel, shapeLabel } from "./i18n";
import { createId, createPanelRow } from "./storage";
import type { EmotionSize, FaceSize, PanelRow, Project, Shape, Size } from "./types";

type RawSheetRow = Record<string, unknown>;
const STORYBOARD_TEMPLATE_URL = "/templates/storyboard-template.xlsx";
const jsonSheetName = "JSON";
const jsonChunkLength = 30000;
const exportDpi = 600;
const millimetersPerInch = 25.4;
const exportPageWidth = 5196;
const exportPageHeight = 7322;
const manuscriptPageWidthMm = 220;
const manuscriptPageHeightMm = 310;
const manuscriptInnerWidthMm = 180;
const manuscriptInnerHeightMm = 270;
const manuscriptInnerMarginMm = 20;

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

export async function downloadStoryboardXlsx(rows: PanelRow[], title: string, project?: Project) {
  const table = storyboardRowsForExport(rows);
  const data = await buildStoryboardTemplateXlsx(table, project);
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

async function buildStoryboardTemplateXlsx(table: string[][], project?: Project) {
  const fflate = await import("fflate");
  const response = await fetch(STORYBOARD_TEMPLATE_URL);
  if (!response.ok) throw new Error("文字ネームテンプレートを読み込めませんでした");

  const archive = fflate.unzipSync(new Uint8Array(await response.arrayBuffer()));
  const sheetPath = firstWorksheetPath(archive);
  const sheetXml = fflate.strFromU8(archive[sheetPath]);
  const nextSheetXml = replaceSheetData(sheetXml, table);

  archive[sheetPath] = fflate.strToU8(nextSheetXml);
  if (project) addJsonSheetToArchive(archive, JSON.stringify(project, null, 2), fflate);
  return fflate.zipSync(archive, { level: 6 });
}

function addJsonSheetToArchive(
  archive: Record<string, Uint8Array>,
  json: string,
  fflate: typeof import("fflate"),
) {
  const workbookPath = "xl/workbook.xml";
  const workbookRelsPath = "xl/_rels/workbook.xml.rels";
  const contentTypesPath = "[Content_Types].xml";
  const workbookXml = fflate.strFromU8(archive[workbookPath]);
  const workbookRelsXml = fflate.strFromU8(archive[workbookRelsPath]);
  const contentTypesXml = fflate.strFromU8(archive[contentTypesPath]);
  const existingJsonSheet = findWorkbookSheet(workbookXml, jsonSheetName);
  if (existingJsonSheet) {
    const target = findRelationshipTarget(workbookRelsXml, existingJsonSheet.relationshipId);
    if (target) {
      archive[workbookTargetPath(target)] = fflate.strToU8(jsonWorksheetXml(json));
      archive[workbookPath] = fflate.strToU8(ensureSheetHidden(workbookXml, existingJsonSheet.sheetXml));
      return;
    }
  }

  const existingSheetNames = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"/g)).map((match) =>
    unescapeXmlAttribute(match[1]),
  );
  const sheetName = uniqueSheetName(jsonSheetName, existingSheetNames);
  const nextSheetNumber = nextNumericSuffix(Object.keys(archive), /^xl\/worksheets\/sheet(\d+)\.xml$/);
  const nextSheetId = nextNumericSuffixFromXml(workbookXml, /sheetId="(\d+)"/g);
  const nextRelationshipId = uniqueRelationshipId(workbookRelsXml);

  archive[`xl/worksheets/sheet${nextSheetNumber}.xml`] = fflate.strToU8(jsonWorksheetXml(json));
  archive[workbookPath] = fflate.strToU8(
    workbookXml.replace(
      "</sheets>",
      `<sheet name="${escapeXmlAttribute(sheetName)}" sheetId="${nextSheetId}" state="hidden" r:id="${nextRelationshipId}"/></sheets>`,
    ),
  );
  archive[workbookRelsPath] = fflate.strToU8(
    workbookRelsXml.replace(
      "</Relationships>",
      `<Relationship Id="${nextRelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${nextSheetNumber}.xml"/></Relationships>`,
    ),
  );
  archive[contentTypesPath] = fflate.strToU8(
    contentTypesXml.replace(
      "</Types>",
      `<Override PartName="/xl/worksheets/sheet${nextSheetNumber}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
    ),
  );
}

function findWorkbookSheet(workbookXml: string, sheetName: string) {
  const sheetXmls = Array.from(workbookXml.matchAll(/<sheet\b[^>]*\/>/g)).map((match) => match[0]);
  const sheetXml = sheetXmls.find((xml) => normalizeSheetName(unescapeXmlAttribute(xmlAttribute(xml, "name") ?? "")) === normalizeSheetName(sheetName));
  const relationshipId = sheetXml ? xmlAttribute(sheetXml, "r:id") : undefined;
  return sheetXml && relationshipId ? { sheetXml, relationshipId } : undefined;
}

function findRelationshipTarget(workbookRelsXml: string, relationshipId: string) {
  const relationshipXmls = Array.from(workbookRelsXml.matchAll(/<Relationship\b[^>]*\/>/g)).map((match) => match[0]);
  const relationshipXml = relationshipXmls.find((xml) => xmlAttribute(xml, "Id") === relationshipId);
  return relationshipXml ? xmlAttribute(relationshipXml, "Target") : undefined;
}

function workbookTargetPath(target: string) {
  if (target.startsWith("/xl/")) return target.slice(1);
  if (target.startsWith("xl/")) return target;
  return `xl/${target.replace(/^\.\//, "")}`;
}

function ensureSheetHidden(workbookXml: string, sheetXml: string) {
  const hiddenSheetXml = /\bstate=/.test(sheetXml)
    ? sheetXml.replace(/\bstate="[^"]*"/, `state="hidden"`)
    : sheetXml.replace(/\s*\/>$/, ` state="hidden"/>`);
  return workbookXml.replace(sheetXml, hiddenSheetXml);
}

function xmlAttribute(xml: string, name: string) {
  return xml.match(new RegExp(`\\b${name}="([^"]*)"`))?.[1];
}

function jsonWorksheetXml(json: string) {
  const chunks = json.match(new RegExp(`[\\s\\S]{1,${jsonChunkLength}}`, "g")) ?? [""];
  const rows = chunks
    .map((chunk, index) => {
      const rowNumber = index + 1;
      return `<row r="${rowNumber}"><c r="A${rowNumber}" t="inlineStr"><is><t>${escapeXml(chunk)}</t></is></c></row>`;
    })
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:A${chunks.length}"/><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><cols><col min="1" max="1" width="120" customWidth="1"/></cols><sheetData>${rows}</sheetData></worksheet>`;
}

function nextNumericSuffix(paths: string[], pattern: RegExp) {
  const max = paths.reduce((found, path) => Math.max(found, Number(path.match(pattern)?.[1] ?? 0)), 0);
  return max + 1;
}

function nextNumericSuffixFromXml(xml: string, pattern: RegExp) {
  const max = Array.from(xml.matchAll(pattern)).reduce((found, match) => Math.max(found, Number(match[1])), 0);
  return max + 1;
}

function uniqueRelationshipId(workbookRelsXml: string) {
  const used = new Set(Array.from(workbookRelsXml.matchAll(/Id="([^"]+)"/g)).map((match) => match[1]));
  let index = nextNumericSuffixFromXml(workbookRelsXml, /Id="rId(\d+)"/g);
  while (used.has(`rId${index}`)) index += 1;
  return `rId${index}`;
}

function uniqueSheetName(baseName: string, existingNames: string[]) {
  if (!existingNames.includes(baseName)) return baseName;

  let index = 2;
  while (existingNames.includes(`${baseName}${index}`)) index += 1;
  return `${baseName}${index}`;
}

function escapeXmlAttribute(value: string) {
  return escapeXml(value).replace(/'/g, "&apos;");
}

function unescapeXmlAttribute(value: string) {
  return value
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, '"')
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
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
    `<conditionalFormatting sqref="B1:B1048576 F1:F1048576"><cfRule type="cellIs" dxfId="16" priority="1" operator="equal"><formula>"特大"</formula></cfRule><cfRule type="cellIs" dxfId="20" priority="2" operator="equal"><formula>"極小"</formula></cfRule></conditionalFormatting>`,
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

export async function readSheetProject(file: File): Promise<Project | null> {
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (ext !== "xlsx" && ext !== "xls") return null;

  const xlsx = await import("xlsx");
  const data = await file.arrayBuffer();
  const workbook = xlsx.read(data);
  const sheetName = workbook.SheetNames.find((name) => normalizeSheetName(name) === normalizeSheetName(jsonSheetName));
  if (!sheetName) return null;

  const table = xlsx.utils.sheet_to_json<string[]>(workbook.Sheets[sheetName], { header: 1, defval: "", raw: false });
  const json = table.flat().join("").trim();
  if (!json) return null;

  return JSON.parse(json) as Project;
}

export async function readSheetRows(file: File): Promise<PanelRow[]> {
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (ext === "csv" || ext === "tsv" || ext === "txt") {
    return rowsFromTable(parseCsv((await file.text()).replace(/^\uFEFF/, ""), ext === "tsv" ? "\t" : ","));
  }

  const xlsx = await import("xlsx");
  const data = await file.arrayBuffer();
  const workbook = xlsx.read(data);
  const firstSheetName = workbook.SheetNames.find((name) => normalizeSheetName(name) !== normalizeSheetName(jsonSheetName));
  if (!firstSheetName) return [];

  const firstSheet = workbook.Sheets[firstSheetName];
  const table = xlsx.utils.sheet_to_json<string[]>(firstSheet, { header: 1, defval: "", raw: false });
  return rowsFromTable(table);
}

function normalizeSheetName(value: string) {
  return value.trim().toLowerCase();
}

export async function exportPagePng(pageEl: HTMLElement, fileName: string) {
  const canvas = renderPageToCanvas(pageEl);
  const blob = await canvasToPngBlob(canvas);
  if (blob) downloadBlob(fileName, blob, "image/png");
}

export async function exportPagesPngZip(pages: { pageEl: HTMLElement; fileName: string }[], zipFileName: string) {
  if (pages.length === 0) return;

  const fflate = await import("fflate");
  const archive: Record<string, Uint8Array> = {};

  for (const page of pages) {
    const canvas = renderPageToCanvas(page.pageEl);
    const blob = await canvasToPngBlob(canvas);
    if (!blob) continue;
    archive[page.fileName] = new Uint8Array(await blob.arrayBuffer());
  }

  if (Object.keys(archive).length === 0) return;
  downloadBlob(zipFileName, fflate.zipSync(archive, { level: 6 }), "application/zip");
}

export async function exportPagesPdf(pageEls: HTMLElement[], fileName: string) {
  if (pageEls.length === 0) return;

  const { jsPDF } = await import("jspdf");
  const canvases = pageEls.map(renderPageToCanvas);
  const firstCanvas = canvases[0];
  const pdfPageWidth = pixelsToMillimeters(firstCanvas.width);
  const pdfPageHeight = pixelsToMillimeters(firstCanvas.height);
  const pdf = new jsPDF({
    orientation: firstCanvas.width >= firstCanvas.height ? "landscape" : "portrait",
    unit: "mm",
    format: [pdfPageWidth, pdfPageHeight],
  });

  canvases.forEach((canvas, index) => {
    if (index > 0) pdf.addPage();
    const image = canvas.toDataURL("image/jpeg", 0.92);
    pdf.addImage(image, "JPEG", 0, 0, pdfPageWidth, pdfPageHeight);
  });

  pdf.save(fileName);
}

function renderPageToCanvas(pageEl: HTMLElement) {
  const rect = pageEl.getBoundingClientRect();
  const pageStyle = getComputedStyle(pageEl);
  const paddingLeft = parseFloat(pageStyle.paddingLeft) || 0;
  const paddingRight = parseFloat(pageStyle.paddingRight) || 0;
  const paddingTop = parseFloat(pageStyle.paddingTop) || 0;
  const paddingBottom = parseFloat(pageStyle.paddingBottom) || 0;
  const contentLeft = rect.left + paddingLeft;
  const contentTop = rect.top + paddingTop;
  const contentWidth = Math.max(rect.width - paddingLeft - paddingRight, 1);
  const contentHeight = Math.max(rect.height - paddingTop - paddingBottom, 1);
  const exportInnerX = exportPageWidth * (manuscriptInnerMarginMm / manuscriptPageWidthMm);
  const exportInnerY = exportPageHeight * (manuscriptInnerMarginMm / manuscriptPageHeightMm);
  const exportInnerWidth = exportPageWidth * (manuscriptInnerWidthMm / manuscriptPageWidthMm);
  const exportInnerHeight = exportPageHeight * (manuscriptInnerHeightMm / manuscriptPageHeightMm);
  const scaleX = exportInnerWidth / contentWidth;
  const scaleY = exportInnerHeight / contentHeight;
  const offsetX = exportInnerX / scaleX;
  const offsetY = exportInnerY / scaleY;
  const exportLogicalWidth = exportPageWidth / scaleX;
  const exportLogicalHeight = exportPageHeight / scaleY;
  const canvas = document.createElement("canvas");
  canvas.width = exportPageWidth;
  canvas.height = exportPageHeight;
  const ctx = canvas.getContext("2d");
  if (!ctx) return canvas;

  ctx.scale(scaleX, scaleY);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, canvas.width / scaleX, canvas.height / scaleY);

  pageEl.querySelectorAll<HTMLElement>(".panel-frame").forEach((panelEl) => {
    const bounds = panelEl.getBoundingClientRect();
    const x = bounds.left - contentLeft + offsetX;
    const y = bounds.top - contentTop + offsetY;
    const width = bounds.width;
    const height = bounds.height;
    const isFrameHidden = panelEl.classList.contains("frame-hidden");

    ctx.fillStyle = "#ffffff";
    ctx.strokeStyle = "#181611";
    ctx.lineWidth = 2;
    ctx.fillRect(x, y, width, height);
    if (!isFrameHidden) {
      strokePanelFrame(ctx, panelEl, x, y, width, height, exportLogicalWidth, exportLogicalHeight);
    }

    drawExportPanelText(ctx, panelEl, x, y, width, height);

    drawFaceSizeMarker(ctx, panelEl, x, y, width, height);
  });

  drawPageSideMark(ctx, pageEl, 0, 0, canvas.width / scaleX, canvas.height / scaleY);

  return canvas;
}

function drawExportPanelText(
  ctx: CanvasRenderingContext2D,
  panelEl: HTMLElement,
  x: number,
  y: number,
  width: number,
  height: number,
) {
  const padding = 10;
  const fields = Array.from(panelEl.querySelectorAll<HTMLElement>(".panel-fields span")).map((node) =>
    node.innerText.trim(),
  );
  const content = panelEl.querySelector<HTMLElement>("p[data-export-text]")?.innerText.trim() ?? "";
  const textX = x + padding;
  const textY = y + padding;
  const textWidth = Math.max(width - padding * 2, 1);
  const textHeight = Math.max(height - padding * 2, 1);
  const headerBaseline = textY + 18;

  ctx.save();
  ctx.beginPath();
  ctx.rect(textX, textY, textWidth, textHeight);
  ctx.clip();

  ctx.fillStyle = "#181611";
  ctx.font = "12px sans-serif";

  fields.forEach((field, index) => {
    if (!field) return;
    ctx.font = index === 0 ? "bold 13px sans-serif" : "12px sans-serif";
    if (index === 0) {
      ctx.fillText(field, textX, headerBaseline);
      return;
    }

    const rightAlignedFields = fields.slice(1).filter(Boolean);
    const fieldIndex = index - 1;
    const fieldGap = 8;
    const fieldWidths = rightAlignedFields.map((value) => ctx.measureText(value).width);
    const totalWidth = fieldWidths.reduce((total, value) => total + value, 0) + fieldGap * Math.max(0, rightAlignedFields.length - 1);
    const fieldX =
      x +
      width -
      padding -
      totalWidth +
      fieldWidths.slice(0, fieldIndex).reduce((total, value) => total + value + fieldGap, 0);

    ctx.fillText(field, fieldX, headerBaseline);
  });

  if (content) {
    ctx.font = "12px sans-serif";
    const lineHeight = 16;
    const contentMetrics = ctx.measureText("あg");
    const contentDescent = Math.ceil(contentMetrics.actualBoundingBoxDescent || 3);
    const maxContentLines = Math.max(Math.floor((textHeight - 28 - contentDescent) / lineHeight), 1);
    const contentLines = wrapTextLines(ctx, content, textWidth).slice(0, maxContentLines);
    const contentY = textY + textHeight - contentDescent - (contentLines.length - 1) * lineHeight;
    contentLines.forEach((line, index) => {
      ctx.fillText(line, textX, contentY + index * lineHeight);
    });
  }

  ctx.restore();
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
  pageWidth: number,
  pageHeight: number,
) {
  const hasTopBleed = panelEl.classList.contains("bleed-edge-top");
  const hasRightBleed = panelEl.classList.contains("bleed-edge-right");
  const hasBottomBleed = panelEl.classList.contains("bleed-edge-bottom");
  const hasLeftBleed = panelEl.classList.contains("bleed-edge-left");
  const right = x + width;
  const bottom = y + height;

  ctx.beginPath();
  if (!hasTopBleed) {
    ctx.moveTo(x, y);
    ctx.lineTo(right, y);
  }
  if (!hasRightBleed) {
    ctx.moveTo(right, y);
    ctx.lineTo(right, bottom);
  }
  if (!hasBottomBleed) {
    ctx.moveTo(right, bottom);
    ctx.lineTo(x, bottom);
  }
  if (!hasLeftBleed) {
    ctx.moveTo(x, bottom);
    ctx.lineTo(x, y);
  }

  if (hasTopBleed) {
    if (!hasLeftBleed) {
      ctx.moveTo(x, 0);
      ctx.lineTo(x, y);
    }
    if (!hasRightBleed) {
      ctx.moveTo(right, 0);
      ctx.lineTo(right, y);
    }
  }
  if (hasRightBleed) {
    if (!hasTopBleed) {
      ctx.moveTo(right, y);
      ctx.lineTo(pageWidth, y);
    }
    if (!hasBottomBleed) {
      ctx.moveTo(right, bottom);
      ctx.lineTo(pageWidth, bottom);
    }
  }
  if (hasBottomBleed) {
    if (!hasRightBleed) {
      ctx.moveTo(right, bottom);
      ctx.lineTo(right, pageHeight);
    }
    if (!hasLeftBleed) {
      ctx.moveTo(x, bottom);
      ctx.lineTo(x, pageHeight);
    }
  }
  if (hasLeftBleed) {
    if (!hasBottomBleed) {
      ctx.moveTo(0, bottom);
      ctx.lineTo(x, bottom);
    }
    if (!hasTopBleed) {
      ctx.moveTo(0, y);
      ctx.lineTo(x, y);
    }
  }

  ctx.stroke();
}

function drawPageSideMark(
  ctx: CanvasRenderingContext2D,
  pageEl: HTMLElement,
  offsetX: number,
  offsetY: number,
  width: number,
  height: number,
) {
  const isRight = pageEl.classList.contains("side-mark-right");
  const centerX = offsetX + (isRight ? width - 20 : 20);
  const centerY = offsetY + height / 2;
  const radius = 13;

  ctx.save();
  ctx.strokeStyle = "#181611";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(centerX - radius, centerY - radius);
  ctx.lineTo(centerX + radius, centerY + radius);
  ctx.moveTo(centerX + radius, centerY - radius);
  ctx.lineTo(centerX - radius, centerY + radius);
  ctx.stroke();
  ctx.restore();
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

async function canvasToPngBlob(canvas: HTMLCanvasElement) {
  const blob = await canvasToBlob(canvas, "image/png");
  if (!blob) return null;
  return new Blob([addPngDpiMetadata(new Uint8Array(await blob.arrayBuffer()), exportDpi)], { type: "image/png" });
}

function pixelsToMillimeters(pixels: number) {
  return (pixels / exportDpi) * millimetersPerInch;
}

function addPngDpiMetadata(png: Uint8Array, dpi: number) {
  const pixelsPerMeter = Math.round(dpi / 0.0254);
  const chunkData = new Uint8Array(9);
  writeUint32(chunkData, 0, pixelsPerMeter);
  writeUint32(chunkData, 4, pixelsPerMeter);
  chunkData[8] = 1;

  const chunk = createPngChunk("pHYs", chunkData);
  const withoutPhys = removePngChunk(png, "pHYs");
  const output = new Uint8Array(withoutPhys.length + chunk.length);
  output.set(withoutPhys.slice(0, 33), 0);
  output.set(chunk, 33);
  output.set(withoutPhys.slice(33), 33 + chunk.length);
  return output;
}

function removePngChunk(png: Uint8Array, chunkType: string) {
  const chunks: Uint8Array[] = [png.slice(0, 8)];
  let offset = 8;

  while (offset < png.length) {
    const length = readUint32(png, offset);
    const type = bytesToString(png.slice(offset + 4, offset + 8));
    const nextOffset = offset + 12 + length;
    if (type !== chunkType) chunks.push(png.slice(offset, nextOffset));
    offset = nextOffset;
  }

  const totalLength = chunks.reduce((total, chunk) => total + chunk.length, 0);
  const output = new Uint8Array(totalLength);
  let cursor = 0;
  chunks.forEach((chunk) => {
    output.set(chunk, cursor);
    cursor += chunk.length;
  });
  return output;
}

function createPngChunk(type: string, data: Uint8Array) {
  const typeBytes = stringToBytes(type);
  const chunk = new Uint8Array(12 + data.length);
  writeUint32(chunk, 0, data.length);
  chunk.set(typeBytes, 4);
  chunk.set(data, 8);
  writeUint32(chunk, 8 + data.length, crc32(chunk.slice(4, 8 + data.length)));
  return chunk;
}

function readUint32(bytes: Uint8Array, offset: number) {
  return ((bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3]) >>> 0;
}

function writeUint32(bytes: Uint8Array, offset: number, value: number) {
  bytes[offset] = (value >>> 24) & 0xff;
  bytes[offset + 1] = (value >>> 16) & 0xff;
  bytes[offset + 2] = (value >>> 8) & 0xff;
  bytes[offset + 3] = value & 0xff;
}

function stringToBytes(value: string) {
  return Uint8Array.from(Array.from(value).map((char) => char.charCodeAt(0)));
}

function bytesToString(bytes: Uint8Array) {
  return String.fromCharCode(...bytes);
}

function crc32(bytes: Uint8Array) {
  let crc = 0xffffffff;
  bytes.forEach((byte) => {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  });
  return (crc ^ 0xffffffff) >>> 0;
}

function wrapTextLines(ctx: CanvasRenderingContext2D, text: string, maxWidth: number) {
  const lines: string[] = [];
  let line = "";

  Array.from(text).forEach((char) => {
    const testLine = line + char;
    if (ctx.measureText(testLine).width > maxWidth && line) {
      lines.push(line);
      line = char;
    } else {
      line = testLine;
    }
  });

  if (line) lines.push(line);
  return lines;
}
