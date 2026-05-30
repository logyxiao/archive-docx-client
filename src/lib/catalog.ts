import ExcelJS from "exceljs";
import type { ArchiveItem, ArchiveRecord, GeneratedFile } from "./types";

export const CATALOG_OUTPUT_NAME = "1.2、案卷目录、卷内目录-著录台账（及打印模板）.xlsx";

const CATALOG_TEMPLATE_PATH = "/templates/catalog.xlsx";
const CATALOG_DATA_START_ROW = 3;
const DETAIL_DATA_START_ROW = 4;
const DETAIL_ROWS_PER_SHEET = 14;
const DETAIL_LAST_PRINT_ROW = DETAIL_DATA_START_ROW + DETAIL_ROWS_PER_SHEET - 1;
const DETAIL_DATA_ROW_HEIGHT = 44;
const CATALOG_COLUMNS = 7;

interface CatalogGenerationOptions {
  selectedCodes: string[];
  outputDir: string;
}

export async function loadCatalogTemplate(): Promise<ArrayBuffer> {
  const response = await fetch(CATALOG_TEMPLATE_PATH);
  if (!response.ok) {
    throw new Error(`无法加载模板：${CATALOG_TEMPLATE_PATH}`);
  }

  return response.arrayBuffer();
}

export async function generateArchiveCatalog(
  records: ArchiveRecord[],
  options: CatalogGenerationOptions,
  writeFile: (path: string, bytes: Uint8Array) => Promise<void>,
): Promise<GeneratedFile> {
  const selected = records.filter((record) => options.selectedCodes.includes(record.archiveCode));
  if (selected.length === 0) {
    throw new Error("请先勾选需要生成台账的案卷");
  }

  const bytes = await renderCatalogWorkbook(await loadCatalogTemplate(), selected);
  const path = joinPath(options.outputDir, CATALOG_OUTPUT_NAME);
  await writeFile(path, bytes);
  return { name: CATALOG_OUTPUT_NAME, path };
}

export async function renderCatalogWorkbook(template: ArrayBuffer | Uint8Array, records: ArchiveRecord[]): Promise<Uint8Array> {
  const templateWorkbook = new ExcelJS.Workbook();
  await templateWorkbook.xlsx.load(toArrayBuffer(template));

  const catalogTemplate = templateWorkbook.getWorksheet("案卷目录") ?? templateWorkbook.worksheets[0];
  const detailTemplate = templateWorkbook.worksheets.find((sheet) => sheet.name !== "案卷目录");
  if (!catalogTemplate || !detailTemplate) {
    throw new Error("台账模板缺少案卷目录或卷内目录样式工作表");
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = templateWorkbook.creator || "archive-docx-client";
  workbook.created = new Date();
  workbook.modified = new Date();

  const catalogSheet = workbook.addWorksheet("案卷目录");
  applyWorksheetTemplate(catalogTemplate, catalogSheet);
  fillCatalogSheet(catalogSheet, catalogTemplate, records);

  const usedSheetNames = new Set(["案卷目录"]);
  for (const record of records) {
    const groups = chunk(record.items.length > 0 ? record.items : [emptyItem()], DETAIL_ROWS_PER_SHEET);
    for (const [groupIndex, items] of groups.entries()) {
      const baseName = groupIndex === 0 ? record.archiveCode : `${record.archiveCode} (${groupIndex + 1})`;
      const sheet = workbook.addWorksheet(uniqueSheetName(baseName, usedSheetNames));
      applyWorksheetTemplate(detailTemplate, sheet);
      fillDetailSheet(sheet, detailTemplate, record, items, groupIndex);
    }
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(buffer);
}

function fillCatalogSheet(sheet: ExcelJS.Worksheet, template: ExcelJS.Worksheet, records: ArchiveRecord[]) {
  copyStaticRows(template, sheet, 1, 2);
  const rowCount = Math.max(template.rowCount, CATALOG_DATA_START_ROW + records.length - 1);

  for (let rowNumber = CATALOG_DATA_START_ROW; rowNumber <= rowCount; rowNumber += 1) {
    copyRowStyle(template, sheet, CATALOG_DATA_START_ROW, rowNumber, CATALOG_COLUMNS);
    const record = records[rowNumber - CATALOG_DATA_START_ROW];
    const values = record
      ? [
          rowNumber - CATALOG_DATA_START_ROW + 1,
          record.archiveCode,
          record.fullTitle,
          record.totalPages,
          record.retentionPeriod,
          record.filingUnit,
          "",
        ]
      : ["", "", "", "", "", "", ""];
    setRowValues(sheet, rowNumber, values);
  }

  sheet.pageSetup.printTitlesRow = "1:2";
}

function fillDetailSheet(
  sheet: ExcelJS.Worksheet,
  template: ExcelJS.Worksheet,
  record: ArchiveRecord,
  items: ArchiveItem[],
  groupIndex: number,
) {
  copyStaticRows(template, sheet, 1, 3);
  sheet.getCell("A2").value = ` 档号：${record.archiveCode}`;
  sheet.getCell("F3").value = items.some((item) => item.note === "页数") ? "页数" : "页号";

  for (let offset = 0; offset < DETAIL_ROWS_PER_SHEET; offset += 1) {
    const rowNumber = DETAIL_DATA_START_ROW + offset;
    const item = items[offset];
    copyRowStyle(template, sheet, rowNumber, rowNumber, CATALOG_COLUMNS);
    sheet.getRow(rowNumber).height = DETAIL_DATA_ROW_HEIGHT;
    const sequence = groupIndex * DETAIL_ROWS_PER_SHEET + offset + 1;
    const values = item && hasMeaningfulItem(item)
      ? [
          item.sequence || sequence,
          item.fileCode,
          item.owner,
          item.title,
          item.fileDate,
          item.pageNo,
          item.note === "页数" ? "" : item.note,
        ]
      : ["", "", "", "", "", "", ""];
    setRowValues(sheet, rowNumber, values);
  }

  sheet.pageSetup.printTitlesRow = "1:3";
  sheet.pageSetup.fitToPage = true;
  sheet.pageSetup.fitToWidth = 1;
  sheet.pageSetup.fitToHeight = 1;
  sheet.pageSetup.scale = undefined;
  sheet.pageSetup.printArea = `A1:G${DETAIL_LAST_PRINT_ROW}`;
}

function applyWorksheetTemplate(source: ExcelJS.Worksheet, target: ExcelJS.Worksheet) {
  target.properties = clonePlain(source.properties);
  target.pageSetup = clonePlain(source.pageSetup);
  (target as ExcelJS.Worksheet & { pageMargins?: unknown }).pageMargins = clonePlain(
    (source as ExcelJS.Worksheet & { pageMargins?: unknown }).pageMargins,
  );
  target.headerFooter = clonePlain(source.headerFooter);
  target.views = clonePlain(source.views);

  for (let index = 1; index <= Math.max(source.columnCount, CATALOG_COLUMNS); index += 1) {
    const sourceColumn = source.getColumn(index);
    const targetColumn = target.getColumn(index);
    targetColumn.width = sourceColumn.width;
    targetColumn.hidden = sourceColumn.hidden;
    targetColumn.outlineLevel = sourceColumn.outlineLevel;
    targetColumn.style = clonePlain(sourceColumn.style);
  }

  for (let rowNumber = 1; rowNumber <= source.rowCount; rowNumber += 1) {
    copyRowStyle(source, target, rowNumber, rowNumber, Math.max(source.columnCount, CATALOG_COLUMNS));
  }
}

function copyStaticRows(source: ExcelJS.Worksheet, target: ExcelJS.Worksheet, start: number, end: number) {
  for (let rowNumber = start; rowNumber <= end; rowNumber += 1) {
    copyRowStyle(source, target, rowNumber, rowNumber, CATALOG_COLUMNS);
    const values = Array.from({ length: CATALOG_COLUMNS }, (_, index) => source.getCell(rowNumber, index + 1).value);
    setRowValues(target, rowNumber, values);
  }
}

function copyRowStyle(
  source: ExcelJS.Worksheet,
  target: ExcelJS.Worksheet,
  sourceRowNumber: number,
  targetRowNumber: number,
  columnCount: number,
) {
  const sourceRow = source.getRow(sourceRowNumber);
  const targetRow = target.getRow(targetRowNumber);
  targetRow.height = sourceRow.height;
  targetRow.hidden = sourceRow.hidden;
  targetRow.outlineLevel = sourceRow.outlineLevel;
  (targetRow as ExcelJS.Row & { style?: unknown }).style = clonePlain((sourceRow as ExcelJS.Row & { style?: unknown }).style);

  for (let columnNumber = 1; columnNumber <= columnCount; columnNumber += 1) {
    const sourceCell = source.getCell(sourceRowNumber, columnNumber);
    const targetCell = target.getCell(targetRowNumber, columnNumber);
    targetCell.style = clonePlain(sourceCell.style);
    targetCell.numFmt = sourceCell.numFmt;
    targetCell.alignment = clonePlain(sourceCell.alignment);
    targetCell.border = clonePlain(sourceCell.border);
    targetCell.fill = clonePlain(sourceCell.fill);
    targetCell.font = clonePlain(sourceCell.font);
  }
}

function setRowValues(sheet: ExcelJS.Worksheet, rowNumber: number, values: Array<ExcelJS.CellValue | string | number>) {
  for (const [index, value] of values.entries()) {
    sheet.getCell(rowNumber, index + 1).value = value === "/" ? "/" : value;
  }
}

function hasMeaningfulItem(item: ArchiveItem): boolean {
  return Boolean(item.sequence || item.fileCode || item.owner || item.title || item.fileDate || item.pageNo || item.note);
}

function emptyItem(): ArchiveItem {
  return {
    sequence: "",
    fileCode: "",
    owner: "",
    title: "",
    fileDate: "",
    pageNo: "",
    note: "",
  };
}

function uniqueSheetName(baseName: string, used: Set<string>): string {
  const normalized = sanitizeSheetName(baseName);
  let sheetName = normalized.slice(0, 31);
  let suffix = 2;

  while (used.has(sheetName)) {
    const suffixText = ` (${suffix})`;
    sheetName = `${normalized.slice(0, 31 - suffixText.length)}${suffixText}`;
    suffix += 1;
  }

  used.add(sheetName);
  return sheetName;
}

function sanitizeSheetName(name: string): string {
  const sanitized = name.replace(/[:\\/?*\[\]]/g, " ").trim();
  return sanitized || "Sheet";
}

function chunk<T>(items: T[], size: number): T[][] {
  const groups: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    groups.push(items.slice(index, index + size));
  }
  return groups;
}

function clonePlain<T>(value: T): T {
  if (value === undefined || value === null) {
    return value;
  }

  return JSON.parse(JSON.stringify(value)) as T;
}

function toArrayBuffer(input: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (input instanceof ArrayBuffer) {
    return input;
  }

  return input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength);
}

function joinPath(dir: string, fileName: string): string {
  const separator = dir.includes("\\") ? "\\" : "/";
  return `${dir.replace(/[\\/]$/, "")}${separator}${fileName}`;
}
