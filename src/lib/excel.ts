import * as XLSX from "xlsx";
import type { ArchiveItem, ArchiveRecord } from "./types";

const HEADER_ROW_INDEX = 1;
const DATA_START_ROW_INDEX = 3;
const HEADER_ALIASES = {
  categoryCode: ["分类号"],
  archiveCode: ["档号"],
  fullTitle: ["案卷题名"],
  sequence: ["卷内序号"],
  fileCode: ["文件编号"],
  owner: ["责任者"],
  itemTitle: ["文件题名"],
  fileDate: ["文件日期"],
  pageNo: ["页号"],
  note: ["备注"],
  retentionPeriod: ["保管期限"],
  totalPages: ["总页数"],
  filingUnit: ["立卷单位", "立卷单位（填写案卷形成单位）"],
  startDate: ["开始日期", "开始日期（卷内文件最早日期）"],
  endDate: ["结束日期", "结束日期（卷内文件最晚日期）"],
} as const;

type HeaderKey = keyof typeof HEADER_ALIASES;
type Row = Array<string | number | null | undefined>;

interface DraftRecord {
  categoryCode: string;
  archiveCode: string;
  fullTitle: string;
  owner: string;
  filingUnit: string;
  retentionPeriod: string;
  declaredTotalPages: number;
  declaredStartDate: string;
  declaredEndDate: string;
  items: ArchiveItem[];
}

export function parseArchiveWorkbook(input: ArrayBuffer | Uint8Array): ArchiveRecord[] {
  const workbook = XLSX.read(input, { type: "array", cellDates: false });
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!firstSheet) {
    throw new Error("Excel 中没有可读取的工作表");
  }

  const rows = XLSX.utils.sheet_to_json<Row>(firstSheet, {
    header: 1,
    blankrows: false,
    raw: true,
    defval: "",
  });

  const headerRow = rows[HEADER_ROW_INDEX];
  if (!headerRow) {
    throw new Error("Excel 缺少第 2 行表头");
  }

  const headers = buildHeaderMap(headerRow);
  const drafts: DraftRecord[] = [];
  let current: DraftRecord | null = null;

  for (const row of rows.slice(DATA_START_ROW_INDEX)) {
    const archiveCode = cell(row, headers.archiveCode);
    const hasNewArchive = archiveCode.length > 0;

    if (hasNewArchive) {
      current = {
        categoryCode: cell(row, headers.categoryCode),
        archiveCode,
        fullTitle: normalizeTitle(cell(row, headers.fullTitle)),
        owner: cell(row, headers.owner),
        filingUnit: cell(row, headers.filingUnit),
        retentionPeriod: cell(row, headers.retentionPeriod),
        declaredTotalPages: numberCell(row, headers.totalPages),
        declaredStartDate: normalizeDate(cell(row, headers.startDate)),
        declaredEndDate: normalizeDate(cell(row, headers.endDate)),
        items: [],
      };
      drafts.push(current);
    }

    if (!current) {
      continue;
    }

    const item = rowToItem(row, headers);
    if (hasMeaningfulItem(item)) {
      current.items.push(item);
    }
  }

  return drafts.map(finalizeRecord);
}

export function excelSerialToDate(serial: number): string {
  if (!Number.isFinite(serial) || serial <= 0) {
    return "";
  }

  const utcDays = Math.floor(serial - 25569);
  const date = new Date(utcDays * 86400 * 1000);
  return formatDateParts(
    date.getUTCFullYear(),
    date.getUTCMonth() + 1,
    date.getUTCDate(),
  );
}

export function inferTextPages(items: ArchiveItem[], declaredTotalPages = 0): number {
  return inferPages(items, declaredTotalPages);
}

export function inferDrawingPages(items: ArchiveItem[]): number {
  return inferPages(items.filter(isDrawingItem), 0);
}

export function isDrawingItem(item: ArchiveItem): boolean {
  return item.title.includes("图纸");
}

function inferPages(items: ArchiveItem[], declaredTotalPages = 0): number {
  const pageValues = items.map((item) => item.pageNo).filter(Boolean);
  const ranges = pageValues.map(parsePageRange).filter((range): range is PageRange => !!range);

  if (ranges.length === 0) {
    return declaredTotalPages;
  }

  if (ranges.every((range) => range.isPageCount)) {
    return ranges.reduce((total, range) => total + range.end, 0);
  }

  const hasSequentialSignals = ranges.some((range) => range.isRange) || isMostlySequential(ranges);
  if (hasSequentialSignals) {
    return Math.max(...ranges.map((range) => range.end));
  }

  const hasExplicitPageCount = ranges.some((range) => range.isPageCount);
  if (hasExplicitPageCount) {
    return ranges.reduce((total, range) => total + (range.isPageCount ? range.end : 0), 0);
  }

  const sum = ranges.reduce((total, range) => total + range.end, 0);
  return sum || declaredTotalPages;
}

function finalizeRecord(draft: DraftRecord): ArchiveRecord {
  const itemDates = draft.items
    .map((item) => normalizeDate(item.fileDate))
    .filter(Boolean)
    .sort();
  const startDate = itemDates[0] || draft.declaredStartDate;
  const endDate = itemDates[itemDates.length - 1] || draft.declaredEndDate || startDate;
  const drawingItems = draft.items.filter(isDrawingItem);
  const textItems = draft.items.filter((item) => !isDrawingItem(item));
  const drawingPages = inferPages(drawingItems, 0);
  const textPages = inferPages(textItems, draft.declaredTotalPages && drawingItems.length === 0 ? draft.declaredTotalPages : 0);
  const { projectName, volumeTitle } = splitArchiveTitle(draft.fullTitle);
  const filingUnit = draft.filingUnit || draft.owner || firstNonEmpty(draft.items.map((item) => item.owner));
  const retentionPeriod = draft.retentionPeriod || firstNonEmpty(draft.items.map((item) => item.note));

  return {
    categoryCode: draft.categoryCode,
    archiveCode: draft.archiveCode,
    fullTitle: draft.fullTitle,
    projectName,
    volumeTitle,
    owner: draft.owner,
    filingUnit,
    retentionPeriod,
    startDate,
    endDate,
    dateRange: startDate && endDate ? `${startDate}-${endDate}` : startDate || endDate,
    totalPages: textPages + drawingPages,
    drawingPages,
    textPages,
    items: draft.items,
  };
}

function buildHeaderMap(headerRow: Row): Record<HeaderKey, number> {
  const normalized = headerRow.map((value) => normalizeHeader(String(value ?? "")));
  const result = {} as Record<HeaderKey, number>;

  for (const key of Object.keys(HEADER_ALIASES) as HeaderKey[]) {
    const aliases = HEADER_ALIASES[key].map(normalizeHeader);
    result[key] = normalized.findIndex((header) => aliases.includes(header));
  }

  for (const required of ["archiveCode", "fullTitle", "itemTitle"] as HeaderKey[]) {
    if (result[required] < 0) {
      throw new Error(`Excel 表头缺少必要字段：${HEADER_ALIASES[required][0]}`);
    }
  }

  return result;
}

function rowToItem(row: Row, headers: Record<HeaderKey, number>): ArchiveItem {
  return {
    sequence: cell(row, headers.sequence),
    fileCode: cell(row, headers.fileCode),
    owner: cell(row, headers.owner),
    title: cell(row, headers.itemTitle),
    fileDate: normalizeDate(cell(row, headers.fileDate)),
    pageNo: cell(row, headers.pageNo),
    note: cell(row, headers.note),
  };
}

function hasMeaningfulItem(item: ArchiveItem): boolean {
  return Boolean(item.sequence || item.fileCode || item.owner || item.title || item.fileDate || item.pageNo);
}

function cell(row: Row, index: number): string {
  if (index < 0) {
    return "";
  }

  const value = row[index];
  if (value === null || value === undefined) {
    return "";
  }

  return String(value).trim();
}

function numberCell(row: Row, index: number): number {
  const value = Number(cell(row, index));
  return Number.isFinite(value) ? value : 0;
}

function normalizeHeader(value: string): string {
  return value.replace(/\s+/g, "").replace(/[()（）]/g, "");
}

function normalizeTitle(value: string): string {
  return value.replace(/\s*\n\s*/g, " ").replace(/\s+/g, " ").trim();
}

function normalizeDate(value: string): string {
  const text = String(value ?? "").trim();
  if (!text) {
    return "";
  }

  const compact = text.replace(/[./年月日-]/g, "");
  if (/^\d{8}$/.test(compact)) {
    return compact;
  }

  if (/^\d+(\.\d+)?$/.test(text)) {
    return excelSerialToDate(Number(text));
  }

  return text;
}

function formatDateParts(year: number, month: number, day: number): string {
  return `${year}${String(month).padStart(2, "0")}${String(day).padStart(2, "0")}`;
}

interface PageRange {
  start: number;
  end: number;
  isRange: boolean;
  isPageCount: boolean;
}

function parsePageRange(value: string): PageRange | null {
  const text = value.trim();
  const pageCountMatch = text.match(/页数\s*[（(]\s*(\d+)\s*[）)]/);
  if (pageCountMatch) {
    const pageCount = Number(pageCountMatch[1]);
    return { start: pageCount, end: pageCount, isRange: false, isPageCount: true };
  }

  const rangeMatch = text.match(/^(\d+)\s*[~-]\s*(\d+)$/);
  if (rangeMatch) {
    const start = Number(rangeMatch[1]);
    const end = Number(rangeMatch[2]);
    return { start, end: Math.max(start, end), isRange: true, isPageCount: false };
  }

  const singleMatch = text.match(/^(\d+)$/);
  if (singleMatch) {
    const end = Number(singleMatch[1]);
    return { start: end, end, isRange: false, isPageCount: false };
  }

  return null;
}

function isMostlySequential(ranges: PageRange[]): boolean {
  if (ranges.length < 2) {
    return false;
  }

  let sequentialPairs = 0;
  for (let index = 1; index < ranges.length; index += 1) {
    const previous = ranges[index - 1];
    const current = ranges[index];
    if (current.start === previous.end + 1 || current.end === previous.end + 1) {
      sequentialPairs += 1;
    }
  }

  return sequentialPairs >= Math.max(1, ranges.length - 2);
}

function splitArchiveTitle(fullTitle: string): { projectName: string; volumeTitle: string } {
  const normalized = normalizeTitle(fullTitle);
  const marker = "项目";
  const markerIndex = normalized.indexOf(marker);
  if (markerIndex >= 0) {
    return {
      projectName: normalized.slice(0, markerIndex + marker.length).trim(),
      volumeTitle: normalized.slice(markerIndex + marker.length).trim(),
    };
  }

  const firstSpace = normalized.indexOf(" ");
  if (firstSpace >= 0) {
    return {
      projectName: normalized.slice(0, firstSpace).trim(),
      volumeTitle: normalized.slice(firstSpace + 1).trim(),
    };
  }

  return { projectName: normalized, volumeTitle: "" };
}

function firstNonEmpty(values: string[]): string {
  return values.find((value) => value.trim()) ?? "";
}
