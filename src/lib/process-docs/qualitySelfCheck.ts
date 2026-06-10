import ExcelJS from "exceljs";
import { isMergedSlave } from "./workbookCells";

const QUALITY_STANDARD_HEADER = "质量标准";
const SELF_CHECK_HEADERS = ["施工单位自检记录", "质量验收结果", "质量检查验收结果"];
const SAMPLE_COUNT = 10;

interface ColumnRange {
  left: number;
  right: number;
}

interface NumericRange {
  min: number;
  max: number;
  decimals: number;
}

export function fillRandomSelfCheckValues(sheet: ExcelJS.Worksheet) {
  const layout = findQualitySelfCheckLayout(sheet);
  if (!layout) {
    return;
  }

  for (let rowNumber = layout.headerRow + 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
    const qualityText = rangeText(sheet, rowNumber, layout.qualityColumns);
    const range = parseNumericQualityRange(qualityText);
    if (!range || !shouldFillSelfCheckRow(sheet, rowNumber, layout.selfCheckColumns)) {
      continue;
    }

    fillSelfCheckRow(sheet, rowNumber, layout.selfCheckColumns, range);
  }
}

function findQualitySelfCheckLayout(sheet: ExcelJS.Worksheet): { headerRow: number; qualityColumns: ColumnRange; selfCheckColumns: ColumnRange } | null {
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    const qualityCells: ExcelJS.Cell[] = [];
    const selfCheckCells: ExcelJS.Cell[] = [];
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (isMergedSlave(cell)) {
        return;
      }

      const text = normalizeText(cell.value);
      if (text === QUALITY_STANDARD_HEADER) {
        qualityCells.push(cell);
      }
      if (SELF_CHECK_HEADERS.includes(text)) {
        selfCheckCells.push(cell);
      }
    });

    const qualityCell = qualityCells[0];
    const selfCheckCell = selfCheckCells[0];
    if (qualityCell && selfCheckCell) {
      return {
        headerRow: Number(row.number),
        qualityColumns: mergedColumnRange(sheet, qualityCell),
        selfCheckColumns: mergedColumnRange(sheet, selfCheckCell),
      };
    }
  }
  return null;
}

function isSelfCheckNumericLike(value: unknown): boolean {
  if (value === null || value === undefined) {
    return true;
  }
  const str = normalizeText(value).trim();
  if (str === "") {
    return true;
  }
  if (!Number.isNaN(Number(str))) {
    return true;
  }
  const parts = str.split(",");
  return parts.length > 0 && parts.every((part) => !Number.isNaN(Number(part.trim())));
}

function shouldFillSelfCheckRow(sheet: ExcelJS.Worksheet, rowNumber: number, range: ColumnRange): boolean {
  const existingValues: unknown[] = [];
  for (let column = range.left; column <= range.right; column += 1) {
    const cell = sheet.getCell(rowNumber, column);
    if (!isMergedSlave(cell) && cell.value !== null && cell.value !== undefined && cell.value !== "") {
      existingValues.push(cell.value);
    }
  }

  return existingValues.length === 0 || existingValues.every(isSelfCheckNumericLike);
}

function fillSelfCheckRow(sheet: ExcelJS.Worksheet, rowNumber: number, range: ColumnRange, numericRange: NumericRange) {
  const writableCells: ExcelJS.Cell[] = [];
  for (let column = range.left; column <= range.right; column += 1) {
    const cell = sheet.getCell(rowNumber, column);
    if (!isMergedSlave(cell)) {
      writableCells.push(cell);
    }
  }

  if (writableCells.length === 1 && range.right > range.left) {
    writableCells[0].value = distributedValuesInRange(numericRange, SAMPLE_COUNT).join(",");
    return;
  }

  const values = distributedValuesInRange(numericRange, writableCells.length);
  writableCells.forEach((cell, index) => {
    cell.value = values[index];
  });
}

function distributedValuesInRange(range: NumericRange, count: number): number[] {
  if (count <= 0) {
    return [];
  }

  const scale = 10 ** range.decimals;
  const min = Math.round(range.min * scale);
  const max = Math.round(range.max * scale);
  const span = max - min;
  if (span <= 0) {
    return Array.from({ length: count }, () => min / scale);
  }

  if (span + 1 <= count) {
    const allValues = Array.from({ length: span + 1 }, (_value, index) => min + index);
    const rest = Array.from({ length: count - allValues.length }, () => randomInteger(min, max));
    return shuffle([...allValues, ...rest]).map((value) => value / scale);
  }

  const values = Array.from({ length: count }, (_value, index) => {
    const segmentStart = min + Math.floor((span * index) / count);
    const segmentEnd = min + Math.floor((span * (index + 1)) / count);
    return randomInteger(segmentStart, Math.max(segmentStart, segmentEnd));
  });

  return shuffle(values).map((value) => value / scale);
}

export function parseNumericQualityRange(text: string): NumericRange | null {
  const normalized = text.replace(/\s+/g, "");
  const absolutePlusMinus = normalized.match(/(?:不应大于|不大于|不超过)±(-?\d+(?:\.\d+)?)/);
  if (absolutePlusMinus) {
    const limit = Number(absolutePlusMinus[1]);
    return { min: 0, max: Math.abs(limit), decimals: decimalsFor(limit) };
  }

  const plusMinus = normalized.match(/±(-?\d+(?:\.\d+)?)/);
  if (plusMinus) {
    const limit = Number(plusMinus[1]);
    return { min: -limit, max: limit, decimals: decimalsFor(limit) };
  }

  const bounded = normalized.match(/(-?\d+(?:\.\d+)?)\s*(?:~|～|至)\s*(-?\d+(?:\.\d+)?)/);
  if (bounded) {
    const min = Number(bounded[1]);
    const max = Number(bounded[2]);
    return { min: Math.min(min, max), max: Math.max(min, max), decimals: Math.max(decimalsFor(min), decimalsFor(max)) };
  }

  const upper = normalized.match(/(?:≤|<=|不大于|不超过|不应大于)(-?\d+(?:\.\d+)?)/);
  if (upper) {
    const max = Number(upper[1]);
    return { min: 0, max, decimals: decimalsFor(max) };
  }

  const strictUpper = normalized.match(/(?:<|＜|小于|应小于)(-?\d+(?:\.\d+)?)/);
  if (strictUpper) {
    const limit = Number(strictUpper[1]);
    const decimals = Math.max(decimalsFor(limit), 1);
    const step = 1 / (10 ** decimals);
    return { min: 0, max: limit - step, decimals };
  }

  const lower = normalized.match(/(?:≥|>=|不小于|不少于|不应小于)(-?\d+(?:\.\d+)?)/);
  if (lower) {
    const min = Number(lower[1]);
    return { min, max: lowerBoundMax(min), decimals: decimalsFor(min) };
  }

  const strictLower = normalized.match(/(?:>|＞|大于|应大于)(-?\d+(?:\.\d+)?)/);
  if (strictLower) {
    const limit = Number(strictLower[1]);
    const decimals = Math.max(decimalsFor(limit), 1);
    const step = 1 / (10 ** decimals);
    const min = limit + step;
    return { min, max: lowerBoundMax(limit), decimals };
  }

  return null;
}

export function qualityResultTextForStandard(text: string, count = SAMPLE_COUNT): string | null {
  const range = parseNumericQualityRange(text);
  if (!range) {
    return null;
  }

  return distributedValuesInRange(range, count).join(",");
}

function lowerBoundMax(min: number): number {
  return min + 3;
}

function randomInteger(min: number, max: number): number {
  return min + Math.floor(Math.random() * (max - min + 1));
}

function shuffle<T>(values: T[]): T[] {
  const shuffled = [...values];
  for (let index = shuffled.length - 1; index > 0; index -= 1) {
    const swapIndex = Math.floor(Math.random() * (index + 1));
    [shuffled[index], shuffled[swapIndex]] = [shuffled[swapIndex], shuffled[index]];
  }
  return shuffled;
}

function rangeText(sheet: ExcelJS.Worksheet, rowNumber: number, range: ColumnRange): string {
  const values: string[] = [];
  for (let column = range.left; column <= range.right; column += 1) {
    const cell = sheet.getCell(rowNumber, column);
    if (!isMergedSlave(cell)) {
      const text = normalizeText(cell.value);
      if (text) {
        values.push(text);
      }
    }
  }
  return values.join("");
}

function mergedColumnRange(sheet: ExcelJS.Worksheet, cell: ExcelJS.Cell): ColumnRange {
  const range = findMergeRange(sheet, cell);
  return {
    left: range?.left ?? Number(cell.col),
    right: range?.right ?? Number(cell.col),
  };
}

function findMergeRange(sheet: ExcelJS.Worksheet, cell: ExcelJS.Cell): { left: number; right: number } | null {
  const model = sheet.model as ExcelJS.WorksheetModel & { merges?: string[] };
  for (const merge of model.merges ?? []) {
    const range = decodeRange(merge);
    const row = Number(cell.row);
    const column = Number(cell.col);
    if (row >= range.top && row <= range.bottom && column >= range.left && column <= range.right) {
      return range;
    }
  }
  return null;
}

function decodeRange(range: string): { top: number; bottom: number; left: number; right: number } {
  const [start, end = start] = range.split(":");
  const startCell = decodeCellAddress(start);
  const endCell = decodeCellAddress(end);
  return {
    top: Math.min(startCell.row, endCell.row),
    bottom: Math.max(startCell.row, endCell.row),
    left: Math.min(startCell.column, endCell.column),
    right: Math.max(startCell.column, endCell.column),
  };
}

function decodeCellAddress(address: string): { row: number; column: number } {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return { row: 1, column: 1 };
  }
  const [, columnText, rowText] = match;
  let column = 0;
  for (const char of columnText) {
    column = column * 26 + char.charCodeAt(0) - 64;
  }
  return { row: Number(rowText), column };
}

function decimalsFor(value: number): number {
  return Number.isInteger(value) ? 0 : String(value).split(".")[1]?.length ?? 0;
}

function normalizeText(value: unknown): string {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "string") {
    return value.replace(/\s+/g, "");
  }
  if (typeof value === "number") {
    return String(value);
  }
  if (typeof value === "object") {
    const richTextValue = value as { richText?: Array<{ text?: string }> };
    if (Array.isArray(richTextValue.richText)) {
      return richTextValue.richText
        .map((item) => (item && typeof item.text === "string" ? item.text : ""))
        .join("")
        .replace(/\s+/g, "");
    }
  }
  return "";
}
