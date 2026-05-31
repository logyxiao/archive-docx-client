import ExcelJS from "exceljs";

export function fillAfterLabel(sheet: ExcelJS.Worksheet, labels: string[], value: string) {
  if (!value) {
    return;
  }

  for (const cell of cellsMatchingLabels(sheet, labels)) {
    const target = cell.value && String(cell.value).includes("：") && !String(cell.value).trim().endsWith("：")
      ? cell
      : cellToRightOfMergedRange(sheet, cell);
    if (target === cell) {
      target.value = replaceAfterColon(String(cell.value), value);
    } else {
      target.value = value;
    }
  }
}

export function fillAfterLabelExcept(sheet: ExcelJS.Worksheet, labels: string[], excludedLabels: string[], value: string) {
  if (!value) {
    return;
  }

  for (const cell of cellsMatchingLabelsExcept(sheet, labels, excludedLabels)) {
    const target = cell.value && String(cell.value).includes("：") && !String(cell.value).trim().endsWith("：")
      ? cell
      : cellToRightOfMergedRange(sheet, cell);
    if (target === cell) {
      target.value = replaceAfterColon(String(cell.value), value);
    } else {
      target.value = value;
    }
  }
}

export function clearAfterLabel(sheet: ExcelJS.Worksheet, labels: string[]) {
  for (const cell of cellsMatchingLabels(sheet, labels)) {
    const target = cellToRightOfMergedRange(sheet, cell);
    if (target !== cell) {
      target.value = "";
    }
  }
}

export function rowHasExactLabel(row: ExcelJS.Row, label: string): boolean {
  let hasLabel = false;
  row.eachCell({ includeEmpty: false }, (cell) => {
    if (isMergedSlave(cell)) {
      return;
    }
    if (normalizeCellText(cell) === label) {
      hasLabel = true;
    }
  });
  return hasLabel;
}

export function fillOrClearAfterLabelInRow(row: ExcelJS.Row, labels: string[], value: string) {
  const targets: ExcelJS.Cell[] = [];
  row.eachCell({ includeEmpty: false }, (cell) => {
    if (isMergedSlave(cell)) {
      return;
    }
    const text = normalizeCellText(cell);
    if (!labels.some((label) => text === label)) {
      return;
    }

    const target = cellToRightOfMergedRange(row.worksheet, cell);
    if (target !== cell) {
      targets.push(target);
    }
  });

  for (const target of targets) {
    target.value = value || "";
  }
}

export function isMergedSlave(cell: ExcelJS.Cell): boolean {
  const maybeMerged = cell as ExcelJS.Cell & { master?: ExcelJS.Cell };
  return cell.isMerged && Boolean(maybeMerged.master) && maybeMerged.master !== cell;
}

function cellsMatchingLabels(sheet: ExcelJS.Worksheet, labels: string[]): ExcelJS.Cell[] {
  const cells: ExcelJS.Cell[] = [];
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (isMergedSlave(cell)) {
        return;
      }
      const value = typeof cell.value === "string" ? cell.value.replace(/\s+/g, "") : "";
      if (labels.some((label) => value.includes(label))) {
        cells.push(cell);
      }
    });
  }
  return cells;
}

function cellsMatchingLabelsExcept(sheet: ExcelJS.Worksheet, labels: string[], excludedLabels: string[]): ExcelJS.Cell[] {
  const cells: ExcelJS.Cell[] = [];
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    row.eachCell({ includeEmpty: false }, (cell) => {
      if (isMergedSlave(cell)) {
        return;
      }
      const value = normalizeCellText(cell);
      if (excludedLabels.some((label) => value.includes(label))) {
        return;
      }
      if (labels.some((label) => value.includes(label))) {
        cells.push(cell);
      }
    });
  }
  return cells;
}

function normalizeCellText(cell: ExcelJS.Cell): string {
  return typeof cell.value === "string" ? cell.value.replace(/\s+/g, "") : "";
}

function cellToRightOfMergedRange(sheet: ExcelJS.Worksheet, cell: ExcelJS.Cell): ExcelJS.Cell {
  const range = findMergeRange(sheet, cell);
  return sheet.getCell(cell.row, (range?.right ?? Number(cell.col)) + 1);
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

function replaceAfterColon(text: string, value: string): string {
  return text.replace(/([：:]).*$/, `$1${value}`);
}
