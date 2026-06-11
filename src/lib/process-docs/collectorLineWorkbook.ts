import ExcelJS from "exceljs";
import type { ArchiveItem } from "../types";
import { collectorLineTrialCablePair, collectorLineWorkbookReplacements } from "./collectorLine";
import type { ProcessTemplateModule } from "./types";

export function applyCollectorLineWorkbookValues(
  workbook: ExcelJS.Workbook,
  item: ArchiveItem,
  templateModule: ProcessTemplateModule,
) {
  if (templateModule !== "collector-line") {
    return;
  }

  const replacements = collectorLineWorkbookReplacements(item.title);
  if (replacements.length === 0) {
    return;
  }

  const cablePair = collectorLineTrialCablePair(item.title);
  for (const sheet of workbook.worksheets) {
    if (cablePair && item.title.includes("电缆带电试运签证")) {
      sheet.name = `${cablePair}电缆带电试运签证`.slice(0, 31);
    }

    for (let rowNumber = 1; rowNumber <= sheet.rowCount; rowNumber += 1) {
      const row = sheet.getRow(rowNumber);
      for (let columnNumber = 1; columnNumber <= sheet.columnCount; columnNumber += 1) {
        replaceCollectorLineCellValue(row.getCell(columnNumber), replacements);
      }
    }
  }
}

function replaceCollectorLineCellValue(
  cell: ExcelJS.Cell,
  replacements: ReturnType<typeof collectorLineWorkbookReplacements>,
) {
  if (typeof cell.value === "string") {
    cell.value = applyCollectorLineTextReplacements(cell.value, replacements);
    return;
  }

  if (cell.value && typeof cell.value === "object" && "richText" in cell.value && Array.isArray(cell.value.richText)) {
    const originalText = cell.value.richText.map((part) => part.text).join("");
    const nextText = applyCollectorLineTextReplacements(originalText, replacements);
    if (nextText !== originalText) {
      cell.value = nextText;
    }
  }
}

function applyCollectorLineTextReplacements(
  value: string,
  replacements: ReturnType<typeof collectorLineWorkbookReplacements>,
): string {
  return replacements.reduce((text, replacement) => text.replace(replacement.search, replacement.replacement), value);
}
