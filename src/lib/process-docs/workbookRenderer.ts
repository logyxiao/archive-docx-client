import ExcelJS from "exceljs";
import type { ArchiveItem, ArchiveRecord } from "../types";
import { PRESERVED_USER_FIELD_LABELS, SOURCE_MISSING_LABELS } from "./constants";
import { resolveProcessFields } from "./fields";
import { replaceBusinessText } from "./textReplacement";
import type { ProcessUserFields, ResolvedProcessFields } from "./types";
import { archiveProjectCode, formatChineseDate, toArrayBuffer } from "./utils";
import { clearAfterLabel, fillAfterLabel, fillOrClearAfterLabelInRow, isMergedSlave, rowHasExactLabel } from "./workbookCells";
import { fillRandomSelfCheckValues } from "./qualitySelfCheck";

export async function renderProcessWorkbook(
  template: ArrayBuffer | Uint8Array,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields = {},
): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(toArrayBuffer(template));

  for (const sheet of workbook.worksheets) {
    applyWorkbookValues(sheet, record, item, userFields);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(buffer);
}

function applyWorkbookValues(sheet: ExcelJS.Worksheet, record: ArchiveRecord, item: ArchiveItem, userFields: ProcessUserFields) {
  const fields = resolveProcessFields(userFields, item.owner || record.filingUnit);
  const projectName = userFields.projectName?.trim() || record.projectName;
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (isMergedSlave(cell)) {
        return;
      }
      if (typeof cell.value === "string") {
        cell.value = replaceBusinessText(cell.value, record, item, userFields);
      }
    });
  }

  fillAfterLabel(sheet, ["工程名称", "工程项目名称"], projectName);
  fillAfterLabel(sheet, ["工程编号"], archiveProjectCode(record.archiveCode));
  fillAfterLabel(sheet, ["抽样单位"], item.owner || record.filingUnit);
  fillAfterLabel(sheet, ["抽样日期"], formatChineseDate(item.fileDate));
  fillUnitScopedFields(sheet, fields);
  fillOrClearProfessionalForeman(sheet, fields.constructionTechnicalLeader);
  fillRandomSelfCheckValues(sheet);
  clearAfterLabel(
    sheet,
    SOURCE_MISSING_LABELS.filter((label) => !PRESERVED_USER_FIELD_LABELS.includes(label)),
  );
}

function fillUnitScopedFields(sheet: ExcelJS.Worksheet, fields: ResolvedProcessFields) {
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    if (rowHasExactLabel(row, "总承包单位")) {
      fillOrClearAfterLabelInRow(row, ["总承包单位"], fields.generalContractorUnit);
      fillOrClearAfterLabelInRow(row, ["项目负责人"], fields.generalContractorProjectManager);
      fillOrClearAfterLabelInRow(row, ["项目技术负责人"], fields.generalContractorTechnicalLeader);
    }
    if (rowHasExactLabel(row, "施工单位")) {
      fillOrClearAfterLabelInRow(row, ["施工单位"], fields.constructionUnit);
      fillOrClearAfterLabelInRow(row, ["项目负责人"], fields.constructionProjectManager);
      fillOrClearAfterLabelInRow(row, ["项目技术负责人"], fields.constructionTechnicalLeader);
    }
    if (rowHasExactLabel(row, "分包单位")) {
      fillOrClearAfterLabelInRow(row, ["分包单位"], fields.subcontractorUnit);
      fillOrClearAfterLabelInRow(row, ["分包项目负责人", "项目负责人"], fields.subcontractorProjectManager);
      fillOrClearAfterLabelInRow(row, ["分包内容"], fields.subcontractorContent);
    }
  }
}

function fillOrClearProfessionalForeman(sheet: ExcelJS.Worksheet, value: string) {
  const labels = ["专业工长", "施工员"];
  if (value) {
    fillAfterLabel(sheet, labels, value);
  } else {
    clearAfterLabel(sheet, labels);
  }
}
