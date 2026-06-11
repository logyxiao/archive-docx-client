import ExcelJS from "exceljs";
import type { ArchiveItem, ArchiveRecord } from "../types";
import { PRESERVED_USER_FIELD_LABELS, SOURCE_MISSING_LABELS } from "./constants";
import { resolveProcessFields } from "./fields";
import {
  isSwitchStationArchiveRecord,
  switchStationInspectionLotContext,
} from "./switchStation";
import { inspectionApplicationFullSubject, replaceBusinessText, stripProjectPrefix } from "./textReplacement";
import type { ProcessTemplateModule, ProcessUserFields, ResolvedProcessFields } from "./types";
import { archiveProjectCode, formatChineseDate, toArrayBuffer } from "./utils";
import { clearAfterLabel, fillAfterLabel, fillAfterLabelExcept, fillOrClearAfterLabelInRow, isMergedSlave, rowHasExactLabel } from "./workbookCells";
import { applyCollectorLineWorkbookValues } from "./collectorLineWorkbook";
import { preserveProcessWorkbookPrintLayout } from "./workbookLayout";
import { replaceWorkbookQualityPlaceholders } from "./workbookQuality";
import { fillRandomSelfCheckValues } from "./qualitySelfCheck";

export async function renderProcessWorkbook(
  template: ArrayBuffer | Uint8Array,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields = {},
  templateModule: ProcessTemplateModule = "process",
): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(toArrayBuffer(template));

  for (const sheet of workbook.worksheets) {
    applyWorkbookValues(sheet, record, item, userFields, templateModule);
  }
  applyCollectorLineWorkbookValues(workbook, item, templateModule);

  const buffer = await workbook.xlsx.writeBuffer();
  return preserveProcessWorkbookPrintLayout(new Uint8Array(buffer), template, templateModule);
}

function applyWorkbookValues(
  sheet: ExcelJS.Worksheet,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields,
  templateModule: ProcessTemplateModule,
) {
  const fields = resolveProcessFields(userFields, item.owner || record.filingUnit);
  const projectName = userFields.projectName?.trim() || record.projectName;
  const isSwitchStation = templateModule === "switch-station" || isSwitchStationArchiveRecord(record);
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (isMergedSlave(cell)) {
        return;
      }
      replaceWorkbookCellText(cell, record, item, userFields);
    });
  }

  if (!shouldSkipProjectNameAutoFill(item, isSwitchStation)) {
    fillAfterLabelExcept(sheet, ["工程名称", "工程项目名称"], ["单位（子单位）工程名称", "分部（子分部）工程名称", "分项工程名称"], projectName);
  }
  fillAfterLabel(sheet, ["工程编号"], archiveProjectCode(record.archiveCode));
  fillAfterLabel(sheet, ["抽样单位"], item.owner || record.filingUnit);
  fillAfterLabel(sheet, ["抽样日期"], formatChineseDate(item.fileDate));
  fillInspectionLotProjectNames(sheet, record, item, isSwitchStation);
  fillUnitScopedFields(sheet, fields);
  fillOrClearProfessionalForeman(sheet, fields.constructionTechnicalLeader);
  fillRandomSelfCheckValues(sheet);
  clearAfterLabel(
    sheet,
    SOURCE_MISSING_LABELS.filter((label) => !PRESERVED_USER_FIELD_LABELS.includes(label)),
  );
}

function replaceWorkbookCellText(
  cell: ExcelJS.Cell,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields,
) {
  if (typeof cell.value === "string") {
    cell.value = replaceWorkbookText(cell.value, record, item, userFields);
    return;
  }

  if (cell.value && typeof cell.value === "object" && "richText" in cell.value && Array.isArray(cell.value.richText)) {
    const nextRichText = cell.value.richText.map((part) => ({
      ...part,
      text: replaceWorkbookText(part.text ?? "", record, item, userFields),
    }));
    const originalText = cell.value.richText.map((part) => part.text ?? "").join("");
    const nextText = nextRichText.map((part) => part.text).join("");
    if (nextText !== originalText) {
      cell.value = {
        ...cell.value,
        richText: nextRichText,
      };
    }
    return;
  }

  if (cell.value && typeof cell.value === "object" && "text" in cell.value && typeof cell.value.text === "string") {
    const nextText = replaceWorkbookText(cell.value.text, record, item, userFields);
    if (nextText !== cell.value.text) {
      cell.value = {
        ...cell.value,
        text: nextText,
      };
    }
  }
}

function replaceWorkbookText(
  value: string,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields,
): string {
  return replaceWorkbookQualityPlaceholders(replaceBusinessText(value, record, item, userFields));
}

function fillInspectionLotProjectNames(sheet: ExcelJS.Worksheet, record: ArchiveRecord, item: ArchiveItem, isSwitchStation: boolean) {
  if (!item.title.includes("检验批质量验收记录")) {
    return;
  }

  const context = isSwitchStation ? switchStationInspectionLotContext(record, item) : inspectionLotContext(record, item);
  fillAfterLabel(sheet, ["单位（子单位）工程名称"], context.unitProjectName);
  fillAfterLabel(sheet, ["分部（子分部）工程名称", "分部工程名称"], context.divisionName);
  fillAfterLabel(sheet, ["分项工程名称"], context.subitemName);
}

function shouldSkipProjectNameAutoFill(item: ArchiveItem, isSwitchStation: boolean): boolean {
  return /隐蔽工程|检查记录|施工记录|测量记录|签证|位置记录/.test(item.title)
    || (isSwitchStation && /隐蔽工程|检查记录|施工记录|测量记录/.test(item.title));
}

function inspectionLotContext(record: ArchiveRecord, item: ArchiveItem): {
  unitProjectName: string;
  divisionName: string;
  subitemName: string;
} {
  const beforeItems = record.items.slice(0, record.items.indexOf(item)).reverse();
  const division = beforeItems.find((candidate) => isDivisionQualityItem(candidate.title));
  const subitem = beforeItems.find((candidate) => isSubitemQualityItem(candidate.title));
  const subitemName = subitem ? qualitySubject(subitem.title).replace(/\s*分项工程\s*$/, "").trim() : inspectionLotSubitemName(item.title);

  return {
    unitProjectName: unitProjectNameFromRecord(record),
    divisionName: division ? qualitySubject(division.title).replace(/\s*分部工程\s*$/, "").trim() : divisionNameFromRecord(record, subitemName || item.title),
    subitemName,
  };
}

function unitProjectNameFromRecord(record: ArchiveRecord): string {
  return record.volumeTitle
    .split(/[，,、]/u)[0]
    ?.replace(/\s*开工报审.*$/u, "")
    .replace(/\s*子单位工程\s*$/u, "")
    .trim() || record.projectName;
}

function divisionNameFromRecord(record: ArchiveRecord, hint: string): string {
  const divisionNames = record.volumeTitle
    .split(/[，,、]/u)
    .map((part) => part.trim())
    .filter((part) => part.includes("分部"))
    .map((part) => part
      .replace(/\s*开工报审.*$/u, "")
      .replace(/\s*分部工程.*$/u, "")
      .replace(/\s*分部.*$/u, "")
      .trim())
    .filter(Boolean);
  const electricalLineDivision = divisionNames.find((name) => name.includes("电气线路安装"));
  if ((hint.includes("电气二次系统") || hint.includes("直流配电柜")) && electricalLineDivision) {
    return electricalLineDivision;
  }

  return divisionNames.length === 1 ? divisionNames[0] : "";
}

function qualitySubject(title: string): string {
  return inspectionApplicationFullSubject(title);
}

function inspectionLotSubitemName(title: string): string {
  return stripProjectPrefix(title)
    .replace(/\s*检验批质量验收记录\s*$/u, "")
    .replace(/\d+#厂房/g, "")
    .replace(/综合楼/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function isDivisionQualityItem(title: string): boolean {
  return /分部工程质量(?:报验申请|报审表)及验收记录/.test(title) && !title.includes("子单位");
}

function isSubitemQualityItem(title: string): boolean {
  return /分项工程(?:工程)?质量(?:报验申请|报审表)及验收记录/.test(title)
    || /分项工程验收记录$/.test(title);
}

function fillUnitScopedFields(sheet: ExcelJS.Worksheet, fields: ResolvedProcessFields) {
  let inSignatureSection = false;
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    if (!inSignatureSection) {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (isMergedSlave(cell)) return;
        const val = typeof cell.value === "string" ? cell.value.replace(/\s+/g, "") : "";
        if (
          val.includes("签字") ||
          val.includes("验收单位") ||
          val.includes("验收结论") ||
          val.includes("签字栏") ||
          val.includes("项目监理机构")
        ) {
          inSignatureSection = true;
        }
      });
    }

    if (inSignatureSection) {
      continue;
    }

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
