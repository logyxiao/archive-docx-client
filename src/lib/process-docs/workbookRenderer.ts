import ExcelJS from "exceljs";
import PizZip from "pizzip";
import type { ArchiveItem, ArchiveRecord } from "../types";
import { PRESERVED_USER_FIELD_LABELS, SOURCE_MISSING_LABELS } from "./constants";
import { resolveProcessFields } from "./fields";
import { inspectionApplicationFullSubject, replaceBusinessText, stripProjectPrefix } from "./textReplacement";
import type { ProcessUserFields, ResolvedProcessFields } from "./types";
import { archiveProjectCode, formatChineseDate, toArrayBuffer } from "./utils";
import { clearAfterLabel, fillAfterLabel, fillAfterLabelExcept, fillOrClearAfterLabelInRow, isMergedSlave, rowHasExactLabel } from "./workbookCells";
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
  return preserveProcessWorkbookPrintLayout(new Uint8Array(buffer));
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

  fillAfterLabelExcept(sheet, ["工程名称", "工程项目名称"], ["单位（子单位）工程名称", "分部（子分部）工程名称", "分项工程名称"], projectName);
  fillAfterLabel(sheet, ["工程编号"], archiveProjectCode(record.archiveCode));
  fillAfterLabel(sheet, ["抽样单位"], item.owner || record.filingUnit);
  fillAfterLabel(sheet, ["抽样日期"], formatChineseDate(item.fileDate));
  fillInspectionLotProjectNames(sheet, record, item);
  fillUnitScopedFields(sheet, fields);
  fillOrClearProfessionalForeman(sheet, fields.constructionTechnicalLeader);
  fillRandomSelfCheckValues(sheet);
  clearAfterLabel(
    sheet,
    SOURCE_MISSING_LABELS.filter((label) => !PRESERVED_USER_FIELD_LABELS.includes(label)),
  );
}

function fillInspectionLotProjectNames(sheet: ExcelJS.Worksheet, record: ArchiveRecord, item: ArchiveItem) {
  if (!item.title.includes("检验批质量验收记录")) {
    return;
  }

  const context = inspectionLotContext(record, item);
  fillAfterLabel(sheet, ["单位（子单位）工程名称"], context.unitProjectName);
  fillAfterLabel(sheet, ["分部（子分部）工程名称"], context.divisionName);
  fillAfterLabel(sheet, ["分项工程名称"], context.subitemName);
}

function inspectionLotContext(record: ArchiveRecord, item: ArchiveItem): {
  unitProjectName: string;
  divisionName: string;
  subitemName: string;
} {
  const beforeItems = record.items.slice(0, record.items.indexOf(item)).reverse();
  const division = beforeItems.find((candidate) => isDivisionQualityItem(candidate.title));
  const subitem = beforeItems.find((candidate) => isSubitemQualityItem(candidate.title));

  return {
    unitProjectName: unitProjectNameFromRecord(record),
    divisionName: division ? qualitySubject(division.title).replace(/\s*分部工程\s*$/, "").trim() : "",
    subitemName: subitem ? qualitySubject(subitem.title).replace(/\s*分项工程\s*$/, "").trim() : inspectionLotSubitemName(item.title),
  };
}

function unitProjectNameFromRecord(record: ArchiveRecord): string {
  return record.volumeTitle
    .split(/[，,、]/u)[0]
    ?.replace(/\s*开工报审.*$/u, "")
    .trim() || record.projectName;
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
  return /分项工程质量(?:报验申请|报审表)及验收记录/.test(title);
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

function preserveProcessWorkbookPrintLayout(bytes: Uint8Array): Uint8Array {
  const zip = new PizZip(bytes);
  const worksheetPaths = Object.keys(zip.files).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name));

  for (const path of worksheetPaths) {
    const sheet = zip.file(path);
    if (!sheet) {
      continue;
    }
    zip.file(path, forceOnePageWorksheetLayout(sheet.asText()));
  }

  preserveWorkbookPrintAreas(zip, worksheetPaths);
  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function forceOnePageWorksheetLayout(xml: string): string {
  let nextXml = upsertFitToPage(xml);
  nextXml = insertSelfClosingElementIfMissing(
    nextXml,
    "pageMargins",
    '<pageMargins left="0.3" right="0.3" top="0.3" bottom="0.3" header="0.2" footer="0.2"/>',
    "</worksheet>",
  );
  nextXml = upsertPageSetup(nextXml);
  return nextXml;
}

function upsertFitToPage(xml: string): string {
  if (/<sheetPr\b[^>]*>[\s\S]*?<\/sheetPr>/.test(xml)) {
    return xml.replace(/<sheetPr\b([^>]*)>([\s\S]*?)<\/sheetPr>/, (_match, attrs: string, body: string) => {
      const nextBody = /<pageSetUpPr\b[^>]*\/>/.test(body)
        ? body.replace(/<pageSetUpPr\b([^>]*)\/>/, (_pageSetupMatch: string, pageSetupAttrs: string) => {
            return `<pageSetUpPr${upsertXmlAttribute(pageSetupAttrs, "fitToPage", "1")}/>`;
          })
        : `<pageSetUpPr fitToPage="1"/>${body}`;
      return `<sheetPr${attrs}>${nextBody}</sheetPr>`;
    });
  }

  if (/<sheetPr\b[^>]*\/>/.test(xml)) {
    return xml.replace(/<sheetPr\b([^>]*)\/>/, (_match, attrs: string) => `<sheetPr${attrs}><pageSetUpPr fitToPage="1"/></sheetPr>`);
  }

  return xml.replace(/(<worksheet\b[^>]*>)/, '$1<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>');
}

function upsertPageSetup(xml: string): string {
  const existing = /<pageSetup\b([^>]*)\/>/;
  if (existing.test(xml)) {
    return xml.replace(existing, (_match, attrs: string) => {
      const nextAttrs = upsertXmlAttribute(
        upsertXmlAttribute(
          upsertXmlAttribute(
            upsertXmlAttribute(attrs, "paperSize", "9"),
            "fitToWidth",
            "1",
          ),
          "fitToHeight",
          "1",
        ),
        "usePrinterDefaults",
        "0",
      );
      return `<pageSetup${nextAttrs}/>`;
    });
  }

  return upsertSelfClosingElement(
    xml,
    "pageSetup",
    '<pageSetup paperSize="9" fitToWidth="1" fitToHeight="1" usePrinterDefaults="0"/>',
    "</worksheet>",
  );
}

function preserveWorkbookPrintAreas(zip: PizZip, worksheetPaths: string[]) {
  const workbook = zip.file("xl/workbook.xml");
  if (!workbook) {
    return;
  }

  const workbookXml = workbook.asText();
  const sheetNames = Array.from(workbookXml.matchAll(/<sheet\b[^>]*name="([^"]+)"[^>]*>/g)).map((match) => match[1]);
  const printAreas = worksheetPaths
    .map((path, index) => {
      const sheet = zip.file(path);
      const sheetName = sheetNames[index];
      const dimension = sheet?.asText().match(/<dimension\b[^>]*ref="([^"]+)"/)?.[1];
      if (!sheetName || !dimension) {
        return null;
      }
      return `<definedName name="_xlnm.Print_Area" localSheetId="${index}">${sheetName}!${absoluteCellRange(dimension)}</definedName>`;
    })
    .filter((value): value is string => Boolean(value));

  if (printAreas.length === 0) {
    return;
  }

  const workbookWithoutPrintAreas = workbookXml.replace(/<definedName name="_xlnm\.Print_Area"[^>]*>[\s\S]*?<\/definedName>/g, "");
  const nextWorkbookXml = /<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/.test(workbookWithoutPrintAreas)
    ? workbookWithoutPrintAreas.replace(/<definedNames\b([^>]*)>([\s\S]*?)<\/definedNames>/, (_match, attrs: string, body: string) => {
        return `<definedNames${attrs}>${body}${printAreas.join("")}</definedNames>`;
      })
    : workbookWithoutPrintAreas.replace("</workbook>", `<definedNames>${printAreas.join("")}</definedNames></workbook>`);

  zip.file("xl/workbook.xml", nextWorkbookXml);
}

function upsertSelfClosingElement(xml: string, tagName: string, elementXml: string, insertBefore: string): string {
  const existing = new RegExp(`<${tagName}\\b[^>]*/>`);
  if (existing.test(xml)) {
    return xml.replace(existing, elementXml);
  }

  const insertIndex = xml.indexOf(insertBefore);
  if (insertIndex === -1) {
    return xml;
  }
  return `${xml.slice(0, insertIndex)}${elementXml}${xml.slice(insertIndex)}`;
}

function insertSelfClosingElementIfMissing(xml: string, tagName: string, elementXml: string, insertBefore: string): string {
  const existing = new RegExp(`<${tagName}\\b[^>]*/>`);
  if (existing.test(xml)) {
    return xml;
  }

  const insertIndex = xml.indexOf(insertBefore);
  if (insertIndex === -1) {
    return xml;
  }
  return `${xml.slice(0, insertIndex)}${elementXml}${xml.slice(insertIndex)}`;
}

function upsertXmlAttribute(attrs: string, name: string, value: string): string {
  const pattern = new RegExp(`\\s${name}="[^"]*"`);
  if (pattern.test(attrs)) {
    return attrs.replace(pattern, ` ${name}="${value}"`);
  }
  return `${attrs} ${name}="${value}"`;
}

function absoluteCellRange(range: string): string {
  return range.split(":").map(absoluteCellRef).join(":");
}

function absoluteCellRef(ref: string): string {
  return ref.replace(/^([A-Z]+)(\d+)$/i, (_match, column: string, row: string) => `$${column.toUpperCase()}$${row}`);
}
