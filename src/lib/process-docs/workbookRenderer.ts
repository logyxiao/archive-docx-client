import ExcelJS from "exceljs";
import PizZip from "pizzip";
import type { ArchiveItem, ArchiveRecord } from "../types";
import { collectorLineTrialCablePair, collectorLineWorkbookReplacements } from "./collectorLine";
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
      if (typeof cell.value === "string") {
        cell.value = replaceBusinessText(cell.value, record, item, userFields);
      }
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

function fillInspectionLotProjectNames(sheet: ExcelJS.Worksheet, record: ArchiveRecord, item: ArchiveItem, isSwitchStation: boolean) {
  if (!item.title.includes("检验批质量验收记录")) {
    return;
  }

  const context = isSwitchStation ? switchStationInspectionLotContext(record, item) : inspectionLotContext(record, item);
  fillAfterLabel(sheet, ["单位（子单位）工程名称"], context.unitProjectName);
  fillAfterLabel(sheet, ["分部（子分部）工程名称"], context.divisionName);
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

function applyCollectorLineWorkbookValues(
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

function preserveProcessWorkbookPrintLayout(
  bytes: Uint8Array,
  template: ArrayBuffer | Uint8Array,
  templateModule: ProcessTemplateModule,
): Uint8Array {
  const zip = new PizZip(bytes);
  const templateZip = new PizZip(template);
  const worksheetPaths = Object.keys(zip.files).filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/.test(name));

  for (const path of worksheetPaths) {
    const sheet = zip.file(path);
    if (!sheet) {
      continue;
    }
    const templateSheet = templateZip.file(path);
    let sheetXml = templateModule === "collector-line" ? sheet.asText() : forceOnePageWorksheetLayout(sheet.asText());
    if (templateSheet) {
      sheetXml = restoreTemplateSheetFormatting(sheetXml, templateSheet.asText());
    }
    zip.file(path, sheetXml);
  }

  const templateStylesFile = templateZip.file("xl/styles.xml");
  const generatedStylesFile = zip.file("xl/styles.xml");
  if (templateStylesFile && generatedStylesFile) {
    zip.file("xl/styles.xml", restoreTemplateDefaultFont(generatedStylesFile.asText(), templateStylesFile.asText()));
  }

  preserveProcessWorkbookPrintArea(zip, templateZip);

  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function forceOnePageWorksheetLayout(xml: string): string {
  let nextXml = upsertFitToPage(xml);
  nextXml = upsertPageSetup(nextXml);
  return nextXml;
}

/**
 * Copy <sheetFormatPr> and <cols> from template to generated sheet XML
 * so ExcelJS-injected defaultColWidth and altered column widths are reverted.
 */
function restoreTemplateSheetFormatting(generatedXml: string, templateXml: string): string {
  let result = generatedXml;

  // Restore <sheetFormatPr .../> — remove ExcelJS's version, insert template's (or nothing)
  const templateSheetFormatPr = templateXml.match(/<sheetFormatPr[^>]*\/>/) ?.[0];
  if (templateSheetFormatPr) {
    // Replace generated sheetFormatPr with template's
    if (/<sheetFormatPr[^>]*\/>/.test(result)) {
      result = result.replace(/<sheetFormatPr[^>]*\/>/, templateSheetFormatPr);
    } else {
      result = result.replace(/(<sheetData\b)/, `${templateSheetFormatPr}$1`);
    }
  } else {
    // Template has no sheetFormatPr — strip ExcelJS's injected one
    result = result.replace(/<sheetFormatPr[^>]*\/>/, '');
  }

  // Restore <cols>...</cols> — replace generated with template's
  const templateCols = templateXml.match(/<cols>[\s\S]*?<\/cols>/) ?.[0];
  const generatedHasCols = /<cols>[\s\S]*?<\/cols>/.test(result);
  if (templateCols) {
    if (generatedHasCols) {
      result = result.replace(/<cols>[\s\S]*?<\/cols>/, templateCols);
    } else {
      result = result.replace(/(<sheetData\b)/, `${templateCols}$1`);
    }
  } else if (generatedHasCols) {
    result = result.replace(/<cols>[\s\S]*?<\/cols>/, '');
  }

  return result;
}

function restoreTemplateDefaultFont(generatedStylesXml: string, templateStylesXml: string): string {
  const templateDefaultFont = firstFontXml(templateStylesXml);
  if (!templateDefaultFont) {
    return generatedStylesXml;
  }

  return generatedStylesXml.replace(/(<fonts\b[^>]*>)([\s\S]*?<\/font>)/, `$1${templateDefaultFont}`);
}

function firstFontXml(stylesXml: string): string | undefined {
  return stylesXml.match(/<fonts\b[^>]*>\s*(<font>[\s\S]*?<\/font>)/)?.[1];
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
      const cleanAttrs = attrs.replace(/\s+scale="[^"]*"/g, "");
      const nextAttrs = upsertXmlAttribute(
        upsertXmlAttribute(
          upsertXmlAttribute(
            upsertXmlAttribute(cleanAttrs, "paperSize", "9"),
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

function upsertXmlAttribute(attrs: string, name: string, value: string): string {
  const pattern = new RegExp(`\\s${name}="[^"]*"`);
  if (pattern.test(attrs)) {
    return attrs.replace(pattern, ` ${name}="${value}"`);
  }
  return `${attrs} ${name}="${value}"`;
}

/** Convert e.g. "A1:AI25" into "$A$1:$AI$25" */
function toAbsoluteRange(range: string): string {
  const toAbsCell = (cell: string) => {
    const m = cell.match(/^([A-Z]+)(\d+)$/);
    return m ? `$${m[1]}$${m[2]}` : `$${cell}`;
  };
  const [start, end] = range.split(":");
  return end ? `${toAbsCell(start)}:${toAbsCell(end)}` : toAbsCell(start);
}

function preserveProcessWorkbookPrintArea(zip: PizZip, templateZip: PizZip) {
  const workbook = zip.file("xl/workbook.xml");
  if (!workbook) {
    return;
  }

  const workbookXml = workbook.asText();
  const sheets = workbookXml.match(/<sheet\s+[^>]*name="([^"]+)"[^>]*>/g) || [];
  if (sheets.length === 0) {
    return;
  }

  const printAreas: string[] = [];
  sheets.forEach((sheetTag, index) => {
    const nameMatch = sheetTag.match(/name="([^"]+)"/);
    if (!nameMatch) {
      return;
    }
    const sheetName = nameMatch[1];
    
    const sheetPath = `xl/worksheets/sheet${index + 1}.xml`;
    const sheetFile = zip.file(sheetPath);
    if (!sheetFile) {
      return;
    }
    const sheetXml = sheetFile.asText();

    // Prefer dimension from template sheet (avoids phantom columns added by ExcelJS)
    const templateSheetFile = templateZip.file(sheetPath);
    const sourceXml = templateSheetFile ? templateSheetFile.asText() : sheetXml;
    const dimensionMatch = sourceXml.match(/<dimension\s+ref="([^"]+)"/);
    if (!dimensionMatch) {
      return;
    }
    const range = dimensionMatch[1]; // e.g. "A1:AI25"
    if (!/^[A-Z]+\d+(?::[A-Z]+\d+)?$/.test(range)) {
      return;
    }
    const absRange = toAbsoluteRange(range);
    
    printAreas.push(`<definedName name="_xlnm.Print_Area" localSheetId="${index}">${quoteSheetName(sheetName)}!${absRange}</definedName>`);
  });

  if (printAreas.length === 0) {
    zip.file("xl/workbook.xml", removePrintAreas(workbookXml));
    return;
  }

  const printAreasXml = printAreas.join("");
  let nextWorkbookXml = removePrintAreas(workbookXml);

  if (/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/.test(nextWorkbookXml)) {
    nextWorkbookXml = nextWorkbookXml.replace(/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/, `<definedNames>${printAreasXml}</definedNames>`);
  } else if (/<definedNames\/>/.test(nextWorkbookXml)) {
    nextWorkbookXml = nextWorkbookXml.replace(/<definedNames\/>/, `<definedNames>${printAreasXml}</definedNames>`);
  } else {
    nextWorkbookXml = nextWorkbookXml.replace("</workbook>", `<definedNames>${printAreasXml}</definedNames></workbook>`);
  }

  zip.file("xl/workbook.xml", nextWorkbookXml);
}

function removePrintAreas(workbookXml: string): string {
  return workbookXml
    .replace(/<definedNames\b[^>]*>[\s\S]*?<\/definedNames>/g, "")
    .replace(/<definedNames>\s*<\/definedNames>/g, "");
}

function quoteSheetName(sheetName: string): string {
  if (/^[A-Za-z0-9_\u4e00-\u9fa5]+$/.test(sheetName)) {
    return sheetName;
  }
  return `'${sheetName.replace(/'/g, "''")}'`;
}
