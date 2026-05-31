import PizZip from "pizzip";
import type { ArchiveItem, ArchiveRecord } from "../types";
import { collectorLineDivisionEvaluationItems } from "./collectorLine";
import { resolveProcessFields } from "./fields";
import {
  followingSwitchStationDivisionChildren,
  isSwitchStationSummaryTemplate,
  normalizeSwitchStationQualityName,
  switchStationContextRecords,
  switchStationDivisionName,
  switchStationSubunitName,
  switchStationSubunitParts,
  switchStationUnitProjectName,
} from "./switchStation";
import { inspectionApplicationFullSubject, inspectionApplicationSubject, subunitProjectName } from "./textReplacement";
import type { ProcessTemplate, ProcessTemplateModule, ProcessUserFields } from "./types";
import { archiveProjectCode, escapeRegExp, escapeXml } from "./utils";

export function renderSummaryWorkbook(
  template: ArrayBuffer | Uint8Array,
  record: ArchiveRecord,
  userFields: ProcessUserFields = {},
  item?: ArchiveItem,
  processTemplate?: ProcessTemplate,
  templateModule: ProcessTemplateModule = "process",
  contextRecords: ArchiveRecord[] = [record],
): Uint8Array {
  const zip = new PizZip(template);
  const sheet = zip.file("xl/worksheets/sheet1.xml");
  if (!sheet) {
    throw new Error("无法读取汇总表模板工作表");
  }

  const fields = resolveProcessFields(userFields, record.filingUnit || record.owner);
  let xml = sheet.asText();
  const isDivisionEvaluation = isDivisionEvaluationTemplate(processTemplate);
  const isSwitchStation = templateModule === "switch-station";
  const isCollectorLine = templateModule === "collector-line";
  xml = setInlineStringCell(xml, "F4", archiveProjectCode(record.archiveCode));
  if (item) {
    xml = isDivisionEvaluation
      ? fillDivisionEvaluationWorkbook(xml, record, item, isSwitchStation, isCollectorLine)
      : isSubunitSummaryTemplate(processTemplate)
      ? fillSubunitQualityWorkbook(xml, record, item, processTemplate, contextRecords)
      : fillDivisionQualityWorkbook(xml, record, item, isSwitchStation);
    if (!isDivisionEvaluation) {
      xml = preserveSummaryQualityPrintLayout(xml);
    }
  }
  if (!isDivisionEvaluation) {
    xml = setInlineStringCell(xml, "G7", fields.generalContractorUnit);
    xml = setInlineStringCell(xml, "U7", fields.generalContractorProjectManager);
    xml = setInlineStringCell(xml, "AF7", fields.generalContractorTechnicalLeader);
    xml = setInlineStringCell(xml, "G8", fields.constructionUnit);
    xml = setInlineStringCell(xml, "U8", fields.constructionProjectManager);
    xml = setInlineStringCell(xml, "AF8", fields.constructionTechnicalLeader);
    xml = setInlineStringCell(xml, "G9", fields.subcontractorUnit);
    xml = setInlineStringCell(xml, "U9", fields.subcontractorProjectManager);
    xml = setInlineStringCell(xml, "AF9", fields.subcontractorContent);
  }

  zip.file("xl/worksheets/sheet1.xml", xml);
  if (item && !isDivisionEvaluation) {
    preserveSummaryWorkbookPrintArea(zip);
  }
  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function fillSubunitQualityWorkbook(
  xml: string,
  record: ArchiveRecord,
  item: ArchiveItem,
  template: ProcessTemplate | undefined,
  contextRecords: ArchiveRecord[],
): string {
  const summary = subunitQualitySummary(record, item, template, contextRecords);
  let nextXml = xml;
  nextXml = setInlineStringCell(nextXml, "I2", summary.subunitProjectName);
  nextXml = setInlineStringCell(nextXml, "G6", summary.unitProjectName);
  nextXml = setInlineStringCell(nextXml, "U6", String(summary.parts.length));
  nextXml = setInlineStringCell(nextXml, "AF6", String(summary.parts.reduce((total, part) => total + part.count, 0)));

  for (let index = 0; index < 12; index += 1) {
    const row = 11 + index;
    const part = summary.parts[index];
    nextXml = setInlineStringCell(nextXml, `B${row}`, part ? String(index + 1) : "");
    nextXml = setInlineStringCell(nextXml, `D${row}`, part?.name ?? "");
    nextXml = setInlineStringCell(nextXml, `J${row}`, part ? String(part.count) : "");
    nextXml = setInlineStringCell(nextXml, `V${row}`, part ? "自检检查符合规范及设计要求" : "");
    nextXml = setInlineStringCell(nextXml, `AD${row}`, "");
  }

  nextXml = setInlineStringCell(nextXml, "V23", "齐全完整");
  nextXml = setInlineStringCell(nextXml, "V24", "自检合格");
  nextXml = setInlineStringCell(nextXml, "V25", "自检合格");
  return nextXml;
}

function fillDivisionQualityWorkbook(xml: string, record: ArchiveRecord, item: ArchiveItem, isSwitchStation = false): string {
  const summary = divisionQualitySummary(record, item, isSwitchStation);
  let nextXml = xml;
  nextXml = setInlineStringCell(nextXml, "I2", summary.divisionName);
  nextXml = setInlineStringCell(nextXml, "G6", summary.unitProjectName);
  nextXml = setInlineStringCell(nextXml, "U6", summary.subdivisionCount);
  nextXml = setInlineStringCell(nextXml, "AF6", String(summary.items.length));

  for (let index = 0; index < 12; index += 1) {
    const row = 11 + index;
    const part = summary.items[index];
    nextXml = setInlineStringCell(nextXml, `B${row}`, part ? String(index + 1) : "");
    nextXml = setInlineStringCell(nextXml, `D${row}`, "");
    nextXml = setInlineStringCell(nextXml, `J${row}`, part?.name ?? "");
    nextXml = setInlineStringCell(nextXml, `S${row}`, part ? String(part.inspectionLotCount) : "");
    nextXml = setInlineStringCell(nextXml, `V${row}`, part ? "自检检查符合规范及设计要求" : "");
    nextXml = setInlineStringCell(nextXml, `AD${row}`, "");
  }

  nextXml = setInlineStringCell(nextXml, "V23", "齐全完整");
  nextXml = setInlineStringCell(nextXml, "V24", "自检合格");
  nextXml = setInlineStringCell(nextXml, "V25", "自检合格");
  return nextXml;
}

function fillDivisionEvaluationWorkbook(
  xml: string,
  record: ArchiveRecord,
  item: ArchiveItem,
  isSwitchStation = false,
  isCollectorLine = false,
): string {
  const collectorItems = isCollectorLine ? collectorLineDivisionEvaluationItems() : undefined;
  const summaryItems = isCollectorLine ? undefined : divisionQualitySummary(record, item, isSwitchStation).items;
  let nextXml = xml;

  for (let index = 0; index < 15; index += 1) {
    const row = 7 + index;
    const collectorPart = collectorItems?.[index];
    const summaryPart = summaryItems?.[index];
    const name = collectorPart?.name ?? summaryPart?.name;
    const totalItems = collectorItems?.length ?? summaryItems?.length ?? 0;
    nextXml = setInlineStringCell(nextXml, `B${row}`, name ? String(index + 1) : "");
    nextXml = setInlineStringCell(nextXml, `D${row}`, name ?? (index === totalItems ? "以下空白" : ""));
    nextXml = setInlineStringCell(nextXml, `T${row}`, name ? (collectorPart?.nature ?? "主要") : "");
    nextXml = setInlineStringCell(nextXml, `W${row}`, name ? (collectorPart?.grade ?? "合格") : "");
    nextXml = setInlineStringCell(nextXml, `AC${row}`, "");
  }

  return nextXml;
}

function preserveSummaryQualityPrintLayout(xml: string): string {
  let nextXml = xml.replace(
    /<pageSetUpPr(?:\s[^>]*)?\/>/,
    '<pageSetUpPr fitToPage="1"/>',
  );

  if (!/<pageSetUpPr\b/.test(nextXml)) {
    nextXml = nextXml.replace(/<sheetPr>/, '<sheetPr><pageSetUpPr fitToPage="1"/>');
  }

  nextXml = upsertSelfClosingElement(
    nextXml,
    "pageMargins",
    '<pageMargins left="0.786805555555556" right="0" top="0.590277777777778" bottom="0" header="0.511805555555556" footer="0.511805555555556"/>',
    "</worksheet>",
  );
  nextXml = upsertPageSetup(nextXml);
  return nextXml;
}

function preserveSummaryWorkbookPrintArea(zip: PizZip) {
  const workbook = zip.file("xl/workbook.xml");
  if (!workbook) {
    return;
  }

  const workbookXml = workbook.asText();
  const sheetName = workbookXml.match(/<sheet[^>]*name="([^"]+)"/)?.[1] ?? "Sheet1";
  const printArea = `<definedName name="_xlnm.Print_Area" localSheetId="0">${sheetName}!$A$1:$AL$30</definedName>`;
  let nextWorkbookXml = workbookXml;
  if (nextWorkbookXml.includes("_xlnm.Print_Area")) {
    nextWorkbookXml = nextWorkbookXml.replace(/<definedName name="_xlnm\.Print_Area"[^>]*>[\s\S]*?<\/definedName>/, printArea);
  } else if (/<definedNames\/>|<definedNames><\/definedNames>/.test(nextWorkbookXml)) {
    nextWorkbookXml = nextWorkbookXml.replace(/<definedNames\/>|<definedNames><\/definedNames>/, `<definedNames>${printArea}</definedNames>`);
  } else {
    nextWorkbookXml = nextWorkbookXml.replace("</workbook>", `<definedNames>${printArea}</definedNames></workbook>`);
  }

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

function upsertPageSetup(xml: string): string {
  const existing = /<pageSetup\b([^>]*)\/>/;
  if (existing.test(xml)) {
    return xml.replace(existing, (_match, attrs: string) => {
      const cleanAttrs = attrs.replace(/\s+scale="[^"]*"/g, "");
      const nextAttrs = upsertXmlAttribute(
        upsertXmlAttribute(
          upsertXmlAttribute(
            upsertXmlAttribute(
              upsertXmlAttribute(cleanAttrs, "paperSize", "9"),
              "fitToWidth",
              "1",
            ),
            "fitToHeight",
            "1",
          ),
          "horizontalDpi",
          "600",
        ),
        "verticalDpi",
        "600",
      );
      return `<pageSetup${nextAttrs}/>`;
    });
  }

  return upsertSelfClosingElement(
    xml,
    "pageSetup",
    '<pageSetup paperSize="9" fitToWidth="1" fitToHeight="1" horizontalDpi="600" verticalDpi="600"/>',
    "</worksheet>",
  );
}

function upsertXmlAttribute(attrs: string, name: string, value: string): string {
  const pattern = new RegExp(`\\s${name}="[^"]*"`);
  if (pattern.test(attrs)) {
    return attrs.replace(pattern, ` ${name}="${value}"`);
  }
  return `${attrs} ${name}="${value}"`;
}

function subunitQualitySummary(record: ArchiveRecord, item: ArchiveItem, template: ProcessTemplate | undefined, contextRecords: ArchiveRecord[]): {
  subunitProjectName: string;
  unitProjectName: string;
  parts: Array<{ name: string; count: number }>;
} {
  if (isSwitchStationSummaryTemplate(template)) {
    return {
      subunitProjectName: switchStationSubunitName(item.title),
      unitProjectName: switchStationUnitProjectName(),
      parts: switchStationSubunitParts(switchStationContextRecords(contextRecords)),
    };
  }

  const followingItems = record.items.slice(record.items.indexOf(item) + 1);
  const parts = followingItems
    .slice(0, nextSubunitBoundaryIndex(followingItems))
    .filter((candidate) => isDivisionQualityItem(candidate.title))
    .map((candidate) => ({
      name: divisionName(candidate.title),
      count: countSubitemsInDivision(record, candidate),
    }));

  return {
    subunitProjectName: subunitProjectName(item.title),
    unitProjectName: unitProjectName(record),
    parts,
  };
}

function divisionQualitySummary(record: ArchiveRecord, item: ArchiveItem, isSwitchStation = false): {
  divisionName: string;
  unitProjectName: string;
  subdivisionCount: string;
  items: Array<{ name: string; inspectionLotCount: number | "/" }>;
} {
  const divisionName = isSwitchStation ? switchStationDivisionName(item.title) : inspectionApplicationSubject(item.title);
  const followingItems = record.items.slice(record.items.indexOf(item) + 1);
  const children = isSwitchStation
    ? followingSwitchStationDivisionChildren(record, item)
    : followingItems
      .slice(0, nextDivisionBoundaryIndex(followingItems))
      .filter((candidate) => isDivisionSummaryChild(candidate.title));
  const items = groupedDivisionSummaryChildren(record, children, isSwitchStation);

  return {
    divisionName,
    unitProjectName: unitProjectName(record),
    subdivisionCount: "/",
    items,
  };
}

function groupedDivisionSummaryChildren(
  record: ArchiveRecord,
  children: ArchiveItem[],
  isSwitchStation = false,
): Array<{ name: string; inspectionLotCount: number | "/" }> {
  const groups = new Map<string, { name: string; inspectionLotCount: number }>();
  for (const child of children) {
    const name = isSwitchStation ? normalizeSwitchStationQualityName(qualityItemName(child.title)) : qualityItemName(child.title);
    const current = groups.get(name) ?? { name, inspectionLotCount: 0 };
    current.inspectionLotCount += countInspectionLots(record, child);
    groups.set(name, current);
  }

  return Array.from(groups.values()).map((group) => ({
    name: group.name,
    inspectionLotCount: group.inspectionLotCount > 0 ? group.inspectionLotCount : "/",
  }));
}

function nextDivisionBoundaryIndex(items: ArchiveItem[]): number {
  const index = items.findIndex((item) =>
    item.title.includes("开工报审") || isDivisionQualityItem(item.title) || isSubunitQualityItem(item.title),
  );
  return index === -1 ? items.length : index;
}

function isDivisionSummaryChild(title: string): boolean {
  return /分项工程(?:工程)?质量(?:报验申请|报审表)及验收记录/.test(title)
    || /分项工程验收记录$/.test(title);
}

function qualityItemName(title: string): string {
  return inspectionApplicationFullSubject(title)
    .replace(/\s*分项工程\s*$/, "")
    .trim();
}

function countInspectionLots(record: ArchiveRecord, item: ArchiveItem): number {
  const followingItems = record.items.slice(record.items.indexOf(item) + 1);
  const children = followingItems.slice(0, nextSubitemBoundaryIndex(followingItems));
  return children.filter((candidate) => candidate.title.includes("检验批质量验收记录")).length;
}

function nextSubitemBoundaryIndex(items: ArchiveItem[]): number {
  const index = items.findIndex((item) =>
    item.title.includes("开工报审")
    || isDivisionQualityItem(item.title)
    || isDivisionSummaryChild(item.title),
  );
  return index === -1 ? items.length : index;
}

function nextSubunitBoundaryIndex(items: ArchiveItem[]): number {
  const index = items.findIndex((item) =>
    isSubunitQualityItem(item.title),
  );
  return index === -1 ? items.length : index;
}

function isSubunitQualityItem(title: string): boolean {
  return /子单位(?:工程)?/.test(title) && /质量(?:报验申请|报审表)及验收记录/.test(title);
}

function isDivisionQualityItem(title: string): boolean {
  return /分部工程质量(?:报验申请|报审表)及验收记录/.test(title) && !title.includes("子单位");
}

function isSubunitSummaryTemplate(template: ProcessTemplate | undefined): boolean {
  return Boolean(template?.originalName.includes("子单位") && template.originalName.includes("质量验收记录"));
}

function isDivisionEvaluationTemplate(template: ProcessTemplate | undefined): boolean {
  return Boolean(template?.kind === "xlsx" && /分部工程质量(?:检查)?验收评定表/.test(template.originalName));
}

function divisionName(title: string): string {
  return inspectionApplicationFullSubject(title)
    .replace(/\s*分部工程\s*$/, "")
    .trim();
}

function countSubitemsInDivision(record: ArchiveRecord, item: ArchiveItem): number {
  const followingItems = record.items.slice(record.items.indexOf(item) + 1);
  const children = followingItems
    .slice(0, nextDivisionBoundaryIndex(followingItems))
    .filter((candidate) => isDivisionSummaryChild(candidate.title));
  const groups = new Set(children.map(subitemGroupKey));
  return groups.size;
}

function subitemGroupKey(item: ArchiveItem): string {
  return `${subitemParentCode(item.fileCode)}|${normalizedSubitemTitle(qualityItemName(item.title))}`;
}

function subitemParentCode(fileCode: string): string {
  const code = fileCode.trim();
  if (!code || code === "/") {
    return "";
  }

  return code.replace(/-\d{3}$/, "");
}

function normalizedSubitemTitle(title: string): string {
  return title
    .replace(/^#\d+\s*/, "")
    .replace(/开关柜G\d+柜/g, "开关柜")
    .trim();
}

function unitProjectName(record: ArchiveRecord): string {
  const startReport = record.items.find((item) => item.title.includes("单位工程开工报审"));
  return (startReport?.title ?? record.volumeTitle)
    .replace(/^\s*\d+[、.．\-\s]*/, "")
    .replace(/^.*?项目\s*/, "")
    .replace(/\s*单位工程开工报审\s*$/, "")
    .trim();
}

function setInlineStringCell(xml: string, address: string, value: string): string {
  const escapedAddress = escapeRegExp(address);
  const selfClosingPattern = new RegExp(`<c([^>]*\\br="${escapedAddress}"[^>]*)\\/>`);
  if (selfClosingPattern.test(xml)) {
    return xml.replace(selfClosingPattern, (_match, attrs: string) => inlineStringCellXml(attrs, value));
  }

  const cellPattern = new RegExp(`<c([^>]*\\br="${escapedAddress}"[^>]*)>[\\s\\S]*?<\\/c>`);
  return xml.replace(cellPattern, (_match, attrs: string) => inlineStringCellXml(attrs, value));
}

function inlineStringCellXml(attrs: string, value: string): string {
  const cleanAttrs = attrs.replace(/\s+t="[^"]*"/, "").replace(/\s*\/\s*$/, "");
  return `<c${cleanAttrs} t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
}
