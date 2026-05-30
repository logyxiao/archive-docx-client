import PizZip from "pizzip";
import type { ArchiveItem, ArchiveRecord } from "../types";
import { resolveProcessFields } from "./fields";
import type { ProcessUserFields } from "./types";
import { archiveProjectCode, escapeRegExp, escapeXml } from "./utils";

export function renderSummaryWorkbook(
  template: ArrayBuffer | Uint8Array,
  record: ArchiveRecord,
  userFields: ProcessUserFields = {},
  item?: ArchiveItem,
): Uint8Array {
  const zip = new PizZip(template);
  const sheet = zip.file("xl/worksheets/sheet1.xml");
  if (!sheet) {
    throw new Error("无法读取汇总表模板工作表");
  }

  const fields = resolveProcessFields(userFields, record.filingUnit || record.owner);
  let xml = sheet.asText();
  xml = setInlineStringCell(xml, "F4", archiveProjectCode(record.archiveCode));
  if (item) {
    xml = fillSubunitQualityWorkbook(xml, record, item);
  }
  xml = setInlineStringCell(xml, "G7", fields.generalContractorUnit);
  xml = setInlineStringCell(xml, "U7", fields.generalContractorProjectManager);
  xml = setInlineStringCell(xml, "AF7", fields.generalContractorTechnicalLeader);
  xml = setInlineStringCell(xml, "G8", fields.constructionUnit);
  xml = setInlineStringCell(xml, "U8", fields.constructionProjectManager);
  xml = setInlineStringCell(xml, "AF8", fields.constructionTechnicalLeader);
  xml = setInlineStringCell(xml, "G9", fields.subcontractorUnit);
  xml = setInlineStringCell(xml, "U9", fields.subcontractorProjectManager);
  xml = setInlineStringCell(xml, "AF9", fields.subcontractorContent);

  zip.file("xl/worksheets/sheet1.xml", xml);
  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function fillSubunitQualityWorkbook(xml: string, record: ArchiveRecord, item: ArchiveItem): string {
  const summary = subunitQualitySummary(record, item);
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

function subunitQualitySummary(record: ArchiveRecord, item: ArchiveItem): {
  subunitProjectName: string;
  unitProjectName: string;
  parts: Array<{ name: string; count: number }>;
} {
  return {
    subunitProjectName: subunitProjectName(item.title),
    unitProjectName: unitProjectName(record),
    parts: [
      { name: "子方阵支架及组件安装", count: 10 },
      { name: "通用工程", count: 5 },
      { name: "主体工程", count: 2 },
    ],
  };
}

function subunitProjectName(title: string): string {
  return title
    .replace(/^\s*\d+[、.．\-\s]*/, "")
    .replace(/^.*?项目\s*/, "")
    .replace(/\s*质量报验申请及验收记录\s*$/, "")
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
