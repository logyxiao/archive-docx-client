import PizZip from "pizzip";
import type { ArchiveRecord } from "../types";
import { resolveProcessFields } from "./fields";
import type { ProcessUserFields } from "./types";
import { archiveProjectCode, escapeRegExp, escapeXml } from "./utils";

export function renderSummaryWorkbook(
  template: ArrayBuffer | Uint8Array,
  record: ArchiveRecord,
  userFields: ProcessUserFields = {},
): Uint8Array {
  const zip = new PizZip(template);
  const sheet = zip.file("xl/worksheets/sheet1.xml");
  if (!sheet) {
    throw new Error("无法读取汇总表模板工作表");
  }

  const fields = resolveProcessFields(userFields, record.filingUnit || record.owner);
  let xml = sheet.asText();
  xml = setInlineStringCell(xml, "F4", archiveProjectCode(record.archiveCode));
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
