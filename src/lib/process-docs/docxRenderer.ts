import PizZip from "pizzip";
import type { ArchiveItem, ArchiveRecord } from "../types";
import { replaceBusinessText, replaceStartReportScopeText } from "./textReplacement";
import type { ProcessUserFields } from "./types";
import { escapeXml, unescapeXml } from "./utils";

export function renderProcessDocx(
  template: ArrayBuffer | Uint8Array,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields = {},
): Uint8Array {
  const zip = new PizZip(template);
  for (const path of ["word/document.xml", ...Object.keys(zip.files).filter((name) => /^word\/(header|footer)\d+\.xml$/.test(name))]) {
    const file = zip.file(path);
    if (!file) {
      continue;
    }
    zip.file(path, replaceDocxParagraphs(replaceBusinessText(file.asText(), record, item, userFields), record, item, userFields));
  }

  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function replaceDocxParagraphs(xml: string, record: ArchiveRecord, item: ArchiveItem, userFields: ProcessUserFields): string {
  const projectName = userFields.projectName?.trim() || record.projectName;
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const text = paragraphText(paragraph);
    let nextText = text;

    if (item.fileCode && item.fileCode !== "/" && text.includes("编号：") && text.includes("5028G01")) {
      nextText = nextText.replace(/编号：.*$/, `编号：${item.fileCode}`);
    }

    if (text.includes("我方承担的") && text.includes("已完成了")) {
      nextText = replaceStartReportScopeText(nextText, projectName, item.title);
    }

    return nextText === text ? paragraph : replaceParagraphText(paragraph, nextText);
  });
}

function paragraphText(paragraph: string): string {
  return Array.from(paragraph.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g))
    .map((match) => unescapeXml(match[1]))
    .join("");
}

function replaceParagraphText(paragraph: string, text: string): string {
  let isFirstTextNode = true;
  return paragraph.replace(/<w:t([^>]*)>[\s\S]*?<\/w:t>/g, (_match, attrs: string) => {
    if (isFirstTextNode) {
      isFirstTextNode = false;
      return `<w:t${attrs}>${escapeXml(text)}</w:t>`;
    }
    return `<w:t${attrs}></w:t>`;
  });
}
