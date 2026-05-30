import PizZip from "pizzip";
import type { ArchiveItem, ArchiveRecord } from "../types";
import { replaceBusinessText } from "./textReplacement";
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
    zip.file(path, replaceDocxParagraphs(replaceBusinessText(file.asText(), record, item, userFields), item));
  }

  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function replaceDocxParagraphs(xml: string, item: ArchiveItem): string {
  if (!item.fileCode || item.fileCode === "/") {
    return xml;
  }

  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const text = paragraphText(paragraph);
    if (!text.includes("编号：") || !text.includes("5028G01")) {
      return paragraph;
    }

    return replaceParagraphText(paragraph, text.replace(/编号：.*$/, `编号：${item.fileCode}`));
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
