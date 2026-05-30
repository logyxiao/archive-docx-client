import PizZip from "pizzip";
import type { ArchiveItem, ArchiveRecord } from "../types";
import {
  inspectionApplicationSubject,
  replaceBusinessText,
  replaceStartReportScopeText,
  subunitInspectionSubject,
  subunitProjectName,
} from "./textReplacement";
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
    const renderedXml = replaceDocxParagraphs(replaceBusinessText(file.asText(), record, item, userFields), record, item, userFields);
    zip.file(path, path === "word/document.xml" ? compactDocxPageLayout(renderedXml) : renderedXml);
  }

  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function replaceDocxParagraphs(xml: string, record: ArchiveRecord, item: ArchiveItem, userFields: ProcessUserFields): string {
  const projectName = userFields.projectName?.trim() || record.projectName;
  const subunitName = subunitProjectName(item.title);
  const subunitSubject = subunitInspectionSubject(item.title);
  const inspectionSubject = inspectionApplicationSubject(item.title);
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const text = paragraphText(paragraph);
    let nextText = text;
    let nextParagraph = paragraph;

    if (item.fileCode && item.fileCode !== "/" && text.includes("编号：") && text.includes("5028G01")) {
      nextText = nextText.replace(/编号：.*$/, `编号：${item.fileCode}`);
    }

    if (text.includes("我方承担的") && text.includes("已完成了")) {
      nextText = replaceStartReportScopeText(nextText, projectName, item.title);
    }

    if (subunitName) {
      nextParagraph = replaceTextNodes(nextParagraph, "光伏方阵安装子单位工程", subunitName);
    }

    if (subunitSubject && text.includes("根据施工承包合同的规定") && text.includes("子单位工程现已施工完毕")) {
      nextParagraph = replaceTextNodes(nextParagraph, subunitName, subunitSubject);
    }

    if (subunitSubject && text.includes("附件：") && text.includes("子单位工程质量验收记录")) {
      nextParagraph = replaceTextNodes(nextParagraph, subunitName, subunitSubject);
    }

    if (inspectionSubject && text.includes("根据施工承包合同的规定") && /分部（子分部）工程现已施工完毕|分项工程现已施工完毕/.test(text)) {
      nextParagraph = replaceUnderlinedTextInParagraph(nextParagraph, inspectionSubject);
    }

    if (inspectionSubject && text.includes("附件：") && /分部（子分部）工程质量验收记录|分项工程质量验收记录/.test(text)) {
      nextParagraph = replaceUnderlinedTextInParagraph(nextParagraph, inspectionSubject);
    }

    return nextText === text ? nextParagraph : replaceParagraphText(nextParagraph, nextText);
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

function replaceTextNodes(xml: string, search: string, replacement: string): string {
  return xml.replace(/<w:t([^>]*)>([\s\S]*?)<\/w:t>/g, (_match, attrs: string, value: string) => {
    const text = unescapeXml(value);
    if (!text.includes(search)) {
      return `<w:t${attrs}>${value}</w:t>`;
    }
    return `<w:t${attrs}>${escapeXml(text.split(search).join(replacement))}</w:t>`;
  });
}

function replaceUnderlinedTextInParagraph(xml: string, replacement: string): string {
  let replaced = false;
  return xml.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (run) => {
    if (!/<w:u\b/.test(run)) {
      return run;
    }

    if (!replaced) {
      replaced = true;
      return replaceRunText(run, `     ${replacement}     `);
    }

    return replaceRunText(run, "");
  });
}

function replaceRunText(run: string, text: string): string {
  let replacedTextNode = false;
  return run.replace(/<w:t([^>]*)>[\s\S]*?<\/w:t>/g, (_match, attrs: string) => {
    if (replacedTextNode) {
      return `<w:t${attrs}></w:t>`;
    }

    replacedTextNode = true;
    return `<w:t${ensurePreserveSpace(attrs)}>${escapeXml(text)}</w:t>`;
  });
}

function ensurePreserveSpace(attrs: string): string {
  return attrs.includes("xml:space=") ? attrs : `${attrs} xml:space="preserve"`;
}

function compactDocxPageLayout(xml: string): string {
  return compactFormNoteParagraphs(compactSectionLayout(xml));
}

function compactSectionLayout(xml: string): string {
  return xml
    .replace(/<w:pgMar\b([^>]*)\/>/g, (_match, attrs: string) => {
      const nextAttrs = setXmlAttributes(attrs, {
        top: "720",
        right: "1440",
        bottom: "720",
        left: "1440",
        header: "425",
        footer: "425",
      });
      return `<w:pgMar${nextAttrs}/>`;
    })
    .replace(/<w:docGrid\b([^>]*)\/>/g, (_match, attrs: string) => `<w:docGrid${setXmlAttributes(attrs, { linePitch: "276" })}/>`);
}

function compactFormNoteParagraphs(xml: string): string {
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    if (!paragraphText(paragraph).includes("本表一式")) {
      return paragraph;
    }

    return compactParagraphSpacing(paragraph);
  });
}

function compactParagraphSpacing(paragraph: string): string {
  if (/<w:pPr\b[\s\S]*?<\/w:pPr>/.test(paragraph)) {
    return paragraph.replace(/<w:pPr\b([^>]*)>([\s\S]*?)<\/w:pPr>/, (_match, attrs: string, body: string) => {
      const nextBody = /<w:spacing\b[^>]*\/>/.test(body)
        ? body.replace(/<w:spacing\b([^>]*)\/>/, (_spacingMatch: string, spacingAttrs: string) => {
            return `<w:spacing${setXmlAttributes(spacingAttrs, { before: "0", after: "0", line: "240", lineRule: "auto" })}/>`;
          })
        : `<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>${body}`;
      return `<w:pPr${attrs}>${nextBody}</w:pPr>`;
    });
  }

  return paragraph.replace(/(<w:p\b[^>]*>)/, '$1<w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>');
}

function setXmlAttributes(attrs: string, values: Record<string, string>): string {
  let nextAttrs = attrs;
  for (const [name, value] of Object.entries(values)) {
    const pattern = new RegExp(`\\sw:${name}="[^"]*"`);
    if (pattern.test(nextAttrs)) {
      nextAttrs = nextAttrs.replace(pattern, ` w:${name}="${value}"`);
    } else {
      nextAttrs += ` w:${name}="${value}"`;
    }
  }

  return nextAttrs;
}
