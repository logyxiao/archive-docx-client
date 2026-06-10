import PizZip from "pizzip";
import type { ArchiveItem, ArchiveRecord } from "../types";
import { collectorLineApplicationSubject } from "./collectorLine";
import { switchStationApplicationSubject } from "./switchStation";
import {
  inspectionApplicationSubject,
  replaceBusinessText,
  startReportScope,
  subunitInspectionSubject,
  subunitProjectName,
} from "./textReplacement";
import type { ProcessTemplateModule, ProcessUserFields } from "./types";
import { escapeXml, unescapeXml } from "./utils";

export function renderProcessDocx(
  template: ArrayBuffer | Uint8Array,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields = {},
  templateModule: ProcessTemplateModule = "process",
): Uint8Array {
  const zip = new PizZip(template);
  for (const path of ["word/document.xml", ...Object.keys(zip.files).filter((name) => /^word\/(header|footer)\d+\.xml$/.test(name))]) {
    const file = zip.file(path);
    if (!file) {
      continue;
    }
    const renderedXml = replaceDocxParagraphs(replaceBusinessText(file.asText(), record, item, userFields), record, item, userFields, templateModule);
    zip.file(path, path === "word/document.xml" ? compactDocxPageLayout(renderedXml) : renderedXml);
  }

  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function replaceDocxParagraphs(
  xml: string,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields,
  templateModule: ProcessTemplateModule,
): string {
  const projectName = (userFields.projectName?.trim() || record.projectName).replace(/MWP/g, "MWp");
  const subunitName = subunitProjectName(item.title);
  const subunitSubject = subunitInspectionSubject(item.title);
  const inspectionSubject = templateModule === "switch-station"
    ? switchStationApplicationSubject(item.title)
    : templateModule === "collector-line"
      ? collectorLineApplicationSubject(item.title)
    : inspectionApplicationSubject(item.title);
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const text = paragraphText(paragraph);
    let nextText = text;
    let nextParagraph = paragraph;

    if (text.includes("工程开工报审表")) {
      nextText = startReportTitle(item.title);
    }

    if (item.fileCode && item.fileCode !== "/" && text.includes("编号：") && text.includes("5028G01")) {
      nextText = nextText.replace(/编号：.*$/, `编号：${item.fileCode}`);
    }

    if (text.includes("我方承担的") && text.includes("已完成了")) {
      nextParagraph = replaceStartReportUnderlinedScope(nextParagraph, startReportScope(projectName, item.title));
    }

    if (text.includes("特申请于") && text.includes("日开工")) {
      nextParagraph = replaceStartReportDate(nextParagraph, item.fileDate);
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

    if (text.includes("工程已完成施工任务") && text.includes("现报请查验")) {
      if (templateModule === "switch-station") {
        nextParagraph = replaceTextNodes(nextParagraph, "设备基础模板", switchStationApplicationSubject(item.title));
      } else {
        nextText = replaceHiddenWorkSubjectText(nextText, item.title);
      }
    }

    return nextText === text ? nextParagraph : replaceParagraphText(nextParagraph, nextText);
  });
}

function startReportTitle(title: string): string {
  return title.includes("分部开工报审") || title.includes("分部工程开工报审")
    ? "分部工程开工报审表"
    : "单位工程开工报审表";
}

function replaceStartReportUnderlinedScope(paragraph: string, scope: string): string {
  if (!scope) {
    return paragraph;
  }

  let afterPrefix = false;
  let scopeWritten = false;
  return paragraph.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (run) => {
    const text = runText(run);
    if (!afterPrefix) {
      if (text.includes("我方承担的")) {
        afterPrefix = true;
      }
      return run;
    }

    if (text.includes("，已完成了")) {
      afterPrefix = false;
      return run;
    }

    if (!/<w:u\b/.test(run)) {
      return run;
    }

    if (!scopeWritten) {
      scopeWritten = true;
      return replaceRunText(run, ` ${scope} `);
    }

    return replaceRunText(run, "");
  });
}

function replaceStartReportDate(paragraph: string, date: string): string {
  const parts = compactDateParts(date);
  if (!parts) {
    return paragraph;
  }

  let afterTrigger = false;
  let partIndex = 0;
  return paragraph.replace(/<w:r\b[^>]*>[\s\S]*?<\/w:r>/g, (run) => {
    const text = runText(run);
    if (!afterTrigger) {
      if (text.includes("特申请于")) {
        afterTrigger = true;
      }
      return run;
    }

    if (partIndex >= parts.length || text.includes("日开工")) {
      return run;
    }

    if (!/<w:u\b/.test(run)) {
      return run;
    }

    const replacement = ` ${parts[partIndex]} `;
    partIndex += 1;
    return replaceRunText(run, replacement);
  });
}

function compactDateParts(date: string): [string, string, string] | null {
  const match = date.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (!match) {
    return null;
  }

  return [match[1], String(Number(match[2])), String(Number(match[3]))];
}

function replaceHiddenWorkSubjectText(text: string, title: string): string {
  const subject = hiddenWorkSubject(title);
  if (!subject) {
    return text;
  }

  return text.replace(/：\s*.*?\s*工程已完成施工任务/, `： ${subject} 工程已完成施工任务`);
}

function hiddenWorkSubject(title: string): string {
  return title
    .replace(/^\s*\d+[、.．\-\s]*/, "")
    .replace(/^.*?项目\s*/, "")
    .replace(/\s*隐蔽工程报验申请及质量验收记录\s*$/, "")
    .replace(/\s*隐蔽工程质量报验单及隐蔽工程质量验收记录\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();
}

function paragraphText(paragraph: string): string {
  return Array.from(paragraph.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g))
    .map((match) => unescapeXml(match[1]))
    .join("");
}

function runText(run: string): string {
  return Array.from(run.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g))
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
  return compactProjectCodeParagraphs(compactFormNoteParagraphs(xml));
}

function compactFormNoteParagraphs(xml: string): string {
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    if (!paragraphText(paragraph).includes("本表一式")) {
      return paragraph;
    }

    return compactParagraphSpacing(paragraph);
  });
}

function compactProjectCodeParagraphs(xml: string): string {
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const text = paragraphText(paragraph);
    if (!text.includes("工程名称：") || !text.includes("编号：") || text.length < 42) {
      return paragraph;
    }

    return setParagraphFontSize(compactParagraphSpacing(paragraph), compactProjectCodeFontSize(text));
  });
}

function compactProjectCodeFontSize(text: string): string {
  if (text.length >= 52) {
    return "16";
  }
  if (text.length >= 48) {
    return "18";
  }
  return "19";
}

function setParagraphFontSize(paragraph: string, halfPoints: string): string {
  return paragraph.replace(/<w:rPr\b([^>]*)>([\s\S]*?)<\/w:rPr>/g, (_match, attrs: string, body: string) => {
    const withAsciiSize = setEmptyXmlElementAttributes(body, "w:sz", { val: halfPoints });
    const withComplexSize = setEmptyXmlElementAttributes(withAsciiSize, "w:szCs", { val: halfPoints });
    return `<w:rPr${attrs}>${withComplexSize}</w:rPr>`;
  });
}

function setEmptyXmlElementAttributes(body: string, elementName: string, values: Record<string, string>): string {
  const pattern = new RegExp(`<${elementName}\\b([^>]*)\\/>`);
  if (pattern.test(body)) {
    return body.replace(pattern, (_match, attrs: string) => `<${elementName}${setXmlAttributes(attrs, values)}/>`);
  }

  const attrs = Object.entries(values)
    .map(([name, value]) => ` w:${name}="${value}"`)
    .join("");
  return `<${elementName}${attrs}/>${body}`;
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
