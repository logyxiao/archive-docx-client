import type { ArchiveItem } from "../types";
import { collectorLineOutputTitle } from "./collectorLine";
import { stripProjectPrefix } from "./textReplacement";
import type { ProcessTemplate } from "./types";
import { sanitizeFileName } from "./utils";

export function processOutputName(template: ProcessTemplate, item: ArchiveItem): string {
  const extension = template.outputExtension;
  const sequence = item.sequence || String(template.sequence);
  const code = template.outputFileCodeOverride ?? item.fileCode;
  const fileCode = code && code !== "/" ? code : "";
  const title = outputTitle(template, item.title);

  return sanitizeFileName(`${sequence}、${fileCode}${title}${extension}`);
}

function outputTitle(template: ProcessTemplate, title: string): string {
  const cleanTitle = title.replace(/^\s*\d+[、.．\-\s]*/, "").trim();
  const collectorLineTitle = collectorLineOutputTitle(cleanTitle);
  if (collectorLineTitle && template.kind === "xlsx") {
    return collectorLineTitle;
  }
  const acceptanceStem = cleanTitle
    .replace(/\s*质量(?:报验申请|报审表)(?:及验收记录)?\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();
  const hiddenWorkStem = cleanTitle
    .replace(/\s*隐蔽工程报验申请及质量验收记录\s*$/, "")
    .replace(/\s*隐蔽工程质量报验单及隐蔽工程质量验收记录\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();

  if (template.kind === "docx" && template.templateFile === "隐蔽工程质量报验单.docx") {
    return `${stripProjectPrefix(hiddenWorkStem)}隐蔽工程质量报验单`;
  }

  if (template.kind === "xlsx" && template.originalName.includes("隐蔽工程质量验收记录")) {
    return `${stripProjectPrefix(hiddenWorkStem)}隐蔽工程质量验收记录`;
  }

  if (template.kind === "docx" && template.originalName.includes("报验申请单")) {
    return `${stripProjectPrefix(acceptanceStem)}报验申请单`;
  }

  if (template.kind === "xlsx" && template.originalName.includes("检验批质量验收记录")) {
    return `${stripProjectPrefix(acceptanceStem).replace(/\s*分项工程\s*$/, "").replace(/\s*检验批质量验收记录\s*$/, "")}检验批质量验收记录`;
  }

  if (template.kind === "xlsx" && template.originalName.includes("质量验收记录")) {
    return `${stripProjectPrefix(acceptanceStem)}${template.originalName.includes("汇总用") ? "质量验收记录（汇总用）" : "质量验收记录"}`;
  }

  if (template.kind === "xlsx" && /质量(?:检查)?验收评定表/.test(template.originalName)) {
    return `${stripProjectPrefix(acceptanceStem)}${template.originalName.includes("检查验收评定表") ? "质量检查验收评定表" : "质量验收评定表"}`;
  }

  if (template.kind === "xlsx" && template.originalName.includes("质量验收表") && cleanTitle.includes("分项工程验收记录")) {
    return stripProjectPrefix(cleanTitle).replace(/\s*分项工程验收记录\s*$/, "分项工程质量验收表");
  }

  if (template.kind === "xlsx" && template.originalName.includes("质量验收表")) {
    return `${stripProjectPrefix(acceptanceStem)}质量验收表`;
  }

  return cleanTitle;
}
