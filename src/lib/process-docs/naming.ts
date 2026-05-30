import type { ArchiveItem } from "../types";
import { stripProjectPrefix } from "./textReplacement";
import type { ProcessTemplate } from "./types";
import { sanitizeFileName } from "./utils";

export function processOutputName(template: ProcessTemplate, item: ArchiveItem): string {
  const extension = template.outputExtension;
  const sequence = item.sequence || String(template.sequence);
  const fileCode = item.fileCode && item.fileCode !== "/" ? item.fileCode : "";
  const title = outputTitle(template, item.title);

  return sanitizeFileName(`${sequence}、${fileCode}${title}${extension}`);
}

function outputTitle(template: ProcessTemplate, title: string): string {
  const cleanTitle = title.replace(/^\s*\d+[、.．\-\s]*/, "").trim();
  const acceptanceStem = cleanTitle
    .replace(/\s*质量(?:报验申请|报审表)及验收记录\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();

  if (template.kind === "docx" && template.originalName.includes("报验申请单")) {
    return `${stripProjectPrefix(acceptanceStem)}报验申请单`;
  }

  if (template.kind === "xlsx" && template.originalName.includes("检验批质量验收记录")) {
    return `${stripProjectPrefix(acceptanceStem).replace(/\s*分项工程\s*$/, "").replace(/\s*检验批质量验收记录\s*$/, "")}检验批质量验收记录`;
  }

  if (template.kind === "xlsx" && template.originalName.includes("质量验收记录")) {
    return `${stripProjectPrefix(acceptanceStem)}${template.originalName.includes("汇总用") ? "质量验收记录（汇总用）" : "质量验收记录"}`;
  }

  return cleanTitle;
}
