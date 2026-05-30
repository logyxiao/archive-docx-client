import type { ArchiveItem } from "../types";
import { subunitProjectName } from "./textReplacement";
import type { ProcessTemplate } from "./types";
import { sanitizeFileName } from "./utils";

export function processOutputName(template: ProcessTemplate, item: ArchiveItem): string {
  const extension = template.outputExtension;
  if (template.templateFile === "子单位工程报验申请单.docx") {
    const fileCode = item.fileCode && item.fileCode !== "/" ? item.fileCode : "";
    return sanitizeFileName(`${item.sequence || template.sequence}、${fileCode}${subunitProjectName(item.title)}报验申请单${extension}`);
  }

  const originalStem = template.originalName.replace(/\.(docx|xls)$/i, "").replace(/^\d+、/, "");
  const withFileCode = item.fileCode && item.fileCode !== "/"
    ? originalStem.replace(/5028G01-[A-Z0-9-]+-\d{2,}/, item.fileCode)
    : originalStem;
  return sanitizeFileName(`${item.sequence || template.sequence}、${withFileCode}${extension}`);
}
