import type { ArchiveItem } from "../types";
import type { ProcessTemplate } from "./types";
import { sanitizeFileName } from "./utils";

export function processOutputName(template: ProcessTemplate, item: ArchiveItem): string {
  const extension = template.outputExtension;
  const originalStem = template.originalName.replace(/\.(docx|xls)$/i, "").replace(/^\d+、/, "");
  const withFileCode = item.fileCode && item.fileCode !== "/"
    ? originalStem.replace(/5028G01-[A-Z0-9-]+-\d{2,}/, item.fileCode)
    : originalStem;
  return sanitizeFileName(`${item.sequence || template.sequence}、${withFileCode}${extension}`);
}
