import type { ArchiveRecord } from "../types";
import { PROCESS_RECORD_KEYWORDS, PROCESS_TEMPLATE_ROOT } from "./constants";
import type { ProcessTemplate, ProcessTemplateManifest } from "./types";

export async function loadProcessManifest(): Promise<ProcessTemplateManifest> {
  const response = await fetch(`${PROCESS_TEMPLATE_ROOT}/manifest.json`);
  if (!response.ok) {
    throw new Error("无法加载过程资料模板清单");
  }

  return response.json();
}

export async function loadProcessTemplate(templateFile: string): Promise<ArrayBuffer> {
  const response = await fetch(`${PROCESS_TEMPLATE_ROOT}/${templateFile}`);
  if (!response.ok) {
    throw new Error(`无法加载过程资料模板：${templateFile}`);
  }

  return response.arrayBuffer();
}

export function findStartReportTemplate(templates: ProcessTemplate[]): ProcessTemplate | undefined {
  return templates.find((template) => template.templateFile === "开工报审.docx")
    ?? templates.find((template) => template.kind === "docx" && template.originalName.includes("开工报审"));
}

export function findSubunitQualityTemplate(templates: ProcessTemplate[]): ProcessTemplate | undefined {
  return templates.find((template) => template.templateFile === "子单位工程质量验收记录.xlsx")
    ?? templates.find((template) => template.kind === "xlsx" && template.originalName.includes("子单位工程质量验收记录"));
}

export function isStartReportItemTitle(title: string): boolean {
  return /开工报审表?|开工报审/.test(title);
}

export function isSubunitQualityItemTitle(title: string): boolean {
  return /子单位(?:工程)?/.test(title) && /质量(?:报验申请|报审表)及验收记录/.test(title);
}

export function getProcessRecordApplicability(record: ArchiveRecord): { isApplicable: boolean; matchedKeywords: string[] } {
  const signalText = [record.fullTitle, ...record.items.map((item) => `${item.fileCode} ${item.title}`)].join(" ");
  const matchedKeywords = PROCESS_RECORD_KEYWORDS.filter((keyword) => signalText.includes(keyword));
  return {
    isApplicable: matchedKeywords.length > 0,
    matchedKeywords,
  };
}

export function isSummaryWorkbookTemplate(template: ProcessTemplate): boolean {
  return template.kind === "xlsx" && template.originalName.includes("汇总用");
}

export function isSubunitQualityTemplate(template: ProcessTemplate): boolean {
  return template.templateFile === "子单位工程质量验收记录.xlsx";
}
