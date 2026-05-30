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

export function groupTemplatesBySequence(templates: ProcessTemplate[]): Map<number, ProcessTemplate[]> {
  const grouped = new Map<number, ProcessTemplate[]>();
  for (const template of templates) {
    const group = grouped.get(template.sequence) ?? [];
    group.push(template);
    grouped.set(template.sequence, group);
  }
  return grouped;
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
