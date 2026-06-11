import type {
  ProcessTemplate,
  ProcessTemplateCategoryId,
  ProcessTemplateMatch,
  ProcessTemplateMatchMode,
  ProcessTemplateModule,
} from "../lib/processDocs";
import type { ArchiveRecord } from "../lib/types";

export interface AllProcessTemplateOption {
  key: string;
  label: string;
  searchText: string;
  match?: ProcessTemplateMatch;
}

export interface TemplateDirectoryEntry {
  key: string;
  name: string;
  path: string;
  source: string;
  kind: string;
  searchText: string;
}

export interface UserTemplateDraft {
  displayName: string;
  matchKeywordsText: string;
  matchMode: ProcessTemplateMatchMode;
  templateModule: ProcessTemplateModule;
  category: ProcessTemplateCategoryId | "";
  enabled: boolean;
}

export function allProcessRowKey(archiveCode: string, fileCode: string, sequence: string): string {
  return `${archiveCode}\u0000${fileCode}\u0000${sequence}`;
}

export function processTemplateOptionKey(match: ProcessTemplateMatch): string {
  return [
    match.templateModule,
    match.template.templateFile,
    match.template.originalName,
    match.template.userTemplatePath ?? "",
    match.template.outputFileCodeOverride ?? "",
  ].join("\u0000");
}

export function processTemplateOptionLabel(match: ProcessTemplateMatch): string {
  const code = match.template.outputFileCodeOverride ? `（${match.template.outputFileCodeOverride}）` : "";
  const source = match.template.userTemplatePath ? "导入模板" : processModuleLabel(match.templateModule);
  return `${match.template.templateFile}${code} · ${source}`;
}

export function processTemplateSearchText(match: ProcessTemplateMatch): string {
  return `${processTemplateOptionLabel(match)} ${match.template.originalName}`.toLowerCase();
}

export function templateMatchFromKey(
  key: string | undefined,
  options: AllProcessTemplateOption[],
): ProcessTemplateMatch | undefined {
  return options.find((option) => option.key === key)?.match;
}

export function archiveShortCode(record: ArchiveRecord): string {
  const parts = record.archiveCode.split("-");
  const categoryCode = record.categoryCode || parts[parts.length - 2] || "";
  const sequence = parts[parts.length - 1] || "";
  return [categoryCode, sequence].filter(Boolean).join("-");
}

export function processModuleLabel(templateModule: ProcessTemplateModule): string {
  if (templateModule === "switch-station") {
    return "开关站电气设备安装（子单位工程）";
  }
  if (templateModule === "collector-line") {
    return "集电线路安装工程";
  }
  return "过程资料";
}

export function templateIdentityKey(template: ProcessTemplate): string {
  return [
    template.userTemplatePath ?? "builtin",
    template.templateFile,
    template.originalName,
    template.outputFileCodeOverride ?? "",
  ].join("\u0000");
}

export function userTemplateKey(template: ProcessTemplate): string {
  return template.userTemplatePath ?? templateIdentityKey(template);
}

export function draftFromUserTemplate(template: ProcessTemplate): UserTemplateDraft {
  return {
    displayName: template.displayName?.trim() || template.originalName || template.templateFile,
    matchKeywordsText: (template.matchKeywords ?? []).join("\n"),
    matchMode: template.matchMode === "all" ? "all" : "any",
    templateModule: template.templateModule ?? "process",
    category: template.category ?? "",
    enabled: template.enabled !== false,
  };
}

export function parseTemplateKeywords(value: string): string[] {
  const result: string[] = [];
  for (const keyword of value.split(/[\n,，;；]+/)) {
    const trimmed = keyword.trim();
    if (!trimmed || result.includes(trimmed)) {
      continue;
    }
    result.push(trimmed);
  }
  return result;
}

export function joinTemplatePath(dir: string, fileName: string): string {
  if (!dir) {
    return "";
  }

  const separator = dir.includes("\\") ? "\\" : "/";
  return `${dir.replace(/[\\/]+$/, "")}${separator}${fileName}`;
}

export function parentDirectory(path: string): string {
  const normalized = path.replace(/[\\/]+$/, "");
  const index = Math.max(normalized.lastIndexOf("/"), normalized.lastIndexOf("\\"));
  return index > 0 ? normalized.slice(0, index) : normalized;
}
