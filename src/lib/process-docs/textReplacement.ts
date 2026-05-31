import type { ArchiveItem, ArchiveRecord } from "../types";
import { PROJECT_NAME_PATTERNS } from "./constants";
import { resolveProcessFields } from "./fields";
import type { ProcessUserFields } from "./types";
import { archiveProjectCode } from "./utils";

export function replaceBusinessText(value: string, record: ArchiveRecord, item: ArchiveItem, userFields: ProcessUserFields): string {
  const fields = resolveProcessFields(userFields, item.owner || record.filingUnit);
  const projectName = userFields.projectName?.trim() || record.projectName;
  let result = value;
  for (const pattern of PROJECT_NAME_PATTERNS) {
    result = result.replace(pattern, projectName);
  }

  result = result.replace(/5028G01-[A-Z0-9-]+-\d{2,}/g, (match) => {
    return item.fileCode && item.fileCode !== "/" ? item.fileCode : match;
  });
  result = result.replace(/5028G01-0011/g, archiveProjectCode(record.archiveCode));
  result = result.replace(/中核华辰建筑工程有限公司/g, fields.constructionUnit || fields.generalContractorUnit);
  result = replaceStartReportScope(result, projectName, item.title);

  const sourceMissingReplacements: Record<string, string> = {
    河南中核五院研究设计有限公司: fields.supervisionDepartment,
    谢智敏: fields.constructionProjectManager,
    蒋志炜: fields.generalContractorTechnicalLeader,
    刘彦堂: fields.constructionTechnicalLeader,
    河南誉华美建设工程有限公司: fields.subcontractorUnit,
    胡彦芳: fields.subcontractorProjectManager,
  };

  for (const [missing, replacement] of Object.entries(sourceMissingReplacements)) {
    result = result.split(missing).join(replacement);
  }

  return result;
}

export function startReportScope(projectName: string, title: string): string {
  const suffix = startReportSuffix(title);
  return suffix ? `${projectName} ${suffix}` : "";
}

export function subunitProjectName(title: string): string {
  return title
    .replace(/^\s*\d+[、.．\-\s]*/, "")
    .replace(/^.*?项目\s*/, "")
    .replace(/\s*质量(?:报验申请|报审表)及验收记录\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();
}

export function subunitInspectionSubject(title: string): string {
  return subunitProjectName(title).replace(/\s*子单位工程\s*$/, "").trim();
}

export function inspectionApplicationFullSubject(title: string): string {
  return stripProjectPrefix(title)
    .replace(/\s*质量(?:报验申请|报审表)及验收记录\s*$/, "")
    .replace(/\s*验收记录\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();
}

export function inspectionApplicationSubject(title: string): string {
  return inspectionApplicationFullSubject(title)
    .replace(/\s*子单位工程\s*$/, "")
    .replace(/\s*分部工程\s*$/, "")
    .replace(/\s*分项工程\s*$/, "")
    .trim();
}

export function stripProjectPrefix(title: string): string {
  return title
    .replace(/^\s*\d+[、.．\-\s]*/, "")
    .replace(/^.*?项目\s*/, "")
    .trim();
}

export function replaceStartReportScopeText(text: string, projectName: string, title: string): string {
  const scope = startReportScope(projectName, title);
  if (!scope) {
    return text;
  }

  return text.replace(/我方承担的\s*[^，,]*?，已完成了/g, `我方承担的 ${scope} ，已完成了`);
}

function replaceStartReportScope(text: string, projectName: string, title: string): string {
  return replaceStartReportScopeText(text, projectName, title);
}

function startReportSuffix(title: string): string {
  return stripProjectPrefix(title)
    .replace(/\s*开工报审表?\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();
}
