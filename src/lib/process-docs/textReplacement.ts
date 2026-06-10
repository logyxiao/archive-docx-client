import type { ArchiveItem, ArchiveRecord } from "../types";
import { PROJECT_NAME_PATTERNS } from "./constants";
import { resolveProcessFields } from "./fields";
import type { ProcessUserFields } from "./types";
import { archiveProjectCode, formatChineseDate } from "./utils";

export function replaceBusinessText(value: string, record: ArchiveRecord, item: ArchiveItem, userFields: ProcessUserFields): string {
  const fields = resolveProcessFields(userFields, item.owner || record.filingUnit);
  const projectName = normalizeProjectName(userFields.projectName?.trim() || record.projectName);
  let result = replaceTemplatePlaceholders(value, record, item, userFields);
  for (const pattern of PROJECT_NAME_PATTERNS) {
    result = result.replace(pattern, projectName);
  }
  result = result.replace(/MWP/g, "MWp");

  result = result.replace(/5028G01-[A-Z0-9-]+-\d{2,}/g, (match) => {
    return item.fileCode && item.fileCode !== "/" ? item.fileCode : match;
  });
  result = result.replace(/5028G01-0011/g, archiveProjectCode(record.archiveCode));
  result = result.replace(/中核华辰建筑工程有限公司/g, fields.constructionUnit || fields.generalContractorUnit);

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

export function replaceTemplatePlaceholders(
  value: string,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields,
): string {
  const fields = resolveProcessFields(userFields, item.owner || record.filingUnit);
  const projectName = normalizeProjectName(userFields.projectName?.trim() || record.projectName);
  const fileCode = item.fileCode && item.fileCode !== "/" ? item.fileCode : "";
  const replacements: Record<string, string> = {
    项目名: projectName,
    工程名称: projectName,
    工程编号: archiveProjectCode(record.archiveCode),
    档号: record.archiveCode,
    案卷题名: record.fullTitle,
    文件编号: fileCode,
    文件题名: item.title,
    文件日期: item.fileDate,
    文件日期中文: formatChineseDate(item.fileDate),
    责任者: item.owner || record.filingUnit,
    编制单位: item.owner || record.filingUnit,
    立卷单位: record.filingUnit,
    总承包单位: fields.generalContractorUnit,
    总承包单位项目负责人: fields.generalContractorProjectManager,
    总承包单位项目技术负责人: fields.generalContractorTechnicalLeader,
    施工单位: fields.constructionUnit,
    施工单位项目负责人: fields.constructionProjectManager,
    施工单位项目技术负责人: fields.constructionTechnicalLeader,
    分包单位: fields.subcontractorUnit,
    分包单位项目负责人: fields.subcontractorProjectManager,
    分包内容: fields.subcontractorContent,
    监理单位: fields.supervisionDepartment,
    监理项目部: fields.supervisionDepartment,
    抽样单位: item.owner || record.filingUnit,
    抽样日期: formatChineseDate(item.fileDate),
  };

  return value.replace(/\{\{([^{}]+)\}\}/g, (match, rawKey: string) => {
    const key = rawKey.trim();
    return Object.prototype.hasOwnProperty.call(replacements, key) ? replacements[key] : match;
  });
}

export function startReportScope(projectName: string, title: string): string {
  const suffix = startReportSuffix(title);
  return suffix ? `${projectName} ${suffix}` : "";
}

function normalizeProjectName(value: string): string {
  return value.replace(/MWP/g, "MWp");
}

export function subunitProjectName(title: string): string {
  return title
    .replace(/^\s*\d+[、.．\-\s]*/, "")
    .replace(/^.*?项目\s*/, "")
    .replace(/\s*质量(?:报验申请|报审表)及验收记录\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .replace(/\s*开工报审表?\s*$/, "")
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

function startReportSuffix(title: string): string {
  return stripProjectPrefix(title)
    .replace(/\s*开工报审表?\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();
}
