import type { ArchiveItem, ArchiveRecord } from "../types";
import { PROJECT_NAME_PATTERNS } from "./constants";
import { resolveProcessFields } from "./fields";
import type { ProcessUserFields } from "./types";
import { archiveProjectCode } from "./utils";

export function replaceBusinessText(value: string, record: ArchiveRecord, item: ArchiveItem, userFields: ProcessUserFields): string {
  const fields = resolveProcessFields(userFields, item.owner || record.filingUnit);
  let result = value;
  for (const pattern of PROJECT_NAME_PATTERNS) {
    result = result.replace(pattern, record.projectName);
  }

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
