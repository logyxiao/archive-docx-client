import type { ProcessTemplate } from "./types";

export const PROCESS_TEMPLATE_CATEGORY_IDS = [
  "start-report",
  "subunit-inspection-application",
  "summary-quality-acceptance",
  "division-inspection-application",
  "division-quality-acceptance",
  "subitem-inspection-application",
  "subitem-quality-acceptance",
  "inspection-lot-acceptance",
  "construction-record",
  "other",
] as const;

export type ProcessTemplateCategoryId = (typeof PROCESS_TEMPLATE_CATEGORY_IDS)[number];

export interface ProcessTemplateCategoryOption {
  id: ProcessTemplateCategoryId;
  label: string;
}

export const PROCESS_TEMPLATE_CATEGORIES: ProcessTemplateCategoryOption[] = [
  { id: "start-report", label: "开工报审" },
  { id: "subunit-inspection-application", label: "子单位工程报验申请单" },
  { id: "summary-quality-acceptance", label: "子单位工程质量验收记录" },
  { id: "division-inspection-application", label: "分部工程报验申请单" },
  { id: "division-quality-acceptance", label: "分部工程质量验收记录" },
  { id: "subitem-inspection-application", label: "分项工程报验申请单" },
  { id: "subitem-quality-acceptance", label: "分项工程质量验收记录" },
  { id: "inspection-lot-acceptance", label: "检验批质量验收记录" },
  { id: "construction-record", label: "施工/测量/检查记录" },
  { id: "other", label: "其它模板" },
];

export function normalizeProcessTemplateCategories(
  categoryIds: readonly string[] | undefined,
): Set<ProcessTemplateCategoryId> {
  const validIds = new Set<ProcessTemplateCategoryId>(PROCESS_TEMPLATE_CATEGORY_IDS);
  const normalized = expandLegacyCategoryIds(categoryIds ?? PROCESS_TEMPLATE_CATEGORY_IDS)
    .filter((id) => validIds.has(id));
  return new Set(normalized);
}

function expandLegacyCategoryIds(categoryIds: readonly string[]): ProcessTemplateCategoryId[] {
  const expanded: ProcessTemplateCategoryId[] = [];
  for (const id of categoryIds) {
    if (id === "inspection-application") {
      expanded.push("division-inspection-application", "subitem-inspection-application");
      continue;
    }
    if (id === "quality-acceptance") {
      expanded.push("division-quality-acceptance", "subitem-quality-acceptance");
      continue;
    }
    expanded.push(id as ProcessTemplateCategoryId);
  }
  return expanded;
}

export function getProcessTemplateCategory(template: ProcessTemplate): ProcessTemplateCategoryId {
  const name = template.originalName;

  if (name.includes("开工报审")) {
    return "start-report";
  }

  if (template.templateFile === "子单位工程报验申请单.docx") {
    return "subunit-inspection-application";
  }

  if (name.includes("子单位") && name.includes("质量验收记录")) {
    return "summary-quality-acceptance";
  }

  if (name.includes("分部") && name.includes("报验申请单")) {
    return "division-inspection-application";
  }

  if (name.includes("分部") && /质量(?:检查)?验收(?:记录|评定表)/.test(name)) {
    return "division-quality-acceptance";
  }

  if (template.templateFile === "分项工程报验申请单.docx" || (name.includes("分项") && name.includes("报验申请单"))) {
    return "subitem-inspection-application";
  }

  if (name.includes("检验批质量验收记录")) {
    return "inspection-lot-acceptance";
  }

  if (name.includes("分项") && /质量(?:检查)?验收(?:记录|表|评定表)/.test(name)) {
    return "subitem-quality-acceptance";
  }

  if (/施工记录|测量记录|检查记录|防腐记录|短路电流|开路电压/.test(name)) {
    return "construction-record";
  }

  return "other";
}

export function isProcessTemplateCategorySelected(
  template: ProcessTemplate,
  selectedCategories: Set<ProcessTemplateCategoryId>,
): boolean {
  return selectedCategories.has(getProcessTemplateCategory(template));
}
