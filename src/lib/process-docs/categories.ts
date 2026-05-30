import type { ProcessTemplate } from "./types";

export const PROCESS_TEMPLATE_CATEGORY_IDS = [
  "start-report",
  "summary-quality-acceptance",
  "inspection-application",
  "quality-acceptance",
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
  { id: "summary-quality-acceptance", label: "子单位工程质量验收记录" },
  { id: "inspection-application", label: "报验申请单" },
  { id: "quality-acceptance", label: "分项/分部质量验收" },
  { id: "inspection-lot-acceptance", label: "检验批质量验收记录" },
  { id: "construction-record", label: "施工/测量/检查记录" },
  { id: "other", label: "其它模板" },
];

export function normalizeProcessTemplateCategories(
  categoryIds: ProcessTemplateCategoryId[] | undefined,
): Set<ProcessTemplateCategoryId> {
  const validIds = new Set<ProcessTemplateCategoryId>(PROCESS_TEMPLATE_CATEGORY_IDS);
  const normalized = (categoryIds ?? PROCESS_TEMPLATE_CATEGORY_IDS).filter((id) => validIds.has(id));
  return new Set(normalized);
}

export function getProcessTemplateCategory(template: ProcessTemplate): ProcessTemplateCategoryId {
  const name = template.originalName;

  if (name.includes("开工报审")) {
    return "start-report";
  }

  if (template.templateFile === "子单位工程质量验收记录.xlsx") {
    return "summary-quality-acceptance";
  }

  if (name.includes("报验申请")) {
    return "inspection-application";
  }

  if (name.includes("检验批质量验收记录")) {
    return "inspection-lot-acceptance";
  }

  if (name.includes("质量验收")) {
    return "quality-acceptance";
  }

  if (/施工记录|测量记录|检查记录|防腐记录|短路电流|开路电压|接地/.test(name)) {
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
