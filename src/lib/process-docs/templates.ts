import type { ArchiveItem, ArchiveRecord } from "../types";
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
  const response = await fetch(`${PROCESS_TEMPLATE_ROOT}/${encodeURIComponent(templateFile)}`);
  if (!response.ok) {
    throw new Error(`无法加载过程资料模板：${templateFile}`);
  }

  return response.arrayBuffer();
}

export function isStartReportItemTitle(title: string): boolean {
  return /开工报审表?|开工报审/.test(title);
}

export function isSubunitQualityItemTitle(title: string): boolean {
  return /子单位(?:工程)?/.test(title) && /质量(?:报验申请|报审表)及验收记录/.test(title);
}

export function matchingTemplatesByTitle(item: ArchiveItem, templates: ProcessTemplate[]): ProcessTemplate[] {
  const title = item.title;
  if (isStartReportItemTitle(title)) {
    return findByTemplateFiles(templates, ["开工报审.docx"]);
  }

  if (isSubunitQualityItemTitle(title)) {
    return findByTemplateFiles(templates, ["子单位工程报验申请单.docx", "子单位工程质量验收记录.xlsx"]);
  }

  if (isDivisionQualityApplicationTitle(title)) {
    return matchingDivisionTemplates(title, templates);
  }

  if (isSubitemQualityApplicationTitle(title)) {
    return matchingSubitemTemplates(title, templates);
  }

  if (title.includes("检验批质量验收记录")) {
    return matchingInspectionLotTemplates(title, templates);
  }

  return matchingConstructionRecordTemplates(title, templates);
}

function matchingDivisionTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  if (title.includes("子方阵支架及组件安装")) {
    return findByTemplateFiles(templates, ["分部工程报验申请单.docx", "子方阵支架及组件安装分部工程质量验收记录（汇总用）.xlsx"]);
  }
  if (title.includes("通用工程")) {
    return findByTemplateFiles(templates, ["通用工程分部工程报验申请单.docx", "通用工程分部工程质量验收记录（汇总用）.xlsx"]);
  }
  if (title.includes("主体结构")) {
    return findByTemplateFiles(templates, ["主体结构工程分部工程报验申请单.docx", "主体结构工程分部工程质量验收记录（汇总用）.xlsx"]);
  }
  return findByTemplateFiles(templates, ["分部工程报验申请单.docx", "子方阵支架及组件安装分部工程质量验收记录（汇总用）.xlsx"]);
}

function matchingSubitemTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  const templateFiles = ["分项工程报验申请单.docx"];
  if (title.includes("支架安装")) {
    templateFiles.push(`${buildingTemplatePrefix(title)}支架安装分项工程质量验收表.xlsx`);
    return findByTemplateFiles(templates, templateFiles);
  }
  if (title.includes("光伏组件安装")) {
    templateFiles.push(`${buildingTemplatePrefix(title)}光伏组件安装分项工程质量验收表.xlsx`);
    return findByTemplateFiles(templates, templateFiles);
  }
  if (title.includes("建筑电气")) {
    templateFiles.push("建筑电气工程分项工程质量验收记录.xlsx");
    return findByTemplateFiles(templates, templateFiles);
  }
  if (isElectricalSubitemTitle(title)) {
    templateFiles.push("建筑电气工程分项工程质量验收记录.xlsx", ...inspectionLotTemplateFilesForTitle(title));
    return findByTemplateFiles(templates, templateFiles);
  }
  if (title.includes("紧固件连接")) {
    templateFiles.push("紧固件连接分项工程质量验收记录.xlsx");
    return findByTemplateFiles(templates, templateFiles);
  }
  if (title.includes("组装")) {
    templateFiles.push("墙架檩条支撑系统组装分项工程质量验收记录.xlsx");
    return findByTemplateFiles(templates, templateFiles);
  }
  return [];
}

function matchingInspectionLotTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  return findByTemplateFiles(templates, inspectionLotTemplateFilesForTitle(title));
}

function inspectionLotTemplateFilesForTitle(title: string): string[] {
  const building = buildingTemplatePrefix(title);
  if (title.includes("接地装置安装")) {
    return [`${building}接地装置安装检验批质量验收记录.xlsx`];
  }
  if (title.includes("普通紧固件连接") || /紧固件连接工程检验批质量验收记录/.test(title)) {
    return [`${building}普通紧固件连接工程检验批质量验收记录.xlsx`];
  }
  if (title.includes("高强度螺栓")) {
    return [`${building}高强度螺栓连接工程检验批质量验收记录.xlsx`];
  }
  if (title.includes("墙架") || title.includes("檩条") || title.includes("支撑系统")) {
    return [`${building}墙架檩条支撑系统组装工程检验批质量验收记录.xlsx`];
  }
  return [];
}

function matchingConstructionRecordTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  if (title.includes("光伏组件支架防腐记录")) {
    return findByTemplateFiles(templates, ["光伏组件支架防腐记录.xlsx"]);
  }
  if (title.includes("组串开路电压") || title.includes("短路电流测量记录")) {
    return findByTemplateFiles(templates, ["组串开路电压短路电流测量记录.xlsx"]);
  }
  if (title.includes("组件安装施工记录")) {
    return findByTemplateFiles(templates, ["组件安装施工记录.xlsx"]);
  }
  if (title.includes("组件边缘高差测量记录")) {
    return findByTemplateFiles(templates, ["组件边缘高差测量记录.xlsx"]);
  }
  if (title.includes("组件倾斜角度偏差测量记录")) {
    return findByTemplateFiles(templates, ["组件倾斜角度偏差测量记录.xlsx"]);
  }
  if (title.includes("组件接地检查记录")) {
    return findByTemplateFiles(templates, ["组件接地检查记录.xlsx"]);
  }
  return [];
}

function isDivisionQualityApplicationTitle(title: string): boolean {
  return /分部工程质量(?:报验申请|报审表)及验收记录/.test(title) && !title.includes("子单位");
}

function isSubitemQualityApplicationTitle(title: string): boolean {
  return /分项工程质量(?:报验申请|报审表)及验收记录/.test(title);
}

function isElectricalSubitemTitle(title: string): boolean {
  return /逆变器安装|干式变压器安装|接地装置安装|电缆桥架安装|电缆敷设|电缆防火阻燃|开关柜/.test(title);
}

function findByTemplateFiles(templates: ProcessTemplate[], templateFiles: string[]): ProcessTemplate[] {
  const wanted = new Set(templateFiles);
  return templates.filter((template) => wanted.has(template.templateFile));
}

function buildingTemplatePrefix(title: string): string {
  return title.includes("综合楼") ? "综合楼" : "厂房";
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
