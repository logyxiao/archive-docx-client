import type { ArchiveItem, ArchiveRecord } from "../types";
import { loadUserProcessTemplates, readBinaryFile } from "../tauriFiles";
import { PROCESS_RECORD_KEYWORDS, PROCESS_TEMPLATE_ROOT } from "./constants";
import { collectorLineTemplatesForItem, isCollectorLineRecord, isCollectorLineTemplate } from "./collectorLine";
import { isSwitchStationArchiveRecord } from "./switchStation";
import type { ProcessTemplate, ProcessTemplateManifest, ProcessTemplateModule } from "./types";

export async function loadProcessManifest(): Promise<ProcessTemplateManifest> {
  const response = await fetch(`${PROCESS_TEMPLATE_ROOT}/manifest.json`);
  if (!response.ok) {
    throw new Error("无法加载过程资料模板清单");
  }

  const manifest = await response.json() as ProcessTemplateManifest;
  let userTemplates: ProcessTemplate[] = [];
  try {
    userTemplates = await loadUserProcessTemplates();
  } catch {
    userTemplates = [];
  }

  return {
    templates: mergeProcessTemplates(manifest.templates, userTemplates),
  };
}

export async function loadProcessTemplate(template: ProcessTemplate | string): Promise<ArrayBuffer> {
  if (typeof template !== "string" && template.userTemplatePath) {
    const bytes = await readBinaryFile(template.userTemplatePath);
    return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  }

  const templateFile = typeof template === "string" ? template : template.templateFile;
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
  return /子单位(?:工程)?/.test(title) && /质量(?:报验申请|报审表)(?:及验收记录)?/.test(title);
}

export function matchingTemplatesByTitle(
  item: ArchiveItem,
  templates: ProcessTemplate[],
  templateModule: ProcessTemplateModule = "process",
): ProcessTemplate[] {
  const title = item.title;
  const activeTemplates = templates.filter((template) => isTemplateInModule(template, templateModule));
  const userTemplates = matchingUserTemplates(title, activeTemplates);

  if (isStartReportItemTitle(title)) {
    return withUserTemplates(findByTemplateFiles(activeTemplates, ["开工报审.docx"]), userTemplates);
  }

  if (templateModule === "collector-line") {
    return withUserTemplates(findByTemplateFiles(activeTemplates, collectorLineTemplatesForItem(item)), userTemplates);
  }

  if (isSubunitQualityItemTitle(title)) {
    const qualityTemplate = title.includes("并网点光伏变电系统")
      ? "并网点光伏变电系统子单位工程质量验收记录（汇总用）.xlsx"
      : title.includes("开关站电气设备安装")
      ? "开关站电气设备安装子单位工程质量验收记录（汇总用）.xlsx"
      : "子单位工程质量验收记录.xlsx";
    const templateFiles = ["子单位工程报验申请单.docx"];
    if (title.includes("及验收记录")) {
      templateFiles.push(qualityTemplate);
    }
    return withUserTemplates(findByTemplateFiles(activeTemplates, templateFiles), userTemplates);
  }

  if (isHiddenWorkQualityTitle(title)) {
    return withUserTemplates(matchingHiddenWorkTemplates(title, activeTemplates), userTemplates);
  }

  if (isDivisionQualityApplicationTitle(title)) {
    return withUserTemplates(matchingDivisionTemplates(title, activeTemplates), userTemplates);
  }

  if (isSubitemQualityApplicationTitle(title) || isSubitemQualityRecordTitle(title)) {
    return withUserTemplates(matchingSubitemTemplates(title, activeTemplates), userTemplates);
  }

  if (title.includes("检验批质量验收记录")) {
    return withUserTemplates(matchingInspectionLotTemplates(title, activeTemplates), userTemplates);
  }

  return withUserTemplates(matchingConstructionRecordTemplates(title, activeTemplates), userTemplates);
}

function mergeProcessTemplates(builtInTemplates: ProcessTemplate[], userTemplates: ProcessTemplate[]): ProcessTemplate[] {
  const result = [...builtInTemplates];
  const seenUserPaths = new Set<string>();
  for (const template of userTemplates) {
    if (!template.userTemplatePath || seenUserPaths.has(template.userTemplatePath)) {
      continue;
    }
    seenUserPaths.add(template.userTemplatePath);
    result.push(template);
  }
  return result;
}

function withUserTemplates(baseTemplates: ProcessTemplate[], userTemplates: ProcessTemplate[]): ProcessTemplate[] {
  if (userTemplates.length === 0) {
    return baseTemplates;
  }

  const result = [...baseTemplates];
  const seen = new Set(result.map(templateIdentity));
  for (const template of userTemplates) {
    const identity = templateIdentity(template);
    if (seen.has(identity)) {
      continue;
    }
    seen.add(identity);
    result.push(template);
  }
  return result;
}

function matchingUserTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  const normalizedTitle = normalizeTemplateMatchText(title);
  if (!normalizedTitle) {
    return [];
  }

  return templates.filter((template) => {
    if (!template.userTemplatePath) {
      return false;
    }
    return [template.templateFile, template.originalName].some((name) => {
      const normalizedTemplateName = normalizeTemplateMatchText(fileNameWithoutExtension(name));
      return normalizedTemplateName.length >= 2
        && (normalizedTitle.includes(normalizedTemplateName) || normalizedTemplateName.includes(normalizedTitle));
    });
  });
}

function fileNameWithoutExtension(value: string): string {
  return value.replace(/\.[^.]+$/, "");
}

function normalizeTemplateMatchText(value: string): string {
  return value
    .replace(/\{\{[^}]+}}/g, "")
    .replace(/^[0-9]+[-_、.\s]*/, "")
    .replace(/（.*?）|\(.*?\)/g, "")
    .replace(/[^\p{Script=Han}A-Za-z0-9]/gu, "")
    .toLowerCase();
}

function templateIdentity(template: ProcessTemplate): string {
  return template.userTemplatePath ?? `${template.templateFile}\u0000${template.originalName}\u0000${template.outputFileCodeOverride ?? ""}`;
}

function matchingDivisionTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  if (title.includes("子方阵电气线路安装")) {
    return findByTemplateFiles(templates, ["分部工程报验申请单.docx", "子方阵电气线路安装分部工程质量验收记录（汇总用）.xlsx"]);
  }
  if (title.includes("子方阵场电气安装") || title.includes("子方阵电气安装")) {
    return findByTemplateFiles(templates, ["分部工程报验申请单.docx", "子方阵电气安装分部工程质量验收记录（汇总用）.xlsx"]);
  }
  if (title.includes("子方阵")) {
    return findByTemplateFiles(templates, ["分部工程报验申请单.docx", "子方阵支架及组件安装分部工程质量验收记录（汇总用）.xlsx"]);
  }
  if (title.includes("通用工程")) {
    return findByTemplateFiles(templates, ["通用工程分部工程报验申请单.docx", "通用工程分部工程质量验收记录（汇总用）.xlsx"]);
  }
  if (title.includes("主体结构")) {
    return findByTemplateFiles(templates, ["主体结构工程分部工程报验申请单.docx", "主体结构工程分部工程质量验收记录（汇总用）.xlsx"]);
  }

  const templateFile = firstMatchingTemplateFile(title, [
    ["地基与基础", "地基与基础工程分部工程质量验收记录（汇总用）.xlsx"],
    ["控制、保护及交直流控制电源系统设备安装", "控制保护及交直流控制电源系统设备安装分部工程质量验收评定表.xlsx"],
    ["交直流控制电源系统设备安装", "交直流控制电源系统设备安装分部工程质量验收评定表.xlsx"],
    ["接地装置安装", "接地装置安装分部工程质量检查验收评定表.xlsx"],
    ["计算机监控系统设备安装", "计算机监控系统设备安装分部工程质量验收评定表.xlsx"],
    ["通讯系统设备安装", "通讯系统设备安装分部工程质量验收评定表.xlsx"],
    ["通信系统设备安装", "通讯系统设备安装分部工程质量验收评定表.xlsx"],
    ["开关站电气设备安装", "开关站电气设备安装分部工程质量验收评定表.xlsx"],
  ]);

  return templateFile ? findByTemplateFiles(templates, ["分部工程报验申请单.docx", templateFile]) : [];
}

function matchingSubitemTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  const applicationTemplateFiles = isSubitemQualityApplicationTitle(title) ? ["分项工程报验申请单.docx"] : [];
  const processTemplate = firstMatchingTemplateFile(title, [
    ["逆变器安装", "逆变器安装分项工程质量验收表.xlsx"],
    ["干式变压器安装", "干式变压器安装分项工程质量验收表.xlsx"],
    ["屋外接地装置安装", "屋外接地装置安装分项工程质量验收表.xlsx"],
    ["屋内接地装置安装", "屋内接地装置安装分项工程质量验收表.xlsx"],
    ["电缆桥架安装", "电缆桥架安装分项质量检查验收评定表.xlsx"],
    ["电缆敷设分项工程质量", "电缆敷设分项工程质量检查验收评定表.xlsx"],
    ["电缆终端制作", "电力电缆终端制作安装分项工程质量验收表.xlsx"],
    ["电缆防火与阻燃", "光伏升压变电缆防火阻燃施工分项工程质量验收表.xlsx"],
    ["电气二次系统", "二次回路检查及控制电缆接线分项工程质量检查验收评定表.xlsx"],
    ["光伏升压变电缆防火阻燃施工", "光伏升压变电缆防火阻燃施工分项工程质量验收表.xlsx"],
    ["开关柜G01柜电缆防火阻燃施工", "开关柜电缆防火阻燃施工分项工程质量验收表.xlsx"],
    ["开关柜G03柜电缆防火阻燃施工", "开关柜电缆防火阻燃施工分项工程质量验收表.xlsx"],
    ["开关柜G08柜电缆防火阻燃施工", "开关柜电缆防火阻燃施工分项工程质量验收表.xlsx"],
    ["开关柜G09柜电缆防火阻燃施工", "开关柜电缆防火阻燃施工分项工程质量验收表.xlsx"],
  ]);
  if (processTemplate) {
    return findByTemplateFiles(templates, [processTemplate, ...applicationTemplateFiles]);
  }

  const switchStationTemplate = firstMatchingTemplateFile(title, [
    ["定位及高程控制", "定位及高程控制分项工程质量验收记录.xlsx"],
    ["土方", "土方分项工程质量验收记录.xlsx"],
    ["垫层", "垫层分项工程质量验收记录.xlsx"],
    ["基础模板", "基础模板分项工程质量验收记录.xlsx"],
    ["钢筋", "钢筋分项工程质量验收记录.xlsx"],
    ["混凝土", "混凝土分项工程质量验收记录.xlsx"],
    ["开关柜及附属设备安装", "开关柜及附属设备安装分项工程质量验收评定表.xlsx"],
    ["控制及保护屏台", "控制及保护屏台的安装分项工程质量检查验收评定表.xlsx"],
    ["二次回路检查及控制电缆接线", "二次回路检查及控制电缆接线分项工程质量检查验收评定表.xlsx"],
    ["蓄电池安装", "蓄电池安装分项工程质量检查验收评定表.xlsx"],
    ["UPS装置检查及带负荷试验", "UPS装置检查及带负荷试验分项工程质量检查验收评定表.xlsx"],
    ["柜外接地装置安装", "屋外接地装置安装分项工程质量验收表.xlsx"],
    ["柜内接地装置安装", "屋内接地装置安装分项工程质量验收表.xlsx"],
    ["单个盘（台、箱、柜）安装", "单个盘台箱柜安装分项工程质量检查验收评定表.xlsx"],
    ["监视设备安装", "监视设备安装分项工程质量检查验收评定表.xlsx"],
    ["电缆支架、桥架、竖井制作及安装", "电缆支架桥架竖井制作及安装分项工程质量验收表.xlsx"],
    ["电缆敷设分项工程质量", "电缆敷设分项工程质量检查验收评定表.xlsx"],
    ["#1光伏升压变 电缆构筑物", "光伏升压变电缆敷设分项工程质量验收表.xlsx"],
    ["#2光伏升压变 电缆构筑物", "光伏升压变电缆敷设分项工程质量验收表.xlsx"],
    ["开关柜G01柜电缆构筑物", "开关柜电缆敷设分项工程质量验收表.xlsx"],
    ["开关柜G03柜电缆构筑物", "开关柜电缆敷设分项工程质量验收表.xlsx"],
    ["开关柜G08柜电缆构筑物", "开关柜电缆敷设分项工程质量验收表.xlsx"],
    ["开关柜G09柜电缆构筑物", "开关柜电缆敷设分项工程质量验收表.xlsx"],
    ["通信系统一次设备安装", "通信系统一次设备安装分项工程质量检查验收评定表.xlsx"],
    ["通讯系统一次设备安装", "通信系统一次设备安装分项工程质量检查验收评定表.xlsx"],
  ]);
  if (switchStationTemplate) {
    return findByTemplateFiles(templates, [switchStationTemplate, ...applicationTemplateFiles]);
  }

  if (title.includes("支架安装")) {
    return findByTemplateFiles(templates, [`${buildingTemplatePrefix(title)}支架安装分项工程质量验收表.xlsx`, ...applicationTemplateFiles]);
  }
  if (title.includes("光伏组件安装")) {
    return findByTemplateFiles(templates, [`${buildingTemplatePrefix(title)}光伏组件安装分项工程质量验收表.xlsx`, ...applicationTemplateFiles]);
  }
  if (title.includes("建筑电气")) {
    return findByTemplateFiles(templates, ["建筑电气工程分项工程质量验收记录.xlsx", ...applicationTemplateFiles]);
  }
  if (isElectricalSubitemTitle(title)) {
    return findByTemplateFiles(templates, ["建筑电气工程分项工程质量验收记录.xlsx", ...inspectionLotTemplateFilesForTitle(title), ...applicationTemplateFiles]);
  }
  if (title.includes("紧固件连接")) {
    return findByTemplateFiles(templates, ["紧固件连接分项工程质量验收记录.xlsx", ...applicationTemplateFiles]);
  }
  if (title.includes("组装")) {
    return findByTemplateFiles(templates, ["墙架檩条支撑系统组装分项工程质量验收记录.xlsx", ...applicationTemplateFiles]);
  }
  return [];
}

function matchingInspectionLotTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  return findByTemplateFiles(templates, inspectionLotTemplateFilesForTitle(title));
}

function inspectionLotTemplateFilesForTitle(title: string): string[] {
  const switchStationTemplate = firstMatchingTemplateFile(title, [
    ["定位放线", "定位放线工程检验批质量验收记录.xlsx"],
    ["土方开挖", "土方开挖检验批质量验收记录.xlsx"],
    ["土方回填", "土方回填检验批质量验收记录.xlsx"],
    ["场地平整", "场地平整检验批质量验收记录.xlsx"],
    ["水泥混凝土垫层", "水泥混凝土垫层和陶粒土垫层检验批质量验收记录.xlsx"],
    ["陶粒土垫层", "水泥混凝土垫层和陶粒土垫层检验批质量验收记录.xlsx"],
    ["现浇混凝土模板安装", "现浇混凝土模板安装工程检验批质量验收记录.xlsx"],
    ["钢筋安装", "钢筋安装检验批质量验收记录.xlsx"],
    ["混凝土施工", "混凝土施工检验批质量验收记录.xlsx"],
    ["现浇混凝土结构外观及尺寸偏差", "现浇混凝土结构外观及尺寸偏差检验批质量验收记录.xlsx"],
  ]);
  if (switchStationTemplate) {
    return [switchStationTemplate];
  }

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

function matchingHiddenWorkTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  const templateFile = firstMatchingTemplateFile(title, [
    ["模板拆除", "模板拆除隐蔽工程质量验收记录.xlsx"],
    ["钢筋", "钢筋隐蔽工程质量验收记录.xlsx"],
    ["电缆线路施工", "低压交流电缆隐蔽工程质量验收记录.xlsx"],
    ["电缆", "低压交流电缆隐蔽工程质量验收记录.xlsx"],
  ]);

  return findByTemplateFiles(templates, templateFile ? ["隐蔽工程质量报验单.docx", templateFile] : ["隐蔽工程质量报验单.docx"]);
}

function matchingConstructionRecordTemplates(title: string, templates: ProcessTemplate[]): ProcessTemplate[] {
  if (title.includes("逆变器施工安装记录")) {
    return findByTemplateFiles(templates, ["逆变器施工安装记录.xlsx"]);
  }
  if (title.includes("逆变器手动分合闸检查记录")) {
    return findByTemplateFiles(templates, ["逆变器手动分合闸检查记录.xlsx"]);
  }
  if (title.includes("逆变器通讯调试记录")) {
    return findByTemplateFiles(templates, ["逆变器通讯调试记录.xlsx"]);
  }
  if (title.includes("逆变器外观") && title.includes("检查记录")) {
    return findByTemplateFiles(templates, ["逆变器外观、主要元器件、控制电源、直交流侧接线及极性（相序）、绝缘、接地检查记录.xlsx"]);
  }
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
  if (title.includes("配电柜手动分合闸检查记录")) {
    return findByTemplateFiles(templates, ["配电柜手动分合闸检查记录.xlsx"]);
  }
  if (title.includes("逆变器散热装置") || title.includes("人机界面检查记录")) {
    return findByTemplateFiles(templates, ["逆变器散热装置人机界面检查记录.xlsx"]);
  }
  if (title.includes("配电柜施工安装记录")) {
    return findByTemplateFiles(templates, ["配电柜施工安装记录.xlsx"]);
  }
  if (title.includes("配电柜安装接地检查记录")) {
    return findByTemplateFiles(templates, ["配电柜安装接地检查记录.xlsx"]);
  }
  return [];
}

function isDivisionQualityApplicationTitle(title: string): boolean {
  return /分部工程质量(?:报验申请|报审表)及验收记录/.test(title) && !title.includes("子单位");
}

function isSubitemQualityApplicationTitle(title: string): boolean {
  return /分项工程(?:工程)?质量(?:报验申请|报审表)及验收记录/.test(title);
}

function isSubitemQualityRecordTitle(title: string): boolean {
  return /分项工程验收记录$/.test(title);
}

function isHiddenWorkQualityTitle(title: string): boolean {
  return title.includes("隐蔽");
}

function isElectricalSubitemTitle(title: string): boolean {
  return /逆变器安装|干式变压器安装|接地装置安装|电缆桥架安装|电缆敷设|电缆防火阻燃|开关柜/.test(title);
}

function findByTemplateFiles(templates: ProcessTemplate[], templateFiles: string[]): ProcessTemplate[] {
  return templateFiles.flatMap((templateFile) => templates.filter((template) => template.templateFile === templateFile));
}

function buildingTemplatePrefix(title: string): string {
  return title.includes("综合楼") ? "综合楼" : "厂房";
}

function firstMatchingTemplateFile(title: string, pairs: Array<[string, string]>): string | undefined {
  return pairs.find(([keyword]) => title.includes(keyword))?.[1];
}

export function isTemplateInModule(template: ProcessTemplate, templateModule: ProcessTemplateModule): boolean {
  if (templateModule === "switch-station") {
    return COMMON_PROCESS_TEMPLATE_FILES.has(template.templateFile) || isSwitchStationTemplate(template);
  }
  if (templateModule === "collector-line") {
    return COMMON_PROCESS_TEMPLATE_FILES.has(template.templateFile) || isCollectorLineTemplate(template);
  }
  return !isSwitchStationTemplate(template) && !isCollectorLineOnlyTemplate(template);
}

export function isRecordInTemplateModule(record: ArchiveRecord, templateModule: ProcessTemplateModule): boolean {
  if (templateModule === "switch-station") {
    return isSwitchStationArchiveRecord(record);
  }
  if (templateModule === "collector-line") {
    return isCollectorLineRecord(record);
  }
  return !isSwitchStationOnlyRecord(record) && !isCollectorLineOnlyRecord(record);
}

function isSwitchStationOnlyRecord(record: ArchiveRecord): boolean {
  return record.archiveCode.includes("-8341-") || record.archiveCode.includes("-8342-");
}

function isCollectorLineOnlyRecord(record: ArchiveRecord): boolean {
  return record.archiveCode.includes("-8331-");
}

function isSwitchStationTemplate(template: ProcessTemplate): boolean {
  if (COMMON_PROCESS_TEMPLATE_FILES.has(template.templateFile)) {
    return false;
  }
  return SWITCH_STATION_TEMPLATE_FILES.has(template.templateFile);
}

function isCollectorLineOnlyTemplate(template: ProcessTemplate): boolean {
  if (COMMON_PROCESS_TEMPLATE_FILES.has(template.templateFile)) {
    return false;
  }
  return isCollectorLineTemplate(template);
}

const COMMON_PROCESS_TEMPLATE_FILES = new Set([
  "开工报审.docx",
  "子单位工程报验申请单.docx",
  "分部工程报验申请单.docx",
  "分项工程报验申请单.docx",
  "隐蔽工程质量报验单.docx",
  "模板拆除隐蔽工程质量验收记录.xlsx",
  "钢筋隐蔽工程质量验收记录.xlsx",
  "低压交流电缆隐蔽工程质量验收记录.xlsx",
  "屋外接地装置安装分项工程质量验收表.xlsx",
  "屋内接地装置安装分项工程质量验收表.xlsx",
  "电缆敷设分项工程质量检查验收评定表.xlsx",
  "电力电缆终端制作安装分项工程质量验收表.xlsx",
  "二次回路检查及控制电缆接线分项工程质量检查验收评定表.xlsx",
]);

const SWITCH_STATION_TEMPLATE_FILES = new Set([
  "开关站电气设备安装子单位工程质量验收记录（汇总用）.xlsx",
  "地基与基础工程分部工程质量验收记录（汇总用）.xlsx",
  "定位及高程控制分项工程质量验收记录.xlsx",
  "定位放线工程检验批质量验收记录.xlsx",
  "土方分项工程质量验收记录.xlsx",
  "土方开挖检验批质量验收记录.xlsx",
  "土方回填检验批质量验收记录.xlsx",
  "场地平整检验批质量验收记录.xlsx",
  "垫层分项工程质量验收记录.xlsx",
  "水泥混凝土垫层和陶粒土垫层检验批质量验收记录.xlsx",
  "基础模板分项工程质量验收记录.xlsx",
  "现浇混凝土模板安装工程检验批质量验收记录.xlsx",
  "隐蔽工程质量报验单.docx",
  "模板拆除隐蔽工程质量验收记录.xlsx",
  "钢筋分项工程质量验收记录.xlsx",
  "钢筋安装检验批质量验收记录.xlsx",
  "钢筋隐蔽工程质量验收记录.xlsx",
  "混凝土分项工程质量验收记录.xlsx",
  "混凝土施工检验批质量验收记录.xlsx",
  "现浇混凝土结构外观及尺寸偏差检验批质量验收记录.xlsx",
  "开关站电气设备安装分部工程质量验收评定表.xlsx",
  "开关柜及附属设备安装分项工程质量验收评定表.xlsx",
  "控制保护及交直流控制电源系统设备安装分部工程质量验收评定表.xlsx",
  "控制及保护屏台的安装分项工程质量检查验收评定表.xlsx",
  "二次回路检查及控制电缆接线分项工程质量检查验收评定表.xlsx",
  "交直流控制电源系统设备安装分部工程质量验收评定表.xlsx",
  "蓄电池安装分项工程质量检查验收评定表.xlsx",
  "UPS装置检查及带负荷试验分项工程质量检查验收评定表.xlsx",
  "接地装置安装分部工程质量检查验收评定表.xlsx",
  "屋外接地装置安装分项工程质量验收表.xlsx",
  "屋内接地装置安装分项工程质量验收表.xlsx",
  "计算机监控系统设备安装分部工程质量验收评定表.xlsx",
  "单个盘台箱柜安装分项工程质量检查验收评定表.xlsx",
  "监视设备安装分项工程质量检查验收评定表.xlsx",
  "电缆支架桥架竖井制作及安装分项工程质量验收表.xlsx",
  "电缆敷设分项工程质量检查验收评定表.xlsx",
  "光伏升压变电缆敷设分项工程质量验收表.xlsx",
  "开关柜电缆敷设分项工程质量验收表.xlsx",
  "低压交流电缆隐蔽工程质量验收记录.xlsx",
  "通讯系统设备安装分部工程质量验收评定表.xlsx",
  "通信系统一次设备安装分项工程质量检查验收评定表.xlsx",
  "配电柜手动分合闸检查记录.xlsx",
  "逆变器散热装置人机界面检查记录.xlsx",
  "配电柜施工安装记录.xlsx",
  "配电柜安装接地检查记录.xlsx",
]);

export function getProcessRecordApplicability(record: ArchiveRecord): { isApplicable: boolean; matchedKeywords: string[] } {
  const signalText = [record.fullTitle, ...record.items.map((item) => `${item.fileCode} ${item.title}`)].join(" ");
  const matchedKeywords = PROCESS_RECORD_KEYWORDS.filter((keyword) => signalText.includes(keyword));
  return {
    isApplicable: matchedKeywords.length > 0,
    matchedKeywords,
  };
}

export function isSummaryWorkbookTemplate(template: ProcessTemplate): boolean {
  return template.kind === "xlsx"
    && (template.originalName.includes("汇总用") || /分部工程质量(?:检查)?验收评定表/.test(template.originalName));
}
