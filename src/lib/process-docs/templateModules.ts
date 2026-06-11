import type { ArchiveRecord } from "../types";
import { PROCESS_RECORD_KEYWORDS } from "./constants";
import { isCollectorLineRecord, isCollectorLineTemplate } from "./collectorLine";
import { isSwitchStationArchiveRecord } from "./switchStation";
import type { ProcessTemplate, ProcessTemplateModule } from "./types";

export function isTemplateInModule(template: ProcessTemplate, templateModule: ProcessTemplateModule): boolean {
  if (template.userTemplatePath && template.templateModule) {
    return template.templateModule === templateModule;
  }

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
