import type { ArchiveItem, ArchiveRecord } from "../types";
import type { ProcessTemplate } from "./types";

export interface CollectorLineWorkbookReplacement {
  search: RegExp;
  replacement: string;
}

export function isCollectorLineRecord(record: ArchiveRecord): boolean {
  return record.archiveCode.includes("-8331-") || record.fullTitle.includes("集电线路");
}

export function isCollectorLineTemplate(template: ProcessTemplate): boolean {
  return COLLECTOR_LINE_TEMPLATE_FILES.has(template.templateFile);
}

export function collectorLineTemplatesForItem(item: ArchiveItem): string[] {
  const title = item.title;
  if (title.includes("集电线路安装工程分部开工报审")) {
    return ["开工报审.docx"];
  }
  if (/集电线路安装(?:工程)?分部工程质量/.test(title)) {
    return ["集电线路安装工程分部工程报验申请单.docx", "集电线路安装工程分部工程质量验收评定表.xlsx"];
  }
  if (title.includes("直埋电缆敷设分项工程质量")) {
    return ["直埋电缆敷设分项工程报验申请单.docx"];
  }
  if (title.includes("电缆带电试运签证")) {
    return ["电缆带电试运签证.xlsx"];
  }
  if (title.includes("电缆工程电缆沟开挖及回填工程隐蔽")) {
    return ["直埋电缆敷设施工隐蔽工程质量报验单.docx", "电缆沟开挖及回填工程质量隐蔽验收记录.xlsx"];
  }
  if (title.includes("直埋电缆隐蔽前检查签证")) {
    return ["直埋电缆隐蔽前检查签证.xlsx"];
  }
  if (title.includes("电力电缆终端制作安装分项工程质量") || title.includes("电缆终端制作分项工程质量")) {
    return ["电力电缆终端制作安装分项工程质量验收表.xlsx", "高压电缆终端制作安装分项工程报验申请单.docx"];
  }
  if (title.includes("电力电缆中间接头制作安装分项工程质量")) {
    return ["电力电缆中间接头制作安装分项工程报验申请单.docx", "电力电缆中间接头制作安装分项工程质量验收表.xlsx"];
  }
  if (title.includes("电缆中间接头位置记录")) {
    return ["电缆中间接头位置记录.xlsx"];
  }
  return [];
}

export function collectorLineOutputTitle(title: string): string | undefined {
  if (title.includes("电缆带电试运签证")) {
    return stripProjectPrefix(title);
  }
  if (title.includes("电缆工程电缆沟开挖及回填工程隐蔽")) {
    return "电缆工程电缆沟开挖及回填工程质量隐蔽验收记录";
  }
  if (title.includes("直埋电缆隐蔽前检查签证")) {
    return stripProjectPrefix(title);
  }
  if (title.includes("电缆中间接头位置记录")) {
    return "电缆中间接头位置记录";
  }
  return undefined;
}

export function collectorLineApplicationSubject(title: string): string {
  return stripProjectPrefix(title)
    .replace(/\s*质量(?:报验申请|报审表)及验收记录\s*$/, "")
    .replace(/\s*隐蔽报验申请及验收记录\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim()
    .replace(/\s*分部工程\s*$/, "")
    .replace(/\s*分项工程\s*$/, "");
}

export function collectorLineTrialCablePair(title: string): string | undefined {
  return title.match(/[（(][^）)]*[）)]/)?.[0].replace(/[()]/g, (value) => value === "(" ? "（" : "）");
}

export function collectorLineHiddenCableRoute(title: string): string | undefined {
  const subject = stripProjectPrefix(title).replace(/\s*直埋电缆隐蔽前检查签证\s*$/, "").trim();
  if (subject.includes("#1杆1T1至G01柜601")) {
    return "10KV松围线创楷支线#1杆1T1开关至10KV创楷光伏开关站G01柜601开关";
  }
  if (subject.includes("G01柜601开关至G03柜602开关")) {
    return "10KV广东创楷建设工程有限公司配电专用箱变G01柜601开关至10KV创楷光伏开关站G03柜602开关";
  }
  if (subject.includes("G01柜801开关至G08柜605开关")) {
    return "10KV创楷#1光伏升压变G01柜801开关至10KV创楷光伏开关站G08柜605开关";
  }
  if (subject.includes("G01柜801开关至G09柜606开关")) {
    return "10KV创楷#2光伏升压变G01柜801开关至10KV创楷光伏开关站G09柜606开关";
  }
  return undefined;
}

export function collectorLineEquipmentName(title: string): string | undefined {
  const subject = stripProjectPrefix(title);
  if (subject.includes("#1光伏升压变")) {
    return "10KV创楷#1光伏升压变";
  }
  if (subject.includes("#2光伏升压变")) {
    return "10KV创楷#2光伏升压变";
  }

  const switchgear = subject.match(/开关柜G\d+柜/)?.[0];
  return switchgear ? `10KV创楷光伏${switchgear}` : undefined;
}

export function collectorLineWorkbookReplacements(title: string): CollectorLineWorkbookReplacement[] {
  const replacements: CollectorLineWorkbookReplacement[] = [];
  const cablePair = collectorLineTrialCablePair(title);
  if (cablePair && title.includes("电缆带电试运签证")) {
    replacements.push({
      search: /（00909A-00909B）|\(00909A-00909B\)/g,
      replacement: cablePair,
    });
  }

  const route = collectorLineHiddenCableRoute(title);
  if (route) {
    replacements.push({
      search: /10KV(?:松围线创楷支线#1杆1T1开关至10KV创楷光伏开关站G01柜601开关|广东创楷建设工程有限公司配电专用箱变G01柜601开关至10KV创楷光伏开关站G03柜602开关|创楷#1光伏升压变G01柜801开关至10KV创楷光伏开关站G08柜605开关|创楷#2光伏升压变G01柜801开关至10KV创楷光伏开关站G09柜606开关)/g,
      replacement: route,
    });
  }

  const equipment = collectorLineEquipmentName(title);
  if (equipment) {
    replacements.push({
      search: /10KV创楷(?:#1光伏升压变|#2光伏升压变|光伏开关柜G\d+柜)/g,
      replacement: equipment,
    });
  }

  return replacements;
}

export function collectorLineDivisionEvaluationItems(): Array<{ name: string; nature: string; grade: string }> {
  return [
    { name: "直埋电缆敷设", nature: "主要", grade: "合格" },
    { name: "高压电缆终端制作安装", nature: "主要", grade: "合格" },
    { name: "电力电缆中间接头制作安装", nature: "主要", grade: "合格" },
  ];
}

function stripProjectPrefix(title: string): string {
  return title
    .replace(/^\s*\d+[、.．\-\s]*/, "")
    .replace(/^.*?项目\s*/, "")
    .trim();
}

const COLLECTOR_LINE_TEMPLATE_FILES = new Set([
  "集电线路安装工程分部工程报验申请单.docx",
  "集电线路安装工程分部工程质量验收评定表.xlsx",
  "直埋电缆敷设分项工程报验申请单.docx",
  "电缆带电试运签证.xlsx",
  "直埋电缆敷设施工隐蔽工程质量报验单.docx",
  "电缆沟开挖及回填工程质量隐蔽验收记录.xlsx",
  "直埋电缆隐蔽前检查签证.xlsx",
  "高压电缆终端制作安装分项工程报验申请单.docx",
  "电力电缆终端制作安装分项工程质量验收表.xlsx",
  "电力电缆中间接头制作安装分项工程报验申请单.docx",
  "电力电缆中间接头制作安装分项工程质量验收表.xlsx",
  "电缆中间接头位置记录.xlsx",
]);
