import type { ArchiveItem, ArchiveRecord } from "../types";
import { inspectionApplicationFullSubject, inspectionApplicationSubject, stripProjectPrefix } from "./textReplacement";
import type { ProcessTemplate } from "./types";

export function isSwitchStationArchiveRecord(record: ArchiveRecord): boolean {
  return record.archiveCode.includes("-8341-") || record.archiveCode.includes("-8342-")
    || record.fullTitle.includes("开关站电气设备安装");
}

export function isSwitchStationSummaryTemplate(template: ProcessTemplate | undefined): boolean {
  return Boolean(template?.templateFile.includes("开关站电气设备安装子单位工程质量验收记录"));
}

export function switchStationUnitProjectName(): string {
  return "并网点光伏发电系统";
}

export function switchStationSubunitName(title: string): string {
  return stripProjectPrefix(title)
    .replace(/\s*质量(?:报验申请|报审表)及验收记录\s*$/, "")
    .replace(/\s*子单位工程\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();
}

export function switchStationApplicationSubject(title: string): string {
  if (title.includes("模板拆除隐蔽")) {
    return "设备基础模板";
  }
  if (title.includes("钢筋隐蔽")) {
    return "设备基础钢筋";
  }
  if (title.includes("钢筋分项工程")) {
    return "基础钢筋";
  }

  return normalizeSwitchStationQualityName(inspectionApplicationSubject(title));
}

export function normalizeSwitchStationQualityName(title: string): string {
  let name = stripProjectPrefix(title)
    .replace(/\s*质量(?:报验申请|报审表)及验收记录\s*$/, "")
    .replace(/\s*验收记录\s*$/, "")
    .replace(/[，,。；;：:\s]+$/g, "")
    .trim();

  name = name
    .replace(/分项工程工程/g, "分项工程")
    .replace(/控制及保护屏台的安装/g, "控制及保护屏台安装")
    .replace(/二次回路检查及控制电缆接线/g, "二次回路检查及接线")
    .replace(/电缆支架、桥架、竖井制作及安装/g, "电缆桥架安装")
    .replace(/#1及#2光伏升压变柜外接地装置安装/g, "开关柜外接地装置安装")
    .replace(/#1及#2光伏升压变柜内接地装置安装/g, "开关柜内接地装置安装")
    .replace(/通信系统一次设备安装/g, "通讯系统一次设备安装")
    .replace(/\s+/g, " ")
    .trim();

  return name
    .replace(/\s*分部工程\s*$/, "")
    .replace(/\s*分项工程\s*$/, "")
    .replace(/\s*工程工程\s*$/, "工程")
    .trim();
}

export function switchStationDivisionName(title: string): string {
  return normalizeSwitchStationQualityName(inspectionApplicationFullSubject(title));
}

export function switchStationInspectionLotContext(record: ArchiveRecord, item: ArchiveItem): {
  unitProjectName: string;
  divisionName: string;
  subitemName: string;
} {
  const beforeItems = record.items.slice(0, record.items.indexOf(item)).reverse();
  const division = beforeItems.find((candidate) => isDivisionQualityItem(candidate.title));
  const subitem = beforeItems.find((candidate) => isSubitemQualityItem(candidate.title));

  return {
    unitProjectName: switchStationUnitProjectName(),
    divisionName: division ? switchStationDivisionName(division.title) : "",
    subitemName: normalizeSwitchStationInspectionLotSubitem(subitem
      ? normalizeSwitchStationQualityName(inspectionApplicationFullSubject(subitem.title))
      : normalizeSwitchStationQualityName(item.title.replace(/\s*检验批质量验收记录\s*$/u, ""))),
  };
}

function normalizeSwitchStationInspectionLotSubitem(name: string): string {
  return name === "基础模板" ? "模板" : name;
}

export function switchStationContextRecords(records: ArchiveRecord[]): ArchiveRecord[] {
  return records.filter(isSwitchStationArchiveRecord);
}

export function switchStationSubunitParts(records: ArchiveRecord[]): Array<{ name: string; count: number }> {
  const parts: Array<{ name: string; count: number }> = [];
  for (const record of records) {
    for (const item of record.items) {
      if (!isDivisionQualityItem(item.title)) {
        continue;
      }

      const name = switchStationDivisionName(item.title);
      const count = switchStationSubitemCount(record, item);
      if (count > 0 && !parts.some((part) => part.name === name)) {
        parts.push({ name, count });
      }
    }
  }
  return parts;
}

function switchStationSubitemCount(record: ArchiveRecord, item: ArchiveItem): number {
  const children = followingDivisionChildren(record, item);
  const groups = new Set<string>();
  for (const child of children) {
    const name = normalizeSwitchStationQualityName(inspectionApplicationFullSubject(child.title));
    if (name === "定位及高程控制") {
      continue;
    }
    groups.add(`${subitemParentCode(child.fileCode)}|${name}`);
  }
  return groups.size;
}

export function followingSwitchStationDivisionChildren(record: ArchiveRecord, item: ArchiveItem): ArchiveItem[] {
  return followingDivisionChildren(record, item);
}

function followingDivisionChildren(record: ArchiveRecord, item: ArchiveItem): ArchiveItem[] {
  const followingItems = record.items.slice(record.items.indexOf(item) + 1);
  const end = followingItems.findIndex((candidate) =>
    candidate.title.includes("开工报审")
    || isDivisionQualityItem(candidate.title)
    || isSubunitQualityItem(candidate.title),
  );
  return followingItems
    .slice(0, end === -1 ? followingItems.length : end)
    .filter((candidate) => isSubitemQualityApplicationItem(candidate.title));
}

function isSubunitQualityItem(title: string): boolean {
  return /子单位(?:工程)?/.test(title) && /质量(?:报验申请|报审表)及验收记录/.test(title);
}

function isDivisionQualityItem(title: string): boolean {
  return /分部工程质量(?:报验申请|报审表)及验收记录/.test(title) && !title.includes("子单位");
}

function isSubitemQualityApplicationItem(title: string): boolean {
  return /分项工程(?:工程)?质量(?:报验申请|报审表)及验收记录/.test(title);
}

function isSubitemQualityItem(title: string): boolean {
  return isSubitemQualityApplicationItem(title) || /分项工程验收记录$/.test(title);
}

function subitemParentCode(fileCode: string): string {
  const code = fileCode.trim();
  if (!code || code === "/") {
    return "";
  }
  return code.replace(/-\d{3}$/, "");
}
