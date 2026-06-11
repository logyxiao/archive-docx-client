import { readFileSync } from "node:fs";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import { expect } from "vitest";
import type { ArchiveRecord } from "../lib/types";

export const processRecord = createProcessRecord();
export const collectorLineRecord = createCollectorLineRecord();

export function readPublic(path: string): Buffer {
  return readFileSync(`public${decodeURIComponent(path)}`);
}

export function stubProcessFetch() {
  globalThis.fetch = async (input) => {
    const path = String(input).replace(/^https?:\/\/[^/]+/, "");
    return new Response(readPublic(path));
  };
}

export function fileNames(paths: string[]): string[] {
  return paths.map((path) => {
    const parts = path.split("/");
    return parts[parts.length - 1] || path;
  });
}

export async function docxXml(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file("word/document.xml")!.async("string");
}

export async function workbookFrom(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(bytes));
  return workbook;
}

export async function xlsxXml(bytes: Uint8Array, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file(path)!.async("string");
}

export async function expectValidOfficeFile(path: string, bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  if (path.endsWith(".docx")) {
    expect(zip.file("word/document.xml")).toBeTruthy();
    return;
  }
  if (path.endsWith(".xlsx")) {
    expect(zip.file("xl/workbook.xml")).toBeTruthy();
  }
}

export function cellText(cell: ExcelJS.Cell): string {
  const value = cell.value;
  if (typeof value === "string" || typeof value === "number") {
    return String(value);
  }
  if (value && typeof value === "object") {
    if ("text" in value && typeof value.text === "string") {
      return value.text;
    }
    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText.map((part) => part.text).join("");
    }
  }
  return "";
}

export function createProcessRecord(): ArchiveRecord {
  const items = Array.from({ length: 46 }, (_, index) => {
    const sequence = index + 1;
    const fileCode = sequence <= 44 ? `5028G01-SG-ZHHC-TEST-${String(sequence).padStart(3, "0")}` : "/";
    const titleBySequence: Record<number, string> = {
      1: "高明分布式项目 并网点光伏发电系统单位工程开工报审",
      2: "高明分布式项目 光伏方阵安装子单位工程质量报验申请及验收记录",
      3: "高明分布式项目 子方阵支架及组件安装分部工程开工报审",
      20: "高明分布式项目 通用工程分部工程开工报审",
      29: "高明分布式项目 主体结构分部工程开工报审",
    };
    return {
      sequence: String(sequence),
      fileCode,
      owner: sequence >= 45 ? "中建联设计院（广州）股份有限公司" : "中核华辰建筑工程有限公司",
      title: titleBySequence[sequence] ?? `高明分布式项目 过程资料质量验收记录${sequence}`,
      fileDate: "20250410",
      pageNo: String(sequence),
      note: "",
    };
  });

  return {
    categoryCode: "8312",
    archiveCode: "5028G01-0011-8312-001",
    fullTitle: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目 光伏发电系统单位工程开工报审，支架及组件安装、通用工程、主体结构分部、分项及检验批质量验收记录文件",
    projectName: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目",
    volumeTitle: "光伏发电系统单位工程开工报审，支架及组件安装、通用工程、主体结构分部、分项及检验批质量验收记录文件",
    owner: "中核华辰建筑工程有限公司",
    filingUnit: "中核华辰建筑工程有限公司",
    retentionPeriod: "30年",
    startDate: "20241208",
    endDate: "20250420",
    dateRange: "20241208-20250420",
    totalPages: 94,
    drawingPages: 0,
    textPages: 94,
    items,
  };
}

export function createGridConnectedElectricalRecord(): ArchiveRecord {
  const item = (sequence: number, fileCode: string, title: string, fileDate = "20250420") => ({
    sequence: String(sequence),
    fileCode,
    owner: "中核华辰建筑工程有限公司",
    title,
    fileDate,
    pageNo: String(sequence),
    note: "",
  });

  return {
    categoryCode: "8312",
    archiveCode: "5028G01-0011-8312-002",
    fullTitle: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目 光伏变电系统子单位工程质量报审，分部、分项及检验批质量验收记录文件",
    projectName: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目",
    volumeTitle: "光伏变电系统子单位工程质量报审，分部、分项及检验批质量验收记录文件",
    owner: "中核华辰建筑工程有限公司",
    filingUnit: "中核华辰建筑工程有限公司",
    retentionPeriod: "30年",
    startDate: "20250310",
    endDate: "20250429",
    dateRange: "20250310-20250429",
    totalPages: 80,
    drawingPages: 0,
    textPages: 80,
    items: [
      item(1, "5028G01-SG-ZHHC-02-001", "高明分布式项目 并网点光伏变电系统子单位工程质量报审表及验收记录", "20250429"),
      item(2, "5028G01-SG-ZHHC-KG-05", "高明分布式项目 子方阵场电气安装分部工程开工报审", "20250330"),
      item(3, "5028G01-SG-ZHHC-02-01-001", "高明分布式项目 子方阵场电气安装分部工程质量报验申请及验收记录", "20250428"),
      item(4, "5028G01-SG-ZHHC-02-01-01-001", "高明分布式项目 逆变器安装分项工程质量报验申请及验收记录", "20250428"),
      item(5, "/", "高明分布式项目 逆变器施工安装记录", "20250416"),
      item(6, "/", "高明分布式项目 逆变器手动分合闸检查记录", "20250416"),
      item(7, "/", "高明分布式项目 逆变器通讯调试记录", "20250427"),
      item(8, "/", "高明分布式项目 逆变器外观、主要元器件、控制电源、直交流侧接线及极性（相序）、绝缘、接地检查记录）", "20250427"),
      item(9, "5028G01-SG-ZHHC-02-01-02-001", "高明分布式项目 #1干式变压器安装分项工程质量报验申请及验收记录", "20250427"),
      item(10, "5028G01-SG-ZHHC-02-01-02-002", "高明分布式项目 #2干式变压器安装分项工程质量报验申请及验收记录", "20250427"),
      item(11, "5028G01-SG-ZHHC-02-01-03-001", "高明分布式项目 #1升压变屋外接地装置安装分项工程质量报验申请及验收记录", "20250425"),
      item(12, "5028G01-SG-ZHHC-02-01-03-002", "高明分布式项目 #2升压变屋外接地装置安装分项工程质量报验申请及验收记录", "20250425"),
      item(13, "5028G01-SG-ZHHC-02-01-04-001", "高明分布式项目 #1升压变屋内接地装置安装分项工程质量报验申请及验收记录", "20250425"),
      item(14, "5028G01-SG-ZHHC-02-01-04-002", "高明分布式项目 #2升压变屋内接地装置安装分项工程质量报验申请及验收记录", "20250425"),
      item(15, "5028G01-SG-ZHHC-KG-06", "高明分布式项目 子方阵电气线路安装分部工程开工报审", "20250310"),
      item(16, "5028G01-SG-ZHHC-02-02-001", "高明分布式项目 子方阵电气线路安装分部工程质量报验申请及验收记录", "20250425"),
      item(17, "5028G01-SG-ZHHC-02-02-01-001", "高明分布式项目 电缆桥架安装分项工程质量报验申请及验收记录", "20250420"),
      item(18, "5028G01-SG-ZHHC-02-02-02-001", "高明分布式项目 电缆敷设分项工程质量报验申请及验收记录", "20250423"),
      item(19, "5028G01-SG-ZHHC-02-02-03-001", "高明分布式项目 #1光伏升压变电缆防火阻燃施工分项工程质量报验申请及验收记录", "20250420"),
      item(20, "5028G01-SG-ZHHC-02-02-03-002", "高明分布式项目 #2光伏升压变电缆防火阻燃施工分项工程质量报验申请及验收记录", "20250420"),
      item(21, "5028G01-SG-ZHHC-02-02-03-003", "高明分布式项目 开关柜G01柜电缆防火阻燃施工分项工程质量报验申请及验收记录", "20250423"),
      item(22, "5028G01-SG-ZHHC-02-02-03-004", "高明分布式项目 开关柜G03柜电缆防火阻燃施工分项工程质量报验申请及验收记录", "20250423"),
      item(23, "5028G01-SG-ZHHC-02-02-03-005", "高明分布式项目 开关柜G08柜电缆防火阻燃施工分项工程质量报验申请及验收记录", "20250423"),
      item(24, "5028G01-SG-ZHHC-02-02-03-006", "高明分布式项目 开关柜G09柜电缆防火阻燃施工分项工程质量报验申请及验收记录", "20250423"),
    ],
  };
}

export function createSuixiGridConnectedElectricalRecord(): ArchiveRecord {
  const item = (sequence: number, fileCode: string, title: string, fileDate = "20250417") => ({
    sequence: String(sequence),
    fileCode,
    owner: "中核华辰建筑工程有限公司",
    title,
    fileDate,
    pageNo: String(sequence),
    note: "",
  });

  return {
    categoryCode: "8312",
    archiveCode: "5028G02-0011-8312-004",
    fullTitle: "中核汇能遂溪国恒坤达3.84MW屋顶分布式光伏项目 并网点光伏变电系统子单位工程开工报审、子方阵场电气安装分部、子方阵电气线路安装分部、计算机监控系统设备安装分部、集电线路安装分部工程、分项及检验批质量验收记录文件",
    projectName: "中核汇能遂溪国恒坤达3.84MW屋顶分布式光伏项目",
    volumeTitle: "并网点光伏变电系统子单位工程开工报审、子方阵场电气安装分部、子方阵电气线路安装分部、计算机监控系统设备安装分部、集电线路安装分部工程、分项及检验批质量验收记录文件",
    owner: "中核华辰建筑工程有限公司",
    filingUnit: "中核华辰建筑工程有限公司",
    retentionPeriod: "永久",
    startDate: "20250417",
    endDate: "20250430",
    dateRange: "20250417-20250430",
    totalPages: 80,
    drawingPages: 0,
    textPages: 80,
    items: [
      item(14, "5028G02-SG-ZHHC-03-02-03-001", "遂溪分布式项目 电缆终端制作分项工程质量报验申请及验收记录"),
      item(15, "5028G02-SG-ZHHC-03-02-04-001", "遂溪分布式项目 电缆防火与阻燃分项工程质量报验申请及验收记录"),
      item(16, "5028G02-SG-ZHHC-03-02-05-001", "遂溪分布式项目 电气二次系统分项工程质量报验申请及验收记录"),
      item(17, "5028G02-SG-ZHHC-03-02-05-01-001", "遂溪分布式项目 国恒部分直流配电柜检验批质量验收记录", "20250430"),
      item(18, "5028G02-SG-ZHHC-03-02-05-02-001", "遂溪分布式项目 坤达部分直流配电柜检验批质量验收记录", "20250430"),
      item(19, "5028G02-SG-ZHHC-03-03-001", "遂溪分布式项目 集电线路安装分部工程质量报验申请及验收记录"),
    ],
  };
}

export function createCollectorLineRecord(): ArchiveRecord {
  const item = (
    sequence: number,
    fileCode: string,
    title: string,
    fileDate = "20250501",
    pageNo: string | number = sequence,
  ) => ({
    sequence: String(sequence),
    fileCode,
    owner: "中核华辰建筑工程有限公司",
    title,
    fileDate,
    pageNo: String(pageNo),
    note: "",
  });

  return {
    categoryCode: "8331",
    archiveCode: "5028G01-0011-8331-001",
    fullTitle: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目 集电线路安装工程分部开工报审，分项及隐蔽质量验收记录文件",
    projectName: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目",
    volumeTitle: "集电线路安装工程分部开工报审，分项及隐蔽质量验收记录文件",
    owner: "中核华辰建筑工程有限公司",
    filingUnit: "中核华辰建筑工程有限公司",
    retentionPeriod: "30年",
    startDate: "20250429",
    endDate: "20250502",
    dateRange: "20250429-20250502",
    totalPages: 56,
    drawingPages: 0,
    textPages: 56,
    items: [
      item(1, "5028G01-SG-ZHHC-KG-15", "高明分布式项目 集电线路安装工程分部开工报审"),
      item(2, "5028G01-SG-ZHHC-03-08-001", "高明分布式项目 集电线路安装工程分部工程质量报验申请及验收记录"),
      item(3, "5028G01-SG-ZHHC-03-08-01-001", "高明分布式项目 直埋电缆敷设分项工程质量报验申请及验收记录"),
      item(4, "/", "高明分布式项目 （00909A-00909B）电缆带电试运签证"),
      item(5, "/", "高明分布式项目 （01315A-01315B)电缆带电试运签证"),
      item(6, "/", "高明分布式项目 （01316A-01316B)电缆带电试运签证"),
      item(7, "5028G01-SG-ZHHC-03-08-01-YB-001", "高明分布式项目 电缆工程电缆沟开挖及回填工程隐蔽报验申请及验收记录"),
      item(8, "/", "高明分布式项目 #1杆1T1至G01柜601开关直埋电缆隐蔽前检查签证"),
      item(9, "/", "高明分布式项目 G01柜601开关至G03柜602开关直埋电缆隐蔽前检查签证"),
      item(10, "/", "高明分布式项目 G01柜801开关至G08柜605开关直埋电缆隐蔽前检查签证"),
      item(11, "/", "高明分布式项目 G01柜801开关至G09柜606开关直埋电缆隐蔽前检查签证"),
      item(12, "5028G01-SG-ZHHC-03-08-02-001", "高明分布式项目 #1光伏升压变 电力电缆终端制作安装分项工程质量报验申请及验收记录"),
      item(13, "5028G01-SG-ZHHC-03-08-02-002", "高明分布式项目 #2光伏升压变 电力电缆终端制作安装分项工程质量报验申请及验收记录"),
      item(14, "5028G01-SG-ZHHC-03-08-02-003", "高明分布式项目 开关柜G01柜 电力电缆终端制作安装分项工程质量报验申请及验收记录"),
      item(15, "5028G01-SG-ZHHC-03-08-02-004", "高明分布式项目 开关柜G03柜 电力电缆终端制作安装分项工程质量报验申请及验收记录"),
      item(16, "5028G01-SG-ZHHC-03-08-02-005", "高明分布式项目 开关柜G08柜 电力电缆终端制作安装分项工程质量报验申请及验收记录"),
      item(17, "5028G01-SG-ZHHC-03-08-02-006", "高明分布式项目 开关柜G09柜 电力电缆终端制作安装分项工程质量报验申请及验收记录"),
      item(18, "5028G01-SG-ZHHC-03-08-03-001", "高明分布式项目 #1光伏升压变 电力电缆中间接头制作安装分项工程质量报验申请及验收记录"),
      item(19, "5028G01-SG-ZHHC-03-08-03-002", "高明分布式项目 #2光伏升压变 电力电缆中间接头制作安装分项工程质量报验申请及验收记录"),
      item(20, "5028G01-SG-ZHHC-03-08-03-003", "高明分布式项目 开关柜G01柜 电力电缆中间接头制作安装分项工程质量报验申请及验收记录"),
      item(21, "5028G01-SG-ZHHC-03-08-03-004", "高明分布式项目 开关柜G03柜 电力电缆中间接头制作安装分项工程质量报验申请及验收记录"),
      item(22, "5028G01-SG-ZHHC-03-08-03-005", "高明分布式项目 开关柜G08柜 电力电缆中间接头制作安装分项工程质量报验申请及验收记录"),
      item(23, "5028G01-SG-ZHHC-03-08-03-006", "高明分布式项目 开关柜G09柜 电力电缆中间接头制作安装分项工程质量报验申请及验收记录"),
      item(24, "/", "高明分布式项目 电缆中间接头位置记录", "20250502", "56~56"),
    ],
  };
}

export function createSubunitSummaryFixtureItems() {
  const base = processRecord.items[0];
  const item = (sequence: number, title: string, fileCode = `5028G01-SG-ZHHC-SUM-${String(sequence).padStart(3, "0")}`) => ({
    ...base,
    sequence: String(sequence),
    fileCode,
    title,
  });

  return [
    item(1, "高明分布式项目 并网点光伏发电系统单位工程开工报审", "5028G01-SG-ZHHC-KG-01"),
    item(2, "高明分布式项目 光伏方阵安装子单位工程质量报验申请及验收记录", "5028G01-SG-ZHHC-01-001"),
    item(3, "高明分布式项目 子方阵支架及组件安装分部工程质量报验申请及验收记录"),
    ...Array.from({ length: 10 }, (_, index) =>
      item(4 + index, `高明分布式项目 ${index + 1}#厂房支架安装分项工程质量报验申请及验收记录`),
    ),
    item(14, "高明分布式项目 通用工程分部工程质量报验申请及验收记录"),
    ...Array.from({ length: 5 }, (_, index) =>
      item(15 + index, `高明分布式项目 通用工程第${index + 1}项分项工程质量报验申请及验收记录`),
    ),
    item(20, "高明分布式项目 主体工程分部工程质量报验申请及验收记录"),
    item(21, "高明分布式项目 主体工程普通紧固件连接分项工程质量报验申请及验收记录"),
    item(22, "高明分布式项目 主体工程高强度螺栓连接分项工程质量报验申请及验收记录"),
  ];
}
