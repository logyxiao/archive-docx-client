import { readFileSync } from "node:fs";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { generateProcessDocs, renderProcessDocx, renderProcessWorkbook } from "../lib/processDocs";
import type { ArchiveRecord } from "../lib/types";

const processRecord = createProcessRecord();

function readPublic(path: string): Buffer {
  return readFileSync(`public${decodeURIComponent(path)}`);
}

function stubProcessFetch() {
  globalThis.fetch = async (input) => {
    const path = String(input).replace(/^https?:\/\/[^/]+/, "");
    return new Response(readPublic(path));
  };
}

async function docxXml(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file("word/document.xml")!.async("string");
}

async function workbookFrom(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(bytes));
  return workbook;
}

describe("process docs generation", () => {
  it("matches bundled templates by source row sequence and skips rows without templates", async () => {
    stubProcessFetch();
    const paths: string[] = [];

    const result = await generateProcessDocs(
      [processRecord],
      {
        selectedCodes: [processRecord.archiveCode],
        outputDir: "/tmp/archive-output",
        userFields: {
          generalContractorUnit: "测试总承包单位",
          generalContractorProjectManager: "总包负责人",
          generalContractorTechnicalLeader: "总包技术负责人",
          constructionUnit: "测试施工单位",
          constructionProjectManager: "施工负责人",
          constructionTechnicalLeader: "施工技术负责人",
          subcontractorUnit: "测试分包单位",
          subcontractorProjectManager: "分包负责人",
          subcontractorTechnicalLeader: "分包技术负责人",
          supervisionDepartment: "测试监理项目部",
        },
      },
      async (path) => {
        paths.push(path);
      },
    );

    expect(result.files).toHaveLength(61);
    expect(paths).toHaveLength(61);
    expect(result.skipped).toEqual([
      "5028G01-0011-8312-001 第 45 条：未找到模板",
      "5028G01-0011-8312-001 第 46 条：未找到模板",
    ]);
    expect(result.errors).toEqual([]);
    expect(paths[0]).toContain("/过程资料/5028G01-0011-8312-001");
  });

  it("skips archive records that do not match the process-doc template profile", async () => {
    stubProcessFetch();
    const drawingRecord = {
      ...processRecord,
      archiveCode: "5028G01-0011-813-001",
      fullTitle: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目 DQ-100~109 光伏直流部分施工图",
      items: processRecord.items.map((item, index) => ({
        ...item,
        fileCode: `DQ-${100 + index}`,
        title: `高明分布式项目 施工图图纸${index + 1}`,
      })),
    };

    const result = await generateProcessDocs(
      [drawingRecord],
      {
        selectedCodes: [drawingRecord.archiveCode],
        outputDir: "/tmp/archive-output",
      },
      async () => {},
    );

    expect(result.files).toEqual([]);
    expect(result.skipped).toEqual(["5028G01-0011-813-001：不适用过程资料模板"]);
    expect(result.errors).toEqual([]);
  });

  it("fills docx fields from source data and clears template-only company values", async () => {
    const item = processRecord.items[0];
    const template = readPublic("/templates/process-docs/template-001.docx");
    const xml = await docxXml(renderProcessDocx(template, processRecord, item, { supervisionDepartment: "测试监理项目部" }));

    expect(xml).toContain(processRecord.projectName);
    expect(xml).toContain(item.fileCode);
    expect(xml).toContain("测试监理项目部");
    expect(xml).not.toContain("河南中核五院研究设计有限公司");
  });

  it("fills xlsx source fields and clears source-missing fields", async () => {
    const item = processRecord.items[4];
    const template = readPublic("/templates/process-docs/template-008.xlsx");
    const workbook = await workbookFrom(
      await renderProcessWorkbook(template, processRecord, item, {
        generalContractorUnit: "测试总承包单位",
        generalContractorProjectManager: "总包负责人",
        generalContractorTechnicalLeader: "总包技术负责人",
        constructionUnit: "测试施工单位",
        constructionProjectManager: "施工负责人",
        constructionTechnicalLeader: "施工技术负责人",
        subcontractorUnit: "测试分包单位",
        subcontractorProjectManager: "分包负责人",
      }),
    );
    const sheet = workbook.worksheets[0];

    expect(sheet.getCell("F4").value).toBe(processRecord.projectName);
    expect(sheet.getCell("G6").value).toBe("测试总承包单位");
    expect(sheet.getCell("G7").value).toBe("测试施工单位");
    expect(sheet.getCell("G8").value).toBe("测试分包单位");
    expect(sheet.getCell("T6").value).toBe("总包负责人");
    expect(sheet.getCell("AE6").value).toBe("总包技术负责人");
    expect(sheet.getCell("T7").value).toBe("施工负责人");
    expect(sheet.getCell("AE7").value).toBe("施工技术负责人");
    expect(sheet.getCell("T8").value).toBe("分包负责人");
    expect(sheet.getCell("B5").isMerged).toBe(true);
  });

  it("leaves optional process fields blank when users do not fill them", async () => {
    const item = processRecord.items[4];
    const template = readPublic("/templates/process-docs/template-008.xlsx");
    const workbook = await workbookFrom(await renderProcessWorkbook(template, processRecord, item));
    const sheet = workbook.worksheets[0];

    expect(sheet.getCell("G6").value).toBe(item.owner);
    expect(sheet.getCell("G7").value).toBe(item.owner);
    expect(sheet.getCell("G8").value).toBe("");
    expect(sheet.getCell("T7").value).toBe("");
    expect(sheet.getCell("AE7").value).toBe("");
    expect(sheet.getCell("T8").value).toBe("");
  });

  it("reuses construction technical leader for professional foreman fields", async () => {
    const item = processRecord.items[35];
    const template = readPublic("/templates/process-docs/template-037.xlsx");
    const workbook = await workbookFrom(
      await renderProcessWorkbook(template, processRecord, item, {
        constructionTechnicalLeader: "施工技术负责人",
      }),
    );
    const sheet = workbook.worksheets[0];

    expect(sheet.getCell("AE11").value).toBe("施工技术负责人");
  });
});

function createProcessRecord(): ArchiveRecord {
  const items = Array.from({ length: 46 }, (_, index) => {
    const sequence = index + 1;
    const fileCode = sequence <= 44 ? `5028G01-SG-ZHHC-TEST-${String(sequence).padStart(3, "0")}` : "/";
    return {
      sequence: String(sequence),
      fileCode,
      owner: sequence >= 45 ? "中建联设计院（广州）股份有限公司" : "中核华辰建筑工程有限公司",
      title: `高明分布式项目 过程资料质量验收记录${sequence}`,
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
