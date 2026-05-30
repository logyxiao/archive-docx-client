import { readFileSync } from "node:fs";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { generateProcessDocs, renderProcessDocx, renderProcessWorkbook, renderSummaryWorkbook } from "../lib/processDocs";
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
  describe("template matching", () => {
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
          subcontractorContent: "支架安装",
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

    it("generates only selected process template categories", async () => {
    stubProcessFetch();
    const paths: string[] = [];

    const result = await generateProcessDocs(
      [processRecord],
      {
        selectedCodes: [processRecord.archiveCode],
        outputDir: "/tmp/archive-output",
        selectedTemplateCategories: ["start-report", "summary-quality-acceptance"],
      },
      async (path) => {
        paths.push(path);
      },
    );

    expect(result.files).toHaveLength(8);
    expect(paths).toHaveLength(8);
    expect(paths.some((path) => path.includes("开工报审"))).toBe(true);
    expect(paths.some((path) => path.includes("质量验收记录（汇总用）"))).toBe(true);
    expect(paths.some((path) => path.includes("报验申请"))).toBe(false);
    expect(result.skipped).toEqual([
      "5028G01-0011-8312-001 第 45 条：未找到模板",
      "5028G01-0011-8312-001 第 46 条：未找到模板",
    ]);
    expect(result.errors).toEqual([]);
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
    expect(result.skipped).toEqual(["5028G01-0011-813-001：不适用过程资料模板（未匹配过程资料关键词）"]);
    expect(result.errors).toEqual([]);
  });
  });

  describe("docx rendering", () => {
    it("fills docx fields from source data and clears template-only company values", async () => {
    const item = processRecord.items[0];
    const template = readPublic("/templates/process-docs/template-001.docx");
    const xml = await docxXml(renderProcessDocx(template, processRecord, item, { supervisionDepartment: "测试监理项目部" }));

    expect(xml).toContain(processRecord.projectName);
    expect(xml).toContain(item.fileCode);
    expect(xml).toContain("测试监理项目部");
    expect(xml).not.toContain("河南中核五院研究设计有限公司");
  });
  });

  describe("xlsx rendering", () => {
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
        subcontractorContent: "支架安装",
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
    expect(sheet.getCell("AE8").value).toBe("支架安装");
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

    it("treats blank user fields as not filled", async () => {
    const item = processRecord.items[4];
    const template = readPublic("/templates/process-docs/template-008.xlsx");
    const workbook = await workbookFrom(
      await renderProcessWorkbook(template, processRecord, item, {
        generalContractorUnit: "",
        constructionUnit: "   ",
      }),
    );
    const sheet = workbook.worksheets[0];

    expect(sheet.getCell("G6").value).toBe(item.owner);
    expect(sheet.getCell("G7").value).toBe(item.owner);
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

    it("fills numeric self-check values within quality standard ranges", async () => {
    const item = processRecord.items[43];
    const template = readPublic("/templates/process-docs/template-061.xlsx");
    const workbook = await workbookFrom(await renderProcessWorkbook(template, processRecord, item));
    const sheet = workbook.worksheets[0];
    const values = ["V14", "W14", "X14", "Y14", "Z14", "AA14", "AB14", "AC14", "AD14", "AE14"]
      .map((address) => Number(sheet.getCell(address).value));

    expect(values.every((value) => Number.isFinite(value))).toBe(true);
    expect(values.every((value) => value >= -2 && value <= 2)).toBe(true);
  });

    it("fills merged quality result values when templates use quality acceptance result columns", async () => {
    const item = processRecord.items[4];
    const template = readPublic("/templates/process-docs/template-008.xlsx");
    const workbook = await workbookFrom(await renderProcessWorkbook(template, processRecord, item));
    const sheet = workbook.worksheets[0];
    const values = String(sheet.getCell("V21").value)
      .split(",")
      .map((value) => Number(value));

    expect(values).toHaveLength(10);
    expect(values.every((value) => Number.isFinite(value))).toBe(true);
    expect(values.every((value) => value >= 0 && value <= 2)).toBe(true);
  });
  });

  describe("summary workbook rendering", () => {
    it("keeps summary workbook header merge and style intact", async () => {
    const summaryTemplates = [
      "/templates/process-docs/template-002.xlsx",
      "/templates/process-docs/template-006.xlsx",
      "/templates/process-docs/template-034.xlsx",
      "/templates/process-docs/template-045.xlsx",
    ];

    for (const templatePath of summaryTemplates) {
      const template = readPublic(templatePath);
      const source = await workbookFrom(template);
      const renderedBytes = renderSummaryWorkbook(template, processRecord, {
        generalContractorUnit: "测试总承包单位",
        generalContractorProjectManager: "总包负责人",
        generalContractorTechnicalLeader: "总包技术负责人",
        constructionUnit: "测试施工单位",
        constructionProjectManager: "施工负责人",
        constructionTechnicalLeader: "施工技术负责人",
        subcontractorUnit: "测试分包单位",
        subcontractorProjectManager: "分包负责人",
        subcontractorContent: "光伏方阵安装",
      });
      const rendered = await workbookFrom(renderedBytes);
      const sourceZip = await JSZip.loadAsync(template);
      const renderedZip = await JSZip.loadAsync(renderedBytes);

      expect(rendered.worksheets[0].getCell("B2").style).toEqual(source.worksheets[0].getCell("B2").style);
      expect(rendered.worksheets[0].getCell("B2").isMerged).toBe(source.worksheets[0].getCell("B2").isMerged);
      expect(rendered.worksheets[0].getCell("C2").master.address).toBe(source.worksheets[0].getCell("C2").master.address);
      expect(await renderedZip.file("xl/styles.xml")!.async("string")).toBe(await sourceZip.file("xl/styles.xml")!.async("string"));
      expect(rendered.worksheets[0].getCell("G7").value).toBe("测试总承包单位");
      expect(rendered.worksheets[0].getCell("G8").value).toBe("测试施工单位");
      expect(rendered.worksheets[0].getCell("G9").value).toBe("测试分包单位");
      expect(rendered.worksheets[0].getCell("U7").value).toBe("总包负责人");
      expect(rendered.worksheets[0].getCell("AF8").value).toBe("施工技术负责人");
      expect(rendered.worksheets[0].getCell("AF9").value).toBe("光伏方阵安装");
    }
  });
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
