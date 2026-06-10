import { readFileSync } from "node:fs";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import {
  generateProcessDocs,
  generateSelectedProcessDocs,
  loadProcessManifest,
  matchingAllProcessTemplates,
  renderProcessDocx,
  renderProcessWorkbook,
  renderSummaryWorkbook,
} from "../lib/processDocs";
import type { ArchiveRecord } from "../lib/types";

const processRecord = createProcessRecord();
const collectorLineRecord = createCollectorLineRecord();

function readPublic(path: string): Buffer {
  return readFileSync(`public${decodeURIComponent(path)}`);
}

function stubProcessFetch() {
  globalThis.fetch = async (input) => {
    const path = String(input).replace(/^https?:\/\/[^/]+/, "");
    return new Response(readPublic(path));
  };
}

function fileNames(paths: string[]): string[] {
  return paths.map((path) => {
    const parts = path.split("/");
    return parts[parts.length - 1] || path;
  });
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

async function xlsxXml(bytes: Uint8Array, path: string): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file(path)!.async("string");
}

function cellText(cell: ExcelJS.Cell): string {
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

describe("process docs generation", () => {
  describe("template matching", () => {
    it("matches process templates only by explicit title rules", async () => {
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

      expect(result.files).toHaveLength(6);
      expect(paths).toHaveLength(6);
      expect(fileNames(paths).filter((name) => name.includes("开工报审"))).toHaveLength(4);
      expect(fileNames(paths).filter((name) => name.includes("子单位工程质量验收记录"))).toHaveLength(1);
      expect(fileNames(paths).filter((name) => name.includes("子单位工程报验申请单"))).toHaveLength(1);
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
      expect(paths[0]).toContain("/过程资料/5028G01-0011-8312-001");
    });

    it("resolves all-module matches for preview and generates selected manual templates", async () => {
      stubProcessFetch();
      const manifest = await loadProcessManifest();
      const startReportMatches = matchingAllProcessTemplates(processRecord, processRecord.items[0], manifest.templates);
      const unmatchedItem = {
        ...processRecord.items[4],
        sequence: "99",
        fileCode: "5028G01-SG-ZHHC-TEST-099",
        title: "高明分布式项目 需要人工选择模板的文件",
      };
      const record = {
        ...processRecord,
        items: [unmatchedItem],
      };
      const manualTemplate = manifest.templates.find((template) => template.templateFile === "开工报审.docx")!;
      const paths: string[] = [];

      expect(startReportMatches.map((match) => match.template.templateFile)).toEqual(["开工报审.docx"]);
      expect(matchingAllProcessTemplates(record, unmatchedItem, manifest.templates)).toEqual([]);

      const result = await generateSelectedProcessDocs(
        [record],
        {
          outputDir: "/tmp/archive-output",
          selections: [
            {
              archiveCode: record.archiveCode,
              fileCode: unmatchedItem.fileCode,
              sequence: unmatchedItem.sequence,
              templateModule: "process",
              template: manualTemplate,
            },
          ],
        },
        async (path) => {
          paths.push(path);
        },
      );

      expect(result.files).toHaveLength(1);
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
      expect(fileNames(paths)).toEqual(["99、5028G01-SG-ZHHC-TEST-099高明分布式项目 需要人工选择模板的文件.docx"]);
    });

    it("uses the shared start-report template when the item title contains start report text", async () => {
      stubProcessFetch();
      const startReportRecord: ArchiveRecord = {
        ...processRecord,
        items: [
          {
            ...processRecord.items[4],
            sequence: "5",
            fileCode: "5028G01-SG-ZHHC-KG-99",
            title: "高明分布式项目 光伏方阵安装子单位工程开工报审",
            fileDate: "20250530",
          },
        ],
      };
      const written: Uint8Array[] = [];
      const paths: string[] = [];

      const result = await generateProcessDocs(
        [startReportRecord],
        {
          selectedCodes: [startReportRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          selectedTemplateCategories: ["start-report"],
          userFields: {
            projectName: "测试工程",
            supervisionDepartment: "测试监理项目部",
          },
        },
        async (path, bytes) => {
          paths.push(path);
          written.push(bytes);
        },
      );
      const xml = await docxXml(written[0]);
      const templateXml = await docxXml(readPublic("/templates/process-docs/开工报审.docx"));

      expect(result.files).toHaveLength(1);
      expect(paths[0]).toContain("开工报审");
      expect(xml).toContain("测试监理项目部");
      expect(xml).toContain("我方承担的</w:t>");
      expect(xml).toContain("<w:u w:val=\"single\"/></w:rPr><w:t xml:space=\"preserve\"> 测试工程 光伏方阵安装子单位工程 </w:t>");
      expect(Array.from(xml.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g)).map((match) => match[1]).join("")).toContain(
        "特申请于 2025 年 5 月 30 日开工",
      );
      expect(xml.match(/<w:pgMar[^>]+>/)?.[0]).toBe(templateXml.match(/<w:pgMar[^>]+>/)?.[0]);
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("prioritizes files whose titles contain hidden-work text", async () => {
      stubProcessFetch();
      const manifest = await loadProcessManifest();
      const genericHiddenItem = {
        ...processRecord.items[4],
        sequence: "50",
        fileCode: "5028G01-SG-ZHHC-TEST-YB-001",
        title: "高明分布式项目 设备基础隐蔽检查记录",
      };
      const cableHiddenItem = {
        ...processRecord.items[4],
        sequence: "51",
        fileCode: "5028G01-SG-ZHHC-TEST-YB-002",
        title: "高明分布式项目 电缆隐蔽报验申请及验收记录",
      };

      expect(matchingAllProcessTemplates(processRecord, genericHiddenItem, manifest.templates).map((match) => match.template.templateFile))
        .toEqual(["隐蔽工程质量报验单.docx"]);
      expect(matchingAllProcessTemplates(processRecord, cableHiddenItem, manifest.templates).map((match) => match.template.templateFile))
        .toEqual(["隐蔽工程质量报验单.docx", "低压交流电缆隐蔽工程质量验收记录.xlsx"]);
    });

    it("uses a division start-report title when the source title is a division start report", async () => {
      stubProcessFetch();
      const startReportRecord: ArchiveRecord = {
        ...processRecord,
        items: [
          {
            ...processRecord.items[4],
            sequence: "5",
            fileCode: "5028G01-SG-ZHHC-KG-99",
            title: "高明分布式项目 子方阵支架及组件安装分部工程开工报审",
          },
        ],
      };
      const written: Uint8Array[] = [];

      await generateProcessDocs(
        [startReportRecord],
        {
          selectedCodes: [startReportRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          selectedTemplateCategories: ["start-report"],
        },
        async (_path, bytes) => {
          written.push(bytes);
        },
      );
      const xml = await docxXml(written[0]);

      expect(xml).toContain("分部工程开工报审表");
      expect(xml).not.toContain("单位工程开工报审表");
    });

    it("keeps a unit start-report title when the source title is not a division start report", async () => {
      stubProcessFetch();
      const written: Uint8Array[] = [];

      await generateProcessDocs(
        [processRecord],
        {
          selectedCodes: [processRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          selectedTemplateCategories: ["start-report"],
        },
        async (_path, bytes) => {
          written.push(bytes);
        },
      );
      const xml = await docxXml(written[0]);

      expect(xml).toContain("单位工程开工报审表");
      expect(xml).not.toContain("分部工程开工报审表");
    });

    it("uses the shared subunit quality workbook when the item title matches quality acceptance text", async () => {
      stubProcessFetch();
      const subunitRecord: ArchiveRecord = {
        ...processRecord,
        items: createSubunitSummaryFixtureItems(),
      };
      const written: Uint8Array[] = [];
      const paths: string[] = [];

      const result = await generateProcessDocs(
        [subunitRecord],
        {
          selectedCodes: [subunitRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          selectedTemplateCategories: ["summary-quality-acceptance"],
          userFields: {
            generalContractorUnit: "测试总承包单位",
            generalContractorProjectManager: "总包负责人",
            generalContractorTechnicalLeader: "总包技术负责人",
            constructionUnit: "测试施工单位",
            constructionProjectManager: "施工负责人",
            constructionTechnicalLeader: "施工技术负责人",
            subcontractorUnit: "测试分包单位",
            subcontractorProjectManager: "分包负责人",
            subcontractorContent: "光伏方阵安装",
          },
        },
        async (path, bytes) => {
          paths.push(path);
          written.push(bytes);
        },
      );
      const workbook = await workbookFrom(written[0]);
      const sheet = workbook.worksheets[0];
      const sourceSheetXml = await xlsxXml(readPublic("/templates/process-docs/子单位工程质量验收记录.xlsx"), "xl/worksheets/sheet1.xml");
      const sheetXml = await xlsxXml(written[0], "xl/worksheets/sheet1.xml");
      const workbookXml = await xlsxXml(written[0], "xl/workbook.xml");

      expect(result.files).toHaveLength(1);
      expect(paths[0]).toContain("子单位工程质量验收记录");
      expect(sheet.getCell("I2").value).toBe("光伏方阵安装子单位工程");
      expect(sheet.getCell("F4").value).toBe("5028G01-0011");
      expect(sheet.getCell("G6").value).toBe("并网点光伏发电系统");
      expect(sheet.getCell("U6").value).toBe("3");
      expect(sheet.getCell("AF6").value).toBe("17");
      expect(sheet.getCell("D11").value).toBe("子方阵支架及组件安装");
      expect(sheet.getCell("J11").value).toBe("10");
      expect(sheet.getCell("D12").value).toBe("通用工程");
      expect(sheet.getCell("J12").value).toBe("5");
      expect(sheet.getCell("D13").value).toBe("主体工程");
      expect(sheet.getCell("J13").value).toBe("2");
      expect(sheet.getCell("V11").value).toBe("自检检查符合规范及设计要求");
      expect(sheet.getCell("AD11").value).toBeNull();
      expect(sheet.pageSetup.fitToPage).toBe(true);
      expect(sheet.pageSetup.fitToWidth).toBe(1);
      expect(sheet.pageSetup.fitToHeight).toBe(1);
      expect(sheet.pageSetup.orientation).toBe("portrait");
      expect(sheet.pageSetup.paperSize).toBe(9);
      expect(sheet.pageSetup.scale).toBe(100);
      expect(sheetXml.match(/<pageMargins[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageMargins[^>]+>/)?.[0]);
      expect(sheetXml.indexOf("<pageMargins")).toBeLessThan(sheetXml.indexOf("<pageSetup"));
      expect(workbookXml).toContain('<definedName name="_xlnm.Print_Area" localSheetId="0">光伏方阵安装!$A$1:$AL$30</definedName>');
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("matches subunit quality report titles that use quality review wording", async () => {
      stubProcessFetch();
      const subunitRecord: ArchiveRecord = {
        ...processRecord,
        items: [
          {
            ...processRecord.items[0],
            sequence: "1",
            fileCode: "5028G01-SG-ZHHC-02-001",
            title: "高明分布式项目 光伏变电系统子单位工程质量报审表及验收记录",
          },
        ],
      };
      const written: Uint8Array[] = [];

      const result = await generateProcessDocs(
        [subunitRecord],
        {
          selectedCodes: [subunitRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          selectedTemplateCategories: ["summary-quality-acceptance"],
        },
        async (_path, bytes) => {
          written.push(bytes);
        },
      );
      const workbook = await workbookFrom(written[0]);

      expect(result.files).toHaveLength(1);
      expect(workbook.worksheets[0].getCell("I2").value).toBe("光伏变电系统子单位工程");
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("uses the shared subunit inspection application template when selected", async () => {
      stubProcessFetch();
      const subunitRecord: ArchiveRecord = {
        ...processRecord,
        items: [
          {
            ...processRecord.items[1],
            sequence: "2",
            fileCode: "5028G01-SG-ZHHC-02-001",
            title: "高明分布式项目 并网点光伏变电系统子单位工程质量报审表及验收记录",
            fileDate: "20250530",
          },
        ],
      };
      const written: Uint8Array[] = [];
      const paths: string[] = [];

      const result = await generateProcessDocs(
        [subunitRecord],
        {
          selectedCodes: [subunitRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          selectedTemplateCategories: ["subunit-inspection-application"],
          userFields: {
            projectName: "测试工程名称",
            supervisionDepartment: "测试监理项目部",
            constructionProjectManager: "测试项目经理",
          },
        },
        async (path, bytes) => {
          paths.push(path);
          written.push(bytes);
        },
      );
      const xml = await docxXml(written[0]);

      expect(result.files).toHaveLength(1);
      expect(fileNames(paths)[0]).toBe("2、5028G01-SG-ZHHC-02-001并网点光伏变电系统子单位工程报验申请单.docx");
      expect(xml).toContain("工程名称：测试工程名称");
      expect(xml).toContain("编号：5028G01-SG-ZHHC-02-001");
      expect(xml).toContain("测试监理项目部");
      expect(xml).toContain("工程现已施工完毕");
      expect(xml).toContain("<w:t>子单位工程</w:t>");
      expect(xml).toContain("<w:t>质量验收记录</w:t>");
      expect(xml).not.toContain("项目经理：测试项目经理");
      expect(xml).not.toContain("日    期：2025年 5 月 30 日");
      expect(xml).toContain("<w:u w:val=\"single\"/></w:rPr><w:t xml:space=\"preserve\">        并网点光伏变电系统        </w:t>");
      expect(xml).toContain("<w:u w:val=\"single\"/></w:rPr><w:t xml:space=\"preserve\">     并网点光伏变电系统     </w:t>");
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("generates explicitly supported 8312-002 electrical process templates", async () => {
      stubProcessFetch();
      const record8312002: ArchiveRecord = {
        ...processRecord,
        archiveCode: "5028G01-0011-8312-002",
        fullTitle: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目 光伏变电系统子单位工程质量报审，分部、分项及检验批质量验收记录文件",
        items: [
          {
            ...processRecord.items[2],
            sequence: "3",
            fileCode: "5028G01-SG-ZHHC-02-01-001",
            title: "高明分布式项目 子方阵场电气安装分部工程质量报验申请及验收记录",
          },
          {
            ...processRecord.items[19],
            sequence: "20",
            fileCode: "5028G01-SG-ZHHC-02-02-03-002",
            title: "高明分布式项目 #2光伏升压变电缆防火阻燃施工分项工程质量报验申请及验收记录",
          },
        ],
      };
      const paths: string[] = [];

      const result = await generateProcessDocs(
        [record8312002],
        {
          selectedCodes: [record8312002.archiveCode],
          outputDir: "/tmp/archive-output",
        },
        async (path) => {
          paths.push(path);
        },
      );

      expect(result.files).toHaveLength(4);
      expect(fileNames(paths)).toEqual([
        "3、5028G01-SG-ZHHC-02-01-001子方阵场电气安装分部工程报验申请单.docx",
        "3、5028G01-SG-ZHHC-02-01-001子方阵场电气安装分部工程质量验收记录（汇总用）.xlsx",
        "20、5028G01-SG-ZHHC-02-02-03-002#2光伏升压变电缆防火阻燃施工分项工程报验申请单.docx",
        "20、5028G01-SG-ZHHC-02-02-03-002#2光伏升压变电缆防火阻燃施工分项工程质量验收记录.xlsx",
      ]);
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
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

      expect(result.files).toHaveLength(5);
      expect(paths).toHaveLength(5);
      expect(paths.some((path) => path.includes("开工报审"))).toBe(true);
      expect(paths.some((path) => path.includes("子单位工程质量验收记录"))).toBe(true);
      expect(paths.some((path) => path.includes("子单位工程报验申请单"))).toBe(false);
      expect(result.skipped).toEqual([]);
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

    it("does not generate collector-line records from the regular process module", async () => {
      stubProcessFetch();

      const result = await generateProcessDocs(
        [collectorLineRecord],
        {
          selectedCodes: [collectorLineRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          templateModule: "process",
        },
        async () => {},
      );

      expect(result.files).toEqual([]);
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("does not force mixed 8312 records into the collector-line module only", async () => {
      stubProcessFetch();
      const mixedRecord: ArchiveRecord = {
        ...processRecord,
        fullTitle: `${processRecord.fullTitle}、集电线路安装分部工程质量验收记录文件`,
        volumeTitle: `${processRecord.volumeTitle}、集电线路安装分部工程质量验收记录文件`,
      };
      const paths: string[] = [];

      const result = await generateProcessDocs(
        [mixedRecord],
        {
          selectedCodes: [mixedRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          templateModule: "process",
          selectedTemplateCategories: ["start-report"],
        },
        async (path) => {
          paths.push(path);
        },
      );

      expect(result.files).toHaveLength(4);
      expect(fileNames(paths).filter((name) => name.includes("开工报审"))).toHaveLength(4);
      expect(paths.every((path) => path.includes("/过程资料/5028G01-0011-8312-001"))).toBe(true);
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("still generates collector-line templates from mixed non-8331 records in the collector-line module", async () => {
      stubProcessFetch();
      const mixedRecord: ArchiveRecord = {
        ...collectorLineRecord,
        categoryCode: "8312",
        archiveCode: "5028G01-0011-8312-004",
        fullTitle: "中核汇能高明创楷3.58904MWp屋顶分布式光伏项目 子方阵电气安装分部、集电线路安装分部工程质量验收记录文件",
        volumeTitle: "子方阵电气安装分部、集电线路安装分部工程质量验收记录文件",
        items: [
          {
            ...collectorLineRecord.items[0],
            title: "高明分布式项目 集电线路安装分部工程开工报审",
          },
        ],
      };
      const paths: string[] = [];

      const result = await generateProcessDocs(
        [mixedRecord],
        {
          selectedCodes: [mixedRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          templateModule: "collector-line",
          selectedTemplateCategories: ["start-report"],
        },
        async (path) => {
          paths.push(path);
        },
      );

      expect(result.files).toHaveLength(1);
      expect(fileNames(paths)[0]).toContain("集电线路安装分部工程开工报审");
      expect(paths[0]).toContain("/集电线路安装工程/5028G01-0011-8312-004");
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("generates collector-line templates only in the collector-line module", async () => {
      stubProcessFetch();
      const paths: string[] = [];

      const result = await generateProcessDocs(
        [collectorLineRecord],
        {
          selectedCodes: [collectorLineRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          templateModule: "collector-line",
        },
        async (path) => {
          paths.push(path);
        },
      );
      const names = fileNames(paths);

      expect(result.files).toHaveLength(38);
      expect(names.filter((name) => name.endsWith(".docx"))).toHaveLength(16);
      expect(names.filter((name) => name.endsWith(".xlsx"))).toHaveLength(22);
      expect(paths.every((path) => path.includes("/集电线路安装工程/5028G01-0011-8331-001"))).toBe(true);
      expect(names.some((name) => name.includes("集电线路安装工程分部开工报审"))).toBe(true);
      expect(names.some((name) => name.includes("电缆带电试运签证"))).toBe(true);
      expect(names.some((name) => name.includes("电力电缆终端制作安装分项工程质量验收表"))).toBe(true);
      expect(names.some((name) => name.includes("电力电缆中间接头制作安装分项工程质量验收表"))).toBe(true);
      expect(names.some((name) => name.includes("附件（照片）"))).toBe(false);
      expect(names.some((name) => name.toLowerCase().endsWith(".pdf"))).toBe(false);
      expect(names.some((name) => name.includes("AZ-025"))).toBe(false);
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("uses the shared start-report template for collector-line start reports", async () => {
      stubProcessFetch();
      const written: Uint8Array[] = [];

      const result = await generateProcessDocs(
        [collectorLineRecord],
        {
          selectedCodes: [collectorLineRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          templateModule: "collector-line",
          selectedTemplateCategories: ["start-report"],
        },
        async (_path, bytes) => {
          written.push(bytes);
        },
      );
      const xml = await docxXml(written[0]);
      const paragraphs = Array.from(xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)).map((match) => match[0]);
      const headerParagraphs = paragraphs.filter((paragraph) => paragraph.includes("工程名称：") || paragraph.includes("编号："));

      expect(result.files).toHaveLength(1);
      expect(headerParagraphs).toHaveLength(1);
      expect(headerParagraphs[0]).toContain("工程名称：中核汇能高明创楷3.58904MWp屋顶分布式光伏项目");
      expect(headerParagraphs[0]).toContain("编号：5028G01-SG-ZHHC-KG-15");
      expect(xml).toContain("开工报审表");
      expect(xml).toContain("致：");
    });

    it("calibrates collector-line shared xlsx workbook content from source items", async () => {
      stubProcessFetch();
      const written: Array<{ path: string; bytes: Uint8Array }> = [];

      const result = await generateProcessDocs(
        [collectorLineRecord],
        {
          selectedCodes: [collectorLineRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          templateModule: "collector-line",
        },
        async (path, bytes) => {
          written.push({ path, bytes });
        },
      );
      const workbookByName = async (nameIncludes: string) => {
        const entry = written.find((item) => item.path.includes(nameIncludes) && item.path.endsWith(".xlsx"));
        expect(entry?.path).toBeTruthy();
        return workbookFrom(entry!.bytes);
      };

      const trial5 = await workbookByName("01315A-01315B");
      expect(trial5.worksheets[0].name).toContain("01315A-01315B");
      expect(cellText(trial5.worksheets[0].getCell("B2"))).toContain("01315A-01315B");
      expect(cellText(trial5.worksheets[0].getCell("B2"))).not.toContain("00909A-00909B");

      const trial6 = await workbookByName("01316A-01316B");
      expect(trial6.worksheets[0].name).toContain("01316A-01316B");
      expect(cellText(trial6.worksheets[0].getCell("B2"))).toContain("01316A-01316B");

      const hidden9 = await workbookByName("G01柜601开关至G03柜602开关");
      expect(String(hidden9.worksheets[0].getCell("N8").value)).toBe("10KV广东创楷建设工程有限公司配电专用箱变G01柜601开关至10KV创楷光伏开关站G03柜602开关");
      const hidden10 = await workbookByName("G01柜801开关至G08柜605开关");
      expect(String(hidden10.worksheets[0].getCell("N8").value)).toBe("10KV创楷#1光伏升压变G01柜801开关至10KV创楷光伏开关站G08柜605开关");
      const hidden11 = await workbookByName("G01柜801开关至G09柜606开关");
      expect(String(hidden11.worksheets[0].getCell("N8").value)).toBe("10KV创楷#2光伏升压变G01柜801开关至10KV创楷光伏开关站G09柜606开关");

      const terminal12 = await workbookByName("03-08-02-001");
      expect(terminal12.worksheets[0].getCell("F6").value).toBe("10KV创楷#1光伏升压变");
      const terminal14 = await workbookByName("03-08-02-003");
      expect(terminal14.worksheets[0].getCell("F6").value).toBe("10KV创楷光伏开关柜G01柜");
      const joint19 = await workbookByName("03-08-03-002");
      expect(joint19.worksheets[0].getCell("F6").value).toBe("10KV创楷#2光伏升压变");
      const joint20 = await workbookByName("03-08-03-003");
      expect(joint20.worksheets[0].getCell("F6").value).toBe("10KV创楷光伏开关柜G01柜");

      expect(result.files).toHaveLength(38);
      expect(result.errors).toEqual([]);
    });

    it("summarizes collector-line division evaluation workbooks by template groups", async () => {
      stubProcessFetch();
      const written: Array<{ path: string; bytes: Uint8Array }> = [];

      await generateProcessDocs(
        [collectorLineRecord],
        {
          selectedCodes: [collectorLineRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          templateModule: "collector-line",
          selectedTemplateCategories: ["division-quality-acceptance"],
        },
        async (path, bytes) => {
          written.push({ path, bytes });
        },
      );
      const summaryEntry = written.find((item) => item.path.includes("集电线路安装工程分部工程质量验收评定表"));
      expect(summaryEntry?.path).toBeTruthy();
      const workbook = await workbookFrom(summaryEntry!.bytes);
      const sheet = workbook.worksheets[0];

      expect(sheet.getCell("D7").value).toBe("直埋电缆敷设");
      expect(sheet.getCell("D8").value).toBe("高压电缆终端制作安装");
      expect(sheet.getCell("D9").value).toBe("电力电缆中间接头制作安装");
      expect(sheet.getCell("D10").value).toBe("以下空白");
      expect(sheet.getCell("D11").value ?? "").toBe("");
      expect(sheet.getCell("T7").value).toBe("主要");
      expect(sheet.getCell("W9").value).toBe("合格");
    });

  });

  describe("docx rendering", () => {
    it("fills docx fields from source data and clears template-only company values", async () => {
    const item = processRecord.items[0];
    const template = readPublic("/templates/process-docs/开工报审.docx");
    const xml = await docxXml(
      renderProcessDocx(template, processRecord, item, {
        projectName: "测试自定义工程名称",
        supervisionDepartment: "测试监理项目部",
      }),
    );

    expect(xml).toContain("测试自定义工程名称");
    expect(xml).toContain(item.fileCode);
    expect(xml).toContain("测试监理项目部");
    expect(xml).not.toContain("河南中核五院研究设计有限公司");
  });
  });

  describe("xlsx rendering", () => {
    it("fills xlsx source fields and clears source-missing fields", async () => {
    const item = processRecord.items[4];
    const template = readPublic("/templates/process-docs/厂房支架安装分项工程质量验收表.xlsx");
    const workbook = await workbookFrom(
      await renderProcessWorkbook(template, processRecord, item, {
        projectName: "测试自定义工程名称",
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

    expect(sheet.getCell("F4").value).toBe("测试自定义工程名称");
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

    // Verify that the signature block at the bottom of the worksheet is untouched/blank
    let totalSignatureRows = 0;
    for (const sheet of workbook.worksheets) {
      let signatureRows = 0;
      let inSig = false;
      for (let r = 1; r <= sheet.rowCount; r++) {
        const row = sheet.getRow(r);
        if (!inSig) {
          row.eachCell({ includeEmpty: false }, (cell) => {
            if (cell.value && String(cell.value).includes("验收单位")) {
              inSig = true;
            }
          });
        }
        if (inSig) {
          signatureRows++;
          totalSignatureRows++;
          row.eachCell({ includeEmpty: false }, (cell) => {
            const val = typeof cell.value === "string" ? cell.value : "";
            expect(val).not.toContain("测试施工单位");
            expect(val).not.toContain("测试总承包单位");
          });
        }
      }
    }
    expect(totalSignatureRows).toBeGreaterThan(0);
  });

    it("leaves optional process fields blank when users do not fill them", async () => {
    const item = processRecord.items[4];
    const template = readPublic("/templates/process-docs/厂房支架安装分项工程质量验收表.xlsx");
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
    const template = readPublic("/templates/process-docs/厂房支架安装分项工程质量验收表.xlsx");
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
    const template = readPublic("/templates/process-docs/厂房接地装置安装检验批质量验收记录.xlsx");
    const workbook = await workbookFrom(
      await renderProcessWorkbook(template, processRecord, item, {
        constructionTechnicalLeader: "施工技术负责人",
      }),
    );
    const sheet = workbook.worksheets[0];

    expect(sheet.getCell("AE11").value).toBe("施工技术负责人");
  });

    it("fills inspection lot division and subitem names from source items", async () => {
    const division = {
      ...processRecord.items[0],
      sequence: "1",
      title: "高明分布式项目 通用工程分部工程质量报验申请及验收记录",
    };
    const subitem = {
      ...processRecord.items[0],
      sequence: "2",
      title: "高明分布式项目 建筑电气工程分项工程质量报验申请及验收记录",
    };
    const item = {
      ...processRecord.items[0],
      sequence: "3",
      title: "高明分布式项目 1#厂房光伏组件接地装置安装检验批质量验收记录",
    };
    const record = {
      ...processRecord,
      items: [division, subitem, item],
    };
    const template = readPublic("/templates/process-docs/厂房接地装置安装检验批质量验收记录.xlsx");
    const workbook = await workbookFrom(
      await renderProcessWorkbook(template, record, item, {
        projectName: "用户填写工程名称",
      }),
    );
    const sheet = workbook.worksheets[0];

    expect(sheet.getCell("Z6").value).toBe("通用工程");
    expect(sheet.getCell("I7").value).toBe("建筑电气工程");
  });

    it("fills numeric self-check values within quality standard ranges", async () => {
    const item = processRecord.items[43];
    const template = readPublic("/templates/process-docs/厂房墙架檩条支撑系统组装工程检验批质量验收记录.xlsx");
    const workbook = await workbookFrom(await renderProcessWorkbook(template, processRecord, item));
    const sheet = workbook.worksheets[0];
    const values = ["V14", "W14", "X14", "Y14", "Z14", "AA14", "AB14", "AC14", "AD14", "AE14"]
      .map((address) => Number(sheet.getCell(address).value));

    expect(values.every((value) => Number.isFinite(value))).toBe(true);
    expect(values.every((value) => value >= -2 && value <= 2)).toBe(true);
  });

    it("fills merged quality result values when templates use quality acceptance result columns", async () => {
    const item = processRecord.items[4];
    const template = readPublic("/templates/process-docs/厂房支架安装分项工程质量验收表.xlsx");
    const sourceWorkbook = await workbookFrom(template);
    const sourceValue = String(sourceWorkbook.worksheets[0].getCell("V21").value);

    const workbook = await workbookFrom(await renderProcessWorkbook(template, processRecord, item));
    const sheet = workbook.worksheets[0];
    const renderedValue = String(sheet.getCell("V21").value);
    const values = renderedValue.split(",").map((value) => Number(value));

    expect(values).toHaveLength(10);
    expect(values.every((value) => Number.isFinite(value))).toBe(true);
    expect(values.every((value) => value >= 0 && value <= 2)).toBe(true);
    expect(renderedValue).not.toBe(sourceValue); // Verify it actually overwrote the static value
  });

    it("distributes subitem quality acceptance values across upper-limit standards", async () => {
    const item = processRecord.items[4];
    const template = readPublic("/templates/process-docs/厂房支架安装分项工程质量验收表.xlsx");
    const workbook = await workbookFrom(await renderProcessWorkbook(template, processRecord, item));
    const sheet = workbook.worksheets[0];
    const valuesForTwo = String(sheet.getCell("V21").value).split(",").map(Number);
    const valuesForThree = String(sheet.getCell("V22").value).split(",").map(Number);

    expect(valuesForTwo).toHaveLength(10);
    expect(valuesForTwo.every((value) => Number.isInteger(value) && value >= 0 && value <= 2)).toBe(true);
    expect(new Set(valuesForTwo).size).toBeGreaterThan(1);
    expect(valuesForThree).toHaveLength(10);
    expect(valuesForThree.every((value) => Number.isInteger(value) && value >= 0 && value <= 3)).toBe(true);
    expect(new Set(valuesForThree).size).toBeGreaterThan(1);
  });

    it("uses absolute values for not-greater-than plus-minus quality standards", async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    sheet.mergeCells("A1:B1");
    sheet.mergeCells("C1:L1");
    sheet.mergeCells("A2:B2");
    sheet.mergeCells("C2:L2");
    sheet.getCell("A1").value = "质量标准";
    sheet.getCell("C1").value = "质量验收结果";
    sheet.getCell("A2").value = "不应大于±1°";

    const buffer = await workbook.xlsx.writeBuffer();
    const rendered = await workbookFrom(await renderProcessWorkbook(new Uint8Array(buffer), processRecord, processRecord.items[0]));
    const values = String(rendered.worksheets[0].getCell("C2").value).split(",").map(Number);

    expect(values).toHaveLength(10);
    expect(values.every((value) => value === 0 || value === 1)).toBe(true);
    expect(new Set(values).size).toBe(2);
  });

    it("preserves inspection lot workbook print margins and scale", async () => {
    const item = processRecord.items[35];
    const template = readPublic("/templates/process-docs/厂房接地装置安装检验批质量验收记录.xlsx");
    const rendered = await renderProcessWorkbook(template, processRecord, item);
    const sourceXml = await xlsxXml(template, "xl/worksheets/sheet1.xml");
    const renderedXml = await xlsxXml(rendered, "xl/worksheets/sheet1.xml");

    expect(renderedXml.match(/<pageMargins[^>]+>/)?.[0]).toBe(sourceXml.match(/<pageMargins[^>]+>/)?.[0]);
    expect(renderedXml.match(/<pageSetup[^>]+>/)?.[0]).not.toContain('scale="100"');
  });

    it("adds print areas to subitem quality workbooks", async () => {
    const item = processRecord.items[4];
    const template = readPublic("/templates/process-docs/厂房支架安装分项工程质量验收表.xlsx");
    const rendered = await renderProcessWorkbook(template, processRecord, item);
    const sourceWorkbookXml = await xlsxXml(template, "xl/workbook.xml");
    const renderedWorkbookXml = await xlsxXml(rendered, "xl/workbook.xml");
    const sourceSheetXml = await xlsxXml(template, "xl/worksheets/sheet1.xml");
    const renderedSheetXml = await xlsxXml(rendered, "xl/worksheets/sheet1.xml");

    expect(sourceWorkbookXml).not.toContain("_xlnm.Print_Area");
    expect(renderedWorkbookXml).toContain("_xlnm.Print_Area");
    expect(renderedSheetXml.match(/<pageMargins[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageMargins[^>]+>/)?.[0]);
    expect(renderedSheetXml.match(/<pageSetup[^>]+>/)?.[0]).not.toContain('scale="100"');
  });
  });

  describe("summary workbook rendering", () => {
    it("keeps summary workbook header merge and style intact", async () => {
    const summaryTemplates = [
      "/templates/process-docs/子单位工程质量验收记录.xlsx",
      "/templates/process-docs/子方阵支架及组件安装分部工程质量验收记录（汇总用）.xlsx",
      "/templates/process-docs/通用工程分部工程质量验收记录（汇总用）.xlsx",
      "/templates/process-docs/主体结构工程分部工程质量验收记录（汇总用）.xlsx",
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
      const sourceSheetXml = await sourceZip.file("xl/worksheets/sheet1.xml")!.async("string");
      const renderedSheetXml = await renderedZip.file("xl/worksheets/sheet1.xml")!.async("string");

      expect(rendered.worksheets[0].getCell("B2").style).toEqual(source.worksheets[0].getCell("B2").style);
      expect(rendered.worksheets[0].getCell("B2").isMerged).toBe(source.worksheets[0].getCell("B2").isMerged);
      expect(rendered.worksheets[0].getCell("C2").master.address).toBe(source.worksheets[0].getCell("C2").master.address);
      expect(await renderedZip.file("xl/styles.xml")!.async("string")).toBe(await sourceZip.file("xl/styles.xml")!.async("string"));
      expect(renderedSheetXml.match(/<pageMargins[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageMargins[^>]+>/)?.[0]);
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

function createCollectorLineRecord(): ArchiveRecord {
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

function createSubunitSummaryFixtureItems() {
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
