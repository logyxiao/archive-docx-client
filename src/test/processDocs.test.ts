import { readFileSync } from "node:fs";
import JSZip from "jszip";
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import {
  countProcessGenerationFiles,
  generateProcessDocs,
  generateSelectedProcessDocs,
  loadProcessManifest,
  matchingAllProcessTemplates,
  renderProcessDocx,
  renderProcessWorkbook,
  renderSummaryWorkbook,
  type ProcessTemplate,
} from "../lib/processDocs";
import { replaceTemplatePlaceholders } from "../lib/process-docs/textReplacement";
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

async function expectValidOfficeFile(path: string, bytes: Uint8Array) {
  const zip = await JSZip.loadAsync(bytes);
  if (path.endsWith(".docx")) {
    expect(zip.file("word/document.xml")).toBeTruthy();
    return;
  }
  if (path.endsWith(".xlsx")) {
    expect(zip.file("xl/workbook.xml")).toBeTruthy();
  }
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

    it("prioritizes enabled user templates by configured keywords", () => {
      const record = createSuixiGridConnectedElectricalRecord();
      const item = record.items[0];
      const userTemplate: ProcessTemplate = {
        sequence: 0,
        kind: "docx",
        originalName: "自定义电缆终端制作模板.docx",
        templateFile: "自定义电缆终端制作模板.docx",
        outputExtension: ".docx",
        userTemplatePath: "/tmp/process-templates/自定义电缆终端制作模板.docx",
        displayName: "自定义电缆终端制作模板",
        matchKeywords: ["电缆终端制作"],
        matchMode: "any",
        templateModule: "process",
        enabled: true,
      };
      const builtInTemplate: ProcessTemplate = {
        sequence: 0,
        kind: "xlsx",
        originalName: "电力电缆终端制作安装分项工程质量验收表.xlsx",
        templateFile: "电力电缆终端制作安装分项工程质量验收表.xlsx",
        outputExtension: ".xlsx",
      };

      const matches = matchingAllProcessTemplates(record, item, [builtInTemplate, userTemplate]);

      expect(matches.map((match) => match.template.templateFile)).toEqual([
        "自定义电缆终端制作模板.docx",
        "电力电缆终端制作安装分项工程质量验收表.xlsx",
      ]);
    });

    it("ignores disabled user templates and supports legacy filename matching", () => {
      const item = {
        ...processRecord.items[4],
        title: "高明分布式项目 需要人工选择模板的文件",
      };
      const record = { ...processRecord, items: [item] };
      const disabledTemplate: ProcessTemplate = {
        sequence: 0,
        kind: "docx",
        originalName: "需要人工选择模板.docx",
        templateFile: "需要人工选择模板.docx",
        outputExtension: ".docx",
        userTemplatePath: "/tmp/process-templates/需要人工选择模板.docx",
        matchKeywords: ["需要人工选择模板"],
        templateModule: "process",
        enabled: false,
      };
      const legacyTemplate: ProcessTemplate = {
        sequence: 0,
        kind: "docx",
        originalName: "需要人工选择模板.docx",
        templateFile: "需要人工选择模板.docx",
        outputExtension: ".docx",
        userTemplatePath: "/tmp/process-templates/legacy/需要人工选择模板.docx",
      };

      const matches = matchingAllProcessTemplates(record, item, [disabledTemplate, legacyTemplate]);

      expect(matches.map((match) => match.template.userTemplatePath)).toEqual([
        "/tmp/process-templates/legacy/需要人工选择模板.docx",
      ]);
    });

    it("requires every keyword when user templates use all-keyword matching", () => {
      const record = createSuixiGridConnectedElectricalRecord();
      const item = record.items[0];
      const template: ProcessTemplate = {
        sequence: 0,
        kind: "docx",
        originalName: "严格匹配模板.docx",
        templateFile: "严格匹配模板.docx",
        outputExtension: ".docx",
        userTemplatePath: "/tmp/process-templates/严格匹配模板.docx",
        matchKeywords: ["电缆终端制作", "不存在关键词"],
        matchMode: "all",
        templateModule: "process",
      };

      expect(matchingAllProcessTemplates(record, item, [template])).toEqual([]);
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
      const projectName = "中核汇能遂溪国恒坤达3.84MW屋顶分布式光伏项目";
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
            projectName,
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
      const templateXml = await docxXml(readPublic("/templates/process-docs/子单位工程报验申请单.docx"));
      const projectCodeParagraph = Array.from(xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g))
        .map((match) => match[0])
        .find((paragraph) => paragraph.includes("工程名称：") && paragraph.includes("编号："));

      expect(result.files).toHaveLength(1);
      expect(fileNames(paths)[0]).toBe("2、5028G01-SG-ZHHC-02-001并网点光伏变电系统子单位工程报验申请单.docx");
      expect(xml).toContain(`工程名称：${projectName}`);
      expect(xml).toContain("编号：5028G01-SG-ZHHC-02-001");
      expect(xml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0]).toBe(templateXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0]);
      expect(projectCodeParagraph).toMatch(/<w:spacing\b(?=[^>]*w:before="0")(?=[^>]*w:after="0")(?=[^>]*w:line="240")/);
      expect(projectCodeParagraph).toContain('<w:sz w:val="16"/>');
      expect(projectCodeParagraph).toContain('<w:szCs w:val="16"/>');
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

    it("uses the subunit inspection application template for standalone subunit quality review titles", async () => {
      stubProcessFetch();
      const subunitRecord: ArchiveRecord = {
        ...processRecord,
        items: [
          {
            ...processRecord.items[1],
            sequence: "2",
            fileCode: "5028G01-SG-ZHHC-02-001",
            title: "高明分布式项目 并网点光伏变电系统子单位工程质量报审表",
          },
        ],
      };
      const paths: string[] = [];

      const result = await generateProcessDocs(
        [subunitRecord],
        {
          selectedCodes: [subunitRecord.archiveCode],
          outputDir: "/tmp/archive-output",
        },
        async (path) => {
          paths.push(path);
        },
      );

      expect(result.files).toHaveLength(1);
      expect(fileNames(paths)).toEqual(["2、5028G01-SG-ZHHC-02-001并网点光伏变电系统子单位工程报验申请单.docx"]);
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("matches and generates the full 8312-002 electrical process template set", async () => {
      stubProcessFetch();
      const record8312002 = createGridConnectedElectricalRecord();
      const manifest = await loadProcessManifest();
      const defaultMatches = record8312002.items.flatMap((item) =>
        matchingAllProcessTemplates(record8312002, item, manifest.templates),
      );
      const defaultTemplateFiles = defaultMatches.map((match) => match.template.templateFile);
      const written: Array<{ path: string; bytes: Uint8Array }> = [];

      expect(defaultMatches).toHaveLength(42);
      expect(countProcessGenerationFiles(
        [record8312002],
        [record8312002.archiveCode],
        manifest.templates,
      )).toBe(42);
      expect(defaultTemplateFiles).not.toContain("建筑电气工程分项工程质量验收记录.xlsx");
      expect(defaultTemplateFiles).toEqual(expect.arrayContaining([
        "并网点光伏变电系统子单位工程质量验收记录（汇总用）.xlsx",
        "子方阵电气安装分部工程质量验收记录（汇总用）.xlsx",
        "逆变器安装分项工程质量验收表.xlsx",
        "逆变器施工安装记录.xlsx",
        "逆变器手动分合闸检查记录.xlsx",
        "逆变器通讯调试记录.xlsx",
        "逆变器外观、主要元器件、控制电源、直交流侧接线及极性（相序）、绝缘、接地检查记录.xlsx",
        "干式变压器安装分项工程质量验收表.xlsx",
        "屋外接地装置安装分项工程质量验收表.xlsx",
        "屋内接地装置安装分项工程质量验收表.xlsx",
        "子方阵电气线路安装分部工程质量验收记录（汇总用）.xlsx",
        "电缆桥架安装分项质量检查验收评定表.xlsx",
        "电缆敷设分项工程质量检查验收评定表.xlsx",
        "光伏升压变电缆防火阻燃施工分项工程质量验收表.xlsx",
        "开关柜电缆防火阻燃施工分项工程质量验收表.xlsx",
      ]));

      const result = await generateProcessDocs(
        [record8312002],
        {
          selectedCodes: [record8312002.archiveCode],
          outputDir: "/tmp/archive-output",
        },
        async (path, bytes) => {
          written.push({ path, bytes });
        },
      );

      expect(result.files).toHaveLength(42);
      expect(written).toHaveLength(42);
      expect(fileNames(written.map((item) => item.path))).toEqual(expect.arrayContaining([
        "3、5028G01-SG-ZHHC-02-01-001子方阵场电气安装分部工程质量验收记录（汇总用）.xlsx",
        "4、5028G01-SG-ZHHC-02-01-01-001逆变器安装分项工程质量验收表.xlsx",
        "8、高明分布式项目 逆变器外观、主要元器件、控制电源、直交流侧接线及极性（相序）、绝缘、接地检查记录）.xlsx",
        "17、5028G01-SG-ZHHC-02-02-01-001电缆桥架安装分项工程质量检查验收评定表.xlsx",
        "24、5028G01-SG-ZHHC-02-02-03-006开关柜G09柜电缆防火阻燃施工分项工程质量验收表.xlsx",
      ]));
      for (const file of written) {
        await expectValidOfficeFile(file.path, file.bytes);
      }
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);
    });

    it("matches Suixi short electrical titles to reusable quality workbook templates", async () => {
      stubProcessFetch();
      const suixiRecord = createSuixiGridConnectedElectricalRecord();
      const manifest = await loadProcessManifest();
      const templateFilesForSequence = (sequence: string) => {
        const item = suixiRecord.items.find((entry) => entry.sequence === sequence)!;
        return matchingAllProcessTemplates(suixiRecord, item, manifest.templates).map((match) => match.template.templateFile);
      };
      const written: Array<{ path: string; bytes: Uint8Array }> = [];

      expect(templateFilesForSequence("14")[0]).toBe("电力电缆终端制作安装分项工程质量验收表.xlsx");
      expect(templateFilesForSequence("14")).toContain("电力电缆终端制作安装分项工程质量验收表.xlsx");
      expect(templateFilesForSequence("14")).not.toContain("建筑电气工程分项工程质量验收记录.xlsx");
      expect(templateFilesForSequence("15")[0]).toBe("光伏升压变电缆防火阻燃施工分项工程质量验收表.xlsx");
      expect(templateFilesForSequence("16")[0]).toBe("二次回路检查及控制电缆接线分项工程质量检查验收评定表.xlsx");
      const cableTrayFabricationMatches = matchingAllProcessTemplates(
        suixiRecord,
        {
          ...suixiRecord.items[3],
          title: "遂溪分布式项目 电缆桥架制作及安装分项工程质量报验申请及验收记录",
        },
        manifest.templates,
      ).map((match) => match.template.templateFile);
      expect(cableTrayFabricationMatches[0]).toBe("电缆桥架安装分项质量检查验收评定表.xlsx");
      expect(cableTrayFabricationMatches).toContain("分项工程报验申请单.docx");
      const cableLineItem = {
        ...suixiRecord.items[3],
        title: "遂溪分布式项目 电缆线路施工安装分项工程质量报验申请及验收记录",
      };
      const cableLineMatches = matchingAllProcessTemplates(
        suixiRecord,
        cableLineItem,
        manifest.templates,
      ).map((match) => match.template.templateFile);
      expect(cableLineMatches[0]).toBe("电缆线路施工安装分项质量检查验收评定表.xlsx");
      expect(cableLineMatches).toContain("分项工程报验申请单.docx");
      expect(cableLineMatches).not.toContain("电缆桥架安装分项质量检查验收评定表.xlsx");
      const cableLineTemplate = readPublic("/templates/process-docs/电缆线路施工安装分项质量检查验收评定表.xlsx");
      const cableLineWorkbook = await workbookFrom(await renderProcessWorkbook(cableLineTemplate, suixiRecord, cableLineItem));
      const cableLineSheet = cableLineWorkbook.getWorksheet("电缆线路施工安装")!;
      expect(cellText(cableLineSheet.getCell("B2"))).toContain("电缆线路施工安装分项工程质量检查验收评定表");
      expect(cableLineSheet.getCell("G9").value).toBe("电缆线路施工安装");
      expect(templateFilesForSequence("17")).toEqual(["直流配电柜检验批质量验收记录.xlsx"]);
      expect(templateFilesForSequence("18")).toEqual(["直流配电柜检验批质量验收记录.xlsx"]);
      expect(matchingAllProcessTemplates(
        suixiRecord,
        {
          ...suixiRecord.items[3],
          title: "遂溪分布式项目 国恒部分直流配电柜安装检验批质量验收记录",
        },
        manifest.templates,
      ).map((match) => match.template.templateFile)).not.toContain("直流配电柜检验批质量验收记录.xlsx");
      expect(templateFilesForSequence("19")).toEqual([
        "集电线路安装工程分部工程报验申请单.docx",
        "集电线路安装工程分部工程质量验收评定表.xlsx",
      ]);

      const result = await generateProcessDocs(
        [suixiRecord],
        {
          selectedCodes: [suixiRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          selectedTemplateCategories: ["subitem-quality-acceptance"],
        },
        async (path, bytes) => {
          written.push({ path, bytes });
        },
      );

      expect(result.files).toHaveLength(3);
      expect(fileNames(written.map((item) => item.path))).toEqual(expect.arrayContaining([
        "14、5028G02-SG-ZHHC-03-02-03-001电缆终端制作分项工程质量验收表.xlsx",
        "15、5028G02-SG-ZHHC-03-02-04-001电缆防火与阻燃分项工程质量验收表.xlsx",
        "16、5028G02-SG-ZHHC-03-02-05-001电气二次系统分项工程质量检查验收评定表.xlsx",
      ]));
      for (const file of written) {
        await expectValidOfficeFile(file.path, file.bytes);
      }
      expect(result.skipped).toEqual([]);
      expect(result.errors).toEqual([]);

      const lotWritten: Array<{ path: string; bytes: Uint8Array }> = [];
      const lotResult = await generateProcessDocs(
        [suixiRecord],
        {
          selectedCodes: [suixiRecord.archiveCode],
          outputDir: "/tmp/archive-output",
          selectedTemplateCategories: ["inspection-lot-acceptance"],
        },
        async (path, bytes) => {
          lotWritten.push({ path, bytes });
        },
      );
      const lotNames = fileNames(lotWritten.map((item) => item.path));
      const guohengEntry = lotWritten.find((item) => item.path.includes("国恒部分直流配电柜"));
      const kundaEntry = lotWritten.find((item) => item.path.includes("坤达部分直流配电柜"));

      expect(lotResult.files).toHaveLength(2);
      expect(lotNames).toEqual(expect.arrayContaining([
        "17、5028G02-SG-ZHHC-03-02-05-01-001国恒部分直流配电柜检验批质量验收记录.xlsx",
        "18、5028G02-SG-ZHHC-03-02-05-02-001坤达部分直流配电柜检验批质量验收记录.xlsx",
      ]));
      expect(guohengEntry?.path).toBeTruthy();
      expect(kundaEntry?.path).toBeTruthy();
      for (const file of lotWritten) {
        await expectValidOfficeFile(file.path, file.bytes);
      }
      const guohengWorkbook = await workbookFrom(guohengEntry!.bytes);
      const guohengSheet = guohengWorkbook.getWorksheet("第1页")!;
      const strictUpperValues = String(guohengSheet.getCell("V12").value).split(",").map(Number);

      expect(cellText(guohengSheet.getCell("B2"))).toContain("国恒部分直流配电柜 检验批质量检查验收评定表");
      expect(guohengSheet.getCell("F3").value).toBe("5028G02-0011");
      expect(guohengSheet.getCell("G5").value).toBe("并网点光伏变电系统");
      expect(guohengSheet.getCell("T5").value).toBe("子方阵电气线路安装");
      expect(guohengSheet.getCell("AE5").value).toBe("4台");
      expect(guohengSheet.getCell("G9").value).toBe("电气二次系统");
      expect(strictUpperValues).toHaveLength(10);
      expect(strictUpperValues.every((value) => Number.isFinite(value) && value >= 0 && value < 1.5)).toBe(true);
      const kundaWorkbook = await workbookFrom(kundaEntry!.bytes);
      expect(cellText(kundaWorkbook.getWorksheet("第1页")!.getCell("B2"))).toContain("坤达部分直流配电柜 检验批质量检查验收评定表");
      expect(lotResult.skipped).toEqual([]);
      expect(lotResult.errors).toEqual([]);
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

    it("replaces known placeholders and keeps unknown placeholders unchanged", () => {
    const item = processRecord.items[0];
    const text = replaceTemplatePlaceholders("{{项目名}} {{文件题名}} {{未知字段}}", processRecord, item, {
      projectName: "测试自定义工程名称",
    });
    const suixiRecord = createSuixiGridConnectedElectricalRecord();
    const inspectionLotItem = suixiRecord.items.find((entry) => entry.sequence === "17")!;
    const subjectText = replaceTemplatePlaceholders(
      "{{文件题名去项目}}|{{文件题名主题}}|{{验收项目名称}}",
      suixiRecord,
      inspectionLotItem,
      {},
    );

    expect(text).toBe(`测试自定义工程名称 ${item.title} {{未知字段}}`);
    expect(subjectText).toBe("国恒部分直流配电柜检验批质量验收记录|国恒部分直流配电柜|国恒部分直流配电柜");
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
    sheet.getCell("C1").value = "质量检查验收结果";
    sheet.getCell("A2").value = "不应大于±1°";

    const buffer = await workbook.xlsx.writeBuffer();
    const rendered = await workbookFrom(await renderProcessWorkbook(new Uint8Array(buffer), processRecord, processRecord.items[0]));
    const values = String(rendered.worksheets[0].getCell("C2").value).split(",").map(Number);

    expect(values).toHaveLength(10);
    expect(values.every((value) => value === 0 || value === 1)).toBe(true);
    expect(new Set(values).size).toBe(2);
  });

    it("fills lower-bound quality standards in a bounded random range", async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    sheet.mergeCells("A1:B1");
    sheet.mergeCells("C1:L1");
    sheet.mergeCells("A2:B2");
    sheet.mergeCells("C2:L2");
    sheet.getCell("A1").value = "质量标准";
    sheet.getCell("C1").value = "质量验收结果";
    sheet.getCell("A2").value = "≥6";

    const buffer = await workbook.xlsx.writeBuffer();
    const rendered = await workbookFrom(await renderProcessWorkbook(new Uint8Array(buffer), processRecord, processRecord.items[0]));
    const values = String(rendered.worksheets[0].getCell("C2").value).split(",").map(Number);

    expect(values).toHaveLength(10);
    expect(values.every((value) => Number.isInteger(value) && value >= 6 && value <= 9)).toBe(true);
    expect(new Set(values).size).toBeGreaterThan(1);
  });

    it("fills strict upper-bound quality standards in a bounded random range", async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    sheet.mergeCells("A1:B1");
    sheet.mergeCells("C1:L1");
    sheet.mergeCells("A2:B2");
    sheet.mergeCells("C2:L2");
    sheet.getCell("A1").value = "质量标准";
    sheet.getCell("C1").value = "质量验收结果";
    sheet.getCell("A2").value = "<1.5mm/m";

    const buffer = await workbook.xlsx.writeBuffer();
    const rendered = await workbookFrom(await renderProcessWorkbook(new Uint8Array(buffer), processRecord, processRecord.items[0]));
    const values = String(rendered.worksheets[0].getCell("C2").value).split(",").map(Number);

    expect(values).toHaveLength(10);
    expect(values.every((value) => Number.isFinite(value) && value >= 0 && value < 1.5)).toBe(true);
    expect(new Set(values).size).toBeGreaterThan(1);
  });

    it("fills quality result placeholders from lower-bound standards", async () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Sheet1");
    sheet.getCell("A1").value = "{{质量验收结果:≥6}}";
    sheet.getCell("A2").value = "{{质量验收结果:6~10}}";
    sheet.getCell("A3").value = "{{质量验收结果:<2}}";

    const buffer = await workbook.xlsx.writeBuffer();
    const rendered = await workbookFrom(await renderProcessWorkbook(new Uint8Array(buffer), processRecord, processRecord.items[0]));
    const lowerValues = String(rendered.worksheets[0].getCell("A1").value).split(",").map(Number);
    const rangeValues = String(rendered.worksheets[0].getCell("A2").value).split(",").map(Number);
    const strictUpperValues = String(rendered.worksheets[0].getCell("A3").value).split(",").map(Number);

    expect(lowerValues).toHaveLength(10);
    expect(lowerValues.every((value) => Number.isInteger(value) && value >= 6 && value <= 9)).toBe(true);
    expect(rangeValues).toHaveLength(10);
    expect(rangeValues.every((value) => Number.isInteger(value) && value >= 6 && value <= 10)).toBe(true);
    expect(strictUpperValues).toHaveLength(10);
    expect(strictUpperValues.every((value) => Number.isFinite(value) && value >= 0 && value < 2)).toBe(true);
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

function createGridConnectedElectricalRecord(): ArchiveRecord {
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

function createSuixiGridConnectedElectricalRecord(): ArchiveRecord {
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
