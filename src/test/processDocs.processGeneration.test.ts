import { describe, expect, it } from "vitest";
import {
  countProcessGenerationFiles,
  generateProcessDocs,
  loadProcessManifest,
  matchingAllProcessTemplates,
  renderProcessWorkbook,
} from "../lib/processDocs";
import type { ArchiveRecord } from "../lib/types";
import {
  cellText,
  createGridConnectedElectricalRecord,
  createSubunitSummaryFixtureItems,
  createSuixiGridConnectedElectricalRecord,
  docxXml,
  expectValidOfficeFile,
  fileNames,
  processRecord,
  readPublic,
  stubProcessFetch,
  workbookFrom,
  xlsxXml,
} from "./processDocsTestUtils";

describe("template matching", () => {
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
    const sourceTemplate = readPublic("/templates/process-docs/子单位工程质量验收记录.xlsx");
    const sourceSheetXml = await xlsxXml(sourceTemplate, "xl/worksheets/sheet1.xml");
    const sourceWorkbookXml = await xlsxXml(sourceTemplate, "xl/workbook.xml");
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
    expect(sheetXml.match(/<pageSetUpPr[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageSetUpPr[^>]+>/)?.[0]);
    expect(sheetXml.match(/<printOptions[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<printOptions[^>]+>/)?.[0]);
    expect(sheetXml.match(/<pageMargins[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageMargins[^>]+>/)?.[0]);
    expect(sheetXml.match(/<pageSetup[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageSetup[^>]+>/)?.[0]);
    expect(sheetXml.indexOf("<pageMargins")).toBeLessThan(sheetXml.indexOf("<pageSetup"));
    expect(workbookXml.match(/<definedNames>[\s\S]*<\/definedNames>/)?.[0]).toBe(sourceWorkbookXml.match(/<definedNames>[\s\S]*<\/definedNames>/)?.[0]);
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
});
