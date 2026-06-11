import { describe, expect, it } from "vitest";
import { generateProcessDocs } from "../lib/processDocs";
import type { ArchiveRecord } from "../lib/types";
import { cellText, collectorLineRecord, docxXml, fileNames, processRecord, stubProcessFetch, workbookFrom } from "./processDocsTestUtils";

describe("template matching", () => {
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
