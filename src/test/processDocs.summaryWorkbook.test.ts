import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { renderSummaryWorkbook } from "../lib/processDocs";
import { processRecord, readPublic, workbookFrom } from "./processDocsTestUtils";

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
    const sourceWorkbookXml = await sourceZip.file("xl/workbook.xml")!.async("string");
    const renderedWorkbookXml = await renderedZip.file("xl/workbook.xml")!.async("string");

    expect(rendered.worksheets[0].getCell("B2").style).toEqual(source.worksheets[0].getCell("B2").style);
    expect(rendered.worksheets[0].getCell("B2").isMerged).toBe(source.worksheets[0].getCell("B2").isMerged);
    expect(rendered.worksheets[0].getCell("C2").master.address).toBe(source.worksheets[0].getCell("C2").master.address);
    expect(await renderedZip.file("xl/styles.xml")!.async("string")).toBe(await sourceZip.file("xl/styles.xml")!.async("string"));
    expect(renderedSheetXml.match(/<printOptions[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<printOptions[^>]+>/)?.[0]);
    expect(renderedSheetXml.match(/<pageMargins[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageMargins[^>]+>/)?.[0]);
    expect(renderedSheetXml.match(/<pageSetup[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageSetup[^>]+>/)?.[0]);
    expect(renderedWorkbookXml.match(/<definedNames>[\s\S]*<\/definedNames>/)?.[0]).toBe(sourceWorkbookXml.match(/<definedNames>[\s\S]*<\/definedNames>/)?.[0]);
    expect(rendered.worksheets[0].getCell("G7").value).toBe("测试总承包单位");
    expect(rendered.worksheets[0].getCell("G8").value).toBe("测试施工单位");
    expect(rendered.worksheets[0].getCell("G9").value).toBe("测试分包单位");
    expect(rendered.worksheets[0].getCell("U7").value).toBe("总包负责人");
    expect(rendered.worksheets[0].getCell("AF8").value).toBe("施工技术负责人");
    expect(rendered.worksheets[0].getCell("AF9").value).toBe("光伏方阵安装");
  }
});
});
