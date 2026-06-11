import { describe, expect, it } from "vitest";
import { renderProcessDocx } from "../lib/processDocs";
import { replaceTemplatePlaceholders } from "../lib/process-docs/textReplacement";
import { createSuixiGridConnectedElectricalRecord, docxXml, processRecord, readPublic } from "./processDocsTestUtils";

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
