import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { renderProcessWorkbook } from "../lib/processDocs";
import { cellText, processRecord, readPublic, workbookFrom, xlsxXml } from "./processDocsTestUtils";

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

  it("replaces file code placeholders in rich text workbook cells", async () => {
  const item = processRecord.items[4];
  const template = readPublic("/templates/process-docs/厂房支架安装分项工程质量验收表.xlsx");
  const workbook = await workbookFrom(await renderProcessWorkbook(template, processRecord, item));
  const firstSheet = workbook.getWorksheet("1#厂房")!;
  const secondSheet = workbook.getWorksheet("第2页")!;

  expect(cellText(firstSheet.getCell("AB4"))).toContain(`编号：${item.fileCode}`);
  expect(cellText(firstSheet.getCell("AB4"))).not.toContain("{{文件编号}}");
  expect(cellText(secondSheet.getCell("AA4"))).toContain(`编号：${item.fileCode}`);
  expect(cellText(secondSheet.getCell("AA4"))).not.toContain("{{文件编号}}");
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

  expect(renderedXml.match(/<printOptions[^>]+>/)?.[0]).toBe(sourceXml.match(/<printOptions[^>]+>/)?.[0]);
  expect(renderedXml.match(/<pageMargins[^>]+>/)?.[0]).toBe(sourceXml.match(/<pageMargins[^>]+>/)?.[0]);
  expect(renderedXml.match(/<pageSetup[^>]+>/)?.[0]).toBe(sourceXml.match(/<pageSetup[^>]+>/)?.[0]);
});

  it("does not add print areas when the template does not define them", async () => {
  const item = processRecord.items[4];
  const template = readPublic("/templates/process-docs/厂房支架安装分项工程质量验收表.xlsx");
  const rendered = await renderProcessWorkbook(template, processRecord, item);
  const sourceWorkbookXml = await xlsxXml(template, "xl/workbook.xml");
  const renderedWorkbookXml = await xlsxXml(rendered, "xl/workbook.xml");
  const sourceSheetXml = await xlsxXml(template, "xl/worksheets/sheet1.xml");
  const renderedSheetXml = await xlsxXml(rendered, "xl/worksheets/sheet1.xml");

  expect(sourceWorkbookXml).not.toContain("_xlnm.Print_Area");
  expect(renderedWorkbookXml).not.toContain("_xlnm.Print_Area");
  expect(renderedSheetXml.match(/<sheetViews>[\s\S]*?<\/sheetViews>/)?.[0]).toBe(sourceSheetXml.match(/<sheetViews>[\s\S]*?<\/sheetViews>/)?.[0]);
  expect(renderedSheetXml.match(/<mergeCells\b[\s\S]*?<\/mergeCells>/)?.[0]).toBe(sourceSheetXml.match(/<mergeCells\b[\s\S]*?<\/mergeCells>/)?.[0]);
  expect(renderedSheetXml.match(/<printOptions[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<printOptions[^>]+>/)?.[0]);
  expect(renderedSheetXml.match(/<pageMargins[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageMargins[^>]+>/)?.[0]);
  expect(renderedSheetXml.match(/<pageSetup[^>]+>/)?.[0]).toBe(sourceSheetXml.match(/<pageSetup[^>]+>/)?.[0]);
});
});
