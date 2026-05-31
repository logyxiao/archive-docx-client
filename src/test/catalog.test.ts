import { readFileSync } from "node:fs";
import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import { CATALOG_OUTPUT_NAME, generateArchiveCatalog, renderCatalogWorkbook } from "../lib/catalog";
import { parseArchiveWorkbook } from "../lib/excel";
import { createArchiveWorkbookFixture } from "./fixtures";

const records = parseArchiveWorkbook(createArchiveWorkbookFixture());
const template = readFileSync("public/templates/catalog.xlsx");

async function readWorkbook(bytes: Uint8Array): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(Buffer.from(bytes));
  return workbook;
}

describe("renderCatalogWorkbook", () => {
  it("renders archive catalog and detail sheets from selected records", async () => {
    const record = records.find((item) => item.archiveCode === "5028G01-0011-842-001")!;
    const workbook = await readWorkbook(await renderCatalogWorkbook(template, [record]));
    const catalog = workbook.worksheets[0];
    const detail = workbook.getWorksheet(record.archiveCode);

    expect(catalog.name).toBe("案卷目录");
    expect(catalog.getCell("B3").value).toBe(record.archiveCode);
    expect(catalog.getCell("C3").value).toBe(record.fullTitle);
    expect(catalog.getCell("D3").value).toBe(record.totalPages);
    expect(catalog.getCell("E3").value).toBe(record.retentionPeriod);
    expect(catalog.getCell("F3").value).toBe(record.filingUnit);
    expect(catalog.pageSetup.printArea).toBe("A1:G3");
    expect(catalog.getCell("B4").value).toBeNull();
    expect(detail?.getCell("A2").value).toBe(` 档号：${record.archiveCode}`);
    expect(detail?.getCell("F3").value).toBe("页号");
    expect(detail?.getCell("D4").value).toBe(record.items[0].title);
    expect(detail?.getCell("E4").value).toBe("20250412");
    expect(detail?.getCell("F14").value).toBe("11");
    expect(detail?.getCell("A15").value).toBe("");
    expect(detail?.getRow(4).height).toBe(44);
    expect(detail?.getRow(17).height).toBe(44);
    expect(detail?.pageSetup.fitToPage).toBe(true);
    expect(detail?.pageSetup.fitToHeight).toBe(1);
    expect(detail?.pageSetup.fitToWidth).toBe(1);
    expect(detail?.pageSetup.printArea).toBe("A1:G17");
  });

  it("splits long detail lists into continuation sheets", async () => {
    const record = records.find((item) => item.archiveCode === "5028G01-0011-8312-001")!;
    const workbook = await readWorkbook(await renderCatalogWorkbook(template, [record]));
    const first = workbook.getWorksheet(record.archiveCode);
    const second = workbook.getWorksheet(`${record.archiveCode} (2)`);

    expect(first).toBeDefined();
    expect(second).toBeDefined();
    expect(first?.getCell("A17").value).toBe("14");
    expect(second?.getCell("A4").value).toBe("15");
    expect(second?.getCell("D5").value).toBe("光伏发电系统验收文件16");
    expect(second?.getCell("A6").value).toBe("");
  });

  it("limits archive catalog print area to actual records", async () => {
    const manyRecords = Array.from({ length: 32 }, (_, index) => ({
      ...records[0],
      archiveCode: `5028G02-0011-${String(index + 1).padStart(3, "0")}`,
    }));
    const workbook = await readWorkbook(await renderCatalogWorkbook(template, manyRecords));
    const catalog = workbook.worksheets[0];

    expect(catalog.pageSetup.printArea).toBe("A1:G34");
    expect(catalog.getCell("A34").value).toBe(32);
    expect(catalog.getCell("A35").value).toBeNull();
  });

  it("uses page-count heading when detail rows are marked as page counts", async () => {
    const record = records.find((item) => item.archiveCode === "5028G01-0011-845-001")!;
    const workbook = await readWorkbook(await renderCatalogWorkbook(template, [record]));
    const detail = workbook.getWorksheet(record.archiveCode);

    expect(detail?.getCell("F3").value).toBe("页数");
    expect(detail?.getCell("F4").value).toBe("74");
    expect(detail?.getCell("G4").value).toBe("");
  });

  it("keeps generated sheet names within Excel limits", async () => {
    const record = {
      ...records[0],
      archiveCode: "5028G01-0011-12345678901234567890-001",
    };
    const workbook = await readWorkbook(await renderCatalogWorkbook(template, [record, record]));

    expect(workbook.worksheets[1].name.length).toBeLessThanOrEqual(31);
    expect(workbook.worksheets[2].name.length).toBeLessThanOrEqual(31);
    expect(workbook.worksheets[1].name).not.toBe(workbook.worksheets[2].name);
  });
});

describe("generateArchiveCatalog", () => {
  it("writes the catalog workbook with the fixed output file name", async () => {
    globalThis.fetch = async (input) => {
      const path = `public${String(input)}`;
      return new Response(readFileSync(path));
    };
    const paths: string[] = [];

    const result = await generateArchiveCatalog(
      records,
      {
        selectedCodes: ["5028G01-0011-842-001"],
        outputDir: "/tmp/archive-output",
      },
      async (path) => {
        paths.push(path);
      },
    );

    expect(result.name).toBe(CATALOG_OUTPUT_NAME);
    expect(paths).toEqual([`/tmp/archive-output/${CATALOG_OUTPUT_NAME}`]);
  });
});
