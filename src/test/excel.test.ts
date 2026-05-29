import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { excelSerialToDate, inferDrawingPages, inferTextPages, parseArchiveWorkbook } from "../lib/excel";

const workbookBytes = readFileSync("../预立卷档案总目录（高明）(5).xlsx");

describe("parseArchiveWorkbook", () => {
  it("groups rows by archive code and derives archive fields", () => {
    const records = parseArchiveWorkbook(workbookBytes);
    const record = records.find((item) => item.archiveCode === "5028G01-0011-842-001");

    expect(records.length).toBeGreaterThan(20);
    expect(record).toBeDefined();
    expect(record?.fullTitle).toBe("中核汇能高明创楷3.58904MWp分布式光伏项目 二次设备试验");
    expect(record?.projectName).toBe("中核汇能高明创楷3.58904MWp分布式光伏项目");
    expect(record?.volumeTitle).toBe("二次设备试验");
    expect(record?.startDate).toBe("20250412");
    expect(record?.endDate).toBe("20250418");
    expect(record?.dateRange).toBe("20250412-20250418");
    expect(record?.textPages).toBe(11);
    expect(record?.drawingPages).toBe(0);
    expect(record?.totalPages).toBe(11);
    expect(record?.retentionPeriod).toBe("永久");
    expect(record?.items).toHaveLength(11);
  });
});

describe("excelSerialToDate", () => {
  it("converts Excel serial dates to compact dates", () => {
    expect(excelSerialToDate(45759)).toBe("20250412");
    expect(excelSerialToDate(45765)).toBe("20250418");
  });
});

describe("inferTextPages", () => {
  it("uses max page for sequential page numbers and ranges", () => {
    expect(
      inferTextPages([
        { sequence: "1", fileCode: "", owner: "", title: "", fileDate: "", pageNo: "1", note: "" },
        { sequence: "2", fileCode: "", owner: "", title: "", fileDate: "", pageNo: "2", note: "" },
        { sequence: "3", fileCode: "", owner: "", title: "", fileDate: "", pageNo: "3~5", note: "" },
      ]),
    ).toBe(5);
  });

  it("sums non-sequential page counts", () => {
    expect(
      inferTextPages([
        { sequence: "1", fileCode: "", owner: "", title: "", fileDate: "", pageNo: "74", note: "" },
        { sequence: "2", fileCode: "", owner: "", title: "", fileDate: "", pageNo: "26", note: "" },
      ]),
    ).toBe(100);
  });

  it("separates drawing pages by file title", () => {
    const items = [
      { sequence: "1", fileCode: "", owner: "", title: "图纸目录", fileDate: "", pageNo: "页数（5）", note: "" },
      { sequence: "2", fileCode: "", owner: "", title: "施工图纸", fileDate: "", pageNo: "页数（7）", note: "" },
      { sequence: "3", fileCode: "", owner: "", title: "验收记录", fileDate: "", pageNo: "3", note: "" },
    ];

    expect(inferDrawingPages(items)).toBe(12);
    expect(inferTextPages(items.filter((item) => !item.title.includes("图纸")))).toBe(3);
  });

  it("derives drawing pages for 941 from 页数 markers", () => {
    const records = parseArchiveWorkbook(workbookBytes);
    const record = records.find((item) => item.archiveCode === "5028G01-0011-941-001");

    expect(record?.drawingPages).toBe(49);
    expect(record?.textPages).toBe(1478);
    expect(record?.totalPages).toBe(1527);
  });
});
