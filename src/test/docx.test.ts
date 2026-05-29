import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { coverData, noteData, noteFileName, renderDocx, spineData } from "../lib/docx";
import { parseArchiveWorkbook } from "../lib/excel";

const records = parseArchiveWorkbook(readFileSync("../预立卷档案总目录（高明）(5).xlsx"));
const record = records.find((item) => item.archiveCode === "5028G01-0011-842-001")!;

async function docxText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file("word/document.xml")!.async("string");
}

describe("renderDocx", () => {
  it("renders cover without placeholder leftovers", async () => {
    const template = readFileSync("public/templates/cover.docx");
    const bytes = renderDocx(template, coverData(record));
    const xml = await docxText(bytes);

    expect(xml).toContain("5028G01-0011-842-001");
    expect(xml).toContain("中核汇能高明创楷3.58904MWp分布式光伏项目");
    expect(xml).not.toContain("{{");
  });

  it("renders note fields from archive data", async () => {
    const template = readFileSync("public/templates/note.docx");
    const bytes = renderDocx(template, noteData(record, "测试说明"));
    const xml = await docxText(bytes);

    expect(xml).toContain("测试说明");
    expect(xml).toContain("11");
    expect(xml).not.toContain("{{");
  });

  it("names note files with archive code and title", () => {
    expect(noteFileName(record)).toBe(
      "5028G01-0011-842-001中核汇能高明创楷3.58904MWp分布式光伏项目 二次设备试验备考表.docx",
    );
  });

  it("renders spine slots and leaves empty slots blank", async () => {
    const template = readFileSync("public/templates/spine.docx");
    const bytes = renderDocx(template, spineData([record]));
    const xml = await docxText(bytes);

    expect(xml).toContain("5028G01-0011-842-001");
    expect(xml).toContain("二次设备试验");
    expect(xml).not.toContain("{{");
  });
});
