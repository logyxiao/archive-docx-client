import { readFileSync } from "node:fs";
import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import {
  coverData,
  coverFileName,
  formatSpineDocx,
  generateArchiveDocs,
  noteData,
  noteFileName,
  renderDocx,
  spineColumnWidthTwips,
  spineData,
  spineFileName,
} from "../lib/docx";
import { parseArchiveWorkbook } from "../lib/excel";
import { createArchiveWorkbookFixture } from "./fixtures";

const records = parseArchiveWorkbook(createArchiveWorkbookFixture());
const record = records.find((item) => item.archiveCode === "5028G01-0011-842-001")!;
const largeRecord = records.find((item) => item.archiveCode === "5028G01-0011-941-001")!;

async function docxText(bytes: Uint8Array): Promise<string> {
  const zip = await JSZip.loadAsync(bytes);
  return zip.file("word/document.xml")!.async("string");
}

describe("renderDocx", () => {
  it("renders cover without placeholder leftovers", async () => {
    const template = readFileSync("public/templates/cover.docx");
    const bytes = renderDocx(template, coverData(record));
    const xml = await docxText(bytes);

    expect(xml).toContain("5028G01");
    expect(xml).toContain("-0011");
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

  it("names cover files with archive code and title", () => {
    expect(coverFileName(record)).toBe(
      "5028G01-0011-842-001中核汇能高明创楷3.58904MWp分布式光伏项目 二次设备试验案卷大封面.docx",
    );
  });

  it("groups cover and note for the same archive code in one folder", async () => {
    globalThis.fetch = async (input) => {
      const path = `public${String(input)}`;
      return new Response(readFileSync(path));
    };
    const paths: string[] = [];

    await generateArchiveDocs(
      [record],
      {
        selectedCodes: [record.archiveCode],
        backupNote: "",
        outputDir: "/tmp/archive-output",
        generateCover: true,
        generateNote: true,
        generateSpine: false,
      },
      async (path) => {
        paths.push(path);
      },
    );

    expect(paths).toEqual([
      `/tmp/archive-output/案卷大封面和备考表/${record.archiveCode}/${coverFileName(record)}`,
      `/tmp/archive-output/案卷大封面和备考表/${record.archiveCode}/${noteFileName(record)}`,
    ]);
  });

  it("generates cover, note, and spine when all document types are selected", async () => {
    globalThis.fetch = async (input) => {
      const path = `public${String(input)}`;
      return new Response(readFileSync(path));
    };
    const paths: string[] = [];

    await generateArchiveDocs(
      [record],
      {
        selectedCodes: [record.archiveCode],
        backupNote: "",
        outputDir: "/tmp/archive-output",
        generateCover: true,
        generateNote: true,
        generateSpine: true,
      },
      async (path) => {
        paths.push(path);
      },
    );

    expect(paths).toEqual([
      `/tmp/archive-output/案卷大封面和备考表/${record.archiveCode}/${coverFileName(record)}`,
      `/tmp/archive-output/案卷大封面和备考表/${record.archiveCode}/${noteFileName(record)}`,
      `/tmp/archive-output/${spineFileName([record])}`,
    ]);
  });

  it("names spine files with archive code and title for a single record", () => {
    expect(spineFileName([record])).toBe(
      "5028G01-0011-842-001中核汇能高明创楷3.58904MWp分布式光伏项目 二次设备试验案卷脊背.docx",
    );
  });

  it("renders spine slots and leaves empty slots blank", async () => {
    const template = readFileSync("public/templates/spine.docx");
    const bytes = renderDocx(template, spineData([record]));
    const xml = await docxText(bytes);

    expect(xml).toContain("5028G01");
    expect(xml).toContain("-0011");
    expect(xml).toContain("-842");
    expect(xml).toContain("-001");
    expect(xml).toContain("二次设备试验");
    expect(xml).not.toContain("{{");
  });

  it("formats spine width by page count", async () => {
    const template = readFileSync("public/templates/spine.docx");
    const bytes = formatSpineDocx(renderDocx(template, spineData([largeRecord, record])), [largeRecord, record]);
    const xml = await docxText(bytes);

    expect(spineColumnWidthTwips(largeRecord)).toBe(2155);
    expect(spineColumnWidthTwips(record)).toBe(1077);
    expect(xml).toContain('<w:gridCol w:w="2155"/>');
    expect(xml).toContain('<w:gridCol w:w="1077"/>');
    expect(xml).toContain('<w:textDirection w:val="tbLrV"/>');
  });
});
