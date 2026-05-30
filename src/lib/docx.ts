import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import type { ArchiveRecord, GeneratedFile, GenerationOptions, GenerationResult } from "./types";

type TemplateName = "cover" | "note" | "spine";

const TEMPLATE_PATHS: Record<TemplateName, string> = {
  cover: "/templates/cover.docx",
  note: "/templates/note.docx",
  spine: "/templates/spine.docx",
};
const COVER_NOTE_GROUP_DIR = "案卷大封面和备考表";

export async function loadTemplate(name: TemplateName): Promise<ArrayBuffer> {
  const response = await fetch(TEMPLATE_PATHS[name]);
  if (!response.ok) {
    throw new Error(`无法加载模板：${TEMPLATE_PATHS[name]}`);
  }

  return response.arrayBuffer();
}

export function renderDocx(template: ArrayBuffer | Uint8Array, data: Record<string, unknown>): Uint8Array {
  const zip = new PizZip(template);
  const doc = new Docxtemplater(zip, {
    delimiters: { start: "{{", end: "}}" },
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: () => "",
  });

  doc.render(data);
  return doc.getZip().generate({ type: "uint8array", compression: "DEFLATE" });
}

export async function generateArchiveDocs(
  records: ArchiveRecord[],
  options: GenerationOptions,
  writeFile: (path: string, bytes: Uint8Array) => Promise<void>,
): Promise<GenerationResult> {
  const selected = records.filter((record) => options.selectedCodes.includes(record.archiveCode));
  const files: GeneratedFile[] = [];
  const errors: string[] = [];

  const templates = {
    cover: options.generateCover ? await loadTemplate("cover") : null,
    note: options.generateNote ? await loadTemplate("note") : null,
    spine: options.generateSpine ? await loadTemplate("spine") : null,
  };

  for (const record of selected) {
    const recordOutputDir =
      templates.cover && templates.note
        ? joinPath(joinPath(options.outputDir, COVER_NOTE_GROUP_DIR), sanitizeFileName(record.archiveCode + record.fullTitle))
        : options.outputDir;

    if (templates.cover) {
      await safeWrite(
        files,
        errors,
        joinPath(recordOutputDir, coverFileName(record)),
        () => renderDocx(templates.cover!, coverData(record)),
        writeFile,
      );
    }

    if (templates.note) {
      await safeWrite(
        files,
        errors,
        joinPath(recordOutputDir, noteFileName(record)),
        () => renderDocx(templates.note!, noteData(record, options.backupNote)),
        writeFile,
      );
    }
  }

  if (templates.spine) {
    const groups = chunk(selected, 7);
    for (const [index, group] of groups.entries()) {
      const fileName = spineFileName(group, index);
      await safeWrite(
        files,
        errors,
        joinPath(options.outputDir, fileName),
        () => formatSpineDocx(renderDocx(templates.spine!, spineData(group)), group),
        writeFile,
      );
    }
  }

  return { files, errors };
}

export function coverData(record: ArchiveRecord): Record<string, string | number> {
  return {
    档号: record.archiveCode,
    项目名称: record.projectName,
    案卷标题: record.volumeTitle,
    案卷题名: record.fullTitle,
    立卷单位: record.filingUnit,
    起止日期: record.dateRange,
    保管期限: record.retentionPeriod,
    密级: "",
  };
}

export function noteData(record: ArchiveRecord, backupNote: string): Record<string, string | number> {
  return {
    档号: record.archiveCode,
    总页数: record.totalPages,
    图样页数: record.drawingPages,
    文字材料页数: record.textPages,
    其它情况说明: backupNote,
  };
}

export function coverFileName(record: ArchiveRecord): string {
  return `${sanitizeFileName(record.archiveCode + record.fullTitle)}案卷大封面.docx`;
}

export function noteFileName(record: ArchiveRecord): string {
  return `${sanitizeFileName(record.archiveCode + record.fullTitle)}备考表.docx`;
}

export function spineFileName(records: ArchiveRecord[], index = 0): string {
  if (records.length === 1) {
    const record = records[0];
    return `${sanitizeFileName(record.archiveCode + record.fullTitle)}案卷脊背.docx`;
  }

  const first = records[0]?.archiveCode ?? String(index + 1).padStart(3, "0");
  const last = records[records.length - 1]?.archiveCode;
  const range = last && last !== first ? `${first}-${last}` : first;
  return `${sanitizeFileName(range)}案卷脊背.docx`;
}

export function spineData(records: ArchiveRecord[]): Record<string, string | number> {
  const data: Record<string, string | number> = {};

  for (let index = 0; index < 7; index += 1) {
    const slot = index + 1;
    const record = records[index];
    data[`保管期限${slot}`] = record?.retentionPeriod ?? "";
    data[`档号${slot}`] = record ? formatArchiveCodeForSpine(record.archiveCode) : "";
    data[`案卷题名${slot}`] = record?.fullTitle ?? "";
  }

  return data;
}

export function spineColumnWidthTwips(record?: ArchiveRecord): number {
  return record && record.totalPages > 200 ? 2268 : 1134;
}

export function formatSpineDocx(bytes: Uint8Array, records: ArchiveRecord[]): Uint8Array {
  const zip = new PizZip(bytes);
  const document = zip.file("word/document.xml");
  if (!document) {
    return bytes;
  }

  const widths = Array.from({ length: 7 }, (_, index) => spineColumnWidthTwips(records[index]));
  let xml = document.asText();
  xml = xml.replace(/<w:tblLayout w:type="[^"]+"\s*\/>/, '<w:tblLayout w:type="fixed"/>');

  let gridIndex = 0;
  xml = xml.replace(/<w:gridCol w:w="\d+"\s*\/>/g, () => {
    const width = widths[gridIndex] ?? 1134;
    gridIndex += 1;
    return `<w:gridCol w:w="${width}"/>`;
  });

  let cellIndex = 0;
  xml = xml.replace(/<w:tcW w:w="\d+" w:type="dxa"\s*\/>/g, () => {
    const width = widths[cellIndex % 7] ?? 1134;
    cellIndex += 1;
    return `<w:tcW w:w="${width}" w:type="dxa"/>`;
  });

  xml = xml.replace(/<w:textDirection w:val="[^"]+"\s*\/>/g, '<w:textDirection w:val="tbLrV"/>');
  xml = fixSpineRowHeights(xml);
  xml = formatSpineCodeRow(xml, records);
  xml = formatSpineTitleHeaderRow(xml, records);
  xml = formatSpineTitleContentRow(xml, records);
  zip.file("word/document.xml", xml);
  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function formatSpineCodeRow(xml: string, records: ArchiveRecord[]): string {
  let rowIndex = 0;
  return xml.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (row) => {
    rowIndex += 1;
    if (rowIndex !== 3) {
      return row;
    }

    let cellIndex = 0;
    return row.replace(/<w:tc>[\s\S]*?<\/w:tc>/g, (cell) => {
      const record = records[cellIndex];
      cellIndex += 1;
      const size = record && record.totalPages > 200 ? 32 : 24;
      return removeEmptyParagraphs(setWordFontSize(cell, size)).replace(/<w:vAlign w:val="[^"]+"\s*\/>/g, '<w:vAlign w:val="center"/>');
    });
  });
}

function fixSpineRowHeights(xml: string): string {
  const rowHeights = [2551, 1134, 2835, 1134, 6520, 1686];
  let rowIndex = 0;
  return xml.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (row) => {
    const height = rowHeights[rowIndex];
    rowIndex += 1;
    if (!height) {
      return row;
    }

    return row.replace(/<w:trHeight w:val="\d+" w:hRule="[^"]+"\s*\/>/, `<w:trHeight w:val="${height}" w:hRule="exact"/>`);
  });
}

function formatSpineTitleHeaderRow(xml: string, records: ArchiveRecord[]): string {
  let rowIndex = 0;
  return xml.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (row) => {
    rowIndex += 1;
    if (rowIndex !== 4) {
      return row;
    }

    let cellIndex = 0;
    return row.replace(/<w:tc>[\s\S]*?<\/w:tc>/g, (cell) => {
      const record = records[cellIndex];
      cellIndex += 1;
      return replaceCellParagraphs(cell, record && record.totalPages > 200 ? ["案卷题名"] : ["案卷", "题名"], 36, true);
    });
  });
}

function formatSpineTitleContentRow(xml: string, records: ArchiveRecord[]): string {
  let rowIndex = 0;
  return xml.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (row) => {
    rowIndex += 1;
    if (rowIndex !== 5) {
      return row;
    }

    let cellIndex = 0;
    return row.replace(/<w:tc>[\s\S]*?<\/w:tc>/g, (cell) => {
      const record = records[cellIndex];
      cellIndex += 1;
      const size = spineTitleFontSize(record);
      return setWordLineSpacing(setWordFontSize(cell, size), spineTitleLineSpacing(size))
        .replace(/<w:vAlign w:val="[^"]+"\s*\/>/g, '<w:vAlign w:val="center"/>')
        .replace(/<w:jc w:val="[^"]+"\s*\/>/g, '<w:jc w:val="center"/>');
    });
  });
}

function setWordFontSize(xml: string, size: number): string {
  return xml
    .replace(/<w:sz w:val="\d+"\s*\/>/g, `<w:sz w:val="${size}"/>`)
    .replace(/<w:szCs w:val="\d+"\s*\/>/g, `<w:szCs w:val="${size}"/>`);
}

function setWordLineSpacing(xml: string, spacing: number): string {
  return xml.replace(/<w:spacing w:line="\d+" w:lineRule="exact"\s*\/>/g, `<w:spacing w:line="${spacing}" w:lineRule="exact"/>`);
}

function removeEmptyParagraphs(xml: string): string {
  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    return /<w:t[^>]*>[^<]+<\/w:t>/.test(paragraph) ? paragraph : "";
  });
}

function replaceCellParagraphs(cell: string, lines: string[], size: number, bold: boolean): string {
  const tcPrEnd = cell.indexOf("</w:tcPr>");
  if (tcPrEnd < 0) {
    return cell;
  }

  const prefix = cell.slice(0, tcPrEnd + "</w:tcPr>".length);
  return `${prefix}${lines.map((line) => wordParagraph(line, size, bold)).join("")}</w:tc>`;
}

function wordParagraph(text: string, size: number, bold: boolean): string {
  const boldXml = bold ? "<w:b/>" : "";
  return `<w:p><w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/>${boldXml}<w:color w:val="auto"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:hint="eastAsia" w:ascii="宋体" w:hAnsi="宋体" w:eastAsia="宋体"/>${boldXml}<w:color w:val="auto"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr><w:t>${escapeXml(text)}</w:t></w:r></w:p>`;
}

function spineTitleFontSize(record?: ArchiveRecord): number {
  if (!record) {
    return 32;
  }

  const length = record.fullTitle.replace(/\s+/g, "").length;
  const capacity = record.totalPages > 200 ? 72 : 36;
  if (length > capacity * 1.8) {
    return 20;
  }
  if (length > capacity * 1.35) {
    return 24;
  }
  if (length > capacity) {
    return 28;
  }
  return 32;
}

function spineTitleLineSpacing(fontSize: number): number {
  if (fontSize <= 20) {
    return 280;
  }
  if (fontSize <= 24) {
    return 340;
  }
  if (fontSize <= 28) {
    return 420;
  }
  return 500;
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}


function chunk<T>(items: T[], size: number): T[][] {
  const groups: T[][] = [];
  for (let index = 0; index < items.length; index += size) {
    groups.push(items.slice(index, index + size));
  }
  return groups;
}

async function safeWrite(
  files: GeneratedFile[],
  errors: string[],
  path: string,
  render: () => Uint8Array,
  writeFile: (path: string, bytes: Uint8Array) => Promise<void>,
): Promise<void> {
  try {
    const bytes = render();
    await writeFile(path, bytes);
    files.push({ name: basename(path), path });
  } catch (error) {
    errors.push(`${basename(path)}：${error instanceof Error ? error.message : String(error)}`);
  }
}

function joinPath(dir: string, fileName: string): string {
  const separator = dir.includes("\\") ? "\\" : "/";
  return `${dir.replace(/[\\/]+$/, "")}${separator}${fileName}`;
}

function basename(path: string): string {
  return path.split(/[\\/]/).pop() || path;
}

function sanitizeFileName(value: string): string {
  return value.replace(/[<>:"/\\|?*]+/g, "_");
}

function formatArchiveCodeForSpine(value: string): string {
  return value.replace(/-/g, "\n-");
}
