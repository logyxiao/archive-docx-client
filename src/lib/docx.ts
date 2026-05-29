import Docxtemplater from "docxtemplater";
import PizZip from "pizzip";
import type { ArchiveRecord, GeneratedFile, GenerationOptions, GenerationResult } from "./types";

type TemplateName = "cover" | "note" | "spine";

const TEMPLATE_PATHS: Record<TemplateName, string> = {
  cover: "/templates/cover.docx",
  note: "/templates/note.docx",
  spine: "/templates/spine.docx",
};

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
    if (templates.cover) {
      await safeWrite(
        files,
        errors,
        joinPath(options.outputDir, coverFileName(record)),
        () => renderDocx(templates.cover!, coverData(record)),
        writeFile,
      );
    }

    if (templates.note) {
      await safeWrite(
        files,
        errors,
        joinPath(options.outputDir, noteFileName(record)),
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
        () => renderDocx(templates.spine!, spineData(group)),
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
    data[`档号${slot}`] = record?.archiveCode ?? "";
    data[`案卷题名${slot}`] = record?.fullTitle ?? "";
  }

  return data;
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
