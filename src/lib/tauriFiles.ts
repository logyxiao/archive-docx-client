import { invoke } from "@tauri-apps/api/core";
import PizZip from "pizzip";
import type { ProcessTemplate } from "./process-docs/types";

interface ImportedProcessTemplate {
  template: ProcessTemplate;
  directory: string;
}

export async function readBinaryFile(path: string): Promise<Uint8Array> {
  const contentBase64 = await invoke<string>("read_binary_file_base64", { path });
  return base64ToBytes(contentBase64);
}

export async function writeBinaryFile(path: string, bytes: Uint8Array): Promise<void> {
  validateOfficeFile(path, bytes);
  await invoke("write_binary_file_base64", {
    path,
    contentBase64: bytesToBase64(bytes),
    expectedLen: bytes.byteLength,
  });
}

export async function openSystemPath(path: string): Promise<void> {
  await invoke("open_system_path", { path });
}

export async function processTemplateUserDir(): Promise<string> {
  return invoke<string>("process_template_user_dir");
}

export async function processBuiltinTemplateDir(): Promise<string> {
  return invoke<string>("process_builtin_template_dir");
}

export async function loadUserProcessTemplates(): Promise<ProcessTemplate[]> {
  return invoke<ProcessTemplate[]>("load_user_process_templates");
}

export async function importProcessTemplate(sourcePath: string): Promise<ImportedProcessTemplate> {
  return invoke<ImportedProcessTemplate>("import_process_template", { sourcePath });
}

function validateOfficeFile(path: string, bytes: Uint8Array) {
  const lowerPath = path.toLowerCase();
  if (!lowerPath.endsWith(".docx") && !lowerPath.endsWith(".xlsx")) {
    return;
  }
  if (bytes.byteLength < 4 || bytes[0] !== 0x50 || bytes[1] !== 0x4b) {
    throw new Error(`生成文件不是有效的 Office 压缩包：${path}`);
  }

  const zip = new PizZip(bytes);
  if (lowerPath.endsWith(".docx") && !zip.file("word/document.xml")) {
    throw new Error(`生成的 Word 文件缺少 document.xml：${path}`);
  }
  if (lowerPath.endsWith(".xlsx") && !zip.file("xl/workbook.xml")) {
    throw new Error(`生成的 Excel 文件缺少 workbook.xml：${path}`);
  }
  if (lowerPath.endsWith(".xlsx")) {
    validateWorkbookStyles(zip, path);
  }
}

function validateWorkbookStyles(zip: PizZip, path: string) {
  const stylesXml = zip.file("xl/styles.xml")?.asText();
  if (!stylesXml) {
    return;
  }

  const cellXfsCount = Number(stylesXml.match(/<cellXfs\b[^>]*count="(\d+)"/)?.[1] ?? 0);
  if (!Number.isFinite(cellXfsCount) || cellXfsCount <= 0) {
    return;
  }

  let maxStyleIndex = 0;
  for (const fileName of Object.keys(zip.files)) {
    if (!/^xl\/worksheets\/sheet\d+\.xml$/.test(fileName)) {
      continue;
    }
    const sheetXml = zip.file(fileName)?.asText() ?? "";
    for (const match of sheetXml.matchAll(/<c\b[^>]*\bs="(\d+)"/g)) {
      maxStyleIndex = Math.max(maxStyleIndex, Number(match[1]));
    }
  }

  if (maxStyleIndex >= cellXfsCount) {
    throw new Error(`生成的 Excel 文件样式索引无效：${path}`);
  }
}

function bytesToBase64(bytes: Uint8Array): string {
  const chunkSize = 0x8000;
  let binary = "";
  for (let index = 0; index < bytes.byteLength; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}

function base64ToBytes(value: string): Uint8Array {
  const binary = atob(value);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  return bytes;
}
