import ExcelJS from "exceljs";
import PizZip from "pizzip";
import type { ArchiveItem, ArchiveRecord, GeneratedFile } from "./types";

const PROCESS_TEMPLATE_ROOT = "/templates/process-docs";
const PROCESS_OUTPUT_DIR = "过程资料";
const PROJECT_NAME_PATTERNS = [
  /中核汇能高明创楷3\.58904MWp屋顶分布式光伏项目/g,
  /中核汇能高明创楷3\.58904MWp分布式光伏项目/g,
  /中核汇能高明创楷3\.58904MWP分布式光伏项目/g,
];
const SOURCE_MISSING_LABELS = [
  "项目负责人",
  "项目技术负责人",
  "分包单位",
  "分包项目负责人",
  "分包内容",
  "专业监理工程师",
  "总监理工程师",
  "项目经理",
];

interface ProcessTemplateManifest {
  templates: ProcessTemplate[];
}

interface ProcessTemplate {
  sequence: number;
  kind: "docx" | "xlsx";
  originalName: string;
  templateFile: string;
  outputExtension: ".docx" | ".xlsx";
}

interface GenerateProcessOptions {
  selectedCodes: string[];
  outputDir: string;
  userFields?: ProcessUserFields;
}

export interface ProcessGenerationResult {
  files: GeneratedFile[];
  skipped: string[];
  errors: string[];
}

export interface ProcessUserFields {
  projectManager?: string;
  projectTechnicalLeader?: string;
  generalContractorUnit?: string;
  generalContractorProjectManager?: string;
  generalContractorTechnicalLeader?: string;
  constructionUnit?: string;
  constructionProjectManager?: string;
  constructionTechnicalLeader?: string;
  subcontractorUnit?: string;
  subcontractorProjectManager?: string;
  subcontractorTechnicalLeader?: string;
  supervisionDepartment?: string;
}

export async function generateProcessDocs(
  records: ArchiveRecord[],
  options: GenerateProcessOptions,
  writeFile: (path: string, bytes: Uint8Array) => Promise<void>,
): Promise<ProcessGenerationResult> {
  const manifest = await loadProcessManifest();
  const templatesBySequence = groupTemplatesBySequence(manifest.templates);
  const selected = records.filter((record) => options.selectedCodes.includes(record.archiveCode));
  const files: GeneratedFile[] = [];
  const skipped: string[] = [];
  const errors: string[] = [];

  for (const record of selected) {
    if (!isApplicableProcessRecord(record)) {
      skipped.push(`${record.archiveCode}：不适用过程资料模板`);
      continue;
    }

    const recordOutputDir = joinPath(
      joinPath(options.outputDir, PROCESS_OUTPUT_DIR),
      sanitizeFileName(record.archiveCode + record.fullTitle),
    );

    for (const item of record.items) {
      const sequence = Number(item.sequence);
      const templates = templatesBySequence.get(sequence) ?? [];
      if (templates.length === 0) {
        skipped.push(`${record.archiveCode} 第 ${item.sequence || "?"} 条：未找到模板`);
        continue;
      }

      for (const template of templates) {
        const outputName = processOutputName(template, item);
        const outputPath = joinPath(recordOutputDir, outputName);
        try {
          const bytes = await renderProcessTemplate(template, record, item, options.userFields ?? {});
          await writeFile(outputPath, bytes);
          files.push({ name: outputName, path: outputPath });
        } catch (error) {
          errors.push(`${outputName}：${error instanceof Error ? error.message : String(error)}`);
        }
      }
    }
  }

  return { files, skipped, errors };
}

export async function loadProcessManifest(): Promise<ProcessTemplateManifest> {
  const response = await fetch(`${PROCESS_TEMPLATE_ROOT}/manifest.json`);
  if (!response.ok) {
    throw new Error("无法加载过程资料模板清单");
  }

  return response.json();
}

async function renderProcessTemplate(
  template: ProcessTemplate,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields,
): Promise<Uint8Array> {
  const bytes = await loadProcessTemplate(template.templateFile);
  return template.kind === "docx"
    ? renderProcessDocx(bytes, record, item, userFields)
    : renderProcessWorkbook(bytes, record, item, userFields);
}

async function loadProcessTemplate(templateFile: string): Promise<ArrayBuffer> {
  const response = await fetch(`${PROCESS_TEMPLATE_ROOT}/${templateFile}`);
  if (!response.ok) {
    throw new Error(`无法加载过程资料模板：${templateFile}`);
  }

  return response.arrayBuffer();
}

function isApplicableProcessRecord(record: ArchiveRecord): boolean {
  const signalText = [record.fullTitle, ...record.items.map((item) => `${item.fileCode} ${item.title}`)].join(" ");
  return /SG-ZHHC|报验|质量验收|检验批|施工记录|测量记录|检查记录/.test(signalText);
}

export function renderProcessDocx(
  template: ArrayBuffer | Uint8Array,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields = {},
): Uint8Array {
  const zip = new PizZip(template);
  for (const path of ["word/document.xml", ...Object.keys(zip.files).filter((name) => /^word\/(header|footer)\d+\.xml$/.test(name))]) {
    const file = zip.file(path);
    if (!file) {
      continue;
    }
    zip.file(path, replaceDocxParagraphs(replaceBusinessText(file.asText(), record, item, userFields), item));
  }

  return zip.generate({ type: "uint8array", compression: "DEFLATE" });
}

function replaceDocxParagraphs(xml: string, item: ArchiveItem): string {
  if (!item.fileCode || item.fileCode === "/") {
    return xml;
  }

  return xml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraph) => {
    const text = paragraphText(paragraph);
    if (!text.includes("编号：") || !text.includes("5028G01")) {
      return paragraph;
    }

    return replaceParagraphText(paragraph, text.replace(/编号：.*$/, `编号：${item.fileCode}`));
  });
}

function paragraphText(paragraph: string): string {
  return Array.from(paragraph.matchAll(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g))
    .map((match) => unescapeXml(match[1]))
    .join("");
}

function replaceParagraphText(paragraph: string, text: string): string {
  let isFirstTextNode = true;
  return paragraph.replace(/<w:t([^>]*)>[\s\S]*?<\/w:t>/g, (_match, attrs: string) => {
    if (isFirstTextNode) {
      isFirstTextNode = false;
      return `<w:t${attrs}>${escapeXml(text)}</w:t>`;
    }
    return `<w:t${attrs}></w:t>`;
  });
}

export async function renderProcessWorkbook(
  template: ArrayBuffer | Uint8Array,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields = {},
): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(toArrayBuffer(template));

  for (const sheet of workbook.worksheets) {
    applyWorkbookValues(sheet, record, item, userFields);
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return new Uint8Array(buffer);
}

function applyWorkbookValues(sheet: ExcelJS.Worksheet, record: ArchiveRecord, item: ArchiveItem, userFields: ProcessUserFields) {
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    row.eachCell({ includeEmpty: true }, (cell) => {
      if (typeof cell.value === "string") {
        cell.value = replaceBusinessText(cell.value, record, item, userFields);
      }
    });
  }

  fillAfterLabel(sheet, ["工程名称", "工程项目名称"], record.projectName);
  fillAfterLabel(sheet, ["工程编号"], archiveProjectCode(record.archiveCode));
  fillAfterLabel(sheet, ["抽样单位"], item.owner || record.filingUnit);
  fillAfterLabel(sheet, ["抽样日期"], formatChineseDate(item.fileDate));
  fillUnitScopedFields(sheet, userFields, item.owner || record.filingUnit);
  fillOrClearProfessionalForeman(sheet, userFields.constructionTechnicalLeader ?? userFields.projectTechnicalLeader ?? "");
  clearAfterLabel(
    sheet,
    SOURCE_MISSING_LABELS.filter(
      (label) => !["总承包单位", "施工单位", "分包单位", "项目负责人", "项目技术负责人", "分包项目负责人"].includes(label),
    ),
  );
}

function replaceBusinessText(value: string, record: ArchiveRecord, item: ArchiveItem, userFields: ProcessUserFields): string {
  let result = value;
  for (const pattern of PROJECT_NAME_PATTERNS) {
    result = result.replace(pattern, record.projectName);
  }

  result = result.replace(/5028G01-[A-Z0-9-]+-\d{2,}/g, (match) => {
    return item.fileCode && item.fileCode !== "/" ? item.fileCode : match;
  });
  result = result.replace(/5028G01-0011/g, archiveProjectCode(record.archiveCode));
  result = result.replace(/中核华辰建筑工程有限公司/g, userFields.constructionUnit ?? userFields.generalContractorUnit ?? item.owner ?? record.filingUnit);

  const sourceMissingReplacements: Record<string, string> = {
    河南中核五院研究设计有限公司: userFields.supervisionDepartment ?? "",
    谢智敏: userFields.constructionProjectManager ?? userFields.projectManager ?? "",
    蒋志炜: userFields.generalContractorTechnicalLeader ?? userFields.projectTechnicalLeader ?? "",
    刘彦堂: userFields.constructionTechnicalLeader ?? userFields.projectTechnicalLeader ?? "",
    河南誉华美建设工程有限公司: userFields.subcontractorUnit ?? "",
    胡彦芳: userFields.subcontractorProjectManager ?? "",
  };

  for (const [missing, replacement] of Object.entries(sourceMissingReplacements)) {
    result = result.split(missing).join(replacement);
  }

  return result;
}

function fillUnitScopedFields(sheet: ExcelJS.Worksheet, userFields: ProcessUserFields, defaultUnit: string) {
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    if (rowHasExactLabel(row, "总承包单位")) {
      fillOrClearAfterLabelInRow(row, ["总承包单位"], userFields.generalContractorUnit ?? defaultUnit);
      fillOrClearAfterLabelInRow(row, ["项目负责人"], userFields.generalContractorProjectManager ?? userFields.projectManager ?? "");
      fillOrClearAfterLabelInRow(row, ["项目技术负责人"], userFields.generalContractorTechnicalLeader ?? userFields.projectTechnicalLeader ?? "");
    }
    if (rowHasExactLabel(row, "施工单位")) {
      fillOrClearAfterLabelInRow(row, ["施工单位"], userFields.constructionUnit ?? defaultUnit);
      fillOrClearAfterLabelInRow(row, ["项目负责人"], userFields.constructionProjectManager ?? userFields.projectManager ?? "");
      fillOrClearAfterLabelInRow(row, ["项目技术负责人"], userFields.constructionTechnicalLeader ?? userFields.projectTechnicalLeader ?? "");
    }
    if (rowHasExactLabel(row, "分包单位")) {
      fillOrClearAfterLabelInRow(row, ["分包单位"], userFields.subcontractorUnit ?? "");
      fillOrClearAfterLabelInRow(row, ["分包项目负责人", "项目负责人"], userFields.subcontractorProjectManager ?? "");
      fillOrClearAfterLabelInRow(row, ["分包项目技术负责人", "项目技术负责人"], userFields.subcontractorTechnicalLeader ?? "");
    }
  }
}

function fillOrClearProfessionalForeman(sheet: ExcelJS.Worksheet, value: string) {
  const labels = ["专业工长", "施工员"];
  if (value) {
    fillAfterLabel(sheet, labels, value);
  } else {
    clearAfterLabel(sheet, labels);
  }
}

function rowHasExactLabel(row: ExcelJS.Row, label: string): boolean {
  let hasLabel = false;
  row.eachCell({ includeEmpty: false }, (cell) => {
    if (normalizeCellText(cell) === label) {
      hasLabel = true;
    }
  });
  return hasLabel;
}

function fillOrClearAfterLabelInRow(row: ExcelJS.Row, labels: string[], value: string) {
  const targets: ExcelJS.Cell[] = [];
  row.eachCell({ includeEmpty: false }, (cell) => {
    const text = normalizeCellText(cell);
    if (!labels.some((label) => text === label)) {
      return;
    }

    const target = cellToRightOfMergedRange(row.worksheet, cell);
    if (target !== cell) {
      targets.push(target);
    }
  });

  for (const target of targets) {
    target.value = value || "";
  }
}

function normalizeCellText(cell: ExcelJS.Cell): string {
  return typeof cell.value === "string" ? cell.value.replace(/\s+/g, "") : "";
}

function fillAfterLabel(sheet: ExcelJS.Worksheet, labels: string[], value: string) {
  if (!value) {
    return;
  }

  for (const cell of cellsMatchingLabels(sheet, labels)) {
    const target = cell.value && String(cell.value).includes("：") && !String(cell.value).trim().endsWith("：")
      ? cell
      : cellToRightOfMergedRange(sheet, cell);
    if (target === cell) {
      target.value = replaceAfterColon(String(cell.value), value);
    } else {
      target.value = value;
    }
  }
}

function clearAfterLabel(sheet: ExcelJS.Worksheet, labels: string[]) {
  for (const cell of cellsMatchingLabels(sheet, labels)) {
    const target = cellToRightOfMergedRange(sheet, cell);
    if (target !== cell) {
      target.value = "";
    }
  }
}

function cellsMatchingLabels(sheet: ExcelJS.Worksheet, labels: string[]): ExcelJS.Cell[] {
  const cells: ExcelJS.Cell[] = [];
  for (const row of sheet.getRows(1, sheet.rowCount) ?? []) {
    row.eachCell({ includeEmpty: false }, (cell) => {
      const value = typeof cell.value === "string" ? cell.value.replace(/\s+/g, "") : "";
      if (labels.some((label) => value.includes(label))) {
        cells.push(cell);
      }
    });
  }
  return cells;
}

function cellToRightOfMergedRange(sheet: ExcelJS.Worksheet, cell: ExcelJS.Cell): ExcelJS.Cell {
  const range = findMergeRange(sheet, cell);
  return sheet.getCell(cell.row, (range?.right ?? Number(cell.col)) + 1);
}

function findMergeRange(sheet: ExcelJS.Worksheet, cell: ExcelJS.Cell): { left: number; right: number } | null {
  const model = sheet.model as ExcelJS.WorksheetModel & { merges?: string[] };
  for (const merge of model.merges ?? []) {
    const range = decodeRange(merge);
    const row = Number(cell.row);
    const column = Number(cell.col);
    if (row >= range.top && row <= range.bottom && column >= range.left && column <= range.right) {
      return range;
    }
  }
  return null;
}

function decodeRange(range: string): { top: number; bottom: number; left: number; right: number } {
  const [start, end = start] = range.split(":");
  const startCell = decodeCellAddress(start);
  const endCell = decodeCellAddress(end);
  return {
    top: Math.min(startCell.row, endCell.row),
    bottom: Math.max(startCell.row, endCell.row),
    left: Math.min(startCell.column, endCell.column),
    right: Math.max(startCell.column, endCell.column),
  };
}

function decodeCellAddress(address: string): { row: number; column: number } {
  const match = address.match(/^([A-Z]+)(\d+)$/);
  if (!match) {
    return { row: 1, column: 1 };
  }
  const [, columnText, rowText] = match;
  let column = 0;
  for (const char of columnText) {
    column = column * 26 + char.charCodeAt(0) - 64;
  }
  return { row: Number(rowText), column };
}

function replaceAfterColon(text: string, value: string): string {
  return text.replace(/([：:]).*$/, `$1${value}`);
}

function processOutputName(template: ProcessTemplate, item: ArchiveItem): string {
  const extension = template.outputExtension;
  const originalStem = template.originalName.replace(/\.(docx|xls)$/i, "").replace(/^\d+、/, "");
  const withFileCode = item.fileCode && item.fileCode !== "/"
    ? originalStem.replace(/5028G01-[A-Z0-9-]+-\d{2,}/, item.fileCode)
    : originalStem;
  return sanitizeFileName(`${item.sequence || template.sequence}、${withFileCode}${extension}`);
}

function groupTemplatesBySequence(templates: ProcessTemplate[]): Map<number, ProcessTemplate[]> {
  const grouped = new Map<number, ProcessTemplate[]>();
  for (const template of templates) {
    const group = grouped.get(template.sequence) ?? [];
    group.push(template);
    grouped.set(template.sequence, group);
  }
  return grouped;
}

function archiveProjectCode(archiveCode: string): string {
  return archiveCode.split("-").slice(0, 2).join("-");
}

function formatChineseDate(value: string): string {
  const match = value.match(/^(\d{4})(\d{2})(\d{2})$/);
  if (!match) {
    return value;
  }
  return `${match[1]}年 ${Number(match[2])} 月 ${Number(match[3])} 日`;
}

function toArrayBuffer(input: ArrayBuffer | Uint8Array): ArrayBuffer {
  if (input instanceof ArrayBuffer) {
    return input;
  }
  return input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength);
}

function joinPath(dir: string, fileName: string): string {
  const separator = dir.includes("\\") ? "\\" : "/";
  return `${dir.replace(/[\\/]$/, "")}${separator}${fileName}`;
}

function sanitizeFileName(value: string): string {
  return value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, "").trim();
}

function escapeXml(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function unescapeXml(value: string): string {
  return value
    .replace(/&apos;/g, "'")
    .replace(/&quot;/g, "\"")
    .replace(/&gt;/g, ">")
    .replace(/&lt;/g, "<")
    .replace(/&amp;/g, "&");
}
