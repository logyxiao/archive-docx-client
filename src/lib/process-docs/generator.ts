import type { ArchiveItem, ArchiveRecord } from "../types";
import { isProcessTemplateCategorySelected, normalizeProcessTemplateCategories, type ProcessTemplateCategoryId } from "./categories";
import { COLLECTOR_LINE_OUTPUT_DIR, PROCESS_OUTPUT_DIR, SWITCH_STATION_OUTPUT_DIR } from "./constants";
import { renderProcessDocx } from "./docxRenderer";
import { processOutputName } from "./naming";
import { renderSummaryWorkbook } from "./summaryWorkbookRenderer";
import { switchStationContextRecords } from "./switchStation";
import {
  getProcessRecordApplicability,
  isRecordInTemplateModule,
  isSummaryWorkbookTemplate,
  isTemplateInModule,
  loadProcessManifest,
  loadProcessTemplate,
  matchingTemplatesByTitle,
} from "./templates";
import type {
  GenerateProcessOptions,
  GenerateSelectedProcessOptions,
  ProcessGenerationResult,
  ProcessTemplate,
  ProcessTemplateMatch,
  ProcessTemplateModule,
  ProcessUserFields,
} from "./types";
import { joinPath, sanitizeFileName } from "./utils";
import { renderProcessWorkbook } from "./workbookRenderer";

export async function generateProcessDocs(
  records: ArchiveRecord[],
  options: GenerateProcessOptions,
  writeFile: (path: string, bytes: Uint8Array) => Promise<void>,
): Promise<ProcessGenerationResult> {
  const manifest = await loadProcessManifest();
  const templateModule = options.templateModule ?? "process";
  const contextRecords = templateModule === "switch-station" ? switchStationContextRecords(records) : records;
  const selectedTemplateCategories = normalizeProcessTemplateCategories(options.selectedTemplateCategories);
  const selected = records.filter((record) =>
    options.selectedCodes.includes(record.archiveCode) && isRecordInTemplateModule(record, templateModule),
  );
  const files: ProcessGenerationResult["files"] = [];
  const skipped: string[] = [];
  const errors: string[] = [];

  for (const record of selected) {
    const applicability = getProcessRecordApplicability(record);
    if (!applicability.isApplicable) {
      skipped.push(`${record.archiveCode}：不适用过程资料模板（未匹配过程资料关键词）`);
      continue;
    }

    const recordOutputDir = joinPath(
      joinPath(options.outputDir, outputDirForTemplateModule(templateModule)),
      sanitizeFileName(record.archiveCode + record.fullTitle),
    );

    for (const item of record.items) {
      const allTemplates = matchingTemplatesByTitle(item, manifest.templates, templateModule);
      if (allTemplates.length === 0) {
        continue;
      }

      const templates = expandSwitchStationTemplates(item, allTemplates.filter((template) =>
        isProcessTemplateCategorySelected(template, selectedTemplateCategories),
      ), templateModule);
      if (templates.length === 0) {
        continue;
      }

      for (const template of templates) {
        const outputName = processOutputName(template, item);
        const outputPath = joinPath(recordOutputDir, outputName);
        try {
          const bytes = await renderProcessTemplate(template, record, item, options.userFields ?? {}, templateModule, contextRecords);
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

export async function generateSelectedProcessDocs(
  records: ArchiveRecord[],
  options: GenerateSelectedProcessOptions,
  writeFile: (path: string, bytes: Uint8Array) => Promise<void>,
): Promise<ProcessGenerationResult> {
  const files: ProcessGenerationResult["files"] = [];
  const skipped: string[] = [];
  const errors: string[] = [];
  const switchStationContext = switchStationContextRecords(records);

  for (const selection of options.selections) {
    const record = records.find((item) => item.archiveCode === selection.archiveCode);
    if (!record) {
      skipped.push(`${selection.archiveCode}：未找到案卷`);
      continue;
    }

    const item = record.items.find((entry) => entry.fileCode === selection.fileCode && entry.sequence === selection.sequence);
    if (!item) {
      skipped.push(`${selection.archiveCode} ${selection.sequence} ${selection.fileCode}：未找到文件题名`);
      continue;
    }

    const recordOutputDir = joinPath(
      joinPath(options.outputDir, outputDirForTemplateModule(selection.templateModule)),
      sanitizeFileName(record.archiveCode + record.fullTitle),
    );
    const outputName = processOutputName(selection.template, item);
    const outputPath = joinPath(recordOutputDir, outputName);

    try {
      const contextRecords = selection.templateModule === "switch-station" ? switchStationContext : records;
      const bytes = await renderProcessTemplate(
        selection.template,
        record,
        item,
        options.userFields ?? {},
        selection.templateModule,
        contextRecords,
      );
      await writeFile(outputPath, bytes);
      files.push({ name: outputName, path: outputPath });
    } catch (error) {
      errors.push(`${outputName}：${error instanceof Error ? error.message : String(error)}`);
    }
  }

  return { files, skipped, errors };
}

export function countProcessGenerationFiles(
  records: ArchiveRecord[],
  selectedCodes: string[],
  templates: ProcessTemplate[],
  selectedTemplateCategories?: readonly ProcessTemplateCategoryId[],
  templateModule: ProcessTemplateModule = "process",
): number {
  const selectedCategories = normalizeProcessTemplateCategories(selectedTemplateCategories);
  const selected = records.filter((record) =>
    selectedCodes.includes(record.archiveCode) && isRecordInTemplateModule(record, templateModule),
  );

  return selected.reduce((total, record) => {
    const applicability = getProcessRecordApplicability(record);
    if (!applicability.isApplicable) {
      return total;
    }

    return total + record.items.reduce((itemTotal, item) => {
      const templatesForItem = matchingTemplatesByTitle(item, templates, templateModule)
        .filter((template) => isProcessTemplateCategorySelected(template, selectedCategories));
      return itemTotal + expandSwitchStationTemplates(item, templatesForItem, templateModule).length;
    }, 0);
  }, 0);
}

export function allProcessTemplateOptions(templates: ProcessTemplate[]): ProcessTemplateMatch[] {
  return dedupeTemplateMatches(
    PROCESS_TEMPLATE_MODULES.flatMap((templateModule) =>
      templates
        .filter((template) => template.enabled !== false && isTemplateInModule(template, templateModule))
        .map((template) => ({ templateModule, template })),
    ),
  );
}

export function matchingAllProcessTemplates(record: ArchiveRecord, item: ArchiveItem, templates: ProcessTemplate[]): ProcessTemplateMatch[] {
  return dedupeTemplateMatches(
    PROCESS_TEMPLATE_MODULES.filter((templateModule) => isRecordInTemplateModule(record, templateModule))
      .flatMap((templateModule) =>
        expandSwitchStationTemplates(item, matchingTemplatesByTitle(item, templates, templateModule), templateModule)
          .map((template) => ({ templateModule, template })),
      ),
  );
}

function outputDirForTemplateModule(templateModule: ProcessTemplateModule): string {
  if (templateModule === "switch-station") {
    return SWITCH_STATION_OUTPUT_DIR;
  }
  if (templateModule === "collector-line") {
    return COLLECTOR_LINE_OUTPUT_DIR;
  }
  return PROCESS_OUTPUT_DIR;
}

const PROCESS_TEMPLATE_MODULES: ProcessTemplateModule[] = ["process", "switch-station", "collector-line"];

function dedupeTemplateMatches(matches: ProcessTemplateMatch[]): ProcessTemplateMatch[] {
  const seen = new Set<string>();
  const result: ProcessTemplateMatch[] = [];

  for (const match of matches) {
    const key = [
      match.template.templateFile,
      match.template.originalName,
      match.template.userTemplatePath ?? "",
      match.template.outputFileCodeOverride ?? "",
    ].join("\u0000");
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(match);
  }

  return result;
}

function expandSwitchStationTemplates(
  item: ArchiveItem,
  templates: ProcessTemplate[],
  templateModule: string,
): ProcessTemplate[] {
  if (templateModule !== "switch-station") {
    return templates;
  }

  const expanded: ProcessTemplate[] = [];
  for (const template of templates) {
    expanded.push(template);
    if (
      template.kind === "xlsx"
      && (template.templateFile === "屋外接地装置安装分项工程质量验收表.xlsx"
        || template.templateFile === "屋内接地装置安装分项工程质量验收表.xlsx")
      && item.fileCode.endsWith("-001")
    ) {
      expanded.push({
        ...template,
        originalName: template.originalName.replace("-001", "-002"),
        outputFileCodeOverride: item.fileCode.replace(/-001$/, "-002"),
      });
    }
  }

  return expanded;
}

async function renderProcessTemplate(
  template: ProcessTemplate,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields,
  templateModule: ProcessTemplateModule = "process",
  contextRecords: ArchiveRecord[] = [record],
): Promise<Uint8Array> {
  const bytes = await loadProcessTemplate(template);
  const renderItem = template.outputFileCodeOverride ? { ...item, fileCode: template.outputFileCodeOverride } : item;
  if (template.kind === "docx") {
    return renderProcessDocx(bytes, record, renderItem, userFields, templateModule);
  }

  return isSummaryWorkbookTemplate(template)
    ? renderSummaryWorkbook(bytes, record, userFields, renderItem, template, templateModule, contextRecords)
    : renderProcessWorkbook(bytes, record, renderItem, userFields, templateModule);
}
