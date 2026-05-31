import type { ArchiveItem, ArchiveRecord } from "../types";
import { isProcessTemplateCategorySelected, normalizeProcessTemplateCategories } from "./categories";
import { PROCESS_OUTPUT_DIR, SWITCH_STATION_OUTPUT_DIR } from "./constants";
import { renderProcessDocx } from "./docxRenderer";
import { processOutputName } from "./naming";
import { renderSummaryWorkbook } from "./summaryWorkbookRenderer";
import { switchStationContextRecords } from "./switchStation";
import {
  getProcessRecordApplicability,
  isRecordInTemplateModule,
  isSummaryWorkbookTemplate,
  loadProcessManifest,
  loadProcessTemplate,
  matchingTemplatesByTitle,
} from "./templates";
import type { GenerateProcessOptions, ProcessGenerationResult, ProcessTemplate, ProcessTemplateModule, ProcessUserFields } from "./types";
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
      joinPath(options.outputDir, templateModule === "switch-station" ? SWITCH_STATION_OUTPUT_DIR : PROCESS_OUTPUT_DIR),
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
  const bytes = await loadProcessTemplate(template.templateFile);
  const renderItem = template.outputFileCodeOverride ? { ...item, fileCode: template.outputFileCodeOverride } : item;
  if (template.kind === "docx") {
    return renderProcessDocx(bytes, record, renderItem, userFields, templateModule);
  }

  return isSummaryWorkbookTemplate(template)
    ? renderSummaryWorkbook(bytes, record, userFields, renderItem, template, templateModule, contextRecords)
    : renderProcessWorkbook(bytes, record, renderItem, userFields, templateModule);
}
