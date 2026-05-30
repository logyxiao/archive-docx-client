import type { ArchiveItem, ArchiveRecord } from "../types";
import { isProcessTemplateCategorySelected, normalizeProcessTemplateCategories } from "./categories";
import { PROCESS_OUTPUT_DIR } from "./constants";
import { renderProcessDocx } from "./docxRenderer";
import { processOutputName } from "./naming";
import { renderSummaryWorkbook } from "./summaryWorkbookRenderer";
import {
  findStartReportTemplate,
  findSubunitQualityTemplate,
  getProcessRecordApplicability,
  isStartReportItemTitle,
  isSubunitQualityItemTitle,
  isSubunitQualityTemplate,
  isSummaryWorkbookTemplate,
  loadProcessManifest,
  loadProcessTemplate,
} from "./templates";
import type { GenerateProcessOptions, ProcessGenerationResult, ProcessTemplate, ProcessUserFields } from "./types";
import { joinPath, sanitizeFileName } from "./utils";
import { renderProcessWorkbook } from "./workbookRenderer";

export async function generateProcessDocs(
  records: ArchiveRecord[],
  options: GenerateProcessOptions,
  writeFile: (path: string, bytes: Uint8Array) => Promise<void>,
): Promise<ProcessGenerationResult> {
  const manifest = await loadProcessManifest();
  const startReportTemplate = findStartReportTemplate(manifest.templates);
  const subunitQualityTemplate = findSubunitQualityTemplate(manifest.templates);
  const selectedTemplateCategories = normalizeProcessTemplateCategories(options.selectedTemplateCategories);
  const selected = records.filter((record) => options.selectedCodes.includes(record.archiveCode));
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
      joinPath(options.outputDir, PROCESS_OUTPUT_DIR),
      sanitizeFileName(record.archiveCode + record.fullTitle),
    );

    for (const item of record.items) {
      const allTemplates = matchingTemplatesForItem(item, startReportTemplate, subunitQualityTemplate);
      if (allTemplates.length === 0) {
        continue;
      }

      const templates = allTemplates.filter((template) =>
        isProcessTemplateCategorySelected(template, selectedTemplateCategories),
      );
      if (templates.length === 0) {
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

function matchingTemplatesForItem(
  item: ArchiveItem,
  startReportTemplate: ProcessTemplate | undefined,
  subunitQualityTemplate: ProcessTemplate | undefined,
): ProcessTemplate[] {
  if (isStartReportItemTitle(item.title) && startReportTemplate) {
    return [startReportTemplate];
  }
  if (isSubunitQualityItemTitle(item.title) && subunitQualityTemplate) {
    return [subunitQualityTemplate];
  }
  return [];
}

async function renderProcessTemplate(
  template: ProcessTemplate,
  record: ArchiveRecord,
  item: ArchiveItem,
  userFields: ProcessUserFields,
): Promise<Uint8Array> {
  const bytes = await loadProcessTemplate(template.templateFile);
  if (template.kind === "docx") {
    return renderProcessDocx(bytes, record, item, userFields);
  }

  return isSummaryWorkbookTemplate(template)
    ? renderSummaryWorkbook(bytes, record, userFields, isSubunitQualityTemplate(template) ? item : undefined)
    : renderProcessWorkbook(bytes, record, item, userFields);
}
