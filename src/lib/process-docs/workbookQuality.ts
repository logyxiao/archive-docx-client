import { qualityResultTextForStandard } from "./qualitySelfCheck";

export function replaceWorkbookQualityPlaceholders(value: string): string {
  return value.replace(
    /\{\{\s*(?:质量验收结果|质量检查验收结果|施工单位自检记录)\s*[:：]\s*([^{}]+?)\s*}}/g,
    (match, standard: string) => qualityResultTextForStandard(standard) ?? match,
  );
}
