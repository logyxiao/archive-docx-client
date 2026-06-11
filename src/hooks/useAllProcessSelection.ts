import { useMemo, useState } from "react";
import { NO_PROCESS_TEMPLATE_KEY } from "../app/appConstants";
import {
  allProcessRowKey,
  processTemplateOptionKey,
  templateMatchFromKey,
  type AllProcessTemplateOption,
} from "../app/templateOptions";
import {
  matchingAllProcessTemplates,
  type ProcessTemplateManifest,
  type ProcessTemplateMatch,
  type ProcessTemplateSelection,
} from "../lib/processDocs";
import type { ArchiveRecord } from "../lib/types";

interface UseAllProcessSelectionOptions {
  selectedRecords: ArchiveRecord[];
  processManifest: ProcessTemplateManifest | null;
  allTemplateOptions: AllProcessTemplateOption[];
}

export function useAllProcessSelection({
  selectedRecords,
  processManifest,
  allTemplateOptions,
}: UseAllProcessSelectionOptions) {
  const [manualAllTemplateValues, setManualAllTemplateValues] = useState<Record<string, string>>({});
  const [allTemplateSearchTerms, setAllTemplateSearchTerms] = useState<Record<string, string>>({});
  const [activeAllTemplateRow, setActiveAllTemplateRow] = useState("");

  const allProcessRows = useMemo(() => {
    if (!processManifest) {
      return [];
    }

    return selectedRecords.flatMap((record) =>
      [...record.items]
        .sort((left, right) => Number(left.sequence || 0) - Number(right.sequence || 0))
        .map((item) => ({
          key: allProcessRowKey(record.archiveCode, item.fileCode, item.sequence),
          record,
          item,
          matches: matchingAllProcessTemplates(record, item, processManifest.templates),
        })),
    );
  }, [processManifest, selectedRecords]);

  const allProcessGroups = useMemo(() => {
    if (!processManifest) {
      return [];
    }

    return selectedRecords.map((record) => ({
      record,
      rows: [...record.items]
        .sort((left, right) => Number(left.sequence || 0) - Number(right.sequence || 0))
        .map((item) => ({
          key: allProcessRowKey(record.archiveCode, item.fileCode, item.sequence),
          record,
          item,
          matches: matchingAllProcessTemplates(record, item, processManifest.templates),
        })),
    }));
  }, [processManifest, selectedRecords]);

  const unresolvedAllProcessRows = useMemo(
    () =>
      allProcessRows.filter((row) => {
        const selectedKey = selectedAllTemplateKey(row.key, row.matches);
        return !selectedKey || (!templateMatchFromKey(selectedKey, allTemplateOptions) && selectedKey !== NO_PROCESS_TEMPLATE_KEY);
      }),
    [allProcessRows, allTemplateOptions, manualAllTemplateValues],
  );

  const allProcessGenerationFileCount = useMemo(
    () =>
      allProcessRows.reduce((count, row) => {
        const selectedKey = selectedAllTemplateKey(row.key, row.matches);
        return templateMatchFromKey(selectedKey, allTemplateOptions) ? count + 1 : count;
      }, 0),
    [allProcessRows, allTemplateOptions, manualAllTemplateValues],
  );

  function selectedAllProcessTemplates(): ProcessTemplateSelection[] {
    return allProcessRows.flatMap((row) => {
      const selectedKey = selectedAllTemplateKey(row.key, row.matches);
      if (selectedKey === NO_PROCESS_TEMPLATE_KEY) {
        return [];
      }

      const match = templateMatchFromKey(selectedKey, allTemplateOptions);
      if (!match) {
        return [];
      }

      return [{
        archiveCode: row.record.archiveCode,
        fileCode: row.item.fileCode,
        sequence: row.item.sequence,
        templateModule: match.templateModule,
        template: match.template,
      }];
    });
  }

  function selectedAllTemplateKey(rowKey: string, matches: ProcessTemplateMatch[] = []): string {
    if (Object.prototype.hasOwnProperty.call(manualAllTemplateValues, rowKey)) {
      return manualAllTemplateValues[rowKey] ?? "";
    }

    return matches[0] ? processTemplateOptionKey(matches[0]) : NO_PROCESS_TEMPLATE_KEY;
  }

  function selectedAllTemplateOption(rowKey: string, matches: ProcessTemplateMatch[] = []): AllProcessTemplateOption | undefined {
    const selectedKey = selectedAllTemplateKey(rowKey, matches);
    return allTemplateOptions.find((option) => option.key === selectedKey);
  }

  function filteredAllTemplateOptions(rowKey: string): AllProcessTemplateOption[] {
    const keyword = (allTemplateSearchTerms[rowKey] ?? "").trim().toLowerCase();
    if (!keyword) {
      return allTemplateOptions;
    }

    return allTemplateOptions.filter((option) => option.searchText.includes(keyword));
  }

  function selectAllTemplate(rowKey: string, option: AllProcessTemplateOption) {
    setManualAllTemplateValues((current) => ({
      ...current,
      [rowKey]: option.key,
    }));
    setAllTemplateSearchTerms((current) => ({
      ...current,
      [rowKey]: "",
    }));
    setActiveAllTemplateRow("");
  }

  return {
    allProcessRows,
    allProcessGroups,
    unresolvedAllProcessRows,
    allProcessGenerationFileCount,
    allTemplateSearchTerms,
    activeAllTemplateRow,
    setAllTemplateSearchTerms,
    setActiveAllTemplateRow,
    selectedAllProcessTemplates,
    selectedAllTemplateOption,
    filteredAllTemplateOptions,
    selectAllTemplate,
  };
}
