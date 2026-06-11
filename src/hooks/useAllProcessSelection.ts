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
  const [manualAllTemplateValues, setManualAllTemplateValues] = useState<Record<string, string[]>>({});
  const [allTemplateSearchTerms, setAllTemplateSearchTerms] = useState<Record<string, string>>({});
  const [activeAllTemplateSlot, setActiveAllTemplateSlot] = useState("");

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
        const selectedKeys = selectedAllTemplateKeys(row.key, row.matches);
        return selectedKeys.length === 0
          || selectedKeys.some((selectedKey) => !templateMatchFromKey(selectedKey, allTemplateOptions) && selectedKey !== NO_PROCESS_TEMPLATE_KEY);
      }),
    [allProcessRows, allTemplateOptions, manualAllTemplateValues],
  );

  const allProcessGenerationFileCount = useMemo(
    () =>
      allProcessRows.reduce((count, row) => {
        const selectedKeys = selectedAllTemplateKeys(row.key, row.matches);
        return count + selectedKeys.filter((selectedKey) => templateMatchFromKey(selectedKey, allTemplateOptions)).length;
      }, 0),
    [allProcessRows, allTemplateOptions, manualAllTemplateValues],
  );

  function selectedAllProcessTemplates(): ProcessTemplateSelection[] {
    return allProcessRows.flatMap((row) => {
      const selectedKeys = selectedAllTemplateKeys(row.key, row.matches);
      return selectedKeys.flatMap((selectedKey) => {
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
    });
  }

  function selectedAllTemplateKeys(rowKey: string, matches: ProcessTemplateMatch[] = []): string[] {
    if (Object.prototype.hasOwnProperty.call(manualAllTemplateValues, rowKey)) {
      return manualAllTemplateValues[rowKey] ?? [];
    }

    return matches.length > 0 ? matches.map(processTemplateOptionKey) : [NO_PROCESS_TEMPLATE_KEY];
  }

  function selectedAllTemplateOptions(rowKey: string, matches: ProcessTemplateMatch[] = []): AllProcessTemplateOption[] {
    const selectedKeys = selectedAllTemplateKeys(rowKey, matches);
    return selectedKeys.flatMap((selectedKey) => {
      const option = allTemplateOptions.find((entry) => entry.key === selectedKey);
      return option ? [option] : [];
    });
  }

  function filteredAllTemplateOptions(slotKey: string): AllProcessTemplateOption[] {
    const keyword = (allTemplateSearchTerms[slotKey] ?? "").trim().toLowerCase();
    if (!keyword) {
      return allTemplateOptions;
    }

    return allTemplateOptions.filter((option) => option.searchText.includes(keyword));
  }

  function selectAllTemplate(rowKey: string, option: AllProcessTemplateOption, matches: ProcessTemplateMatch[] = []) {
    const nextKeys = option.key === NO_PROCESS_TEMPLATE_KEY
      ? [NO_PROCESS_TEMPLATE_KEY]
      : selectedAllTemplateKeys(rowKey, matches)
        .filter((currentKey) => currentKey !== NO_PROCESS_TEMPLATE_KEY && currentKey !== option.key)
        .concat(option.key);
    setManualAllTemplateValues((current) => ({
      ...current,
      [rowKey]: nextKeys,
    }));
    setAllTemplateSearchTerms((current) => ({
      ...current,
      [rowKey]: "",
    }));
    setActiveAllTemplateSlot("");
  }

  function removeAllTemplate(rowKey: string, optionKey: string, matches: ProcessTemplateMatch[] = []) {
    const nextKeys = selectedAllTemplateKeys(rowKey, matches).filter((currentKey) => currentKey !== optionKey);
    setManualAllTemplateValues((current) => ({
      ...current,
      [rowKey]: nextKeys.length > 0 ? nextKeys : [NO_PROCESS_TEMPLATE_KEY],
    }));
  }

  return {
    allProcessRows,
    allProcessGroups,
    unresolvedAllProcessRows,
    allProcessGenerationFileCount,
    allTemplateSearchTerms,
    activeAllTemplateSlot,
    setAllTemplateSearchTerms,
    setActiveAllTemplateSlot,
    selectedAllProcessTemplates,
    selectedAllTemplateOptions,
    filteredAllTemplateOptions,
    selectAllTemplate,
    removeAllTemplate,
  };
}
