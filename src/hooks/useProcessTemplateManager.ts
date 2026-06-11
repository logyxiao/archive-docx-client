import { useEffect, useMemo, useState } from "react";
import { ask, message as showDialogMessage, open } from "@tauri-apps/plugin-dialog";
import {
  ALL_PROCESS_DOCS_TAB,
  COLLECTOR_LINE_TAB,
  PROCESS_DOCS_TAB,
  SWITCH_STATION_TAB,
} from "../app/appConstants";
import { showOperationError } from "../app/dialogs";
import {
  draftFromUserTemplate,
  joinTemplatePath,
  parentDirectory,
  parseTemplateKeywords,
  processTemplateOptionKey,
  processTemplateOptionLabel,
  processTemplateSearchText,
  templateIdentityKey,
  userTemplateKey,
  type AllProcessTemplateOption,
  type TemplateDirectoryEntry,
  type UserTemplateDraft,
} from "../app/templateOptions";
import {
  allProcessTemplateOptions,
  loadProcessManifest,
  type ProcessTemplate,
  type ProcessTemplateManifest,
} from "../lib/processDocs";
import {
  deleteUserProcessTemplate,
  importProcessTemplate,
  openSystemPath,
  processBuiltinTemplateDir,
  processTemplateUserDir,
  updateUserProcessTemplate,
} from "../lib/tauriFiles";

interface UseProcessTemplateManagerOptions {
  activeTab: string;
  isTemplateManagerOpen: boolean;
}

export function useProcessTemplateManager({ activeTab, isTemplateManagerOpen }: UseProcessTemplateManagerOptions) {
  const [processManifest, setProcessManifest] = useState<ProcessTemplateManifest | null>(null);
  const [isLoadingProcessManifest, setIsLoadingProcessManifest] = useState(false);
  const [isImportingProcessTemplate, setIsImportingProcessTemplate] = useState(false);
  const [builtInTemplateDir, setBuiltInTemplateDir] = useState("");
  const [templateDirectoryError, setTemplateDirectoryError] = useState("");
  const [templateDirectoryQuery, setTemplateDirectoryQuery] = useState("");
  const [templateDrafts, setTemplateDrafts] = useState<Record<string, UserTemplateDraft>>({});
  const [savingTemplateKeys, setSavingTemplateKeys] = useState<Record<string, boolean>>({});
  const [deletingTemplateKeys, setDeletingTemplateKeys] = useState<Record<string, boolean>>({});

  const allTemplateOptions = useMemo<AllProcessTemplateOption[]>(
    () => {
      if (!processManifest) {
        return [];
      }

      return [
        {
          key: "__none__",
          label: "无",
          searchText: "无 none",
        },
        ...allProcessTemplateOptions(processManifest.templates).map((match) => ({
          key: processTemplateOptionKey(match),
          label: processTemplateOptionLabel(match),
          searchText: processTemplateSearchText(match),
          match,
        })),
      ];
    },
    [processManifest],
  );

  const templateDirectoryEntries = useMemo<TemplateDirectoryEntry[]>(
    () => {
      if (!processManifest) {
        return [];
      }

      return processManifest.templates
        .map((template) => {
          const isUserTemplate = Boolean(template.userTemplatePath);
          const name = template.displayName?.trim() || template.templateFile;
          const path = template.userTemplatePath ?? joinTemplatePath(builtInTemplateDir, template.templateFile);
          const source = isUserTemplate ? "导入模板" : "内置模板";
          const kind = template.kind.toUpperCase();
          return {
            key: templateIdentityKey(template),
            name,
            path,
            source,
            kind,
            searchText: `${name} ${template.templateFile} ${template.originalName} ${path} ${source} ${kind}`.toLowerCase(),
          };
        })
        .sort((left, right) => `${left.source}${left.name}`.localeCompare(`${right.source}${right.name}`, "zh-Hans-CN"));
    },
    [builtInTemplateDir, processManifest],
  );

  const filteredTemplateDirectoryEntries = useMemo(
    () => {
      const keyword = templateDirectoryQuery.trim().toLowerCase();
      if (!keyword) {
        return templateDirectoryEntries;
      }

      return templateDirectoryEntries.filter((entry) => entry.searchText.includes(keyword));
    },
    [templateDirectoryEntries, templateDirectoryQuery],
  );

  const userProcessTemplates = useMemo(
    () => processManifest?.templates.filter((template) => template.userTemplatePath) ?? [],
    [processManifest],
  );

  useEffect(() => {
    const needsProcessManifest = [
      ALL_PROCESS_DOCS_TAB,
      PROCESS_DOCS_TAB,
      SWITCH_STATION_TAB,
      COLLECTOR_LINE_TAB,
    ].includes(activeTab);
    if (!needsProcessManifest || processManifest) {
      return;
    }

    let cancelled = false;
    setIsLoadingProcessManifest(true);
    loadProcessManifest()
      .then((manifest) => {
        if (!cancelled) {
          setProcessManifest(manifest);
        }
      })
      .catch((error) => {
        if (!cancelled) {
          void showOperationError(error);
        }
      })
      .finally(() => {
        if (!cancelled) {
          setIsLoadingProcessManifest(false);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [activeTab, processManifest]);

  useEffect(() => {
    if (!isTemplateManagerOpen) {
      return;
    }

    let cancelled = false;
    setTemplateDirectoryError("");
    processBuiltinTemplateDir()
      .then((templateDir) => {
        if (!cancelled) {
          setBuiltInTemplateDir(templateDir);
        }
      })
      .catch((error) => {
        if (!cancelled) {
          setTemplateDirectoryError(error instanceof Error ? error.message : String(error));
        }
      });

    return () => {
      cancelled = true;
    };
  }, [isTemplateManagerOpen]);

  useEffect(() => {
    if (!isTemplateManagerOpen) {
      return;
    }

    setTemplateDrafts((current) => {
      const next: Record<string, UserTemplateDraft> = {};
      for (const template of userProcessTemplates) {
        const key = userTemplateKey(template);
        next[key] = current[key] ?? draftFromUserTemplate(template);
      }
      return next;
    });
  }, [isTemplateManagerOpen, userProcessTemplates]);

  async function reloadProcessManifest() {
    const manifest = await loadProcessManifest();
    setProcessManifest(manifest);
    return manifest;
  }

  async function importProcessTemplates() {
    const selected = await open({
      multiple: true,
      filters: [{ name: "过程资料模板", extensions: ["docx", "xlsx"] }],
    });
    const selectedPaths = Array.isArray(selected) ? selected : typeof selected === "string" ? [selected] : [];
    if (selectedPaths.length === 0) {
      return;
    }

    setIsImportingProcessTemplate(true);
    try {
      for (const sourcePath of selectedPaths) {
        await importProcessTemplate(sourcePath);
      }
      await reloadProcessManifest();
      await showDialogMessage(`已导入 ${selectedPaths.length} 个模板。`, {
        title: "导入完成",
        kind: "info",
      });
    } catch (error) {
      await showOperationError(error);
    } finally {
      setIsImportingProcessTemplate(false);
    }
  }

  function updateTemplateDraft(template: ProcessTemplate, patch: Partial<UserTemplateDraft>) {
    const key = userTemplateKey(template);
    setTemplateDrafts((current) => ({
      ...current,
      [key]: {
        ...(current[key] ?? draftFromUserTemplate(template)),
        ...patch,
      },
    }));
  }

  async function saveUserProcessTemplate(template: ProcessTemplate) {
    if (!template.userTemplatePath) {
      return;
    }

    const key = userTemplateKey(template);
    const draft = templateDrafts[key] ?? draftFromUserTemplate(template);
    setSavingTemplateKeys((current) => ({ ...current, [key]: true }));
    try {
      await updateUserProcessTemplate({
        ...template,
        displayName: draft.displayName.trim() || undefined,
        matchKeywords: parseTemplateKeywords(draft.matchKeywordsText),
        matchMode: draft.matchMode,
        templateModule: draft.templateModule,
        category: draft.category || undefined,
        enabled: draft.enabled,
      });
      await reloadProcessManifest();
      await showDialogMessage("模板配置已保存。", { title: "保存成功", kind: "info" });
    } catch (error) {
      await showOperationError(error);
    } finally {
      setSavingTemplateKeys((current) => ({ ...current, [key]: false }));
    }
  }

  async function deleteProcessTemplate(template: ProcessTemplate) {
    if (!template.userTemplatePath) {
      return;
    }

    const name = template.displayName?.trim() || template.templateFile;
    const shouldDelete = await ask(`确定删除导入模板“${name}”吗？`, {
      title: "删除模板",
      kind: "warning",
      okLabel: "删除",
      cancelLabel: "取消",
    });
    if (!shouldDelete) {
      return;
    }

    const key = userTemplateKey(template);
    setDeletingTemplateKeys((current) => ({ ...current, [key]: true }));
    try {
      await deleteUserProcessTemplate(template.userTemplatePath);
      await reloadProcessManifest();
      setTemplateDrafts((current) => {
        const next = { ...current };
        delete next[key];
        return next;
      });
    } catch (error) {
      await showOperationError(error);
    } finally {
      setDeletingTemplateKeys((current) => ({ ...current, [key]: false }));
    }
  }

  async function openBuiltInProcessTemplateDir() {
    try {
      const templateDir = await processBuiltinTemplateDir();
      await openSystemPath(templateDir);
    } catch (error) {
      await showOperationError(error);
    }
  }

  async function openUserProcessTemplateDir() {
    try {
      const templateDir = await processTemplateUserDir();
      await openSystemPath(templateDir);
    } catch (error) {
      await showOperationError(error);
    }
  }

  async function openTemplateDirectory(entry: TemplateDirectoryEntry) {
    if (!entry.path) {
      return;
    }

    try {
      await openSystemPath(parentDirectory(entry.path));
    } catch (error) {
      await showOperationError(error);
    }
  }

  return {
    processManifest,
    allTemplateOptions,
    isLoadingProcessManifest,
    isImportingProcessTemplate,
    templateDirectoryEntries,
    filteredTemplateDirectoryEntries,
    templateDirectoryError,
    templateDirectoryQuery,
    setTemplateDirectoryQuery,
    userProcessTemplates,
    templateDrafts,
    savingTemplateKeys,
    deletingTemplateKeys,
    importProcessTemplates,
    updateTemplateDraft,
    saveUserProcessTemplate,
    deleteProcessTemplate,
    openBuiltInProcessTemplateDir,
    openUserProcessTemplateDir,
    openTemplateDirectory,
  };
}
