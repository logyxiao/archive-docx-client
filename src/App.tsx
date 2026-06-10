import { useEffect, useMemo, useState } from "react";
import { ask, message as showDialogMessage, open } from "@tauri-apps/plugin-dialog";
import { relaunch } from "@tauri-apps/plugin-process";
import { check } from "@tauri-apps/plugin-updater";
import {
  Check,
  ChevronLeft,
  ChevronRight,
  FileSpreadsheet,
  FileText,
  FolderTree,
  FolderOpen,
  Info,
  Loader2,
  RefreshCw,
  Search,
  Settings2,
  SlidersHorizontal,
  Upload,
  X,
} from "lucide-react";
import "./App.css";
import { generateArchiveCatalog } from "./lib/catalog";
import { generateArchiveDocs } from "./lib/docx";
import { parseArchiveWorkbook } from "./lib/excel";
import {
  allProcessTemplateOptions,
  generateProcessDocs,
  generateSelectedProcessDocs,
  loadProcessManifest,
  matchingAllProcessTemplates,
  PROCESS_TEMPLATE_CATEGORIES,
  PROCESS_TEMPLATE_CATEGORY_IDS,
  normalizeProcessTemplateCategories,
  type ProcessTemplateManifest,
  type ProcessTemplateMatch,
  type ProcessTemplateCategoryId,
  type ProcessTemplateModule,
  type ProcessTemplateSelection,
} from "./lib/processDocs";
import {
  importProcessTemplate,
  openSystemPath,
  processBuiltinTemplateDir,
  processTemplateUserDir,
  readBinaryFile,
  writeBinaryFile,
} from "./lib/tauriFiles";
import type { ArchiveRecord } from "./lib/types";

const DEFAULT_BACKUP_NOTE = "";
const LAST_OUTPUT_DIR_KEY = "archive-docx-client:last-output-dir";
const PROCESS_FIELDS_KEY = "archive-docx-client:process-fields";
const PROCESS_TEMPLATE_CATEGORIES_KEY = "archive-docx-client:process-template-categories";
const SWITCH_STATION_TEMPLATE_CATEGORIES_KEY = "archive-docx-client:switch-station-template-categories";
const COLLECTOR_LINE_TEMPLATE_CATEGORIES_KEY = "archive-docx-client:collector-line-template-categories";
const ARCHIVE_DOCX_TAB = "archive-docx";
const ALL_PROCESS_DOCS_TAB = "all-process-docs";
const PROCESS_DOCS_TAB = "process-docs";
const SWITCH_STATION_TAB = "switch-station-process-docs";
const COLLECTOR_LINE_TAB = "collector-line-process-docs";
const NO_PROCESS_TEMPLATE_KEY = "__none__";

interface AllProcessTemplateOption {
  key: string;
  label: string;
  searchText: string;
  match?: ProcessTemplateMatch;
}

interface SavedProcessFields {
  projectName: string;
  generalContractorUnit: string;
  generalContractorManager: string;
  generalContractorTechnicalLeader: string;
  constructionUnit: string;
  constructionManager: string;
  constructionTechnicalLeader: string;
  subcontractorUnit: string;
  subcontractorManager: string;
  subcontractorContent: string;
  supervisionDepartment: string;
}

const EMPTY_PROCESS_FIELDS: SavedProcessFields = {
  projectName: "",
  generalContractorUnit: "",
  generalContractorManager: "",
  generalContractorTechnicalLeader: "",
  constructionUnit: "",
  constructionManager: "",
  constructionTechnicalLeader: "",
  subcontractorUnit: "",
  subcontractorManager: "",
  subcontractorContent: "",
  supervisionDepartment: "",
};

function loadSavedProcessFields(): SavedProcessFields {
  try {
    const rawValue = localStorage.getItem(PROCESS_FIELDS_KEY);
    if (!rawValue) {
      return EMPTY_PROCESS_FIELDS;
    }

    const parsed = JSON.parse(rawValue) as Partial<SavedProcessFields> & { subcontractorTechnicalLeader?: string };
    return {
      ...EMPTY_PROCESS_FIELDS,
      ...Object.fromEntries(
        Object.entries(parsed).map(([key, value]) => [key, typeof value === "string" ? value : ""]),
      ),
      subcontractorContent:
        typeof parsed.subcontractorContent === "string"
          ? parsed.subcontractorContent
          : typeof parsed.subcontractorTechnicalLeader === "string"
            ? parsed.subcontractorTechnicalLeader
            : "",
    };
  } catch {
    return EMPTY_PROCESS_FIELDS;
  }
}

function loadSavedProcessTemplateCategories(storageKey: string): ProcessTemplateCategoryId[] {
  try {
    const rawValue = localStorage.getItem(storageKey);
    if (!rawValue) {
      return [...PROCESS_TEMPLATE_CATEGORY_IDS];
    }

    const parsed = JSON.parse(rawValue);
    if (!Array.isArray(parsed)) {
      return [...PROCESS_TEMPLATE_CATEGORY_IDS];
    }

    return [...normalizeProcessTemplateCategories(parsed)];
  } catch {
    return [...PROCESS_TEMPLATE_CATEGORY_IDS];
  }
}

function allProcessRowKey(archiveCode: string, fileCode: string, sequence: string): string {
  return `${archiveCode}\u0000${fileCode}\u0000${sequence}`;
}

function processTemplateOptionKey(match: ProcessTemplateMatch): string {
  return [
    match.templateModule,
    match.template.templateFile,
    match.template.originalName,
    match.template.userTemplatePath ?? "",
    match.template.outputFileCodeOverride ?? "",
  ].join("\u0000");
}

function processTemplateOptionLabel(match: ProcessTemplateMatch): string {
  const code = match.template.outputFileCodeOverride ? `（${match.template.outputFileCodeOverride}）` : "";
  const source = match.template.userTemplatePath ? "导入模板" : processModuleLabel(match.templateModule);
  return `${match.template.templateFile}${code} · ${source}`;
}

function processTemplateSearchText(match: ProcessTemplateMatch): string {
  return `${processTemplateOptionLabel(match)} ${match.template.originalName}`.toLowerCase();
}

function templateMatchFromKey(key: string | undefined, options: AllProcessTemplateOption[]): ProcessTemplateMatch | undefined {
  return options.find((option) => option.key === key)?.match;
}

function archiveShortCode(record: ArchiveRecord): string {
  const parts = record.archiveCode.split("-");
  const categoryCode = record.categoryCode || parts[parts.length - 2] || "";
  const sequence = parts[parts.length - 1] || "";
  return [categoryCode, sequence].filter(Boolean).join("-");
}

function processModuleLabel(templateModule: ProcessTemplateModule): string {
  if (templateModule === "switch-station") {
    return "开关站电气设备安装（子单位工程）";
  }
  if (templateModule === "collector-line") {
    return "集电线路安装工程";
  }
  return "过程资料";
}

function App() {
  const [activeTab, setActiveTab] = useState(ARCHIVE_DOCX_TAB);
  const savedProcessFields = useMemo(loadSavedProcessFields, []);
  const savedProcessTemplateCategories = useMemo(
    () => loadSavedProcessTemplateCategories(PROCESS_TEMPLATE_CATEGORIES_KEY),
    [],
  );
  const savedSwitchStationTemplateCategories = useMemo(
    () => loadSavedProcessTemplateCategories(SWITCH_STATION_TEMPLATE_CATEGORIES_KEY),
    [],
  );
  const savedCollectorLineTemplateCategories = useMemo(
    () => loadSavedProcessTemplateCategories(COLLECTOR_LINE_TEMPLATE_CATEGORIES_KEY),
    [],
  );
  const [excelPath, setExcelPath] = useState("");
  const [outputDir, setOutputDir] = useState(() => localStorage.getItem(LAST_OUTPUT_DIR_KEY) ?? "");
  const actualOutputDir = useMemo(() => {
    if (!outputDir) {
      return "";
    }
    if (!excelPath) {
      return outputDir;
    }
    const base = excelPath.split(/[\\/]/).pop() || "";
    const lastDot = base.lastIndexOf(".");
    const excelName = lastDot > 0 ? base.slice(0, lastDot) : base;
    if (!excelName) {
      return outputDir;
    }
    const separator = outputDir.includes("\\") ? "\\" : "/";
    return `${outputDir.replace(/[\\/]+$/, "")}${separator}${excelName}`;
  }, [outputDir, excelPath]);
  const [records, setRecords] = useState<ArchiveRecord[]>([]);
  const [selectedCodes, setSelectedCodes] = useState<string[]>([]);
  const [showOnlySelected, setShowOnlySelected] = useState(false);
  const [previewRecord, setPreviewRecord] = useState<ArchiveRecord | null>(null);
  const [query, setQuery] = useState("");
  const [backupNote, setBackupNote] = useState(DEFAULT_BACKUP_NOTE);
  const [generateCover, setGenerateCover] = useState(true);
  const [generateNote, setGenerateNote] = useState(true);
  const [generateSpine, setGenerateSpine] = useState(true);
  const [generateCatalogWorkbook, setGenerateCatalogWorkbook] = useState(true);
  const [processProjectName, setProcessProjectName] = useState(savedProcessFields.projectName);
  const [processGeneralContractorUnit, setProcessGeneralContractorUnit] = useState(savedProcessFields.generalContractorUnit);
  const [processGeneralContractorManager, setProcessGeneralContractorManager] = useState(savedProcessFields.generalContractorManager);
  const [processGeneralContractorTechnicalLeader, setProcessGeneralContractorTechnicalLeader] = useState(savedProcessFields.generalContractorTechnicalLeader);
  const [processConstructionUnit, setProcessConstructionUnit] = useState(savedProcessFields.constructionUnit);
  const [processConstructionManager, setProcessConstructionManager] = useState(savedProcessFields.constructionManager);
  const [processConstructionTechnicalLeader, setProcessConstructionTechnicalLeader] = useState(savedProcessFields.constructionTechnicalLeader);
  const [processSubcontractorUnit, setProcessSubcontractorUnit] = useState(savedProcessFields.subcontractorUnit);
  const [processSubcontractorManager, setProcessSubcontractorManager] = useState(savedProcessFields.subcontractorManager);
  const [processSubcontractorContent, setProcessSubcontractorContent] = useState(savedProcessFields.subcontractorContent);
  const [processSupervisionDepartment, setProcessSupervisionDepartment] = useState(savedProcessFields.supervisionDepartment);
  const [selectedProcessTemplateCategories, setSelectedProcessTemplateCategories] = useState<ProcessTemplateCategoryId[]>(
    savedProcessTemplateCategories,
  );
  const [selectedSwitchStationTemplateCategories, setSelectedSwitchStationTemplateCategories] = useState<ProcessTemplateCategoryId[]>(
    savedSwitchStationTemplateCategories,
  );
  const [selectedCollectorLineTemplateCategories, setSelectedCollectorLineTemplateCategories] = useState<ProcessTemplateCategoryId[]>(
    savedCollectorLineTemplateCategories,
  );
  const [processManifest, setProcessManifest] = useState<ProcessTemplateManifest | null>(null);
  const [manualAllTemplateValues, setManualAllTemplateValues] = useState<Record<string, string>>({});
  const [allTemplateSearchTerms, setAllTemplateSearchTerms] = useState<Record<string, string>>({});
  const [activeAllTemplateRow, setActiveAllTemplateRow] = useState("");
  const [isLoadingProcessManifest, setIsLoadingProcessManifest] = useState(false);
  const [isImportingProcessTemplate, setIsImportingProcessTemplate] = useState(false);
  const [isProcessInfoModalOpen, setIsProcessInfoModalOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isGeneratingProcess, setIsGeneratingProcess] = useState(false);
  const [isCheckingUpdate, setIsCheckingUpdate] = useState(false);
 
  const filteredRecords = useMemo(() => {
    let result = records;
    if (showOnlySelected) {
      result = result.filter((record) => selectedCodes.includes(record.archiveCode));
    }
    const keyword = query.trim().toLowerCase();
    if (!keyword) {
      return result;
    }
 
    return result.filter((record) =>
      [record.archiveCode, record.fullTitle, record.filingUnit, record.retentionPeriod]
        .join(" ")
        .toLowerCase()
        .includes(keyword),
    );
  }, [query, records, showOnlySelected, selectedCodes]);

  const selectedRecords = useMemo(
    () => records.filter((record) => selectedCodes.includes(record.archiveCode)),
    [records, selectedCodes],
  );

  const allTemplateOptions = useMemo<AllProcessTemplateOption[]>(
    () => {
      if (!processManifest) {
        return [];
      }

      return [
        {
          key: NO_PROCESS_TEMPLATE_KEY,
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

  const shouldGenerateDocx = generateCover || generateNote || generateSpine;
  const canGenerate = selectedCodes.length > 0 && actualOutputDir && (shouldGenerateDocx || generateCatalogWorkbook);
  const canGenerateProcess =
    selectedCodes.length > 0 && Boolean(actualOutputDir) && selectedProcessTemplateCategories.length > 0 && !isGeneratingProcess;
  const canGenerateSwitchStation =
    selectedCodes.length > 0 && Boolean(actualOutputDir) && selectedSwitchStationTemplateCategories.length > 0 && !isGeneratingProcess;
  const canGenerateCollectorLine =
    selectedCodes.length > 0 && Boolean(actualOutputDir) && selectedCollectorLineTemplateCategories.length > 0 && !isGeneratingProcess;
  const canGenerateAllProcess =
    selectedCodes.length > 0
    && Boolean(actualOutputDir)
    && Boolean(processManifest)
    && allProcessRows.length > 0
    && unresolvedAllProcessRows.length === 0
    && !isGeneratingProcess;
  const previewIndex = previewRecord
    ? filteredRecords.findIndex((record) => record.archiveCode === previewRecord.archiveCode)
    : -1;

  useEffect(() => {
    const fields: SavedProcessFields = {
      projectName: processProjectName,
      generalContractorUnit: processGeneralContractorUnit,
      generalContractorManager: processGeneralContractorManager,
      generalContractorTechnicalLeader: processGeneralContractorTechnicalLeader,
      constructionUnit: processConstructionUnit,
      constructionManager: processConstructionManager,
      constructionTechnicalLeader: processConstructionTechnicalLeader,
      subcontractorUnit: processSubcontractorUnit,
      subcontractorManager: processSubcontractorManager,
      subcontractorContent: processSubcontractorContent,
      supervisionDepartment: processSupervisionDepartment,
    };
    localStorage.setItem(PROCESS_FIELDS_KEY, JSON.stringify(fields));
  }, [
    processProjectName,
    processGeneralContractorUnit,
    processGeneralContractorManager,
    processGeneralContractorTechnicalLeader,
    processConstructionUnit,
    processConstructionManager,
    processConstructionTechnicalLeader,
    processSubcontractorUnit,
    processSubcontractorManager,
    processSubcontractorContent,
    processSupervisionDepartment,
  ]);

  useEffect(() => {
    localStorage.setItem(PROCESS_TEMPLATE_CATEGORIES_KEY, JSON.stringify(selectedProcessTemplateCategories));
  }, [selectedProcessTemplateCategories]);

  useEffect(() => {
    localStorage.setItem(SWITCH_STATION_TEMPLATE_CATEGORIES_KEY, JSON.stringify(selectedSwitchStationTemplateCategories));
  }, [selectedSwitchStationTemplateCategories]);

  useEffect(() => {
    localStorage.setItem(COLLECTOR_LINE_TEMPLATE_CATEGORIES_KEY, JSON.stringify(selectedCollectorLineTemplateCategories));
  }, [selectedCollectorLineTemplateCategories]);

  useEffect(() => {
    if (activeTab !== ALL_PROCESS_DOCS_TAB || processManifest) {
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

  async function chooseExcel() {
    const selected = await open({
      multiple: false,
      filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }],
    });

    if (typeof selected !== "string") {
      return;
    }

    setIsLoading(true);

    try {
      const bytes = await readBinaryFile(selected);
      const parsed = parseArchiveWorkbook(bytes);
      setExcelPath(selected);
      setRecords(parsed);
      setSelectedCodes([]);
    } catch (error) {
      await showOperationError(error);
    } finally {
      setIsLoading(false);
    }
  }

  async function chooseOutputDir() {
    const selected = await open({ directory: true, multiple: false });
    if (typeof selected === "string") {
      setOutputDir(selected);
      localStorage.setItem(LAST_OUTPUT_DIR_KEY, selected);
    }
  }

  async function openOutputDir() {
    const targetDir = actualOutputDir || outputDir;
    if (!targetDir) {
      return;
    }

    try {
      await openSystemPath(targetDir);
    } catch (error) {
      if (targetDir !== outputDir) {
        try {
          await openSystemPath(outputDir);
          return;
        } catch (innerError) {
          await showOperationError(innerError);
          return;
        }
      }
      await showOperationError(error);
    }
  }

  async function checkForUpdates() {
    setIsCheckingUpdate(true);

    try {
      const update = await check();
      if (!update) {
        await showDialogMessage("当前已经是最新版本。", {
          title: "检查更新",
          kind: "info",
        });
        return;
      }

      const shouldInstall = await ask(
        `发现新版本 ${update.version}。\n\n是否立即下载并安装？安装完成后应用会自动重启。`,
        {
          title: "发现更新",
          kind: "info",
          okLabel: "立即更新",
          cancelLabel: "稍后",
        },
      );

      if (!shouldInstall) {
        return;
      }

      await update.downloadAndInstall();
      await relaunch();
    } catch (error) {
      await showOperationError(error);
    } finally {
      setIsCheckingUpdate(false);
    }
  }

  function toggleCode(code: string) {
    setSelectedCodes((current) =>
      current.includes(code) ? current.filter((item) => item !== code) : [...current, code],
    );
  }

  function toggleVisibleRecords(checked: boolean) {
    const visibleCodes = new Set(filteredRecords.map((record) => record.archiveCode));
    setSelectedCodes((current) => {
      if (checked) {
        return Array.from(new Set([...current, ...visibleCodes]));
      }

      return current.filter((code) => !visibleCodes.has(code));
    });
  }

  function switchPreview(offset: number) {
    if (previewIndex < 0) {
      return;
    }

    const nextRecord = filteredRecords[previewIndex + offset];
    if (nextRecord) {
      setPreviewRecord(nextRecord);
    }
  }

  function toggleProcessTemplateCategory(categoryId: ProcessTemplateCategoryId) {
    setSelectedProcessTemplateCategories((current) =>
      current.includes(categoryId) ? current.filter((id) => id !== categoryId) : [...current, categoryId],
    );
  }

  function toggleSwitchStationTemplateCategory(categoryId: ProcessTemplateCategoryId) {
    setSelectedSwitchStationTemplateCategories((current) =>
      current.includes(categoryId) ? current.filter((id) => id !== categoryId) : [...current, categoryId],
    );
  }

  function toggleCollectorLineTemplateCategory(categoryId: ProcessTemplateCategoryId) {
    setSelectedCollectorLineTemplateCategories((current) =>
      current.includes(categoryId) ? current.filter((id) => id !== categoryId) : [...current, categoryId],
    );
  }

  async function generate() {
    if (!canGenerate) {
      return;
    }

    setIsGenerating(true);

    try {
      let generatedFileCount = 0;
      const errors: string[] = [];

      if (shouldGenerateDocx) {
        const generated = await generateArchiveDocs(
          records,
          {
            selectedCodes,
            backupNote,
            outputDir: actualOutputDir,
            generateCover,
            generateNote,
            generateSpine,
          },
          writeBinaryFile,
        );
        generatedFileCount += generated.files.length;
        errors.push(...generated.errors);
      }

      if (generateCatalogWorkbook) {
        await generateArchiveCatalog(
          records,
          {
            selectedCodes,
            outputDir: actualOutputDir,
          },
          writeBinaryFile,
        );
        generatedFileCount += 1;
      }

      if (errors.length > 0) {
        await showDialogMessage(
          `生成完成：${generatedFileCount} 个文件，${errors.length} 个失败。\n\n${errors.join("\n")}`,
          { title: "生成完成", kind: "warning" },
        );
      } else {
        await showDialogMessage(`生成成功，共生成 ${generatedFileCount} 个文件。`, {
          title: "生成成功",
          kind: "info",
        });
      }
    } catch (error) {
      await showOperationError(error);
    } finally {
      setIsGenerating(false);
    }
  }

  async function generateProcess(templateModule: ProcessTemplateModule = "process") {
    const isSwitchStation = templateModule === "switch-station";
    const isCollectorLine = templateModule === "collector-line";
    if (isSwitchStation ? !canGenerateSwitchStation : isCollectorLine ? !canGenerateCollectorLine : !canGenerateProcess) {
      return;
    }

    setIsGeneratingProcess(true);

    try {
      const result = await generateProcessDocs(
        records,
        {
          selectedCodes,
          outputDir: actualOutputDir,
          selectedTemplateCategories: isSwitchStation
            ? selectedSwitchStationTemplateCategories
            : isCollectorLine
              ? selectedCollectorLineTemplateCategories
              : selectedProcessTemplateCategories,
          templateModule,
          userFields: processUserFields(),
        },
        writeBinaryFile,
      );
      const skippedText = result.skipped.length > 0 ? `\n跳过 ${result.skipped.length} 条：\n${result.skipped.slice(0, 12).join("\n")}` : "";
      const errorText = result.errors.length > 0 ? `\n失败 ${result.errors.length} 个：\n${result.errors.slice(0, 12).join("\n")}` : "";
      await showDialogMessage(`${processModuleLabel(templateModule)}生成完成：${result.files.length} 个文件。${skippedText}${errorText}`, {
        title: result.errors.length > 0 ? "生成完成" : "生成成功",
        kind: result.errors.length > 0 ? "warning" : "info",
      });
    } catch (error) {
      await showOperationError(error);
    } finally {
      setIsGeneratingProcess(false);
    }
  }

  async function generateAllProcess() {
    if (!canGenerateAllProcess) {
      return;
    }

    setIsGeneratingProcess(true);

    try {
      const selections = selectedAllProcessTemplates();
      const result = await generateSelectedProcessDocs(
        records,
        {
          outputDir: actualOutputDir,
          selections,
          userFields: processUserFields(),
        },
        writeBinaryFile,
      );
      const skippedText = result.skipped.length > 0 ? `\n跳过 ${result.skipped.length} 条：\n${result.skipped.slice(0, 12).join("\n")}` : "";
      const errorText = result.errors.length > 0 ? `\n失败 ${result.errors.length} 个：\n${result.errors.slice(0, 12).join("\n")}` : "";
      await showDialogMessage(
        `全部过程资料生成完成：${result.files.length} 个文件。${skippedText}${errorText}`,
        {
          title: result.errors.length > 0 ? "生成完成" : "生成成功",
          kind: result.errors.length > 0 ? "warning" : "info",
        },
      );
    } catch (error) {
      await showOperationError(error);
    } finally {
      setIsGeneratingProcess(false);
    }
  }

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
      const manifest = await loadProcessManifest();
      setProcessManifest(manifest);
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

  function processUserFields() {
    return {
      projectName: processProjectName.trim(),
      generalContractorUnit: processGeneralContractorUnit.trim(),
      generalContractorProjectManager: processGeneralContractorManager.trim(),
      generalContractorTechnicalLeader: processGeneralContractorTechnicalLeader.trim(),
      constructionUnit: processConstructionUnit.trim(),
      constructionProjectManager: processConstructionManager.trim(),
      constructionTechnicalLeader: processConstructionTechnicalLeader.trim(),
      subcontractorUnit: processSubcontractorUnit.trim(),
      subcontractorProjectManager: processSubcontractorManager.trim(),
      subcontractorContent: processSubcontractorContent.trim(),
      supervisionDepartment: processSupervisionDepartment.trim(),
    };
  }

  const controlBand = (
    <section className="control-band">
      <button className="primary-button" onClick={chooseExcel} disabled={isLoading}>
        {isLoading ? <Loader2 className="spin" size={18} /> : <FileSpreadsheet size={18} />}
        选择 Excel
      </button>
      <button className="secondary-button" onClick={chooseOutputDir}>
        <FolderOpen size={18} />
        输出目录
      </button>
      <div className="path-stack">
        <PathLine label="Excel" value={excelPath || "未选择"} />
        <PathLine label="输出" value={actualOutputDir || outputDir || "未选择"} />
      </div>
    </section>
  );

  const recordPane = (
    <aside className="record-pane">
      <div className="pane-title">
        <div>
          <h2>案卷列表</h2>
          <p>
            {filteredRecords.length} 条匹配记录
            {selectedCodes.length > 0 ? (
              <span className="selected-count-text">（已选择 {selectedCodes.length} 条记录）</span>
            ) : null}
          </p>
        </div>
        <label className="select-all">
          <input
            type="checkbox"
            checked={filteredRecords.length > 0 && filteredRecords.every((record) => selectedCodes.includes(record.archiveCode))}
            onChange={(event) => toggleVisibleRecords(event.currentTarget.checked)}
          />
          全选
        </label>
      </div>
      <div className="record-pane-filter">
        <label className="search-box">
          <Search size={17} />
          <input value={query} onChange={(event) => setQuery(event.currentTarget.value)} placeholder="搜索档号、题名、单位" />
        </label>
        {selectedCodes.length > 0 ? (
          <label className="only-selected-toggle">
            <input
              type="checkbox"
              checked={showOnlySelected}
              onChange={(event) => setShowOnlySelected(event.currentTarget.checked)}
            />
            仅显示已选 ({selectedCodes.length})
          </label>
        ) : null}
      </div>
      <div className="record-list">
        {filteredRecords.map((record) => (
          <div
            key={record.archiveCode}
            className={`record-row ${selectedCodes.includes(record.archiveCode) ? "selected" : ""}`}
          >
            <button className="record-toggle" onClick={() => toggleCode(record.archiveCode)}>
              <span className="checkmark">{selectedCodes.includes(record.archiveCode) ? <Check size={14} /> : null}</span>
              <span>
                <strong>{record.archiveCode}</strong>
                <small>{record.fullTitle}</small>
              </span>
            </button>
            <button className="icon-button" onClick={() => setPreviewRecord(record)} title="案卷详情" aria-label={`${record.archiveCode} 案卷详情`}>
              <Info size={17} />
            </button>
          </div>
        ))}
        {records.length === 0 ? <div className="empty-state">请选择 Excel 总目录</div> : null}
      </div>
    </aside>
  );

  return (
    <main className="app-shell">
      <nav className="app-tabs" aria-label="功能栏目">
        <button
          className={`tab-button ${activeTab === ARCHIVE_DOCX_TAB ? "active" : ""}`}
          onClick={() => setActiveTab(ARCHIVE_DOCX_TAB)}
          type="button"
        >
          <FileText size={17} />
          档案文档生成器
        </button>
        <button
          className={`tab-button ${activeTab === PROCESS_DOCS_TAB ? "active" : ""}`}
          onClick={() => setActiveTab(PROCESS_DOCS_TAB)}
          type="button"
        >
          <FolderTree size={17} />
          过程资料生成
        </button>
        <button
          className={`tab-button ${activeTab === ALL_PROCESS_DOCS_TAB ? "active" : ""}`}
          onClick={() => setActiveTab(ALL_PROCESS_DOCS_TAB)}
          type="button"
        >
          <FolderTree size={17} />
          全部
        </button>
        <button
          className={`tab-button ${activeTab === SWITCH_STATION_TAB ? "active" : ""}`}
          onClick={() => setActiveTab(SWITCH_STATION_TAB)}
          type="button"
        >
          <FolderTree size={17} />
          开关站电气设备安装（子单位工程）
        </button>
        <button
          className={`tab-button ${activeTab === COLLECTOR_LINE_TAB ? "active" : ""}`}
          onClick={() => setActiveTab(COLLECTOR_LINE_TAB)}
          type="button"
        >
          <FolderTree size={17} />
          集电线路安装工程
        </button>
        <button className="secondary-button update-button tab-update-button" onClick={checkForUpdates} disabled={isCheckingUpdate} title="检查更新">
          {isCheckingUpdate ? <Loader2 className="spin" size={15} /> : <RefreshCw size={15} />}
          更新
        </button>
      </nav>

      {activeTab === ARCHIVE_DOCX_TAB ? (
        <section className="tab-panel">
          {controlBand}

          <section className="workspace">
            {recordPane}

            <section className="detail-pane">
              <div className="tool-section">
                <div className="section-heading">
                  <Settings2 size={19} />
                  <h2>生成设置</h2>
                </div>
                <div className="switch-row">
                  <Toggle checked={generateCover} onChange={setGenerateCover} label="案卷大封面" />
                  <Toggle checked={generateNote} onChange={setGenerateNote} label="备考表" />
                  <Toggle checked={generateSpine} onChange={setGenerateSpine} label="案卷脊背" />
                  <Toggle checked={generateCatalogWorkbook} onChange={setGenerateCatalogWorkbook} label="目录台账" />
                </div>
                <label className="note-field">
                  <span>备考表其它情况</span>
                  <textarea
                    value={backupNote}
                    onChange={(event) => setBackupNote(event.currentTarget.value)}
                    placeholder="不填写则生成空白说明"
                  />
                </label>
                <div className="generation-actions">
                  <button className="generate-button" onClick={generate} disabled={!canGenerate || isGenerating}>
                    {isGenerating ? <Loader2 className="spin" size={18} /> : <FileText size={18} />}
                    生成文件 {selectedCodes.length > 0 ? `(${selectedCodes.length}条)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </section>
          </section>
        </section>
      ) : null}

      {activeTab === PROCESS_DOCS_TAB ? (
        <section className="tab-panel">
          {controlBand}

          <section className="workspace">
            {recordPane}

            <section className="detail-pane">
              <div className="tool-section">
                <div className="section-heading">
                  <FolderTree size={19} />
                  <h2>过程资料生成</h2>
                </div>
                <div className="process-info-actions">
                  <button className="secondary-button process-info-button" type="button" onClick={() => setIsProcessInfoModalOpen(true)}>
                    <SlidersHorizontal size={17} />
                    填写生成信息
                  </button>
                </div>
                <div className="process-template-section">
                  <div className="template-section-heading">
                    <span>生成模板</span>
                    <div className="mini-actions">
                      <button type="button" onClick={() => setSelectedProcessTemplateCategories([...PROCESS_TEMPLATE_CATEGORY_IDS])}>
                        全选
                      </button>
                      <button type="button" onClick={() => setSelectedProcessTemplateCategories([])}>
                        清空
                      </button>
                    </div>
                  </div>
                  <div className="process-template-grid">
                    {PROCESS_TEMPLATE_CATEGORIES.map((category) => (
                      <label key={category.id} className="template-check">
                        <input
                          type="checkbox"
                          checked={selectedProcessTemplateCategories.includes(category.id)}
                          onChange={() => toggleProcessTemplateCategory(category.id)}
                        />
                        <span>{category.label}</span>
                      </label>
                    ))}
                  </div>
                </div>
                <div className="generation-actions">
                  <button className="generate-button" onClick={() => generateProcess("process")} disabled={!canGenerateProcess}>
                    {isGeneratingProcess ? <Loader2 className="spin" size={18} /> : <FolderTree size={18} />}
                    生成过程资料 {selectedCodes.length > 0 ? `(${selectedCodes.length}条)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </section>
          </section>
        </section>
      ) : null}

      {activeTab === ALL_PROCESS_DOCS_TAB ? (
        <section className="tab-panel">
          {controlBand}

          <section className="workspace">
            {recordPane}

            <section className="detail-pane">
              <div className="tool-section all-process-tool">
                <div className="section-heading">
                  <FolderTree size={19} />
                  <h2>全部过程资料</h2>
                </div>
                <div className="process-info-actions">
                  <button className="secondary-button process-info-button" type="button" onClick={() => setIsProcessInfoModalOpen(true)}>
                    <SlidersHorizontal size={17} />
                    填写生成信息
                  </button>
                </div>
                <div className="process-template-section">
                  <div className="template-section-heading">
                    <span>文件题名与模板匹配</span>
                  </div>
                  <div className="all-template-summary">
                    <span>
                      已选 {selectedCodes.length} 个档号，{allProcessRows.length} 条文件题名
                      {unresolvedAllProcessRows.length > 0 ? `，${unresolvedAllProcessRows.length} 条需选择模板` : ""}
                    </span>
                    <div className="template-summary-actions">
                      <button
                        type="button"
                        className="template-dir-button"
                        onClick={() => void importProcessTemplates()}
                        disabled={isImportingProcessTemplate}
                      >
                        {isImportingProcessTemplate ? <Loader2 className="spin" size={15} /> : <Upload size={15} />}
                        导入模板
                      </button>
                      <button type="button" className="template-dir-button" onClick={() => void openBuiltInProcessTemplateDir()}>
                        <FolderOpen size={15} />
                        打开内置目录
                      </button>
                      <button type="button" className="template-dir-button" onClick={() => void openUserProcessTemplateDir()}>
                        <FolderOpen size={15} />
                        打开导入目录
                      </button>
                    </div>
                  </div>
                </div>
                <div className="all-process-list">
                  {isLoadingProcessManifest ? (
                    <div className="empty-state detail-empty">
                      <Loader2 className="spin" size={18} />
                      正在加载模板
                    </div>
                  ) : selectedCodes.length === 0 ? (
                    <div className="empty-state detail-empty">请选择档号</div>
                  ) : allProcessRows.length === 0 ? (
                    <div className="empty-state detail-empty">未读取到文件题名</div>
                  ) : (
                    allProcessGroups.map((group) => (
                      <section key={group.record.archiveCode} className="all-process-record-group">
                        <div className="all-process-record-heading">
                          <strong>{archiveShortCode(group.record)}</strong>
                          <span>{group.rows.length} 个文件</span>
                        </div>
                        <div className="all-process-record-files">
                          {group.rows.map((row) => {
                            const selectedOption = selectedAllTemplateOption(row.key, row.matches);
                            const isSkipped = selectedOption?.key === NO_PROCESS_TEMPLATE_KEY;
                            const searchTerm = allTemplateSearchTerms[row.key] ?? "";
                            const inputValue = activeAllTemplateRow === row.key ? searchTerm : selectedOption?.label ?? "";
                            const options = activeAllTemplateRow === row.key ? filteredAllTemplateOptions(row.key) : [];
                            return (
                              <article key={row.key} className="all-process-row">
                                <div className="all-process-row-main">
                                  <div className="all-process-sequence">{row.item.sequence || "-"}</div>
                                  <div className="all-process-title">
                                    <strong>{row.item.title || "未命名文件"}</strong>
                                    <span>{row.item.fileCode || "/"}</span>
                                  </div>
                                </div>
                                <div className={`template-combobox ${selectedOption ? "selected" : ""} ${isSkipped ? "skipped" : ""}`}>
                                  <label className="template-search-field">
                                    <Search size={15} />
                                    <input
                                      value={inputValue}
                                      onFocus={() => {
                                        setActiveAllTemplateRow(row.key);
                                        setAllTemplateSearchTerms((current) => ({
                                          ...current,
                                          [row.key]: selectedOption?.label ?? "",
                                        }));
                                      }}
                                      onChange={(event) => {
                                        const value = event.currentTarget.value;
                                        setActiveAllTemplateRow(row.key);
                                        setAllTemplateSearchTerms((current) => ({
                                          ...current,
                                          [row.key]: value,
                                        }));
                                      }}
                                      onBlur={() => {
                                        window.setTimeout(() => {
                                          setActiveAllTemplateRow((current) => current === row.key ? "" : current);
                                          setAllTemplateSearchTerms((current) => {
                                            const next = { ...current };
                                            delete next[row.key];
                                            return next;
                                          });
                                        }, 120);
                                      }}
                                      placeholder="待选择"
                                    />
                                  </label>
                                  {options.length > 0 ? (
                                    <div className="template-option-list">
                                      {options.map((option) => (
                                        <button
                                          key={option.key}
                                          type="button"
                                          className={`template-option ${option.key === selectedOption?.key ? "active" : ""}`}
                                          onMouseDown={(event) => {
                                            event.preventDefault();
                                            selectAllTemplate(row.key, option);
                                          }}
                                        >
                                          {option.label}
                                        </button>
                                      ))}
                                    </div>
                                  ) : activeAllTemplateRow === row.key ? (
                                    <div className="template-option-list">
                                      <div className="template-option-empty">无匹配模板</div>
                                    </div>
                                  ) : null}
                                </div>
                              </article>
                            );
                          })}
                        </div>
                      </section>
                    ))
                  )}
                </div>
                <div className="generation-actions">
                  <button className="generate-button" onClick={generateAllProcess} disabled={!canGenerateAllProcess}>
                    {isGeneratingProcess ? <Loader2 className="spin" size={18} /> : <FolderTree size={18} />}
                    生成 {selectedCodes.length > 0 ? `(${selectedCodes.length}条)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </section>
          </section>
        </section>
      ) : null}

      {activeTab === SWITCH_STATION_TAB ? (
        <section className="tab-panel">
          {controlBand}

          <section className="workspace">
            {recordPane}

            <section className="detail-pane">
              <div className="tool-section">
                <div className="section-heading">
                  <FolderTree size={19} />
                  <h2>开关站电气设备安装（子单位工程）</h2>
                </div>
                <div className="process-info-actions">
                  <button className="secondary-button process-info-button" type="button" onClick={() => setIsProcessInfoModalOpen(true)}>
                    <SlidersHorizontal size={17} />
                    填写生成信息
                  </button>
                </div>
                <div className="process-template-section">
                  <div className="template-section-heading">
                    <span>生成模板</span>
                    <div className="mini-actions">
                      <button type="button" onClick={() => setSelectedSwitchStationTemplateCategories([...PROCESS_TEMPLATE_CATEGORY_IDS])}>
                        全选
                      </button>
                      <button type="button" onClick={() => setSelectedSwitchStationTemplateCategories([])}>
                        清空
                      </button>
                    </div>
                  </div>
                  <div className="process-template-grid">
                    {PROCESS_TEMPLATE_CATEGORIES.map((category) => (
                      <label key={category.id} className="template-check">
                        <input
                          type="checkbox"
                          checked={selectedSwitchStationTemplateCategories.includes(category.id)}
                          onChange={() => toggleSwitchStationTemplateCategory(category.id)}
                        />
                        <span>{category.label}</span>
                      </label>
                    ))}
                  </div>
                </div>
                <div className="generation-actions">
                  <button className="generate-button" onClick={() => generateProcess("switch-station")} disabled={!canGenerateSwitchStation}>
                    {isGeneratingProcess ? <Loader2 className="spin" size={18} /> : <FolderTree size={18} />}
                    生成开关站资料 {selectedCodes.length > 0 ? `(${selectedCodes.length}条)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </section>
          </section>
        </section>
      ) : null}

      {activeTab === COLLECTOR_LINE_TAB ? (
        <section className="tab-panel">
          {controlBand}

          <section className="workspace">
            {recordPane}

            <section className="detail-pane">
              <div className="tool-section">
                <div className="section-heading">
                  <FolderTree size={19} />
                  <h2>集电线路安装工程</h2>
                </div>
                <div className="process-info-actions">
                  <button className="secondary-button process-info-button" type="button" onClick={() => setIsProcessInfoModalOpen(true)}>
                    <SlidersHorizontal size={17} />
                    填写生成信息
                  </button>
                </div>
                <div className="process-template-section">
                  <div className="template-section-heading">
                    <span>生成模板</span>
                    <div className="mini-actions">
                      <button type="button" onClick={() => setSelectedCollectorLineTemplateCategories([...PROCESS_TEMPLATE_CATEGORY_IDS])}>
                        全选
                      </button>
                      <button type="button" onClick={() => setSelectedCollectorLineTemplateCategories([])}>
                        清空
                      </button>
                    </div>
                  </div>
                  <div className="process-template-grid">
                    {PROCESS_TEMPLATE_CATEGORIES.map((category) => (
                      <label key={category.id} className="template-check">
                        <input
                          type="checkbox"
                          checked={selectedCollectorLineTemplateCategories.includes(category.id)}
                          onChange={() => toggleCollectorLineTemplateCategory(category.id)}
                        />
                        <span>{category.label}</span>
                      </label>
                    ))}
                  </div>
                </div>
                <div className="generation-actions">
                  <button className="generate-button" onClick={() => generateProcess("collector-line")} disabled={!canGenerateCollectorLine}>
                    {isGeneratingProcess ? <Loader2 className="spin" size={18} /> : <FolderTree size={18} />}
                    生成集电线路资料 {selectedCodes.length > 0 ? `(${selectedCodes.length}条)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </section>
          </section>
        </section>
      ) : null}

      {isProcessInfoModalOpen ? (
        <div className="modal-backdrop" onClick={() => setIsProcessInfoModalOpen(false)}>
          <section
            className="process-info-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="process-info-title"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <div className="section-heading">
                <SlidersHorizontal size={19} />
                <h2 id="process-info-title">过程资料生成信息</h2>
              </div>
              <button className="icon-button" onClick={() => setIsProcessInfoModalOpen(false)} title="关闭" aria-label="关闭生成信息">
                <X size={18} />
              </button>
            </div>
            <div className="process-field-grid modal-process-field-grid">
              <label className="text-field">
                <span>工程名称</span>
                <input
                  value={processProjectName}
                  onChange={(event) => setProcessProjectName(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>总承包单位</span>
                <input
                  value={processGeneralContractorUnit}
                  onChange={(event) => setProcessGeneralContractorUnit(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>总承包单位项目负责人</span>
                <input
                  value={processGeneralContractorManager}
                  onChange={(event) => setProcessGeneralContractorManager(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>总承包单位项目技术负责人</span>
                <input
                  value={processGeneralContractorTechnicalLeader}
                  onChange={(event) => setProcessGeneralContractorTechnicalLeader(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>施工单位</span>
                <input
                  value={processConstructionUnit}
                  onChange={(event) => setProcessConstructionUnit(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>施工单位项目负责人</span>
                <input
                  value={processConstructionManager}
                  onChange={(event) => setProcessConstructionManager(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>施工单位项目技术负责人</span>
                <input
                  value={processConstructionTechnicalLeader}
                  onChange={(event) => setProcessConstructionTechnicalLeader(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>分包单位</span>
                <input
                  value={processSubcontractorUnit}
                  onChange={(event) => setProcessSubcontractorUnit(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>分包单位项目负责人</span>
                <input
                  value={processSubcontractorManager}
                  onChange={(event) => setProcessSubcontractorManager(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>分包内容</span>
                <input
                  value={processSubcontractorContent}
                  onChange={(event) => setProcessSubcontractorContent(event.currentTarget.value)}
                />
              </label>
              <label className="text-field">
                <span>监理项目部</span>
                <input
                  value={processSupervisionDepartment}
                  onChange={(event) => setProcessSupervisionDepartment(event.currentTarget.value)}
                />
              </label>
            </div>
            <div className="modal-footer">
              <button className="primary-button" type="button" onClick={() => setIsProcessInfoModalOpen(false)}>
                完成
              </button>
            </div>
          </section>
        </div>
      ) : null}

      {previewRecord ? (
        <div className="modal-backdrop" onClick={() => setPreviewRecord(null)}>
          <section className="preview-modal" role="dialog" aria-modal="true" aria-labelledby="preview-title" onClick={(event) => event.stopPropagation()}>
            <div className="modal-heading">
              <div className="section-heading">
                <FileText size={19} />
                <h2 id="preview-title">案卷字段预览</h2>
              </div>
              <div className="modal-actions">
                <button className="secondary-button modal-nav-button" onClick={() => switchPreview(-1)} disabled={previewIndex <= 0}>
                  <ChevronLeft size={17} />
                  上一个
                </button>
                <span className="modal-count">
                  {previewIndex >= 0 ? previewIndex + 1 : 1} / {filteredRecords.length || 1}
                </span>
                <button
                  className="secondary-button modal-nav-button"
                  onClick={() => switchPreview(1)}
                  disabled={previewIndex < 0 || previewIndex >= filteredRecords.length - 1}
                >
                  下一个
                  <ChevronRight size={17} />
                </button>
                <button className="icon-button" onClick={() => setPreviewRecord(null)} title="关闭" aria-label="关闭字段预览">
                  <X size={18} />
                </button>
              </div>
            </div>
            <div className="preview-grid modal-preview-grid">
              {getRecordPreviewFields(previewRecord).map((field) => (
                <PreviewItem key={field.label} label={field.label} value={field.value} />
              ))}
            </div>
            <div className="detail-table-wrap">
              <div className="detail-table-title">
                <h3>卷内明细</h3>
                <span>{previewRecord.items.length} 条</span>
              </div>
              <table className="detail-table">
                <thead>
                  <tr>
                    <th>序号</th>
                    <th>文件编号</th>
                    <th>责任者</th>
                    <th>文件题名</th>
                    <th>文件日期</th>
                    <th>页号</th>
                    <th>备注</th>
                  </tr>
                </thead>
                <tbody>
                  {previewRecord.items.map((item, index) => (
                    <tr key={`${item.sequence}-${item.fileCode}-${index}`}>
                      <td>{item.sequence || "/"}</td>
                      <td>{item.fileCode || "/"}</td>
                      <td>{item.owner || "/"}</td>
                      <td>{item.title || "/"}</td>
                      <td>{item.fileDate || "/"}</td>
                      <td>{item.pageNo || "/"}</td>
                      <td>{item.note || "/"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      ) : null}
    </main>
  );
}

async function showOperationError(error: unknown) {
  await showDialogMessage(error instanceof Error ? error.message : String(error), {
    title: "操作失败",
    kind: "error",
  });
}

function PathLine({ label, value }: { label: string; value: string }) {
  return (
    <div className="path-line">
      <span>{label}</span>
      <strong title={value}>{value}</strong>
    </div>
  );
}

function PreviewItem({ label, value }: { label: string; value: string }) {
  return (
    <div className="preview-item">
      <span>{label}</span>
      <strong>{value || "/"}</strong>
    </div>
  );
}

function getRecordPreviewFields(record: ArchiveRecord): Array<{ label: string; value: string }> {
  return [
    { label: "分类号", value: record.categoryCode },
    { label: "档号", value: record.archiveCode },
    { label: "案卷题名", value: record.fullTitle },
    { label: "项目名称", value: record.projectName },
    { label: "案卷标题", value: record.volumeTitle },
    { label: "责任者", value: record.owner },
    { label: "立卷单位", value: record.filingUnit },
    { label: "保管期限", value: record.retentionPeriod },
    { label: "开始日期", value: record.startDate },
    { label: "结束日期", value: record.endDate },
    { label: "起止日期", value: record.dateRange },
    { label: "本卷共有", value: `${record.totalPages} 页` },
    { label: "文字材料", value: `${record.textPages} 页` },
    { label: "图样", value: `${record.drawingPages} 页` },
    { label: "卷内明细", value: `${record.items.length} 条` },
  ];
}

function Toggle({ checked, onChange, label }: { checked: boolean; onChange: (checked: boolean) => void; label: string }) {
  return (
    <label className="toggle">
      <input type="checkbox" checked={checked} onChange={(event) => onChange(event.currentTarget.checked)} />
      <span>{label}</span>
    </label>
  );
}

export default App;
