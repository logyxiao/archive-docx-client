import { useEffect, useMemo, useState } from "react";
import { ask, message as showDialogMessage, open } from "@tauri-apps/plugin-dialog";
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
  Save,
  Search,
  Settings2,
  SlidersHorizontal,
  Trash2,
  Upload,
  X,
} from "lucide-react";
import "./App.css";
import { generateArchiveCatalog } from "./lib/catalog";
import { generateArchiveDocs } from "./lib/docx";
import { parseArchiveWorkbook } from "./lib/excel";
import {
  allProcessTemplateOptions,
  countProcessGenerationFiles,
  generateProcessDocs,
  generateSelectedProcessDocs,
  loadProcessManifest,
  matchingAllProcessTemplates,
  PROCESS_TEMPLATE_CATEGORIES,
  PROCESS_TEMPLATE_CATEGORY_IDS,
  normalizeProcessTemplateCategories,
  type ProcessTemplateManifest,
  type ProcessTemplateMatch,
  type ProcessTemplate,
  type ProcessTemplateCategoryId,
  type ProcessTemplateMatchMode,
  type ProcessTemplateModule,
  type ProcessTemplateSelection,
} from "./lib/processDocs";
import {
  deleteUserProcessTemplate,
  importProcessTemplate,
  openSystemPath,
  processBuiltinTemplateDir,
  processTemplateUserDir,
  readBinaryFile,
  updateUserProcessTemplate,
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

interface TemplateDirectoryEntry {
  key: string;
  name: string;
  path: string;
  source: string;
  kind: string;
  searchText: string;
}

interface UserTemplateDraft {
  displayName: string;
  matchKeywordsText: string;
  matchMode: ProcessTemplateMatchMode;
  templateModule: ProcessTemplateModule;
  category: ProcessTemplateCategoryId | "";
  enabled: boolean;
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

const TEMPLATE_IMPORT_STEPS = [
  "在 Word 或 Excel 模板中写入占位符，例如 {{项目名}}、{{文件题名}}、{{文件日期中文}}。",
  "Excel 质量结果可写 {{质量验收结果:<2}} 或 {{质量验收结果:≥6}}；下限型标准默认填 6~9，明确范围可写 {{质量验收结果:6~10}}。",
  "点击“导入模板”，选择 .docx 或 .xlsx 文件，软件会复制到导入模板目录。",
  "在导入模板配置里填写匹配关键词；生成时文件题名命中关键词就会默认选中该模板。",
];

const TEMPLATE_CONFIG_FIELDS = [
  "显示名称：模板管理页和下拉框里看到的名称。",
  "所属模块：过程资料、开关站电气设备安装、集电线路安装工程。",
  "模板分类：用于普通过程资料页的分类筛选；不填则自动识别。",
  "匹配关键词：每行一个，也可用逗号分隔。",
  "关键词模式：任一命中或全部命中。",
  "启用状态：关闭后该模板不参与默认匹配和下拉选择。",
  "质量验收结果：支持 ±2、<2、≤2、0~2、≥6 这类标准，生成时自动随机填充。",
];

const PROCESS_TEMPLATE_PLACEHOLDERS = [
  "项目名",
  "工程名称",
  "工程编号",
  "档号",
  "案卷题名",
  "文件编号",
  "文件题名",
  "文件题名去项目",
  "文件题名主题",
  "验收项目名称",
  "文件日期",
  "文件日期中文",
  "责任者",
  "编制单位",
  "立卷单位",
  "总承包单位",
  "施工单位",
  "监理单位",
  "监理项目部",
  "分包单位",
  "分包内容",
  "质量验收结果:<2",
  "质量验收结果:≥6",
  "质量验收结果:6~10",
];

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

async function copyText(value: string) {
  try {
    await navigator.clipboard.writeText(value);
    return;
  } catch {
    const textarea = document.createElement("textarea");
    textarea.value = value;
    textarea.setAttribute("readonly", "true");
    textarea.style.position = "fixed";
    textarea.style.left = "-9999px";
    document.body.appendChild(textarea);
    textarea.select();
    document.execCommand("copy");
    document.body.removeChild(textarea);
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

function templateIdentityKey(template: ProcessTemplate): string {
  return [
    template.userTemplatePath ?? "builtin",
    template.templateFile,
    template.originalName,
    template.outputFileCodeOverride ?? "",
  ].join("\u0000");
}

function userTemplateKey(template: ProcessTemplate): string {
  return template.userTemplatePath ?? templateIdentityKey(template);
}

function draftFromUserTemplate(template: ProcessTemplate): UserTemplateDraft {
  return {
    displayName: template.displayName?.trim() || template.originalName || template.templateFile,
    matchKeywordsText: (template.matchKeywords ?? []).join("\n"),
    matchMode: template.matchMode === "all" ? "all" : "any",
    templateModule: template.templateModule ?? "process",
    category: template.category ?? "",
    enabled: template.enabled !== false,
  };
}

function parseTemplateKeywords(value: string): string[] {
  const result: string[] = [];
  for (const keyword of value.split(/[\n,，;；]+/)) {
    const trimmed = keyword.trim();
    if (!trimmed || result.includes(trimmed)) {
      continue;
    }
    result.push(trimmed);
  }
  return result;
}

function joinTemplatePath(dir: string, fileName: string): string {
  if (!dir) {
    return "";
  }

  const separator = dir.includes("\\") ? "\\" : "/";
  return `${dir.replace(/[\\/]+$/, "")}${separator}${fileName}`;
}

function parentDirectory(path: string): string {
  const normalized = path.replace(/[\\/]+$/, "");
  const index = Math.max(normalized.lastIndexOf("/"), normalized.lastIndexOf("\\"));
  return index > 0 ? normalized.slice(0, index) : normalized;
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
  const [isTemplateManagerOpen, setIsTemplateManagerOpen] = useState(false);
  const [isTemplateGuideOpen, setIsTemplateGuideOpen] = useState(false);
  const [copiedPlaceholder, setCopiedPlaceholder] = useState("");
  const [builtInTemplateDir, setBuiltInTemplateDir] = useState("");
  const [templateDirectoryError, setTemplateDirectoryError] = useState("");
  const [templateDirectoryQuery, setTemplateDirectoryQuery] = useState("");
  const [templateDrafts, setTemplateDrafts] = useState<Record<string, UserTemplateDraft>>({});
  const [savingTemplateKeys, setSavingTemplateKeys] = useState<Record<string, boolean>>({});
  const [deletingTemplateKeys, setDeletingTemplateKeys] = useState<Record<string, boolean>>({});
  const [isLoading, setIsLoading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isGeneratingProcess, setIsGeneratingProcess] = useState(false);

 
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

  const archiveGenerationFileCount = useMemo(() => {
    const recordCount = selectedRecords.length;
    if (recordCount === 0) {
      return 0;
    }

    let count = 0;
    if (generateCover) {
      count += recordCount;
    }
    if (generateNote) {
      count += recordCount;
    }
    if (generateSpine) {
      count += Math.ceil(recordCount / 7);
    }
    if (generateCatalogWorkbook) {
      count += 1;
    }
    return count;
  }, [generateCatalogWorkbook, generateCover, generateNote, generateSpine, selectedRecords]);

  const processGenerationFileCount = useMemo(
    () => processManifest
      ? countProcessGenerationFiles(records, selectedCodes, processManifest.templates, selectedProcessTemplateCategories, "process")
      : 0,
    [processManifest, records, selectedCodes, selectedProcessTemplateCategories],
  );

  const switchStationGenerationFileCount = useMemo(
    () => processManifest
      ? countProcessGenerationFiles(records, selectedCodes, processManifest.templates, selectedSwitchStationTemplateCategories, "switch-station")
      : 0,
    [processManifest, records, selectedCodes, selectedSwitchStationTemplateCategories],
  );

  const collectorLineGenerationFileCount = useMemo(
    () => processManifest
      ? countProcessGenerationFiles(records, selectedCodes, processManifest.templates, selectedCollectorLineTemplateCategories, "collector-line")
      : 0,
    [processManifest, records, selectedCodes, selectedCollectorLineTemplateCategories],
  );

  const allProcessGenerationFileCount = useMemo(
    () =>
      allProcessRows.reduce((count, row) => {
        const selectedKey = selectedAllTemplateKey(row.key, row.matches);
        return templateMatchFromKey(selectedKey, allTemplateOptions) ? count + 1 : count;
      }, 0),
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

  async function copyPlaceholder(value: string) {
    await copyText(value);
    setCopiedPlaceholder(value);
    window.setTimeout(() => {
      setCopiedPlaceholder((current) => current === value ? "" : current);
    }, 1200);
  }

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
      <header className="app-header">
        <div className="app-logo-area">
          <FileText className="logo-icon" size={20} />
          <h1 className="app-title">Docx 归档助手</h1>
        </div>
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
            全部过程资料
          </button>
          <button
            className={`tab-button ${activeTab === SWITCH_STATION_TAB ? "active" : ""}`}
            onClick={() => setActiveTab(SWITCH_STATION_TAB)}
            type="button"
          >
            <FolderTree size={17} />
            开关站电气设备安装
          </button>
          <button
            className={`tab-button ${activeTab === COLLECTOR_LINE_TAB ? "active" : ""}`}
            onClick={() => setActiveTab(COLLECTOR_LINE_TAB)}
            type="button"
          >
            <FolderTree size={17} />
            集电线路安装工程
          </button>
        </nav>
      </header>

      <section className="workspace-container">
        <aside className="left-panel">
          {controlBand}
          {recordPane}
        </aside>

        <section className="right-panel">
          {activeTab === ARCHIVE_DOCX_TAB ? (
            <div className="detail-pane">
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
                    生成文件 {selectedCodes.length > 0 ? `(${archiveGenerationFileCount}个文件)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </div>
          ) : null}

          {activeTab === PROCESS_DOCS_TAB ? (
            <div className="detail-pane">
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
                    生成过程资料 {selectedCodes.length > 0 && processManifest ? `(${processGenerationFileCount}个文件)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </div>
          ) : null}

          {activeTab === ALL_PROCESS_DOCS_TAB ? (
            <div className="detail-pane">
              <div className="tool-section all-process-tool">
                <div className="process-template-section">
                  <div className="all-template-summary">
                    <span>
                      已选 {selectedCodes.length} 个档号，{allProcessRows.length} 条文件题名
                      {unresolvedAllProcessRows.length > 0 ? `，${unresolvedAllProcessRows.length} 条需选择模板` : ""}
                    </span>
                    <div className="template-summary-actions">
                      <button className="template-dir-button" type="button" onClick={() => setIsProcessInfoModalOpen(true)}>
                        <SlidersHorizontal size={15} />
                        填写生成信息
                      </button>
                      <button type="button" className="template-dir-button" onClick={() => setIsTemplateManagerOpen(true)}>
                        <Settings2 size={15} />
                        模板管理
                      </button>
                      <button type="button" className="template-dir-button" onClick={() => void openBuiltInProcessTemplateDir()}>
                        <FolderOpen size={15} />
                        内置目录
                      </button>
                      <button type="button" className="template-dir-button" onClick={() => void openUserProcessTemplateDir()}>
                        <FolderOpen size={15} />
                        导入目录
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
                    <div className="empty-state detail-empty all-process-empty">请选择档号</div>
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
                            const isUnmatched = row.matches.length === 0 && !selectedOption?.match;
                            const searchTerm = allTemplateSearchTerms[row.key] ?? "";
                            const inputValue = activeAllTemplateRow === row.key ? searchTerm : selectedOption?.label ?? "";
                            const options = activeAllTemplateRow === row.key ? filteredAllTemplateOptions(row.key) : [];
                            return (
                              <article key={row.key} className={`all-process-row ${isUnmatched ? "unmatched" : ""}`}>
                                <div className="all-process-row-main">
                                  <div className="all-process-sequence">{row.item.sequence || "-"}</div>
                                  <div className="all-process-title">
                                    <strong>{row.item.title || "未命名文件"}</strong>
                                    <span>
                                      {row.item.fileCode || "/"}
                                      {isUnmatched ? <em className="all-process-unmatched-badge">未匹配</em> : null}
                                    </span>
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
                    生成 {selectedCodes.length > 0 && processManifest ? `(${allProcessGenerationFileCount}个文件)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </div>
          ) : null}

          {activeTab === SWITCH_STATION_TAB ? (
            <div className="detail-pane">
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
                    生成开关站资料 {selectedCodes.length > 0 && processManifest ? `(${switchStationGenerationFileCount}个文件)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </div>
          ) : null}

          {activeTab === COLLECTOR_LINE_TAB ? (
            <div className="detail-pane">
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
                    生成集电线路资料 {selectedCodes.length > 0 && processManifest ? `(${collectorLineGenerationFileCount}个文件)` : ""}
                  </button>
                  <button className="secondary-button open-output-button" onClick={openOutputDir} disabled={!outputDir}>
                    <FolderOpen size={18} />
                    打开输出目录
                  </button>
                </div>
              </div>
            </div>
          ) : null}
        </section>
      </section>

      {isTemplateManagerOpen ? (
        <div className="modal-backdrop" onClick={() => setIsTemplateManagerOpen(false)}>
          <section
            className="template-manager-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="template-manager-title"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <div className="section-heading">
                <Settings2 size={19} />
                <h2 id="template-manager-title">模板管理</h2>
              </div>
              <div className="modal-actions">
                <button
                  type="button"
                  className="secondary-button modal-nav-button"
                  onClick={() => setIsTemplateGuideOpen(true)}
                >
                  <Info size={15} />
                  导入教程
                </button>
                <button
                  type="button"
                  className="secondary-button modal-nav-button"
                  onClick={() => void importProcessTemplates()}
                  disabled={isImportingProcessTemplate}
                >
                  {isImportingProcessTemplate ? <Loader2 className="spin" size={15} /> : <Upload size={15} />}
                  导入模板
                </button>
                <button className="icon-button" onClick={() => setIsTemplateManagerOpen(false)} title="关闭" aria-label="关闭模板管理">
                  <X size={18} />
                </button>
              </div>
            </div>

            <section className="template-directory-section" aria-label="当前模板目录">
              <div className="template-section-heading">
                <span>当前模板目录</span>
                <span>{filteredTemplateDirectoryEntries.length} / {templateDirectoryEntries.length} 个模板</span>
              </div>
              <label className="template-directory-search">
                <Search size={15} />
                <input
                  value={templateDirectoryQuery}
                  onChange={(event) => setTemplateDirectoryQuery(event.currentTarget.value)}
                  placeholder="搜索模板名称、来源、类型或路径"
                />
              </label>
              {templateDirectoryError ? (
                <div className="template-directory-error">{templateDirectoryError}</div>
              ) : null}
              <div className="template-directory-list">
                {templateDirectoryEntries.length === 0 ? (
                  <div className="empty-state detail-empty">暂无模板</div>
                ) : filteredTemplateDirectoryEntries.length === 0 ? (
                  <div className="empty-state detail-empty">没有匹配的模板</div>
                ) : (
                  filteredTemplateDirectoryEntries.map((entry) => (
                    <div key={entry.key} className="template-directory-row">
                      <div className="template-directory-name">
                        <strong>{entry.name}</strong>
                        <span>
                          {entry.source} · {entry.kind}
                        </span>
                      </div>
                      <code>{entry.path || "内置模板路径加载中"}</code>
                      <button
                        type="button"
                        className="template-dir-button"
                        onClick={() => void openTemplateDirectory(entry)}
                        disabled={!entry.path}
                      >
                        <FolderOpen size={14} />
                        打开目录
                      </button>
                    </div>
                  ))
                )}
              </div>
            </section>

            <section className="template-user-config-section" aria-label="导入模板配置">
              <div className="template-section-heading">
                <span>导入模板配置</span>
                <span>{userProcessTemplates.length} 个导入模板</span>
              </div>
              {userProcessTemplates.length === 0 ? (
                <div className="empty-state detail-empty">导入模板后可在这里填写匹配配置</div>
              ) : (
                <div className="template-manager-list">
                  {userProcessTemplates.map((template) => {
                    const key = userTemplateKey(template);
                    const draft = templateDrafts[key] ?? draftFromUserTemplate(template);
                    const isSaving = Boolean(savingTemplateKeys[key]);
                    const isDeleting = Boolean(deletingTemplateKeys[key]);
                    return (
                      <article key={key} className={`template-manager-item ${draft.enabled ? "" : "disabled"}`}>
                        <div className="template-manager-item-heading">
                          <div>
                            <strong>{template.templateFile}</strong>
                            <span>{template.userTemplatePath}</span>
                          </div>
                          <label className="template-enabled-toggle">
                            <input
                              type="checkbox"
                              checked={draft.enabled}
                              onChange={(event) => updateTemplateDraft(template, { enabled: event.currentTarget.checked })}
                            />
                            启用
                          </label>
                        </div>
                        <div className="template-manager-grid">
                          <label className="text-field">
                            <span>显示名称</span>
                            <input
                              value={draft.displayName}
                              onChange={(event) => updateTemplateDraft(template, { displayName: event.currentTarget.value })}
                            />
                          </label>
                          <label className="text-field">
                            <span>所属模块</span>
                            <select
                              value={draft.templateModule}
                              onChange={(event) =>
                                updateTemplateDraft(template, { templateModule: event.currentTarget.value as ProcessTemplateModule })
                              }
                            >
                              <option value="process">过程资料</option>
                              <option value="switch-station">开关站电气设备安装</option>
                              <option value="collector-line">集电线路安装工程</option>
                            </select>
                          </label>
                          <label className="text-field">
                            <span>模板分类</span>
                            <select
                              value={draft.category}
                              onChange={(event) =>
                                updateTemplateDraft(template, { category: event.currentTarget.value as ProcessTemplateCategoryId | "" })
                              }
                            >
                              <option value="">自动识别</option>
                              {PROCESS_TEMPLATE_CATEGORIES.map((category) => (
                                <option key={category.id} value={category.id}>
                                  {category.label}
                                </option>
                              ))}
                            </select>
                          </label>
                          <label className="text-field">
                            <span>关键词模式</span>
                            <select
                              value={draft.matchMode}
                              onChange={(event) =>
                                updateTemplateDraft(template, { matchMode: event.currentTarget.value as ProcessTemplateMatchMode })
                              }
                            >
                              <option value="any">任一命中</option>
                              <option value="all">全部命中</option>
                            </select>
                          </label>
                        </div>
                        <label className="text-field template-keyword-field">
                          <span>匹配关键词</span>
                          <textarea
                            value={draft.matchKeywordsText}
                            onChange={(event) => updateTemplateDraft(template, { matchKeywordsText: event.currentTarget.value })}
                            placeholder="每行一个关键词，也可以用逗号分隔；留空时按模板文件名自动匹配"
                          />
                        </label>
                        <div className="template-manager-actions">
                          <button
                            type="button"
                            className="secondary-button modal-nav-button"
                            onClick={() => void deleteProcessTemplate(template)}
                            disabled={isDeleting || isSaving}
                          >
                            {isDeleting ? <Loader2 className="spin" size={15} /> : <Trash2 size={15} />}
                            删除
                          </button>
                          <button
                            type="button"
                            className="primary-button modal-nav-button"
                            onClick={() => void saveUserProcessTemplate(template)}
                            disabled={isSaving || isDeleting}
                          >
                            {isSaving ? <Loader2 className="spin" size={15} /> : <Save size={15} />}
                            保存
                          </button>
                        </div>
                      </article>
                    );
                  })}
                </div>
              )}
            </section>
          </section>
        </div>
      ) : null}

      {isTemplateGuideOpen ? (
        <div className="modal-backdrop" onClick={() => setIsTemplateGuideOpen(false)}>
          <section
            className="template-guide-modal"
            role="dialog"
            aria-modal="true"
            aria-labelledby="template-guide-title"
            onClick={(event) => event.stopPropagation()}
          >
            <div className="modal-heading">
              <div className="section-heading">
                <Info size={19} />
                <h2 id="template-guide-title">导入教程</h2>
              </div>
              <button className="icon-button" onClick={() => setIsTemplateGuideOpen(false)} title="关闭" aria-label="关闭导入教程">
                <X size={18} />
              </button>
            </div>
            <div className="template-guide-section">
              <div className="template-guide-card">
                <div className="template-section-heading">
                  <span>导入步骤</span>
                </div>
                <ol className="template-guide-list">
                  {TEMPLATE_IMPORT_STEPS.map((step) => (
                    <li key={step}>{step}</li>
                  ))}
                </ol>
              </div>
              <div className="template-guide-card">
                <div className="template-section-heading">
                  <span>可填写配置</span>
                </div>
                <ul className="template-guide-list">
                  {TEMPLATE_CONFIG_FIELDS.map((field) => (
                    <li key={field}>{field}</li>
                  ))}
                </ul>
              </div>
              <div className="template-guide-card">
                <div className="template-section-heading">
                  <span>常用占位符</span>
                </div>
                <div className="placeholder-grid">
                  {PROCESS_TEMPLATE_PLACEHOLDERS.map((placeholder) => {
                    const placeholderText = `{{${placeholder}}}`;
                    const isCopied = copiedPlaceholder === placeholderText;
                    return (
                      <button
                        key={placeholder}
                        type="button"
                        className={`placeholder-copy ${isCopied ? "copied" : ""}`}
                        onClick={() => void copyPlaceholder(placeholderText)}
                        title={`复制 ${placeholderText}`}
                      >
                        <code>{placeholderText}</code>
                        <span>{isCopied ? "已复制" : "复制"}</span>
                      </button>
                    );
                  })}
                </div>
              </div>
            </div>
          </section>
        </div>
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
