import { useEffect, useMemo, useState } from "react";
import { ask, message as showDialogMessage, open } from "@tauri-apps/plugin-dialog";
import { openPath } from "@tauri-apps/plugin-opener";
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
  X,
} from "lucide-react";
import "./App.css";
import { generateArchiveCatalog } from "./lib/catalog";
import { generateArchiveDocs } from "./lib/docx";
import { parseArchiveWorkbook } from "./lib/excel";
import {
  generateProcessDocs,
  PROCESS_TEMPLATE_CATEGORIES,
  PROCESS_TEMPLATE_CATEGORY_IDS,
  type ProcessTemplateCategoryId,
} from "./lib/processDocs";
import { readBinaryFile, writeBinaryFile } from "./lib/tauriFiles";
import type { ArchiveRecord } from "./lib/types";

const DEFAULT_BACKUP_NOTE = "";
const LAST_OUTPUT_DIR_KEY = "archive-docx-client:last-output-dir";
const PROCESS_FIELDS_KEY = "archive-docx-client:process-fields";
const PROCESS_TEMPLATE_CATEGORIES_KEY = "archive-docx-client:process-template-categories";
const ARCHIVE_DOCX_TAB = "archive-docx";
const PROCESS_DOCS_TAB = "process-docs";

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

function loadSavedProcessTemplateCategories(): ProcessTemplateCategoryId[] {
  try {
    const rawValue = localStorage.getItem(PROCESS_TEMPLATE_CATEGORIES_KEY);
    if (!rawValue) {
      return [...PROCESS_TEMPLATE_CATEGORY_IDS];
    }

    const parsed = JSON.parse(rawValue);
    if (!Array.isArray(parsed)) {
      return [...PROCESS_TEMPLATE_CATEGORY_IDS];
    }

    const validIds = new Set<ProcessTemplateCategoryId>(PROCESS_TEMPLATE_CATEGORY_IDS);
    return parsed.filter((value): value is ProcessTemplateCategoryId => validIds.has(value));
  } catch {
    return [...PROCESS_TEMPLATE_CATEGORY_IDS];
  }
}

function App() {
  const [activeTab, setActiveTab] = useState(ARCHIVE_DOCX_TAB);
  const savedProcessFields = useMemo(loadSavedProcessFields, []);
  const savedProcessTemplateCategories = useMemo(loadSavedProcessTemplateCategories, []);
  const [excelPath, setExcelPath] = useState("");
  const [outputDir, setOutputDir] = useState(() => localStorage.getItem(LAST_OUTPUT_DIR_KEY) ?? "");
  const [records, setRecords] = useState<ArchiveRecord[]>([]);
  const [selectedCodes, setSelectedCodes] = useState<string[]>([]);
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
  const [isProcessInfoModalOpen, setIsProcessInfoModalOpen] = useState(false);
  const [isLoading, setIsLoading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [isGeneratingProcess, setIsGeneratingProcess] = useState(false);
  const [isCheckingUpdate, setIsCheckingUpdate] = useState(false);

  const filteredRecords = useMemo(() => {
    const keyword = query.trim().toLowerCase();
    if (!keyword) {
      return records;
    }

    return records.filter((record) =>
      [record.archiveCode, record.fullTitle, record.filingUnit, record.retentionPeriod]
        .join(" ")
        .toLowerCase()
        .includes(keyword),
    );
  }, [query, records]);

  const shouldGenerateDocx = generateCover || generateNote || generateSpine;
  const canGenerate = selectedCodes.length > 0 && outputDir && (shouldGenerateDocx || generateCatalogWorkbook);
  const canGenerateProcess =
    selectedCodes.length > 0 && Boolean(outputDir) && selectedProcessTemplateCategories.length > 0 && !isGeneratingProcess;
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
    if (!outputDir) {
      return;
    }

    try {
      await openPath(outputDir);
    } catch (error) {
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
            outputDir,
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
            outputDir,
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

  async function generateProcess() {
    if (!canGenerateProcess) {
      return;
    }

    setIsGeneratingProcess(true);

    try {
      const result = await generateProcessDocs(
        records,
        {
          selectedCodes,
          outputDir,
          selectedTemplateCategories: selectedProcessTemplateCategories,
          userFields: {
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
          },
        },
        writeBinaryFile,
      );
      const skippedText = result.skipped.length > 0 ? `\n跳过 ${result.skipped.length} 条：\n${result.skipped.slice(0, 12).join("\n")}` : "";
      const errorText = result.errors.length > 0 ? `\n失败 ${result.errors.length} 个：\n${result.errors.slice(0, 12).join("\n")}` : "";
      await showDialogMessage(`过程资料生成完成：${result.files.length} 个文件。${skippedText}${errorText}`, {
        title: result.errors.length > 0 ? "生成完成" : "生成成功",
        kind: result.errors.length > 0 ? "warning" : "info",
      });
    } catch (error) {
      await showOperationError(error);
    } finally {
      setIsGeneratingProcess(false);
    }
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
        <PathLine label="输出" value={outputDir || "未选择"} />
      </div>
    </section>
  );

  const recordPane = (
    <aside className="record-pane">
      <div className="pane-title">
        <div>
          <h2>案卷列表</h2>
          <p>{filteredRecords.length} 条匹配记录</p>
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
      <label className="search-box">
        <Search size={17} />
        <input value={query} onChange={(event) => setQuery(event.currentTarget.value)} placeholder="搜索档号、题名、单位" />
      </label>
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
                    生成文件
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
                  <button className="generate-button" onClick={generateProcess} disabled={!canGenerateProcess}>
                    {isGeneratingProcess ? <Loader2 className="spin" size={18} /> : <FolderTree size={18} />}
                    生成过程资料
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
