import { useEffect, useMemo, useState } from "react";
import { message as showDialogMessage, open } from "@tauri-apps/plugin-dialog";
import "./App.css";
import { AllProcessPanel } from "./components/AllProcessPanel";
import { AppHeader } from "./components/AppHeader";
import { ArchiveGeneratorPanel } from "./components/ArchiveGeneratorPanel";
import { ControlBand } from "./components/ControlBand";
import { PreviewModal } from "./components/PreviewModal";
import { ProcessGenerationPanel } from "./components/ProcessGenerationPanel";
import { ProcessInfoModal } from "./components/ProcessInfoModal";
import { RecordPane } from "./components/RecordPane";
import { TemplateGuideModal } from "./components/TemplateGuideModal";
import { TemplateManagerModal } from "./components/TemplateManagerModal";
import { showOperationError } from "./app/dialogs";
import { useAllProcessSelection } from "./hooks/useAllProcessSelection";
import { useProcessFields } from "./hooks/useProcessFields";
import { useProcessTemplateManager } from "./hooks/useProcessTemplateManager";
import { generateArchiveCatalog } from "./lib/catalog";
import { generateArchiveDocs } from "./lib/docx";
import { parseArchiveWorkbook } from "./lib/excel";
import {
  countProcessGenerationFiles,
  generateProcessDocs,
  generateSelectedProcessDocs,
  PROCESS_TEMPLATE_CATEGORY_IDS,
  type ProcessTemplateCategoryId,
  type ProcessTemplateModule,
} from "./lib/processDocs";
import {
  openSystemPath,
  readBinaryFile,
  writeBinaryFile,
} from "./lib/tauriFiles";
import type { ArchiveRecord } from "./lib/types";
import {
  ALL_PROCESS_DOCS_TAB,
  ARCHIVE_DOCX_TAB,
  COLLECTOR_LINE_TAB,
  COLLECTOR_LINE_TEMPLATE_CATEGORIES_KEY,
  DEFAULT_BACKUP_NOTE,
  LAST_OUTPUT_DIR_KEY,
  PROCESS_DOCS_TAB,
  PROCESS_TEMPLATE_CATEGORIES_KEY,
  SWITCH_STATION_TAB,
  SWITCH_STATION_TEMPLATE_CATEGORIES_KEY,
} from "./app/appConstants";
import { copyText } from "./app/clipboard";
import { loadSavedProcessTemplateCategories } from "./app/appStorage";
import {
  processModuleLabel,
} from "./app/templateOptions";

function App() {
  const [activeTab, setActiveTab] = useState(ARCHIVE_DOCX_TAB);
  const { fields: processFields, setters: processFieldSetters, processUserFields } = useProcessFields();
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
  const [selectedProcessTemplateCategories, setSelectedProcessTemplateCategories] = useState<ProcessTemplateCategoryId[]>(
    savedProcessTemplateCategories,
  );
  const [selectedSwitchStationTemplateCategories, setSelectedSwitchStationTemplateCategories] = useState<ProcessTemplateCategoryId[]>(
    savedSwitchStationTemplateCategories,
  );
  const [selectedCollectorLineTemplateCategories, setSelectedCollectorLineTemplateCategories] = useState<ProcessTemplateCategoryId[]>(
    savedCollectorLineTemplateCategories,
  );
  const [isProcessInfoModalOpen, setIsProcessInfoModalOpen] = useState(false);
  const [isTemplateManagerOpen, setIsTemplateManagerOpen] = useState(false);
  const [isTemplateGuideOpen, setIsTemplateGuideOpen] = useState(false);
  const [copiedPlaceholder, setCopiedPlaceholder] = useState("");
  const {
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
  } = useProcessTemplateManager({ activeTab, isTemplateManagerOpen });
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

  const {
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
  } = useAllProcessSelection({ selectedRecords, processManifest, allTemplateOptions });

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
    localStorage.setItem(PROCESS_TEMPLATE_CATEGORIES_KEY, JSON.stringify(selectedProcessTemplateCategories));
  }, [selectedProcessTemplateCategories]);

  useEffect(() => {
    localStorage.setItem(SWITCH_STATION_TEMPLATE_CATEGORIES_KEY, JSON.stringify(selectedSwitchStationTemplateCategories));
  }, [selectedSwitchStationTemplateCategories]);

  useEffect(() => {
    localStorage.setItem(COLLECTOR_LINE_TEMPLATE_CATEGORIES_KEY, JSON.stringify(selectedCollectorLineTemplateCategories));
  }, [selectedCollectorLineTemplateCategories]);

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

  async function copyPlaceholder(value: string) {
    await copyText(value);
    setCopiedPlaceholder(value);
    window.setTimeout(() => {
      setCopiedPlaceholder((current) => current === value ? "" : current);
    }, 1200);
  }

  return (
    <main className="app-shell">
      <AppHeader activeTab={activeTab} onTabChange={setActiveTab} />

      <section className="workspace-container">
        <aside className="left-panel">
          <ControlBand
            excelPath={excelPath}
            outputPath={actualOutputDir || outputDir}
            isLoading={isLoading}
            onChooseExcel={chooseExcel}
            onChooseOutputDir={chooseOutputDir}
          />
          <RecordPane
            recordsCount={records.length}
            filteredRecords={filteredRecords}
            selectedCodes={selectedCodes}
            query={query}
            showOnlySelected={showOnlySelected}
            onQueryChange={setQuery}
            onShowOnlySelectedChange={setShowOnlySelected}
            onToggleVisibleRecords={toggleVisibleRecords}
            onToggleCode={toggleCode}
            onPreviewRecord={setPreviewRecord}
          />
        </aside>

        <section className="right-panel">
          {activeTab === ARCHIVE_DOCX_TAB ? (
            <ArchiveGeneratorPanel
              backupNote={backupNote}
              generateCover={generateCover}
              generateNote={generateNote}
              generateSpine={generateSpine}
              generateCatalogWorkbook={generateCatalogWorkbook}
              selectedCodesCount={selectedCodes.length}
              generationFileCount={archiveGenerationFileCount}
              canGenerate={canGenerate}
              isGenerating={isGenerating}
              hasOutputDir={Boolean(outputDir)}
              onBackupNoteChange={setBackupNote}
              onGenerateCoverChange={setGenerateCover}
              onGenerateNoteChange={setGenerateNote}
              onGenerateSpineChange={setGenerateSpine}
              onGenerateCatalogWorkbookChange={setGenerateCatalogWorkbook}
              onGenerate={generate}
              onOpenOutputDir={openOutputDir}
            />
          ) : null}

          {activeTab === PROCESS_DOCS_TAB ? (
            <ProcessGenerationPanel
              title="过程资料生成"
              generateLabel="生成过程资料"
              selectedCodesCount={selectedCodes.length}
              generationFileCount={processGenerationFileCount}
              hasProcessManifest={Boolean(processManifest)}
              selectedCategories={selectedProcessTemplateCategories}
              canGenerate={canGenerateProcess}
              isGenerating={isGeneratingProcess}
              hasOutputDir={Boolean(outputDir)}
              onOpenInfo={() => setIsProcessInfoModalOpen(true)}
              onSelectAllCategories={() => setSelectedProcessTemplateCategories([...PROCESS_TEMPLATE_CATEGORY_IDS])}
              onClearCategories={() => setSelectedProcessTemplateCategories([])}
              onToggleCategory={toggleProcessTemplateCategory}
              onGenerate={() => generateProcess("process")}
              onOpenOutputDir={openOutputDir}
            />
          ) : null}

          {activeTab === ALL_PROCESS_DOCS_TAB ? (
            <AllProcessPanel
              selectedCodesCount={selectedCodes.length}
              rowCount={allProcessRows.length}
              unresolvedCount={unresolvedAllProcessRows.length}
              groups={allProcessGroups}
              isLoadingProcessManifest={isLoadingProcessManifest}
              hasProcessManifest={Boolean(processManifest)}
              generationFileCount={allProcessGenerationFileCount}
              canGenerate={canGenerateAllProcess}
              isGenerating={isGeneratingProcess}
              hasOutputDir={Boolean(outputDir)}
              allTemplateSearchTerms={allTemplateSearchTerms}
              activeAllTemplateRow={activeAllTemplateRow}
              onOpenInfo={() => setIsProcessInfoModalOpen(true)}
              onOpenTemplateManager={() => setIsTemplateManagerOpen(true)}
              onOpenBuiltInDir={() => void openBuiltInProcessTemplateDir()}
              onOpenUserDir={() => void openUserProcessTemplateDir()}
              onGenerate={generateAllProcess}
              onOpenOutputDir={openOutputDir}
              setActiveAllTemplateRow={setActiveAllTemplateRow}
              setAllTemplateSearchTerms={setAllTemplateSearchTerms}
              selectedAllTemplateOption={selectedAllTemplateOption}
              filteredAllTemplateOptions={filteredAllTemplateOptions}
              selectAllTemplate={selectAllTemplate}
            />
          ) : null}

          {activeTab === SWITCH_STATION_TAB ? (
            <ProcessGenerationPanel
              title="开关站电气设备安装（子单位工程）"
              generateLabel="生成开关站资料"
              selectedCodesCount={selectedCodes.length}
              generationFileCount={switchStationGenerationFileCount}
              hasProcessManifest={Boolean(processManifest)}
              selectedCategories={selectedSwitchStationTemplateCategories}
              canGenerate={canGenerateSwitchStation}
              isGenerating={isGeneratingProcess}
              hasOutputDir={Boolean(outputDir)}
              onOpenInfo={() => setIsProcessInfoModalOpen(true)}
              onSelectAllCategories={() => setSelectedSwitchStationTemplateCategories([...PROCESS_TEMPLATE_CATEGORY_IDS])}
              onClearCategories={() => setSelectedSwitchStationTemplateCategories([])}
              onToggleCategory={toggleSwitchStationTemplateCategory}
              onGenerate={() => generateProcess("switch-station")}
              onOpenOutputDir={openOutputDir}
            />
          ) : null}

          {activeTab === COLLECTOR_LINE_TAB ? (
            <ProcessGenerationPanel
              title="集电线路安装工程"
              generateLabel="生成集电线路资料"
              selectedCodesCount={selectedCodes.length}
              generationFileCount={collectorLineGenerationFileCount}
              hasProcessManifest={Boolean(processManifest)}
              selectedCategories={selectedCollectorLineTemplateCategories}
              canGenerate={canGenerateCollectorLine}
              isGenerating={isGeneratingProcess}
              hasOutputDir={Boolean(outputDir)}
              onOpenInfo={() => setIsProcessInfoModalOpen(true)}
              onSelectAllCategories={() => setSelectedCollectorLineTemplateCategories([...PROCESS_TEMPLATE_CATEGORY_IDS])}
              onClearCategories={() => setSelectedCollectorLineTemplateCategories([])}
              onToggleCategory={toggleCollectorLineTemplateCategory}
              onGenerate={() => generateProcess("collector-line")}
              onOpenOutputDir={openOutputDir}
            />
          ) : null}
        </section>
      </section>

      {isTemplateManagerOpen ? (
        <TemplateManagerModal
          isImporting={isImportingProcessTemplate}
          templateDirectoryEntries={templateDirectoryEntries}
          filteredTemplateDirectoryEntries={filteredTemplateDirectoryEntries}
          templateDirectoryError={templateDirectoryError}
          templateDirectoryQuery={templateDirectoryQuery}
          userProcessTemplates={userProcessTemplates}
          templateDrafts={templateDrafts}
          savingTemplateKeys={savingTemplateKeys}
          deletingTemplateKeys={deletingTemplateKeys}
          onClose={() => setIsTemplateManagerOpen(false)}
          onOpenGuide={() => setIsTemplateGuideOpen(true)}
          onImportTemplates={() => void importProcessTemplates()}
          onTemplateDirectoryQueryChange={setTemplateDirectoryQuery}
          onOpenTemplateDirectory={(entry) => void openTemplateDirectory(entry)}
          onUpdateTemplateDraft={updateTemplateDraft}
          onDeleteTemplate={(template) => void deleteProcessTemplate(template)}
          onSaveTemplate={(template) => void saveUserProcessTemplate(template)}
        />
      ) : null}

      {isTemplateGuideOpen ? (
        <TemplateGuideModal
          copiedPlaceholder={copiedPlaceholder}
          onCopyPlaceholder={(value) => void copyPlaceholder(value)}
          onClose={() => setIsTemplateGuideOpen(false)}
        />
      ) : null}

      {isProcessInfoModalOpen ? (
        <ProcessInfoModal
          projectName={processFields.projectName}
          generalContractorUnit={processFields.generalContractorUnit}
          generalContractorManager={processFields.generalContractorManager}
          generalContractorTechnicalLeader={processFields.generalContractorTechnicalLeader}
          constructionUnit={processFields.constructionUnit}
          constructionManager={processFields.constructionManager}
          constructionTechnicalLeader={processFields.constructionTechnicalLeader}
          subcontractorUnit={processFields.subcontractorUnit}
          subcontractorManager={processFields.subcontractorManager}
          subcontractorContent={processFields.subcontractorContent}
          supervisionDepartment={processFields.supervisionDepartment}
          onProjectNameChange={processFieldSetters.setProjectName}
          onGeneralContractorUnitChange={processFieldSetters.setGeneralContractorUnit}
          onGeneralContractorManagerChange={processFieldSetters.setGeneralContractorManager}
          onGeneralContractorTechnicalLeaderChange={processFieldSetters.setGeneralContractorTechnicalLeader}
          onConstructionUnitChange={processFieldSetters.setConstructionUnit}
          onConstructionManagerChange={processFieldSetters.setConstructionManager}
          onConstructionTechnicalLeaderChange={processFieldSetters.setConstructionTechnicalLeader}
          onSubcontractorUnitChange={processFieldSetters.setSubcontractorUnit}
          onSubcontractorManagerChange={processFieldSetters.setSubcontractorManager}
          onSubcontractorContentChange={processFieldSetters.setSubcontractorContent}
          onSupervisionDepartmentChange={processFieldSetters.setSupervisionDepartment}
          onClose={() => setIsProcessInfoModalOpen(false)}
        />
      ) : null}

      {previewRecord ? (
        <PreviewModal
          record={previewRecord}
          previewIndex={previewIndex}
          totalCount={filteredRecords.length}
          onPrevious={() => switchPreview(-1)}
          onNext={() => switchPreview(1)}
          onClose={() => setPreviewRecord(null)}
        />
      ) : null}
    </main>
  );
}

export default App;
