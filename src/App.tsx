import { useMemo, useState } from "react";
import { message as showDialogMessage, open } from "@tauri-apps/plugin-dialog";
import { openPath } from "@tauri-apps/plugin-opener";
import {
  Check,
  ChevronLeft,
  ChevronRight,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Info,
  Loader2,
  Search,
  Settings2,
  X,
} from "lucide-react";
import "./App.css";
import { generateArchiveDocs } from "./lib/docx";
import { parseArchiveWorkbook } from "./lib/excel";
import { readBinaryFile, writeBinaryFile } from "./lib/tauriFiles";
import type { ArchiveRecord } from "./lib/types";

const DEFAULT_BACKUP_NOTE = "";
const LAST_OUTPUT_DIR_KEY = "archive-docx-client:last-output-dir";
const ARCHIVE_DOCX_TAB = "archive-docx";

function App() {
  const [activeTab, setActiveTab] = useState(ARCHIVE_DOCX_TAB);
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
  const [isLoading, setIsLoading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);

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

  const canGenerate = selectedCodes.length > 0 && outputDir && (generateCover || generateNote || generateSpine);
  const previewIndex = previewRecord
    ? filteredRecords.findIndex((record) => record.archiveCode === previewRecord.archiveCode)
    : -1;

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

  async function generate() {
    if (!canGenerate) {
      return;
    }

    setIsGenerating(true);

    try {
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
      if (generated.errors.length > 0) {
        await showDialogMessage(
          `生成完成：${generated.files.length} 个文件，${generated.errors.length} 个失败。\n\n${generated.errors.join("\n")}`,
          { title: "生成完成", kind: "warning" },
        );
      } else {
        await showDialogMessage(`生成成功，共生成 ${generated.files.length} 个文件。`, {
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

  return (
    <main className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">Archive Workspace</p>
          <h1>档案工具箱</h1>
        </div>
        <div className="summary-strip">
          <SummaryItem label="案卷" value={records.length} />
          <SummaryItem label="已选" value={selectedCodes.length} />
          <SummaryItem label="待生成" value={estimateFileCount(selectedCodes.length, generateCover, generateNote, generateSpine)} />
        </div>
      </header>

      <nav className="app-tabs" aria-label="功能栏目">
        <button
          className={`tab-button ${activeTab === ARCHIVE_DOCX_TAB ? "active" : ""}`}
          onClick={() => setActiveTab(ARCHIVE_DOCX_TAB)}
          type="button"
        >
          <FileText size={17} />
          档案文档生成器
        </button>
      </nav>

      {activeTab === ARCHIVE_DOCX_TAB ? (
        <section className="tab-panel">
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

          <section className="workspace">
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
                    生成 DOCX
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

function SummaryItem({ label, value }: { label: string; value: number }) {
  return (
    <div className="summary-item">
      <span>{label}</span>
      <strong>{value}</strong>
    </div>
  );
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

function estimateFileCount(count: number, cover: boolean, note: boolean, spine: boolean): number {
  return count * Number(cover) + count * Number(note) + (spine ? Math.ceil(count / 7) : 0);
}

export default App;
