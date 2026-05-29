import { useMemo, useState } from "react";
import { open } from "@tauri-apps/plugin-dialog";
import {
  Check,
  ChevronRight,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Loader2,
  Search,
  Settings2,
} from "lucide-react";
import "./App.css";
import { generateArchiveDocs } from "./lib/docx";
import { parseArchiveWorkbook } from "./lib/excel";
import { readBinaryFile, writeBinaryFile } from "./lib/tauriFiles";
import type { ArchiveRecord, GenerationResult } from "./lib/types";

const DEFAULT_BACKUP_NOTE = "";

function App() {
  const [excelPath, setExcelPath] = useState("");
  const [outputDir, setOutputDir] = useState("");
  const [records, setRecords] = useState<ArchiveRecord[]>([]);
  const [selectedCodes, setSelectedCodes] = useState<string[]>([]);
  const [query, setQuery] = useState("");
  const [backupNote, setBackupNote] = useState(DEFAULT_BACKUP_NOTE);
  const [generateCover, setGenerateCover] = useState(true);
  const [generateNote, setGenerateNote] = useState(true);
  const [generateSpine, setGenerateSpine] = useState(true);
  const [isLoading, setIsLoading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);
  const [message, setMessage] = useState("");
  const [result, setResult] = useState<GenerationResult | null>(null);

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

  const selectedRecords = useMemo(
    () => records.filter((record) => selectedCodes.includes(record.archiveCode)),
    [records, selectedCodes],
  );

  const activePreview = selectedRecords[0] ?? records[0] ?? null;
  const canGenerate = selectedCodes.length > 0 && outputDir && (generateCover || generateNote || generateSpine);

  async function chooseExcel() {
    const selected = await open({
      multiple: false,
      filters: [{ name: "Excel", extensions: ["xlsx", "xls"] }],
    });

    if (typeof selected !== "string") {
      return;
    }

    setIsLoading(true);
    setMessage("");
    setResult(null);

    try {
      const bytes = await readBinaryFile(selected);
      const parsed = parseArchiveWorkbook(bytes);
      setExcelPath(selected);
      setRecords(parsed);
      setSelectedCodes([]);
      setMessage(`已读取 ${parsed.length} 个案卷`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setIsLoading(false);
    }
  }

  async function chooseOutputDir() {
    const selected = await open({ directory: true, multiple: false });
    if (typeof selected === "string") {
      setOutputDir(selected);
      setResult(null);
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

  async function generate() {
    if (!canGenerate) {
      return;
    }

    setIsGenerating(true);
    setResult(null);
    setMessage("");

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
      setResult(generated);
      setMessage(`生成完成：${generated.files.length} 个文件`);
    } catch (error) {
      setMessage(error instanceof Error ? error.message : String(error));
    } finally {
      setIsGenerating(false);
    }
  }

  return (
    <main className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">DOCX Archive Builder</p>
          <h1>档案文档生成器</h1>
        </div>
        <div className="summary-strip">
          <SummaryItem label="案卷" value={records.length} />
          <SummaryItem label="已选" value={selectedCodes.length} />
          <SummaryItem label="待生成" value={estimateFileCount(selectedCodes.length, generateCover, generateNote, generateSpine)} />
        </div>
      </header>

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
              <button
                key={record.archiveCode}
                className={`record-row ${selectedCodes.includes(record.archiveCode) ? "selected" : ""}`}
                onClick={() => toggleCode(record.archiveCode)}
              >
                <span className="checkmark">{selectedCodes.includes(record.archiveCode) ? <Check size={14} /> : null}</span>
                <span>
                  <strong>{record.archiveCode}</strong>
                  <small>{record.fullTitle}</small>
                </span>
                <ChevronRight size={16} />
              </button>
            ))}
            {records.length === 0 ? <div className="empty-state">请选择 Excel 总目录</div> : null}
          </div>
        </aside>

        <section className="detail-pane">
          <div className="tool-section">
            <div className="section-heading">
              <FileText size={19} />
              <h2>字段预览</h2>
            </div>
            {activePreview ? (
              <div className="preview-grid">
                <PreviewItem label="档号" value={activePreview.archiveCode} />
                <PreviewItem label="项目名称" value={activePreview.projectName} />
                <PreviewItem label="案卷标题" value={activePreview.volumeTitle} />
                <PreviewItem label="立卷单位" value={activePreview.filingUnit} />
                <PreviewItem label="起止日期" value={activePreview.dateRange} />
                <PreviewItem label="保管期限" value={activePreview.retentionPeriod} />
                <PreviewItem label="文字材料" value={`${activePreview.textPages} 页`} />
                <PreviewItem label="图样" value={`${activePreview.drawingPages} 页`} />
              </div>
            ) : (
              <div className="empty-state detail-empty">读取 Excel 后会显示推导字段</div>
            )}
          </div>

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
            <button className="generate-button" onClick={generate} disabled={!canGenerate || isGenerating}>
              {isGenerating ? <Loader2 className="spin" size={18} /> : <FileText size={18} />}
              生成 DOCX
            </button>
          </div>

          <div className="status-panel">
            <strong>{message || "准备就绪"}</strong>
            {result?.errors.length ? (
              <ul className="error-list">
                {result.errors.map((error) => (
                  <li key={error}>{error}</li>
                ))}
              </ul>
            ) : null}
            {result?.files.length ? (
              <div className="file-list">
                {result.files.slice(0, 8).map((file) => (
                  <span key={file.path}>{file.name}</span>
                ))}
                {result.files.length > 8 ? <span>另有 {result.files.length - 8} 个文件</span> : null}
              </div>
            ) : null}
          </div>
        </section>
      </section>
    </main>
  );
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
