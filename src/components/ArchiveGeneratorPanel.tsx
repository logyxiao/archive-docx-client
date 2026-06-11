import { FileText, FolderOpen, Loader2, Settings2 } from "lucide-react";
import { Toggle } from "./common";

interface ArchiveGeneratorPanelProps {
  backupNote: string;
  generateCover: boolean;
  generateNote: boolean;
  generateSpine: boolean;
  generateCatalogWorkbook: boolean;
  selectedCodesCount: number;
  generationFileCount: number;
  canGenerate: boolean | string;
  isGenerating: boolean;
  hasOutputDir: boolean;
  onBackupNoteChange: (value: string) => void;
  onGenerateCoverChange: (checked: boolean) => void;
  onGenerateNoteChange: (checked: boolean) => void;
  onGenerateSpineChange: (checked: boolean) => void;
  onGenerateCatalogWorkbookChange: (checked: boolean) => void;
  onGenerate: () => void;
  onOpenOutputDir: () => void;
}

export function ArchiveGeneratorPanel({
  backupNote,
  generateCover,
  generateNote,
  generateSpine,
  generateCatalogWorkbook,
  selectedCodesCount,
  generationFileCount,
  canGenerate,
  isGenerating,
  hasOutputDir,
  onBackupNoteChange,
  onGenerateCoverChange,
  onGenerateNoteChange,
  onGenerateSpineChange,
  onGenerateCatalogWorkbookChange,
  onGenerate,
  onOpenOutputDir,
}: ArchiveGeneratorPanelProps) {
  return (
    <div className="detail-pane">
      <div className="tool-section">
        <div className="section-heading">
          <Settings2 size={19} />
          <h2>生成设置</h2>
        </div>
        <div className="switch-row">
          <Toggle checked={generateCover} onChange={onGenerateCoverChange} label="案卷大封面" />
          <Toggle checked={generateNote} onChange={onGenerateNoteChange} label="备考表" />
          <Toggle checked={generateSpine} onChange={onGenerateSpineChange} label="案卷脊背" />
          <Toggle checked={generateCatalogWorkbook} onChange={onGenerateCatalogWorkbookChange} label="目录台账" />
        </div>
        <label className="note-field">
          <span>备考表其它情况</span>
          <textarea
            value={backupNote}
            onChange={(event) => onBackupNoteChange(event.currentTarget.value)}
            placeholder="不填写则生成空白说明"
          />
        </label>
        <div className="generation-actions">
          <button className="generate-button" onClick={onGenerate} disabled={!canGenerate || isGenerating}>
            {isGenerating ? <Loader2 className="spin" size={18} /> : <FileText size={18} />}
            生成文件 {selectedCodesCount > 0 ? `(${generationFileCount}个文件)` : ""}
          </button>
          <button className="secondary-button open-output-button" onClick={onOpenOutputDir} disabled={!hasOutputDir}>
            <FolderOpen size={18} />
            打开输出目录
          </button>
        </div>
      </div>
    </div>
  );
}
