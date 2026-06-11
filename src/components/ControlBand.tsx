import { FileSpreadsheet, FolderOpen, Loader2 } from "lucide-react";
import { PathLine } from "./common";

interface ControlBandProps {
  excelPath: string;
  outputPath: string;
  isLoading: boolean;
  onChooseExcel: () => void;
  onChooseOutputDir: () => void;
}

export function ControlBand({
  excelPath,
  outputPath,
  isLoading,
  onChooseExcel,
  onChooseOutputDir,
}: ControlBandProps) {
  return (
    <section className="control-band">
      <button className="primary-button" onClick={onChooseExcel} disabled={isLoading}>
        {isLoading ? <Loader2 className="spin" size={18} /> : <FileSpreadsheet size={18} />}
        选择 Excel
      </button>
      <button className="secondary-button" onClick={onChooseOutputDir}>
        <FolderOpen size={18} />
        输出目录
      </button>
      <div className="path-stack">
        <PathLine label="Excel" value={excelPath || "未选择"} />
        <PathLine label="输出" value={outputPath || "未选择"} />
      </div>
    </section>
  );
}
