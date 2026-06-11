import { FileText, FolderTree } from "lucide-react";
import {
  ALL_PROCESS_DOCS_TAB,
  ARCHIVE_DOCX_TAB,
  COLLECTOR_LINE_TAB,
  PROCESS_DOCS_TAB,
  SWITCH_STATION_TAB,
} from "../app/appConstants";

interface AppHeaderProps {
  activeTab: string;
  onTabChange: (tab: string) => void;
}

export function AppHeader({ activeTab, onTabChange }: AppHeaderProps) {
  return (
    <header className="app-header">
      <div className="app-logo-area">
        <FileText className="logo-icon" size={20} />
        <h1 className="app-title">Docx 归档助手</h1>
      </div>
      <nav className="app-tabs" aria-label="功能栏目">
        <button
          className={`tab-button ${activeTab === ARCHIVE_DOCX_TAB ? "active" : ""}`}
          onClick={() => onTabChange(ARCHIVE_DOCX_TAB)}
          type="button"
        >
          <FileText size={17} />
          档案文档生成器
        </button>
        <button
          className={`tab-button ${activeTab === PROCESS_DOCS_TAB ? "active" : ""}`}
          onClick={() => onTabChange(PROCESS_DOCS_TAB)}
          type="button"
        >
          <FolderTree size={17} />
          过程资料生成
        </button>
        <button
          className={`tab-button ${activeTab === ALL_PROCESS_DOCS_TAB ? "active" : ""}`}
          onClick={() => onTabChange(ALL_PROCESS_DOCS_TAB)}
          type="button"
        >
          <FolderTree size={17} />
          全部过程资料
        </button>
        <button
          className={`tab-button ${activeTab === SWITCH_STATION_TAB ? "active" : ""}`}
          onClick={() => onTabChange(SWITCH_STATION_TAB)}
          type="button"
        >
          <FolderTree size={17} />
          开关站电气设备安装
        </button>
        <button
          className={`tab-button ${activeTab === COLLECTOR_LINE_TAB ? "active" : ""}`}
          onClick={() => onTabChange(COLLECTOR_LINE_TAB)}
          type="button"
        >
          <FolderTree size={17} />
          集电线路安装工程
        </button>
      </nav>
    </header>
  );
}
