import { Check, Info, Search } from "lucide-react";
import type { ArchiveRecord } from "../lib/types";

interface RecordPaneProps {
  recordsCount: number;
  filteredRecords: ArchiveRecord[];
  selectedCodes: string[];
  query: string;
  showOnlySelected: boolean;
  onQueryChange: (value: string) => void;
  onShowOnlySelectedChange: (checked: boolean) => void;
  onToggleVisibleRecords: (checked: boolean) => void;
  onToggleCode: (code: string) => void;
  onPreviewRecord: (record: ArchiveRecord) => void;
}

export function RecordPane({
  recordsCount,
  filteredRecords,
  selectedCodes,
  query,
  showOnlySelected,
  onQueryChange,
  onShowOnlySelectedChange,
  onToggleVisibleRecords,
  onToggleCode,
  onPreviewRecord,
}: RecordPaneProps) {
  return (
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
            onChange={(event) => onToggleVisibleRecords(event.currentTarget.checked)}
          />
          全选
        </label>
      </div>
      <div className="record-pane-filter">
        <label className="search-box">
          <Search size={17} />
          <input value={query} onChange={(event) => onQueryChange(event.currentTarget.value)} placeholder="搜索档号、题名、单位" />
        </label>
        {selectedCodes.length > 0 ? (
          <label className="only-selected-toggle">
            <input
              type="checkbox"
              checked={showOnlySelected}
              onChange={(event) => onShowOnlySelectedChange(event.currentTarget.checked)}
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
            <button className="record-toggle" onClick={() => onToggleCode(record.archiveCode)}>
              <span className="checkmark">{selectedCodes.includes(record.archiveCode) ? <Check size={14} /> : null}</span>
              <span>
                <strong>{record.archiveCode}</strong>
                <small>{record.fullTitle}</small>
              </span>
            </button>
            <button className="icon-button" onClick={() => onPreviewRecord(record)} title="案卷详情" aria-label={`${record.archiveCode} 案卷详情`}>
              <Info size={17} />
            </button>
          </div>
        ))}
        {recordsCount === 0 ? <div className="empty-state">请选择 Excel 总目录</div> : null}
      </div>
    </aside>
  );
}
