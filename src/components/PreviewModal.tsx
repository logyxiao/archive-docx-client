import { ChevronLeft, ChevronRight, FileText, X } from "lucide-react";
import type { ArchiveRecord } from "../lib/types";
import { PreviewItem } from "./common";

interface PreviewModalProps {
  record: ArchiveRecord;
  previewIndex: number;
  totalCount: number;
  onPrevious: () => void;
  onNext: () => void;
  onClose: () => void;
}

export function PreviewModal({ record, previewIndex, totalCount, onPrevious, onNext, onClose }: PreviewModalProps) {
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <section
        className="preview-modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="preview-title"
        onClick={(event) => event.stopPropagation()}
      >
        <div className="modal-heading">
          <div className="section-heading">
            <FileText size={19} />
            <h2 id="preview-title">案卷字段预览</h2>
          </div>
          <div className="modal-actions">
            <button className="secondary-button modal-nav-button" onClick={onPrevious} disabled={previewIndex <= 0}>
              <ChevronLeft size={17} />
              上一个
            </button>
            <span className="modal-count">
              {previewIndex >= 0 ? previewIndex + 1 : 1} / {totalCount || 1}
            </span>
            <button
              className="secondary-button modal-nav-button"
              onClick={onNext}
              disabled={previewIndex < 0 || previewIndex >= totalCount - 1}
            >
              下一个
              <ChevronRight size={17} />
            </button>
            <button className="icon-button" onClick={onClose} title="关闭" aria-label="关闭字段预览">
              <X size={18} />
            </button>
          </div>
        </div>
        <div className="preview-grid modal-preview-grid">
          {getRecordPreviewFields(record).map((field) => (
            <PreviewItem key={field.label} label={field.label} value={field.value} />
          ))}
        </div>
        <div className="detail-table-wrap">
          <div className="detail-table-title">
            <h3>卷内明细</h3>
            <span>{record.items.length} 条</span>
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
              {record.items.map((item, index) => (
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
