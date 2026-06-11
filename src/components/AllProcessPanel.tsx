import type { Dispatch, SetStateAction } from "react";
import { FolderOpen, FolderTree, Loader2, Search, Settings2, SlidersHorizontal } from "lucide-react";
import { NO_PROCESS_TEMPLATE_KEY } from "../app/appConstants";
import { archiveShortCode, type AllProcessTemplateOption } from "../app/templateOptions";
import type { ProcessTemplateMatch } from "../lib/processDocs";
import type { ArchiveItem, ArchiveRecord } from "../lib/types";

export interface AllProcessRow {
  key: string;
  record: ArchiveRecord;
  item: ArchiveItem;
  matches: ProcessTemplateMatch[];
}

export interface AllProcessGroup {
  record: ArchiveRecord;
  rows: AllProcessRow[];
}

interface AllProcessPanelProps {
  selectedCodesCount: number;
  rowCount: number;
  unresolvedCount: number;
  groups: AllProcessGroup[];
  isLoadingProcessManifest: boolean;
  hasProcessManifest: boolean;
  generationFileCount: number;
  canGenerate: boolean;
  isGenerating: boolean;
  hasOutputDir: boolean;
  allTemplateSearchTerms: Record<string, string>;
  activeAllTemplateRow: string;
  onOpenInfo: () => void;
  onOpenTemplateManager: () => void;
  onOpenBuiltInDir: () => void;
  onOpenUserDir: () => void;
  onGenerate: () => void;
  onOpenOutputDir: () => void;
  setActiveAllTemplateRow: Dispatch<SetStateAction<string>>;
  setAllTemplateSearchTerms: Dispatch<SetStateAction<Record<string, string>>>;
  selectedAllTemplateOption: (rowKey: string, matches?: ProcessTemplateMatch[]) => AllProcessTemplateOption | undefined;
  filteredAllTemplateOptions: (rowKey: string) => AllProcessTemplateOption[];
  selectAllTemplate: (rowKey: string, option: AllProcessTemplateOption) => void;
}

export function AllProcessPanel({
  selectedCodesCount,
  rowCount,
  unresolvedCount,
  groups,
  isLoadingProcessManifest,
  hasProcessManifest,
  generationFileCount,
  canGenerate,
  isGenerating,
  hasOutputDir,
  allTemplateSearchTerms,
  activeAllTemplateRow,
  onOpenInfo,
  onOpenTemplateManager,
  onOpenBuiltInDir,
  onOpenUserDir,
  onGenerate,
  onOpenOutputDir,
  setActiveAllTemplateRow,
  setAllTemplateSearchTerms,
  selectedAllTemplateOption,
  filteredAllTemplateOptions,
  selectAllTemplate,
}: AllProcessPanelProps) {
  return (
    <div className="detail-pane">
      <div className="tool-section all-process-tool">
        <div className="process-template-section">
          <div className="all-template-summary">
            <span>
              已选 {selectedCodesCount} 个档号，{rowCount} 条文件题名
              {unresolvedCount > 0 ? `，${unresolvedCount} 条需选择模板` : ""}
            </span>
            <div className="template-summary-actions">
              <button className="template-dir-button" type="button" onClick={onOpenInfo}>
                <SlidersHorizontal size={15} />
                填写生成信息
              </button>
              <button type="button" className="template-dir-button" onClick={onOpenTemplateManager}>
                <Settings2 size={15} />
                模板管理
              </button>
              <button type="button" className="template-dir-button" onClick={onOpenBuiltInDir}>
                <FolderOpen size={15} />
                内置目录
              </button>
              <button type="button" className="template-dir-button" onClick={onOpenUserDir}>
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
          ) : selectedCodesCount === 0 ? (
            <div className="empty-state detail-empty all-process-empty">请选择档号</div>
          ) : rowCount === 0 ? (
            <div className="empty-state detail-empty">未读取到文件题名</div>
          ) : (
            groups.map((group) => (
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
          <button className="generate-button" onClick={onGenerate} disabled={!canGenerate}>
            {isGenerating ? <Loader2 className="spin" size={18} /> : <FolderTree size={18} />}
            生成 {selectedCodesCount > 0 && hasProcessManifest ? `(${generationFileCount}个文件)` : ""}
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
