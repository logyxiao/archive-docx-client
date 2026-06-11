import { FolderOpen, FolderTree, Loader2, SlidersHorizontal } from "lucide-react";
import {
  PROCESS_TEMPLATE_CATEGORIES,
  PROCESS_TEMPLATE_CATEGORY_IDS,
  type ProcessTemplateCategoryId,
} from "../lib/processDocs";

interface ProcessGenerationPanelProps {
  title: string;
  generateLabel: string;
  selectedCodesCount: number;
  generationFileCount: number;
  hasProcessManifest: boolean;
  selectedCategories: ProcessTemplateCategoryId[];
  canGenerate: boolean;
  isGenerating: boolean;
  hasOutputDir: boolean;
  onOpenInfo: () => void;
  onSelectAllCategories: () => void;
  onClearCategories: () => void;
  onToggleCategory: (categoryId: ProcessTemplateCategoryId) => void;
  onGenerate: () => void;
  onOpenOutputDir: () => void;
}

export function ProcessGenerationPanel({
  title,
  generateLabel,
  selectedCodesCount,
  generationFileCount,
  hasProcessManifest,
  selectedCategories,
  canGenerate,
  isGenerating,
  hasOutputDir,
  onOpenInfo,
  onSelectAllCategories,
  onClearCategories,
  onToggleCategory,
  onGenerate,
  onOpenOutputDir,
}: ProcessGenerationPanelProps) {
  return (
    <div className="detail-pane">
      <div className="tool-section">
        <div className="section-heading">
          <FolderTree size={19} />
          <h2>{title}</h2>
        </div>
        <div className="process-info-actions">
          <button className="secondary-button process-info-button" type="button" onClick={onOpenInfo}>
            <SlidersHorizontal size={17} />
            填写生成信息
          </button>
        </div>
        <div className="process-template-section">
          <div className="template-section-heading">
            <span>生成模板</span>
            <div className="mini-actions">
              <button type="button" onClick={onSelectAllCategories}>
                全选
              </button>
              <button type="button" onClick={onClearCategories}>
                清空
              </button>
            </div>
          </div>
          <div className="process-template-grid">
            {PROCESS_TEMPLATE_CATEGORIES.map((category) => (
              <label key={category.id} className="template-check">
                <input
                  type="checkbox"
                  checked={selectedCategories.includes(category.id)}
                  onChange={() => onToggleCategory(category.id)}
                />
                <span>{category.label}</span>
              </label>
            ))}
          </div>
        </div>
        <div className="generation-actions">
          <button className="generate-button" onClick={onGenerate} disabled={!canGenerate}>
            {isGenerating ? <Loader2 className="spin" size={18} /> : <FolderTree size={18} />}
            {generateLabel} {selectedCodesCount > 0 && hasProcessManifest ? `(${generationFileCount}个文件)` : ""}
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

export const ALL_PROCESS_TEMPLATE_CATEGORIES = PROCESS_TEMPLATE_CATEGORY_IDS;
