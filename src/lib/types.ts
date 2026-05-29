export interface ArchiveItem {
  sequence: string;
  fileCode: string;
  owner: string;
  title: string;
  fileDate: string;
  pageNo: string;
  note: string;
}

export interface ArchiveRecord {
  categoryCode: string;
  archiveCode: string;
  fullTitle: string;
  projectName: string;
  volumeTitle: string;
  owner: string;
  filingUnit: string;
  retentionPeriod: string;
  startDate: string;
  endDate: string;
  dateRange: string;
  totalPages: number;
  drawingPages: number;
  textPages: number;
  items: ArchiveItem[];
}

export interface GenerationOptions {
  selectedCodes: string[];
  backupNote: string;
  outputDir: string;
  generateCover: boolean;
  generateNote: boolean;
  generateSpine: boolean;
}

export interface GeneratedFile {
  name: string;
  path: string;
}

export interface GenerationResult {
  files: GeneratedFile[];
  errors: string[];
}
