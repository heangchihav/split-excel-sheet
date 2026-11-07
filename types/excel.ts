import type * as XLSX from "xlsx";

export type NamingMode = "original" | "sequence" | "sequence_original";

export type ZeroPadding = "none" | "2" | "3" | "auto";

export type Separator = "_" | "-" | " " | "";

export interface GlobalConfig {
  prefix: string;
  namingMode: NamingMode;
  startNumber: number;
  zeroPadding: ZeroPadding;
  separator: Separator;
  firstSheetIsInfo: boolean;
  infoFileName?: string;
  masterFolderName?: string;
}

export interface SheetItem {
  id: string;
  originalName: string;
  customName?: string;
  selected: boolean;
  rowCount: number;
  colCount: number;
  sampleHeaders?: string[];
  sampleRows?: string[][];
}

export interface UploadedWorkbook {
  id: string;
  fileName: string;
  fileSize: number;
  folderName: string;
  sheets: SheetItem[];
  rawWorkbook: XLSX.WorkBook;
  rawBuffer?: ArrayBuffer;
  rawFile?: File;
  isBinaryXls?: boolean;
  // Per-workbook override settings
  useCustomConfig: boolean;
  customPrefix?: string;
  customNamingMode?: NamingMode;
  customStartNumber?: number;
  customZeroPadding?: ZeroPadding;
  customSeparator?: Separator;
  customFirstSheetIsInfo?: boolean;
  customInfoFileName?: string;
  status: "ready" | "processing" | "done" | "error";
  errorMessage?: string;
}

export interface ResolvedFileItem {
  sheetIndex: number;
  originalName: string;
  resolvedFileName: string;
  folderName: string;
  selected: boolean;
}

