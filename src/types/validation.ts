export interface ValidationError {
  id: string;
  sheet: string;
  row: number;
  column: string;
  field: string;
  message: string;
  severity: 'error' | 'warning' | 'info';
  rule: string;
}

export interface ValidationSummary {
  totalErrors: number;
  totalWarnings: number;
  errorsBySheet: Record<string, number>;
  warningsBySheet: Record<string, number>;
}

export interface SheetData {
  name: string;
  data: any[][];
  headers: string[];
}

export interface ParsedFile {
  fileName: string;
  sheets: SheetData[];
  isValid: boolean;
  validationErrors: ValidationError[];
  validationSummary: ValidationSummary;
}

export interface FileUploadState {
  file: File | null;
  googleSheetUrl: string;
  isProcessing: boolean;
  parsedData: ParsedFile | null;
  error: string | null;
}