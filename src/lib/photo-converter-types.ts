export const PHOTO_FOLDER = "Фото до обработки";
export const OUTPUT_FOLDER = "Фото после обработки";
export const UNKNOWN_FOLDER = "Непонятные фото";
export const MANUAL_FOLDER = "Фото для ручной обработки";
export const WORKBOOK_NAME = "Список.xlsx";
export const LINK_COLUMN = "Ссылка, содержащая название файла до обработки";
export const NAME_COLUMN = "Название файла после обработки";
export const SIZE_LIMIT = 500 * 1024;
export const REPORT_NAME = "processing-report.json";

export type PhotoRow = {
  row: number;
  source: string;
  output: string;
  error?: string;
};

export type PhotoInput = { file: File; path: string };

export type PhotoFolder = {
  name: string;
  files: PhotoInput[];
  size: number;
};

export type PhotoSummary = {
  photos: number;
  rows: number;
  matched: number;
  unmatched: number;
  issues: number;
};

export type PhotoResult = {
  row?: number;
  source: string;
  output?: string;
  status: "converted" | "manual" | "missing" | "unmatched" | "skipped";
  message: string;
  size?: number;
};

export type ConversionProgress = {
  phase: "converting" | "packing";
  completed: number;
  total: number;
  filename: string;
};

export type ConversionReport = {
  converted: number;
  manual: number;
  unmatched: number;
  missing: number;
  issues: number;
  elapsedMs: number;
  results: PhotoResult[];
};

export type WorkerRequest =
  | { type: "inspect"; workbook: File; photos: PhotoInput[] }
  | { type: "convert" };

export type WorkerResponse =
  | { type: "ready"; summary: PhotoSummary }
  | { type: "progress"; progress: ConversionProgress }
  | { type: "complete"; blob: Blob; report: ConversionReport }
  | { type: "error"; message: string };

export class PhotoProcessingError extends Error {}

export function processingErrorMessage(error: unknown, fallback: string) {
  return error instanceof PhotoProcessingError ? error.message : fallback;
}

export function formatDecimal(value: number) {
  return value.toLocaleString("ru-RU", {
    minimumFractionDigits: 1,
    maximumFractionDigits: 1,
  });
}

export function formatBytes(bytes: number) {
  if (bytes < 1024) return `${bytes} Б`;
  if (bytes < 1024 * 1024) return `${formatDecimal(bytes / 1024)} КиБ`;
  return `${formatDecimal(bytes / (1024 * 1024))} МиБ`;
}
