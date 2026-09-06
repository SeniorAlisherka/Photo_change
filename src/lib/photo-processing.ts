import { BlobReader, BlobWriter, TextReader, ZipWriter } from "@zip.js/zip.js";
import { read, utils } from "xlsx";

import { convertPhoto } from "./photo-image";
import { parsePhotoRows } from "./photo-mapping";
import {
  MANUAL_FOLDER,
  OUTPUT_FOLDER,
  PHOTO_FOLDER,
  REPORT_NAME,
  SIZE_LIMIT,
  UNKNOWN_FOLDER,
  WORKBOOK_NAME,
  PhotoProcessingError,
  processingErrorMessage,
  type PhotoInput,
  type PhotoSummary,
  type ConversionProgress,
  type ConversionReport,
  type PhotoResult,
  type PhotoRow,
} from "./photo-converter-types";

export type PhotoBatch = {
  workbook: File;
  photos: Map<string, PhotoInput>;
  rows: PhotoRow[];
  summary: PhotoSummary;
};

export async function inspectPhotos(
  workbookFile: File,
  inputs: PhotoInput[],
): Promise<PhotoBatch> {
  if (!workbookFile.name.toLowerCase().endsWith(".xlsx"))
    throw new PhotoProcessingError("Выберите таблицу Excel в формате .xlsx.");
  const photos = new Map<string, PhotoInput>();
  const names = new Set<string>();
  for (const input of inputs) {
    const path = input.path.replace(/\\/g, "/").normalize("NFC");
    const parts = path.split("/");
    if (parts.some((part) => part.startsWith(".") || part === "__MACOSX"))
      continue;
    if (!/\.(jpe?g|png|heic)$/i.test(path)) continue;
    if (
      path.startsWith("/") ||
      /^[a-z]:/i.test(path) ||
      /[\u0000-\u001f]/.test(path)
    )
      throw new PhotoProcessingError(
        "В папке обнаружен недопустимый путь к файлу.",
      );
    const name = parts.at(-1)!;
    if (names.has(name.toLowerCase()))
      throw new PhotoProcessingError(
        `В выбранной папке несколько фотографий с именем «${name}». Выберите только папку с исходными фото или переименуйте совпадающие файлы.`,
      );
    names.add(name.toLowerCase());
    photos.set(name, { file: input.file, path });
  }
  if (!photos.size)
    throw new PhotoProcessingError(
      "В выбранной папке нет фотографий JPG, JPEG, PNG или HEIC. Выберите папку с исходными фото.",
    );
  const workbook = read(await workbookFile.arrayBuffer(), { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) throw new PhotoProcessingError("В таблице нет листов.");
  const rows = parsePhotoRows(
    utils.sheet_to_json<unknown[]>(sheet, {
      header: 1,
      defval: "",
      raw: false,
      blankrows: true,
    }),
  );
  const referenced = new Set(rows.map((row) => row.source));
  const summary: PhotoSummary = {
    photos: photos.size,
    rows: rows.length,
    matched: rows.filter((row) => !row.error && photos.has(row.source)).length,
    unmatched: [...photos.keys()].filter((name) => !referenced.has(name))
      .length,
    issues: rows.filter((row) => row.error || !photos.has(row.source)).length,
  };
  return { workbook: workbookFile, photos, rows, summary };
}

export async function convertPhotos(
  batch: PhotoBatch,
  onProgress: (progress: ConversionProgress) => void,
): Promise<{ blob: Blob; report: ConversionReport }> {
  const started = performance.now();
  const { rows, photos } = batch;
  const writer = new ZipWriter(new BlobWriter("application/zip"), {
    useWebWorkers: false,
    level: 0,
  });
  const results: PhotoResult[] = [];
  const convertedSources = new Set<string>();
  const referenced = new Set(rows.map((row) => row.source));

  for (const folder of [
    PHOTO_FOLDER,
    OUTPUT_FOLDER,
    UNKNOWN_FOLDER,
    MANUAL_FOLDER,
  ])
    await writer.add(`${folder}/`, undefined, { directory: true });

  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    onProgress({
      phase: "converting",
      completed: index,
      total: rows.length,
      filename: row.source,
    });
    const photo = photos.get(row.source);
    if (row.error) {
      results.push({ ...row, status: "skipped", message: row.error });
      continue;
    }
    if (!photo) {
      results.push({
        ...row,
        status: "missing",
        message: "Фотография не найдена в выбранной папке.",
      });
      continue;
    }

    let jpeg: Blob;
    try {
      jpeg = await convertPhoto(photo.file, row.source);
    } catch (error) {
      results.push({
        ...row,
        status: "manual",
        message: processingErrorMessage(
          error,
          "Не удалось обработать фотографию. Возможно, файл повреждён.",
        ),
      });
      continue;
    }

    // A packaging failure aborts the download instead of returning an incomplete ZIP.
    await writer.add(`${OUTPUT_FOLDER}/${row.output}`, new BlobReader(jpeg));
    convertedSources.add(row.source);
    results.push({
      ...row,
      status: "converted",
      message: "Фотография обработана и переименована.",
      size: jpeg.size,
    });
  }

  let manual = 0;
  let unmatched = 0;
  const copies: PhotoInput[] = [{ path: WORKBOOK_NAME, file: batch.workbook }];
  for (const [name, photo] of photos) {
    copies.push({ path: `${PHOTO_FOLDER}/${photo.path}`, file: photo.file });
    if (convertedSources.has(name)) continue;
    if (referenced.has(name)) {
      manual += 1;
      copies.push({ path: `${MANUAL_FOLDER}/${name}`, file: photo.file });
    } else {
      unmatched += 1;
      copies.push({ path: `${UNKNOWN_FOLDER}/${name}`, file: photo.file });
      results.push({
        source: name,
        status: "unmatched",
        message: "Фотографии нет в таблице.",
      });
    }
  }

  for (let index = 0; index < copies.length; index += 1) {
    const item = copies[index];
    onProgress({
      phase: "packing",
      completed: index,
      total: copies.length,
      filename: item.path.split("/").at(-1) ?? "",
    });
    await writer.add(item.path, new BlobReader(item.file));
  }

  const report: ConversionReport = {
    converted: convertedSources.size,
    manual,
    unmatched,
    missing: results.filter((result) => result.status === "missing").length,
    issues: results.filter((result) =>
      ["manual", "missing", "skipped"].includes(result.status),
    ).length,
    elapsedMs: Math.round(performance.now() - started),
    results,
  };
  await writer.add(
    REPORT_NAME,
    new TextReader(
      JSON.stringify({ sizeLimitBytes: SIZE_LIMIT, ...report }, null, 2),
    ),
  );
  return { blob: await writer.close(), report };
}
