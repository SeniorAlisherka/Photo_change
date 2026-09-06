import {
  LINK_COLUMN,
  NAME_COLUMN,
  PhotoProcessingError,
  processingErrorMessage,
  type PhotoRow,
} from "./photo-converter-types";

export function sourceFilename(value: string) {
  let path = value.trim();

  // Links identify a local file; they are never fetched.
  if (/^https?:\/\//i.test(path)) {
    const url = new URL(path);
    // Yandex Forms/Disk store the real path in the query string. URLSearchParams
    // has already decoded that value; ordinary URL paths need one decode.
    path =
      url.searchParams.get("path") ??
      url.searchParams.get("idDialog") ??
      decodeURIComponent(url.pathname);
  } else {
    path = decodeURIComponent(path);
  }

  const filename = path.replace(/\\/g, "/").split("/").at(-1) ?? "";
  const decoded = filename.normalize("NFC");

  if (!decoded || /[/\\\u0000-\u001f]/.test(decoded)) {
    throw new PhotoProcessingError("В ссылке не найдено корректное имя файла.");
  }

  return decoded;
}

export function outputFilename(value: string) {
  const name = value.trim().normalize("NFC");

  if (
    !name ||
    name === "." ||
    name === ".." ||
    /[<>:"/\\|?*\u0000-\u001f]/.test(name) ||
    /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name)
  ) {
    throw new PhotoProcessingError(
      "Новое имя файла не указано или содержит недопустимые символы.",
    );
  }

  // The spreadsheet supplies the stem, just as in the Python script.
  const output = `${name}.jpg`;
  if (new TextEncoder().encode(output).length > 255) {
    throw new PhotoProcessingError("Новое имя файла слишком длинное.");
  }
  return output;
}

export function parsePhotoRows(data: unknown[][]): PhotoRow[] {
  const headers = (data[0] ?? []).map((value) => String(value ?? "").trim());
  const linkIndex = headers.indexOf(LINK_COLUMN);
  const nameIndex = headers.indexOf(NAME_COLUMN);

  if (linkIndex === -1 || nameIndex === -1) {
    throw new PhotoProcessingError(
      `На первом листе должны быть столбцы «${LINK_COLUMN}» и «${NAME_COLUMN}».`,
    );
  }

  const sources = new Set<string>();
  const outputs = new Set<string>();
  const rows: PhotoRow[] = [];

  for (let index = 1; index < data.length; index += 1) {
    const link = String(data[index][linkIndex] ?? "").trim();
    const name = String(data[index][nameIndex] ?? "").trim();
    if (!link && !name) continue;

    const row: PhotoRow = { row: index + 1, source: "", output: "" };
    try {
      row.source = sourceFilename(link);
      row.output = outputFilename(name);
      if (sources.has(row.source)) {
        throw new PhotoProcessingError(
          "Эта фотография уже указана в предыдущей строке.",
        );
      }
      if (outputs.has(row.output.toLowerCase())) {
        throw new PhotoProcessingError(
          "Такое новое имя файла уже указано в предыдущей строке.",
        );
      }
      sources.add(row.source);
      outputs.add(row.output.toLowerCase());
    } catch (error) {
      row.error = processingErrorMessage(
        error,
        "В строке таблицы указаны некорректные данные или ссылка.",
      );
    }
    rows.push(row);
  }

  if (!rows.length)
    throw new PhotoProcessingError(
      "На первом листе нет списка фотографий для обработки.",
    );
  return rows;
}
