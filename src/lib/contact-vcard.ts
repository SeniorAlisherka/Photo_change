import { read, utils, type CellObject, type WorkBook } from "xlsx";

import {
  CONTACT_COLUMNS,
  ContactProcessingError,
  type ContactConversion,
  type ContactRecord,
} from "./contact-converter-types";

const encoder = new TextEncoder();

function escapeText(value: string) {
  return value
    .replace(/\\/g, "\\\\")
    .replace(/\r\n|\r|\n/g, "\\n")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,");
}

function foldLine(value: string) {
  const lines: string[] = [];
  let line = "";
  let bytes = 0;

  // Iterate Unicode code points so a fold cannot split a UTF-8 character.
  for (const character of value) {
    const length = encoder.encode(character).length;
    if (bytes + length > 75) {
      lines.push(line);
      line = " ";
      bytes = 1;
    }
    line += character;
    bytes += length;
  }
  lines.push(line);
  return lines.join("\r\n");
}

function contactVcard(contact: ContactRecord) {
  // RFC 2426: N is family;given;additional;prefix;suffix, regardless of FN.
  return (
    [
      "BEGIN:VCARD",
      "VERSION:3.0",
      `N:${escapeText(contact.familyName)};${escapeText(contact.givenName)};;;`,
      `FN:${escapeText(contact.formattedName)}`,
      ...(contact.phone ? [`TEL;TYPE=CELL:${escapeText(contact.phone)}`] : []),
      ...(contact.email
        ? [`EMAIL;TYPE=INTERNET:${escapeText(contact.email)}`]
        : []),
      "END:VCARD",
    ]
      .map(foldLine)
      .join("\r\n") + "\r\n"
  );
}

function cellText(cell: CellObject | undefined) {
  return cell ? utils.format_cell(cell).trim() : "";
}

function readWorkbook(buffer: ArrayBuffer): WorkBook {
  if (!buffer.byteLength) {
    throw new ContactProcessingError(
      "Выбранный файл пуст. Выберите таблицу Excel с контактами.",
    );
  }

  try {
    const signature = new Uint8Array(buffer, 0, Math.min(4, buffer.byteLength));
    if (
      signature[0] !== 0x50 ||
      signature[1] !== 0x4b ||
      signature[2] !== 0x03 ||
      signature[3] !== 0x04
    ) {
      throw new Error("Not an XLSX archive");
    }
    return read(buffer, { type: "array", sheets: 0, cellText: true });
  } catch {
    throw new ContactProcessingError(
      "Не удалось прочитать таблицу. Проверьте, что файл открывается в Excel, и сохраните его в формате .xlsx без пароля.",
    );
  }
}

export function convertContactWorkbook(buffer: ArrayBuffer): ContactConversion {
  const workbook = readWorkbook(buffer);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) {
    throw new ContactProcessingError("В таблице нет листов с данными.");
  }
  if (!sheet["!ref"]) {
    throw new ContactProcessingError(
      "Первый лист таблицы пуст. Добавьте названия столбцов и контакты.",
    );
  }

  const range = utils.decode_range(sheet["!ref"]);
  const headers = Array.from({ length: range.e.c + 1 }, (_, column) =>
    cellText(sheet[utils.encode_cell({ r: 0, c: column })]),
  );
  const missing = CONTACT_COLUMNS.filter((column) => !headers.includes(column));
  if (missing.length) {
    throw new ContactProcessingError(
      `В первой строке первого листа не найдены столбцы: ${missing.map((column) => `«${column}»`).join(", ")}.`,
    );
  }
  const duplicate = CONTACT_COLUMNS.find(
    (column) => headers.indexOf(column) !== headers.lastIndexOf(column),
  );
  if (duplicate) {
    throw new ContactProcessingError(
      `Столбец «${duplicate}» повторяется. Оставьте один столбец с этим названием.`,
    );
  }

  const columns = CONTACT_COLUMNS.map((column) => headers.indexOf(column));
  const result: ContactConversion = { contacts: [], issues: [], vcf: "" };

  for (let index = 1; index <= range.e.r; index += 1) {
    const row = Array.from(
      { length: range.e.c + 1 },
      (_, column) =>
        sheet[utils.encode_cell({ r: index, c: column })] as
          CellObject | undefined,
    );
    if (!row.some((cell) => cellText(cell) || cell?.f || cell?.t === "e"))
      continue;

    const cells = columns.map((column) => row[column]);
    const values = cells.map(cellText);
    // Excel may display a long phone number in scientific notation. Never
    // invent a country prefix or try to restore already-lost leading zeroes.
    if (cells[3]?.t === "n") {
      const number = cells[3].v;
      if (
        typeof number === "number" &&
        Number.isSafeInteger(number) &&
        number >= 0
      ) {
        const displayed = values[3].replace(/[\s()\-]/g, "");
        const displayedDigits = displayed.replace(/^\+?0*(?=\d)/, "");
        values[3] =
          /^\+?\d+$/.test(displayed) && displayedDigits === String(number)
            ? displayed
            : String(number);
      } else {
        result.issues.push({
          row: index + 1,
          message:
            "Проверьте номер телефона: сохраните его в Excel как текст, чтобы цифры не терялись.",
        });
        continue;
      }
    }
    const invalid = cells.findIndex(
      (cell, column) =>
        cell?.t === "e" ||
        cell?.t === "b" ||
        (cell?.f && cell.v === undefined) ||
        /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(values[column]),
    );
    const empty = values
      .slice(0, 3)
      .flatMap((value, column) => (value ? [] : [CONTACT_COLUMNS[column]]));

    if (invalid !== -1) {
      result.issues.push({
        row: index + 1,
        message: `Проверьте значение в столбце «${CONTACT_COLUMNS[invalid]}»: ячейка содержит ошибку или недопустимые данные.`,
      });
    } else if (empty.length) {
      result.issues.push({
        row: index + 1,
        message: `Заполните поля: ${empty.map((column) => `«${column}»`).join(", ")}.`,
      });
    } else {
      const [givenName, program, year, phoneText, email] = values;
      const phone = phoneText.replace(/[\s()\-]/g, "");
      if (phoneText && !/^\+?\d+$/.test(phone)) {
        result.issues.push({
          row: index + 1,
          message:
            "Проверьте номер телефона. Укажите один номер: цифры и при необходимости «+» в начале; пробелы, скобки и дефисы допустимы.",
        });
        continue;
      }
      if (email && !/^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/.test(email)) {
        result.issues.push({
          row: index + 1,
          message:
            "Проверьте почту. Укажите один адрес в формате name@example.com.",
        });
        continue;
      }
      const familyName = `${program} ${year}`;
      result.contacts.push({
        row: index + 1,
        givenName,
        familyName,
        formattedName: `${givenName} ${familyName}`,
        ...(phone ? { phone } : {}),
        ...(email ? { email } : {}),
      });
    }
  }

  if (!result.contacts.length && !result.issues.length) {
    throw new ContactProcessingError(
      "На первом листе нет контактов. Добавьте данные под названиями столбцов.",
    );
  }

  result.vcf = result.contacts.map(contactVcard).join("");
  return result;
}
