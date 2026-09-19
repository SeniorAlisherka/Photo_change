import { expect, test } from "@playwright/test";
import { utils, write, type WorkSheet } from "xlsx";

import { CONTACT_COLUMNS } from "../src/lib/contact-converter-types";
import { convertContactWorkbook } from "../src/lib/contact-vcard";

function workbookBuffer(rows: unknown[][], edit?: (sheet: WorkSheet) => void) {
  const workbook = utils.book_new();
  const sheet = utils.aoa_to_sheet(rows);
  edit?.(sheet);
  utils.book_append_sheet(workbook, sheet, "Контакты");
  utils.book_append_sheet(
    workbook,
    utils.aoa_to_sheet([["Игнорировать второй лист"]]),
    "Другой лист",
  );
  return write(workbook, { type: "array", bookType: "xlsx" }) as ArrayBuffer;
}

function unfold(vcf: string) {
  return vcf.replace(/\r\n[ \t]/g, "");
}

test("exports full ФИО as given name and program/year as family name", () => {
  const result = convertContactWorkbook(
    workbookBuffer([
      [
        " Год обучения ",
        "\uFEFFФИО",
        " Короткое название программы ",
        "Номер телефона",
        "Почта",
      ],
      [
        "2025-2026",
        "Иванов Иван Иванович",
        "ОФД",
        "+1 (202) 555-0100",
        "ivan@example.com",
      ],
      ["2026-2027", "Петров Пётр Петрович", "ТЕСТ"],
    ]),
  );
  expect(result.issues).toEqual([]);
  expect(result.contacts[0]).toEqual({
    row: 2,
    givenName: "Иванов Иван Иванович",
    familyName: "ОФД 2025-2026",
    formattedName: "Иванов Иван Иванович ОФД 2025-2026",
    phone: "+12025550100",
    email: "ivan@example.com",
  });
  const vcf = unfold(result.vcf);
  expect(vcf).toContain("N:ОФД 2025-2026;Иванов Иван Иванович;;;\r\n");
  expect(vcf).toContain("FN:Иванов Иван Иванович ОФД 2025-2026\r\n");
  expect(vcf.match(/BEGIN:VCARD/g)).toHaveLength(2);
  expect(vcf.match(/VERSION:3.0/g)).toHaveLength(2);
  expect(vcf.endsWith("END:VCARD\r\n")).toBe(true);
  expect(vcf).toContain("TEL;TYPE=CELL:+12025550100\r\n");
  expect(vcf).toContain("EMAIL;TYPE=INTERNET:ivan@example.com\r\n");
  expect(vcf.match(/^TEL;/gm)).toHaveLength(1);
  expect(vcf.match(/^EMAIL;/gm)).toHaveLength(1);
  expect(vcf.replace(/\r\n/g, "")).not.toMatch(/[\r\n]/);
});

test("escapes field separators, backslashes and line breaks", () => {
  const result = convertContactWorkbook(
    workbookBuffer([
      [...CONTACT_COLUMNS],
      ["Иванов, Иван; Иванович\\тест\r\nНовая строка", "О;Ф,Д\\", "2025-2026"],
    ]),
  );
  const vcf = unfold(result.vcf);
  expect(vcf).toContain(
    "N:О\\;Ф\\,Д\\\\ 2025-2026;Иванов\\, Иван\\; Иванович\\\\тест\\nНовая строка;;;\r\n",
  );
  expect(vcf.split("\r\n")).toHaveLength(6);
});

test("folds Unicode lines to at most 75 UTF-8 bytes without corrupting names", () => {
  const givenName = "Ая😀é".repeat(35);
  const result = convertContactWorkbook(
    workbookBuffer([[...CONTACT_COLUMNS], [givenName, "ОФД", "2025-2026"]]),
  );
  const encoder = new TextEncoder();
  const decoder = new TextDecoder("utf-8", { fatal: true });
  for (const line of result.vcf.split("\r\n")) {
    const bytes = encoder.encode(line);
    expect(bytes.length).toBeLessThanOrEqual(75);
    expect(decoder.decode(bytes)).toBe(line);
  }
  expect(result.vcf).toContain("\r\n ");
  expect(unfold(result.vcf)).toContain(`N:ОФД 2025-2026;${givenName};;;\r\n`);
});

test("preserves formatted spreadsheet values and reports invalid rows with Excel row numbers", () => {
  const result = convertContactWorkbook(
    workbookBuffer(
      [
        [...CONTACT_COLUMNS],
        ["Иванов Иван Иванович", "ОФД", 2025],
        [],
        ["Нет программы", "", "2025-2026"],
        ["Ошибка Excel", "ОФД", "placeholder"],
        ["Неверное значение", true, "2025-2026"],
        ["Недопустимый\u0000символ", "ОФД", "2025-2026"],
      ],
      (sheet) => {
        sheet.C2.z = '0"-2026"';
        sheet.C5 = { t: "e", v: 7 };
      },
    ),
  );
  expect(result.contacts).toHaveLength(1);
  expect(result.contacts[0].familyName).toBe("ОФД 2025-2026");
  expect(result.issues.map((issue) => issue.row)).toEqual([4, 5, 6, 7]);
  expect(result.issues[0].message).toContain("Короткое название программы");
  expect(result.issues[1].message).toContain("Год обучения");
});

test("returns issues without a VCF when no rows are valid", () => {
  const result = convertContactWorkbook(
    workbookBuffer([[...CONTACT_COLUMNS], ["Иванов Иван Иванович", "", ""]]),
  );
  expect(result.contacts).toEqual([]);
  expect(result.vcf).toBe("");
  expect(result.issues).toHaveLength(1);
});

test("rejects missing or duplicate headers and a header placed after the first row", () => {
  expect(() => convertContactWorkbook(workbookBuffer([["ФИО"]]))).toThrow(
    "Год обучения",
  );
  expect(() =>
    convertContactWorkbook(workbookBuffer([[...CONTACT_COLUMNS, "ФИО"]])),
  ).toThrow("Столбец «ФИО» повторяется");
  expect(() =>
    convertContactWorkbook(workbookBuffer([[], [...CONTACT_COLUMNS]])),
  ).toThrow("В первой строке первого листа");
});

test("preserves phone digits and formatted leading plus/zeroes without guessing a country code", () => {
  const result = convertContactWorkbook(
    workbookBuffer(
      [
        [...CONTACT_COLUMNS],
        ["Пример Один", "ОФД", "2025-2026", 12025550100],
        ["Пример Два", "ОФД", "2025-2026", 12025550101],
        ["Пример Три", "ОФД", "2025-2026", "0012025550102"],
        ["Пример Четыре", "ОФД", "2025-2026", "", "test@example.com"],
        ["Пример Пять", "ОФД", "2025-2026"],
        ["Пример Шесть", "ОФД", "2025-2026", 12025550103],
        ["Пример Семь", "ОФД", "2025-2026", 12025550104],
      ],
      (sheet) => {
        sheet.D2.z = "0.00E+00";
        sheet.D3.z = '"+"0';
        sheet.D7.z = "0,";
        sheet.D8.z = "0000000000000";
      },
    ),
  );
  expect(result.issues).toEqual([]);
  expect(result.contacts.map((contact) => contact.phone)).toEqual([
    "12025550100",
    "+12025550101",
    "0012025550102",
    undefined,
    undefined,
    "12025550103",
    "0012025550104",
  ]);
  expect(result.contacts[3].email).toBe("test@example.com");
  expect(unfold(result.vcf)).toContain("TEL;TYPE=CELL:0012025550102\r\n");
});

test("reports invalid phones and emails instead of exporting incorrect contact details", () => {
  const result = convertContactWorkbook(
    workbookBuffer([
      [...CONTACT_COLUMNS],
      ["Пример Один", "ОФД", "2025-2026", "123;456"],
      ["Пример Два", "ОФД", "2025-2026", "", "missing-at.example.com"],
      ["Пример Три", "ОФД", "2025-2026", 123.4],
      ["Пример Четыре", "ОФД", "2025-2026", "", "a@example.com; b@example.com"],
    ]),
  );
  expect(result.contacts).toEqual([]);
  expect(result.issues.map((issue) => issue.row)).toEqual([2, 3, 4, 5]);
  expect(result.issues[0].message).toContain("номер телефона");
  expect(result.issues[1].message).toContain("почту");
  expect(result.vcf).toBe("");
});

test("rejects unreadable, empty and header-only files with Russian messages", () => {
  expect(() => convertContactWorkbook(new ArrayBuffer(0))).toThrow(
    "Выбранный файл пуст",
  );
  expect(() =>
    convertContactWorkbook(
      new TextEncoder().encode("not an Excel file").buffer,
    ),
  ).toThrow("Не удалось прочитать таблицу");
  expect(() => convertContactWorkbook(workbookBuffer([]))).toThrow(
    "Первый лист таблицы пуст",
  );
  expect(() =>
    convertContactWorkbook(workbookBuffer([[...CONTACT_COLUMNS]])),
  ).toThrow("На первом листе нет контактов");
});
