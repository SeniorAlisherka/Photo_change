import { expect, test, type Page } from "@playwright/test";
import { readFile } from "node:fs/promises";
import { utils, write } from "xlsx";

const columns = [
  "ФИО",
  "Короткое название программы",
  "Год обучения",
  "Номер телефона",
  "Почта",
];

function workbookBytes(rows: unknown[][], headers = columns) {
  const workbook = utils.book_new();
  utils.book_append_sheet(
    workbook,
    utils.aoa_to_sheet([headers, ...rows]),
    "Контакты",
  );
  return write(workbook, { type: "buffer", bookType: "xlsx" }) as Buffer;
}

function contactsPanel(page: Page) {
  return page.getByRole("tabpanel", { name: "Контакты", exact: true });
}

async function selectWorkbook(
  page: Page,
  buffer: Buffer,
  name = "Контакты.xlsx",
) {
  await contactsPanel(page)
    .getByLabel("Выбрать таблицу контактов", { exact: true })
    .setInputFiles({
      name,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer,
    });
}

async function downloadContacts(page: Page) {
  const panel = contactsPanel(page);
  await expect(
    panel.getByRole("heading", { name: "Контакты готовы", exact: true }),
  ).toBeVisible();
  const pending = page.waitForEvent("download");
  await panel.locator("a[download]").click();
  const download = await pending;
  expect(download.suggestedFilename()).toMatch(/\.vcf$/i);
  const text = await readFile((await download.path())!, "utf8");
  return { text, unfolded: text.replace(/\r\n[ \t]/g, "") };
}

test("exports Cyrillic contacts locally and preserves both tabs' selections", async ({
  page,
  context,
}) => {
  const forbidden: string[] = [];
  const runtimeErrors: string[] = [];
  page.on("pageerror", (error) => runtimeErrors.push(error.message));
  await context.route("**/*", async (route) => {
    const request = route.request();
    if (
      request.method() !== "GET" ||
      !request.url().startsWith("http://127.0.0.1:4173/")
    ) {
      forbidden.push(`${request.method()} ${request.url()}`);
      await route.abort();
    } else await route.continue();
  });
  await page.goto("/");
  await expect(
    page.getByRole("tab", { name: "Фотообработка", exact: true }),
  ).toHaveAttribute("aria-selected", "true");
  await page
    .getByLabel("Выбрать таблицу Excel", { exact: true })
    .setInputFiles({
      name: "Список фотографий.xlsx",
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer: workbookBytes(
        [["https://example.invalid/photo.jpg", "Фото 001"]],
        [
          "Ссылка, содержащая название файла до обработки",
          "Название файла после обработки",
        ],
      ),
    });
  await page.getByRole("tab", { name: "Контакты", exact: true }).click();
  await selectWorkbook(
    page,
    workbookBytes([
      [
        "Иванов Иван Иванович",
        "ОФД",
        "2025-2026",
        "+7 (000) 000-00-01",
        "ivanov@example.invalid",
      ],
      ["Александрова Александра Александровна", "ДПО", "2026-2027"],
    ]),
  );
  const { text, unfolded } = await downloadContacts(page);
  expect(unfolded.match(/BEGIN:VCARD/g)).toHaveLength(2);
  expect(unfolded).toContain("N:ОФД 2025-2026;Иванов Иван Иванович;;;\r\n");
  expect(unfolded).toContain("FN:Иванов Иван Иванович ОФД 2025-2026\r\n");
  expect(unfolded).toMatch(/\r\nTEL(?:;[^:\r\n]+)?:\+70000000001\r\n/);
  expect(unfolded).toMatch(
    /\r\nEMAIL(?:;[^:\r\n]+)?:ivanov@example\.invalid\r\n/,
  );
  expect(unfolded).toContain(
    "N:ДПО 2026-2027;Александрова Александра Александровна;;;\r\n",
  );
  expect(unfolded).toContain(
    "FN:Александрова Александра Александровна ДПО 2026-2027\r\n",
  );
  expect(text).toMatch(/\r\n /);
  for (const line of text.split("\r\n")) {
    expect(Buffer.byteLength(line, "utf8")).toBeLessThanOrEqual(75);
  }
  const downloadLink = contactsPanel(page).locator("a[download]");
  const originalUrl = await downloadLink.getAttribute("href");
  await page.getByRole("tab", { name: "Фотообработка", exact: true }).click();
  const photoPanel = page.getByRole("tabpanel", {
    name: "Фотообработка",
    exact: true,
  });
  await expect(
    photoPanel.getByText("Список фотографий.xlsx", { exact: true }),
  ).toBeVisible();
  await expect(
    photoPanel.getByRole("button", {
      name: "Выберите папку с фотографиями",
      exact: true,
    }),
  ).toBeDisabled();
  await page.getByRole("tab", { name: "Контакты", exact: true }).click();
  await expect(downloadLink).toHaveAttribute("href", originalUrl!);
  expect((await downloadContacts(page)).text).toBe(text);
  expect(forbidden).toEqual([]);
  expect(runtimeErrors).toEqual([]);
});

test("explains missing headers, accepts a replacement workbook, and resets the result", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("tab", { name: "Контакты", exact: true }).click();
  const panel = contactsPanel(page);
  await selectWorkbook(
    page,
    workbookBytes([["Иванов Иван Иванович", "ОФД"]], columns.slice(0, 2)),
  );
  await expect(panel.getByRole("alert")).toContainText("Год обучения");
  await expect(panel.locator("a[download]")).toHaveCount(0);
  await selectWorkbook(
    page,
    workbookBytes([["Иванов Иван Иванович", "ОФД", "2025-2026"]]),
    "Исправленная таблица.xlsx",
  );
  const { unfolded } = await downloadContacts(page);
  expect(unfolded.match(/BEGIN:VCARD/g)).toHaveLength(1);
  await panel
    .getByRole("button", { name: "Убрать таблицу контактов", exact: true })
    .click();
  await expect(panel.locator("a[download]")).toHaveCount(0);
  await expect(
    panel.getByRole("button", { name: "Выбрать таблицу", exact: true }),
  ).toBeVisible();
  await expect(panel.getByRole("alert")).toHaveCount(0);
});

test("reports incomplete rows and refuses an export with no valid contacts", async ({
  page,
}) => {
  await page.goto("/");
  await page.getByRole("tab", { name: "Контакты", exact: true }).click();
  const panel = contactsPanel(page);
  await selectWorkbook(
    page,
    workbookBytes([
      ["Иванов Иван Иванович", "ОФД", "2025-2026"],
      ["", "ОФД", "2025-2026"],
      ["Без программы", "", "2025-2026"],
      ["Без года", "ОФД", ""],
      ["Некорректный телефон", "ОФД", "2025-2026", "не номер", ""],
      ["Некорректная почта", "ОФД", "2025-2026", "", "не адрес"],
    ]),
  );
  const { unfolded } = await downloadContacts(page);
  await expect(
    panel.getByText("Строк пропущено: 5", { exact: true }),
  ).toBeVisible();
  expect(unfolded.match(/BEGIN:VCARD/g)).toHaveLength(1);
  expect(unfolded).not.toContain("Без программы");
  expect(unfolded).not.toContain("Без года");
  expect(unfolded).not.toContain("Некорректный телефон");
  expect(unfolded).not.toContain("Некорректная почта");
  await selectWorkbook(page, workbookBytes([["Без года", "ОФД", ""]]));
  await expect(
    panel.getByRole("alert").filter({ hasText: "Нет контактов для экспорта" }),
  ).toBeVisible();
  await expect(panel.locator("a[download]")).toHaveCount(0);
});
