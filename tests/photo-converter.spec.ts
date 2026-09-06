import { expect, test, type Page, type TestInfo } from "@playwright/test";
import { readFile, readdir, writeFile, mkdir } from "node:fs/promises";
import { join, dirname } from "node:path";
import { BlobReader, Uint8ArrayWriter, ZipReader } from "@zip.js/zip.js";
import { read, utils, write } from "xlsx";

import {
  LINK_COLUMN,
  NAME_COLUMN,
  MANUAL_FOLDER,
  OUTPUT_FOLDER,
  PHOTO_FOLDER,
  REPORT_NAME,
  UNKNOWN_FOLDER,
  WORKBOOK_NAME,
  SIZE_LIMIT,
  type ConversionReport,
} from "../src/lib/photo-converter-types";

async function selectFolder(
  page: Page,
  testInfo: TestInfo,
  files: Map<string, Uint8Array>,
  name = "Мои фото",
) {
  const directory = testInfo.outputPath(name);
  await mkdir(directory, { recursive: true });
  for (const [path, bytes] of files) {
    const destination = join(directory, path);
    await mkdir(dirname(destination), { recursive: true });
    await writeFile(destination, bytes);
  }
  await page
    .getByLabel("Выбрать папку с фотографиями", { exact: true })
    .setInputFiles(directory);
  return directory;
}

async function selectWorkbook(
  page: Page,
  buffer: Buffer,
  name = "Любое название.xlsx",
) {
  await page
    .getByLabel("Выбрать таблицу Excel", { exact: true })
    .setInputFiles({
      name,
      mimeType:
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      buffer,
    });
}

function workbookBytes(rows: unknown[][]) {
  const workbook = utils.book_new();
  utils.book_append_sheet(
    workbook,
    utils.aoa_to_sheet([[LINK_COLUMN, NAME_COLUMN], ...rows]),
    "Photos",
  );
  return write(workbook, { type: "buffer", bookType: "xlsx" }) as Buffer;
}

async function samplePng(page: Page) {
  const base64 = await page.evaluate(() => {
    const canvas = document.createElement("canvas");
    canvas.width = 64;
    canvas.height = 48;
    const context = canvas.getContext("2d")!;
    context.fillStyle = "#3a8c62";
    context.fillRect(0, 0, 32, 48);
    return canvas.toDataURL("image/png").split(",")[1];
  });
  return Buffer.from(base64, "base64");
}

async function downloadResult(page: Page) {
  await expect(page.getByRole("heading", { name: "Архив готов" })).toBeVisible({
    timeout: 60_000,
  });
  const pending = page.waitForEvent("download");
  await page.locator("a[download]").click();
  const download = await pending;
  const reader = new ZipReader(
    new BlobReader(new Blob([await readFile((await download.path())!)])),
    { useWebWorkers: false },
  );
  const files = new Map<string, Uint8Array>();
  for (const entry of await reader.getEntries()) {
    if (!entry.directory)
      files.set(
        entry.filename,
        await entry.getData(new Uint8ArrayWriter(), { checkSignature: true }),
      );
  }
  await reader.close();
  const reportPath = [...files.keys()].find((name) =>
    name.endsWith(REPORT_NAME),
  )!;
  const report = JSON.parse(
    new TextDecoder().decode(files.get(reportPath)),
  ) as ConversionReport;
  return { files, report };
}

test("converts locally, preserves originals, and reports mapping/image failures", async ({
  page,
  context,
}, testInfo) => {
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
  const png = await samplePng(page);
  const root = "";
  const originals = new Map<string, Uint8Array>([
    [
      `${root}${WORKBOOK_NAME}`,
      workbookBytes([
        ["https://example.invalid/PHOTO.JPG?download=1", "001"],
        [
          "https://forms.yandex.ru/u/files?path=%2Ffolder%2Fwith%20space.PNG",
          "Portrait A.B.",
        ],
        ["broken.png", "Broken"],
        [
          "https://disk.yandex.ru/client/disk/Files?dialog=slider&idDialog=%2Fdisk%2FFiles%2Fdisk.png",
          "Disk",
        ],
        ["duplicate.png", "001"],
        ["missing.jpg", "Missing"],
        ["invalid.png", "../escape"],
        ["PHOTO.JPG", "Second copy"],
        ["", ""],
      ]),
    ],
    [`${root}${PHOTO_FOLDER}/PHOTO.JPG`, png],
    [`${root}${PHOTO_FOLDER}/with space.PNG`, png],
    [`${root}${PHOTO_FOLDER}/broken.png`, Buffer.from("not an image")],
    [`${root}${PHOTO_FOLDER}/Вложенная папка/disk.png`, png],
    [`${root}${PHOTO_FOLDER}/duplicate.png`, png],
    [`${root}${PHOTO_FOLDER}/invalid.png`, png],
    [`${root}${PHOTO_FOLDER}/unlisted.png`, png],
    [`${PHOTO_FOLDER}/__MACOSX/._metadata`, Buffer.from("ignore")],
  ]);
  await selectWorkbook(page, Buffer.from(originals.get(WORKBOOK_NAME)!));
  await expect(
    page.getByRole("button", {
      name: "Выберите папку с фотографиями",
      exact: true,
    }),
  ).toBeDisabled();
  await selectFolder(
    page,
    testInfo,
    new Map(
      [...originals]
        .filter(([path]) => path.startsWith(`${PHOTO_FOLDER}/`))
        .map(([path, bytes]) => [path.slice(PHOTO_FOLDER.length + 1), bytes]),
    ),
  );
  await page.getByRole("button", { name: "Обработать 4 фото" }).click();
  const { files, report } = await downloadResult(page);
  expect(report).toMatchObject({
    converted: 3,
    manual: 3,
    unmatched: 1,
    missing: 1,
    issues: 5,
  });
  expect([...files.keys()].some((path) => path.includes("__MACOSX"))).toBe(
    false,
  );
  expect(
    [...files.keys()].filter((path) => path.startsWith(`${MANUAL_FOLDER}/`)),
  ).toHaveLength(3);
  expect(
    Buffer.from(files.get(`${UNKNOWN_FOLDER}/unlisted.png`)!).equals(png),
  ).toBe(true);
  for (const [path, bytes] of originals) {
    if (
      !path.includes("__MACOSX") &&
      (path.includes(PHOTO_FOLDER) || path.endsWith(WORKBOOK_NAME))
    )
      expect(Buffer.from(files.get(path)!).equals(Buffer.from(bytes))).toBe(
        true,
      );
  }
  for (const name of ["001.jpg", "Portrait A.B..jpg", "Disk.jpg"]) {
    const jpeg = files.get(`${root}${OUTPUT_FOLDER}/${name}`)!;
    expect([...jpeg.slice(0, 2)]).toEqual([0xff, 0xd8]);
    expect(jpeg.length).toBeLessThan(SIZE_LIMIT);
    const dimensions = await page.evaluate(
      async (bytes) => {
        const bitmap = await createImageBitmap(
          new Blob([new Uint8Array(bytes)]),
        );
        const size = [bitmap.width, bitmap.height];
        bitmap.close();
        return size;
      },
      [...jpeg],
    );
    expect(dimensions).toEqual([64, 48]);
  }
  expect(forbidden).toEqual([]);
  expect(runtimeErrors).toEqual([]);
  await page
    .getByRole("button", { name: "Обработать другие фотографии" })
    .click();
  await expect(
    page.getByRole("button", { name: "Выбрать таблицу", exact: true }),
  ).toBeVisible({ timeout: 60_000 });
});

test("selects the folder first, reports an invalid workbook, and recovers without reselecting photos", async ({
  page,
}, testInfo) => {
  await page.goto("/");
  const png = await samplePng(page);
  await selectFolder(page, testInfo, new Map([["photo.png", png]]));
  await expect(
    page.getByRole("button", { name: "Выберите таблицу Excel", exact: true }),
  ).toBeDisabled();
  await selectWorkbook(page, Buffer.from("broken"), "broken.zip");
  await expect(page.getByRole("main").getByRole("alert")).toContainText(
    "формате .xlsx",
  );
  await selectWorkbook(page, Buffer.from("broken"));
  await expect(page.getByRole("main").getByRole("alert")).toBeVisible();
  await expect(
    page.getByRole("button", { name: "Заменить папку", exact: true }),
  ).toBeVisible();
  await selectWorkbook(page, workbookBytes([["photo.png", "one"]]));
  await page
    .getByRole("button", { name: "Обработать 1 фото", exact: true })
    .click();
  const { files, report } = await downloadResult(page);
  expect(report.converted).toBe(1);
  expect(files.has(`${OUTPUT_FOLDER}/one.jpg`)).toBe(true);
});

test("can cancel inspection, replace either selection, and ignore obsolete results", async ({
  page,
}, testInfo) => {
  await page.goto("/");
  const png = await samplePng(page);
  const workbook = workbookBytes([["photo.png", "one"]]);
  await selectWorkbook(page, workbook);
  const directory = await selectFolder(
    page,
    testInfo,
    new Map([["photo.png", png]]),
  );
  await page.getByRole("button", { name: "Убрать папку", exact: true }).click();
  await expect(
    page.getByRole("button", {
      name: "Выберите папку с фотографиями",
      exact: true,
    }),
  ).toBeDisabled();
  await expect(
    page.getByRole("button", { name: "Заменить таблицу", exact: true }),
  ).toBeVisible();
  await page
    .getByLabel("Выбрать папку с фотографиями", { exact: true })
    .setInputFiles(directory);
  await expect(
    page.getByRole("button", { name: "Обработать 1 фото", exact: true }),
  ).toBeVisible();
  await selectWorkbook(page, workbookBytes([["photo.png", "new name"]]));
  await page
    .getByRole("button", { name: "Обработать 1 фото", exact: true })
    .click();
  const { files } = await downloadResult(page);
  expect(files.has(`${OUTPUT_FOLDER}/new name.jpg`)).toBe(true);
  expect(files.has(`${OUTPUT_FOLDER}/one.jpg`)).toBe(false);
});

test("explains empty photo folders and duplicate names, then accepts another folder", async ({
  page,
}, testInfo) => {
  await page.goto("/");
  await selectWorkbook(page, workbookBytes([["photo.png", "one"]]));
  const png = await samplePng(page);
  await selectFolder(
    page,
    testInfo,
    new Map([["notes.txt", Buffer.from("no photos")]]),
    "Без фотографий",
  );
  await expect(page.getByRole("main").getByRole("alert")).toContainText(
    "нет фотографий",
  );
  await selectFolder(
    page,
    testInfo,
    new Map([
      ["a/photo.png", png],
      ["b/photo.png", png],
    ]),
    "Повторяющиеся имена",
  );
  await expect(page.getByRole("main").getByRole("alert")).toContainText(
    "несколько фотографий с именем",
  );
  await selectFolder(
    page,
    testInfo,
    new Map([["photo.png", png]]),
    "Исходные фотографии",
  );
  await expect(
    page.getByRole("button", { name: "Обработать 1 фото", exact: true }),
  ).toBeVisible();
});

test("processes the saved Python dataset including HEIC", async ({
  page,
  context,
}) => {
  const fixtureDirectory = process.env.PHOTO_FIXTURE_DIR;
  test.skip(
    !fixtureDirectory,
    "Set PHOTO_FIXTURE_DIR to the saved Фото самозапись folder.",
  );
  const requests: string[] = [];
  context.on("request", (request) => {
    if (
      request.method() !== "GET" ||
      !request.url().startsWith("http://127.0.0.1:4173/")
    )
      requests.push(request.url());
  });
  const root = "";
  const sourceFiles = new Map<string, Uint8Array>();
  const workbook = await readFile(join(fixtureDirectory!, WORKBOOK_NAME));
  sourceFiles.set(`${root}${WORKBOOK_NAME}`, workbook);
  for (const name of await readdir(join(fixtureDirectory!, PHOTO_FOLDER))) {
    if (!name.startsWith("."))
      sourceFiles.set(
        `${root}${PHOTO_FOLDER}/${name.normalize("NFC")}`,
        await readFile(join(fixtureDirectory!, PHOTO_FOLDER, name)),
      );
  }
  await page.goto("/");
  await page
    .getByLabel("Выбрать таблицу Excel", { exact: true })
    .setInputFiles(join(fixtureDirectory!, WORKBOOK_NAME));
  await page
    .getByLabel("Выбрать папку с фотографиями", { exact: true })
    .setInputFiles(join(fixtureDirectory!, PHOTO_FOLDER));
  await expect(
    page.getByRole("button", { name: "Обработать 93 фото" }),
  ).toBeVisible({ timeout: 15_000 });
  await page.getByRole("button", { name: "Обработать 93 фото" }).click();
  const { files, report } = await downloadResult(page);
  expect(report).toMatchObject({
    converted: 93,
    manual: 0,
    unmatched: 13,
    missing: 0,
    issues: 0,
  });
  expect(
    report.results.some(
      (result) =>
        result.source.toLowerCase().endsWith(".heic") &&
        result.status === "converted",
    ),
  ).toBe(true);
  const processed = [...files.entries()].filter(([name]) =>
    name.startsWith(`${root}${OUTPUT_FOLDER}/`),
  );
  expect(processed).toHaveLength(93);
  for (const [, jpeg] of processed) {
    expect(jpeg.length).toBeLessThan(SIZE_LIMIT);
    expect([...jpeg.slice(0, 2)]).toEqual([0xff, 0xd8]);
  }
  for (const [name, bytes] of sourceFiles)
    expect(Buffer.from(files.get(name)!).equals(Buffer.from(bytes))).toBe(true);
  const sourceWorkbook = read(workbook, { type: "buffer" });
  const rows = utils.sheet_to_json<Record<string, string>>(
    sourceWorkbook.Sheets[sourceWorkbook.SheetNames[0]],
    { raw: false },
  );
  for (const row of rows) {
    if (row[NAME_COLUMN])
      expect(files.has(`${root}${OUTPUT_FOLDER}/${row[NAME_COLUMN]}.jpg`)).toBe(
        true,
      );
  }
  expect(requests).toEqual([]);
  console.log(
    `Saved dataset: ${report.converted} converted, ${report.unmatched} unmatched, ${report.elapsedMs} ms; no uploads.`,
  );
});
