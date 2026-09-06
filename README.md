# Photo Change

A Next.js + TypeScript interface for the original Python photo-processing script. The UI uses shadcn's Base UI components (`base-nova`), including Attachment, Progress, and Accordion. There are no API routes, server actions, file uploads, or external conversion services.

## Run locally

```sh
npm install
npm run dev
```

For a production preview:

```sh
npm run build
npm start
```

## Deploy to Vercel

Import the repository into Vercel as a Next.js project. Use `npm run build` as the build command. `next.config.ts` sets `output: "export"`, so the build creates a static website in `out/`. No environment variables, database, or server functions are required. Deploy application code only; keep private photo fixtures outside the repository.

## Input

Select two sources separately:

1. An Excel workbook in `.xlsx` format (any filename).
2. A folder containing the original JPG/JPEG/PNG/HEIC photos (any folder name).

The browser's native directory picker reads files locally using `webkitdirectory`; the app never submits a form or sends file data over the network. Select the original-photo folder itself, not a parent containing previous results. Subfolders are included and photo basenames must be unique across the selection. Hidden files, macOS metadata, and non-photo files are ignored.

The first worksheet must have these headings in its first row:

- `Ссылка, содержащая название файла до обработки`
- `Название файла после обработки`

The first column contains a photo URL identifying a local photo. The user instructions show fictional Yandex Disk and Yandex Forms links, matching the intended worksheet format. URL paths are decoded, including Yandex Forms/Disk links with `path` or `idDialog` query parameters. Other query parameters are ignored, and URLs are never fetched. The second column contains a filename stem; `.jpg` is appended, matching the Python script. Initials ending in a period are preserved (for example, `Person A.B.` becomes `Person A.B..jpg`). Spreadsheet display formatting is used to preserve identifiers such as `001`.

## Processing and output

A browser Web Worker reads the selected files and Excel workbook, converts JPG/JPEG/PNG/HEIC to JPEG, and builds a downloadable ZIP. The worker is terminated on cancellation and after completion. HEIC decoding is loaded on demand from the application's own assets.

- JPEG quality starts at 95% and decreases in steps of 5 until the result is **smaller than 512,000 bytes** (500 KiB).
- Every quality attempt encodes from the same decoded image. Image dimensions are preserved; no automatic resizing is performed.
- Photos that still exceed the size limit go to manual processing.
- Native image decoding respects image orientation. Transparent areas are filled white. Original metadata is not retained in converted JPEGs. Browser output is not byte-for-byte equivalent to Pillow.
- HEIC uses the first decoded image, matching the original workflow.
- Uppercase extensions work. Blank rows are skipped. Missing sources, duplicate output names (case-insensitive), duplicate source rows, unsafe filenames, unsupported formats, and corrupt images are reported individually.

The result copies selected source photos into `Фото до обработки` (preserving relative subfolders), copies the workbook byte-for-byte as `Список.xlsx`, creates these result folders, and includes `processing-report.json`:

| Folder | Contents |
| --- | --- |
| `Фото после обработки` | Converted and renamed JPEGs |
| `Непонятные фото` | Source photos not referenced by the worksheet |
| `Фото для ручной обработки` | Referenced source photos without a successful conversion |

The original files on the device are never changed. A failure while reading/copying an original or writing the ZIP aborts the download rather than producing an incomplete result.

The app imposes no fixed limits on input bytes, file count, or image resolution. Photos are processed one at a time. Available browser memory, supported image dimensions, workbook parsing, and the accumulated output ZIP still bound what a device can process; arbitrarily large batches are not guaranteed. The 500 KiB output target is part of the photo-conversion requirements, not an input limit.

## Structure

```text
src/app/                         Small page, layout, and global styles
src/components/                  Converter, file/folder pickers, instructions, and results UI
src/components/ui/               Official shadcn Base UI components
src/hooks/use-photo-converter.ts Worker lifecycle and UI state
src/lib/photo-mapping.ts         Worksheet validation and filename mapping
src/lib/photo-image.ts           Browser JPEG encoding and HEIC decoding
src/lib/photo-processing.ts      File inspection, conversion, and ZIP packaging
src/lib/photo-converter.worker.ts Worker message handler
src/lib/photo-converter-types.ts Shared messages, constants, and report types
```

Dependencies were installed from their current releases. SheetJS uses the official CDN release because the npm `xlsx` package is outdated. All runtime libraries are bundled with the app; photos are never sent to a CDN.

## Checks

```sh
npm run lint
npx tsc --noEmit
npx playwright install chromium
npm run test:browser
```

If Chrome is already installed, set `PLAYWRIGHT_CHROMIUM_CHANNEL=chrome` to use it instead of downloading Chromium. To include the original private dataset in the browser test:

```sh
PLAYWRIGHT_CHROMIUM_CHANNEL=chrome \
PHOTO_FIXTURE_DIR='/path/to/Фото самозапись' \
npm run test:browser
```

Tests exercise actual worker conversion and ZIP downloads, preserve originals byte-for-byte, check JPEG dimensions/size, mapping errors, corrupt input, selection order, cancellation/replacement, duplicate source names, folders without photos, and outgoing requests. The optional saved-dataset test checks the original 93 converted / 13 unmatched result, including HEIC. Fixtures are loaded from disk during tests and never deployed.
