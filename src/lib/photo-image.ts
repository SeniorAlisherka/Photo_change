import { SIZE_LIMIT, PhotoProcessingError } from "./photo-converter-types";

type HeifFactory =
  typeof import("libheif-js/libheif-wasm/libheif-bundle.mjs").default;
let heifDecoder:
  | Promise<InstanceType<Awaited<ReturnType<HeifFactory>>["HeifDecoder"]>>
  | undefined;

function getHeifDecoder() {
  heifDecoder ??= import("libheif-js/libheif-wasm/libheif-bundle.mjs")
    .then(({ default: createDecoder }) => createDecoder())
    .then((module) => new module.HeifDecoder());
  return heifDecoder;
}

function createCanvas(width: number, height: number) {
  if (
    !Number.isSafeInteger(width) ||
    !Number.isSafeInteger(height) ||
    width <= 0 ||
    height <= 0
  ) {
    throw new PhotoProcessingError(
      "Не удалось определить размеры изображения.",
    );
  }
  const canvas = new OffscreenCanvas(width, height);
  const context = canvas.getContext("2d", { alpha: false });
  if (!context)
    throw new PhotoProcessingError(
      "Браузер не поддерживает обработку изображений.",
    );
  return { canvas, context };
}

async function decodeHeic(blob: Blob) {
  const decoder = await getHeifDecoder();
  const images = decoder.decode(new Uint8Array(await blob.arrayBuffer()));
  const image = images[0];
  if (!image)
    throw new PhotoProcessingError("Не удалось прочитать изображение HEIC.");

  try {
    const { canvas, context } = createCanvas(
      image.get_width(),
      image.get_height(),
    );
    const pixels = context.createImageData(canvas.width, canvas.height);
    await new Promise<void>((resolve, reject) => {
      image.display(pixels, (data) => {
        if (data) resolve();
        else
          reject(
            new PhotoProcessingError("Не удалось прочитать изображение HEIC."),
          );
      });
    });
    context.putImageData(pixels, 0, 0);
    return canvas;
  } finally {
    for (const frame of images) frame.free();
  }
}

export async function convertPhoto(
  blob: Blob,
  filename: string,
): Promise<Blob> {
  const extension = filename.split(".").at(-1)?.toLowerCase();
  let canvas: OffscreenCanvas;

  if (extension === "heic") {
    canvas = await decodeHeic(blob);
  } else if (["jpg", "jpeg", "png"].includes(extension ?? "")) {
    const bitmap = await createImageBitmap(blob);
    try {
      const surface = createCanvas(bitmap.width, bitmap.height);
      // Make transparency handling deterministic when exporting to opaque JPEG.
      surface.context.fillStyle = "#ffffff";
      surface.context.fillRect(0, 0, bitmap.width, bitmap.height);
      surface.context.drawImage(bitmap, 0, 0);
      canvas = surface.canvas;
    } finally {
      bitmap.close();
    }
  } else {
    throw new PhotoProcessingError(
      `Формат изображения не поддерживается: ${extension || "без расширения"}.`,
    );
  }

  try {
    // Encode from the same decoded original on every attempt, avoiding cumulative JPEG loss.
    for (let quality = 95; quality >= 0; quality -= 5) {
      const output = await canvas.convertToBlob({
        type: "image/jpeg",
        quality: quality / 100,
      });
      if (output.type !== "image/jpeg")
        throw new PhotoProcessingError(
          "Браузер не поддерживает сохранение в JPEG.",
        );
      if (output.size < SIZE_LIMIT) return output;
    }
    throw new PhotoProcessingError(
      "Даже при минимальном качестве файл больше 500 КиБ. Уменьшите размеры фотографии вручную.",
    );
  } finally {
    canvas.width = 1;
    canvas.height = 1;
  }
}
