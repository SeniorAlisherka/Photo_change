import {
  convertPhotos,
  inspectPhotos,
  type PhotoBatch,
} from "./photo-processing";
import {
  PhotoProcessingError,
  processingErrorMessage,
  type WorkerRequest,
  type WorkerResponse,
} from "./photo-converter-types";

let batch: PhotoBatch | undefined;
let busy = false;

function respond(message: WorkerResponse) {
  self.postMessage(message);
}

self.onmessage = async (event: MessageEvent<WorkerRequest>) => {
  if (busy) return;
  busy = true;

  try {
    if (event.data.type === "inspect") {
      batch = await inspectPhotos(event.data.workbook, event.data.photos);
      respond({ type: "ready", summary: batch.summary });
    } else if (batch) {
      if (
        typeof OffscreenCanvas === "undefined" ||
        typeof createImageBitmap === "undefined"
      ) {
        throw new PhotoProcessingError(
          "Для обработки фотографий используйте актуальную версию Chrome, Edge, Firefox или Safari.",
        );
      }
      const result = await convertPhotos(batch, (progress) =>
        respond({ type: "progress", progress }),
      );
      batch = undefined;
      respond({ type: "complete", ...result });
    }
  } catch (error) {
    batch = undefined;
    respond({
      type: "error",
      message: processingErrorMessage(
        error,
        "Не удалось прочитать файлы. Проверьте, что таблица открывается в Excel, а фотографии доступны на устройстве.",
      ),
    });
  } finally {
    busy = false;
  }
};
