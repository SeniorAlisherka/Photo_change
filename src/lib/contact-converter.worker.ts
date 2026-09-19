import {
  ContactProcessingError,
  type ContactWorkerRequest,
  type ContactWorkerResponse,
} from "./contact-converter-types";
import { convertContactWorkbook } from "./contact-vcard";

let busy = false;

function respond(message: ContactWorkerResponse) {
  self.postMessage(message);
}

self.onmessage = async (event: MessageEvent<ContactWorkerRequest>) => {
  if (busy || event.data.type !== "convert") return;
  busy = true;

  try {
    const { file } = event.data;
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      throw new ContactProcessingError(
        "Выберите таблицу Excel в формате .xlsx.",
      );
    }
    const result = convertContactWorkbook(await file.arrayBuffer());
    respond({ type: "complete", result });
  } catch (error) {
    respond({
      type: "error",
      message:
        error instanceof ContactProcessingError
          ? error.message
          : "Не удалось обработать таблицу. Проверьте, что файл доступен на устройстве и открывается в Excel.",
    });
  } finally {
    busy = false;
  }
};
