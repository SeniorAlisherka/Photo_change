"use client";

import { useEffect, useRef, useState } from "react";

import type {
  ContactConversion,
  ContactWorkerRequest,
  ContactWorkerResponse,
} from "@/lib/contact-converter-types";

type ContactState = {
  stage: "idle" | "converting" | "complete" | "error";
  file?: File;
  result?: ContactConversion;
  download?: { url: string; name: string; size: number };
  error?: string;
};

export function useContactConverter() {
  const [state, setState] = useState<ContactState>({ stage: "idle" });
  const workerRef = useRef<Worker | null>(null);
  const downloadRef = useRef<string | null>(null);

  useEffect(
    () => () => {
      workerRef.current?.terminate();
      if (downloadRef.current) URL.revokeObjectURL(downloadRef.current);
    },
    [],
  );

  function release() {
    workerRef.current?.terminate();
    workerRef.current = null;
    if (downloadRef.current) URL.revokeObjectURL(downloadRef.current);
    downloadRef.current = null;
  }

  function reset() {
    release();
    setState({ stage: "idle" });
  }

  function selectFile(file: File) {
    release();
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      setState({
        stage: "error",
        file,
        error: "Выберите таблицу Excel в формате .xlsx.",
      });
      return;
    }
    setState({ stage: "converting", file });

    try {
      const worker = new Worker(
        new URL("../lib/contact-converter.worker.ts", import.meta.url),
        { type: "module" },
      );
      workerRef.current = worker;
      const fail = (message: string) => {
        if (workerRef.current !== worker) return;
        worker.terminate();
        workerRef.current = null;
        setState({ stage: "error", file, error: message });
      };
      worker.onerror = () =>
        fail(
          "Не удалось создать контакты. Попробуйте выбрать таблицу ещё раз.",
        );
      worker.onmessageerror = () =>
        fail("Не удалось получить результат обработки. Попробуйте ещё раз.");
      worker.onmessage = ({ data }: MessageEvent<ContactWorkerResponse>) => {
        if (workerRef.current !== worker) return;
        if (data.type === "error") {
          fail(data.message);
          return;
        }
        let download: ContactState["download"];
        if (data.result.contacts.length) {
          const blob = new Blob([data.result.vcf], {
            type: "text/vcard;charset=utf-8",
          });
          const url = URL.createObjectURL(blob);
          downloadRef.current = url;
          download = {
            url,
            name: `${file.name.replace(/\.xlsx$/i, "")}-контакты.vcf`,
            size: blob.size,
          };
        }
        setState({ stage: "complete", file, result: data.result, download });
        worker.terminate();
        workerRef.current = null;
      };
      worker.postMessage({
        type: "convert",
        file,
      } satisfies ContactWorkerRequest);
    } catch {
      release();
      setState({
        stage: "error",
        file,
        error: "Не удалось запустить обработку. Попробуйте обновить браузер.",
      });
    }
  }

  return { state, selectFile, reset };
}
