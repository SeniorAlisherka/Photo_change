"use client";

import { useEffect, useRef, useState } from "react";

import type {
  PhotoFolder,
  PhotoSummary,
  ConversionProgress,
  ConversionReport,
  WorkerRequest,
  WorkerResponse,
} from "@/lib/photo-converter-types";

type Selection = { workbook?: File; folder?: PhotoFolder };
type ConverterState = Selection & {
  stage: "idle" | "inspecting" | "ready" | "converting" | "complete" | "error";
  summary?: PhotoSummary;
  progress?: ConversionProgress;
  report?: ConversionReport;
  download?: { url: string; name: string; size: number };
  error?: string;
};

export function usePhotoConverter() {
  const [state, setState] = useState<ConverterState>({ stage: "idle" });
  const selectionRef = useRef<Selection>({});
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

  function updateSelection(selection: Selection) {
    release();
    selectionRef.current = selection;
    const { workbook, folder } = selection;
    let error: string | undefined;
    if (workbook && !workbook.name.toLowerCase().endsWith(".xlsx"))
      error = "Выберите таблицу Excel в формате .xlsx.";
    else if (folder && !folder.files.length)
      error = "Выбранная папка пуста. Выберите папку с исходными фото.";
    if (error) {
      setState({ ...selection, stage: "error", error });
      return;
    }
    if (!workbook || !folder) {
      setState({ ...selection, stage: "idle" });
      return;
    }

    setState({ ...selection, stage: "inspecting" });
    try {
      const worker = new Worker(
        new URL("../lib/photo-converter.worker.ts", import.meta.url),
        { type: "module" },
      );
      workerRef.current = worker;
      const fail = (message: string) => {
        if (workerRef.current !== worker) return;
        worker.terminate();
        workerRef.current = null;
        setState((previous) => ({
          ...previous,
          stage: "error",
          error: message,
        }));
      };
      worker.onerror = () =>
        fail(
          "Обработка неожиданно прервалась. Попробуйте выбрать меньше фотографий или обновите страницу.",
        );
      worker.onmessageerror = () =>
        fail(
          "Браузер не смог получить результат обработки. Попробуйте ещё раз.",
        );
      worker.onmessage = ({ data }: MessageEvent<WorkerResponse>) => {
        if (workerRef.current !== worker) return;
        if (data.type === "ready") {
          setState((previous) => ({
            ...previous,
            stage: "ready",
            summary: data.summary,
          }));
        } else if (data.type === "progress") {
          setState((previous) => ({ ...previous, progress: data.progress }));
        } else if (data.type === "error") {
          fail(data.message);
        } else {
          const url = URL.createObjectURL(data.blob);
          downloadRef.current = url;
          setState((previous) => ({
            ...previous,
            stage: "complete",
            report: data.report,
            download: {
              url,
              name: `${folder.name}-обработано.zip`,
              size: data.blob.size,
            },
          }));
          worker.terminate();
          workerRef.current = null;
        }
      };
      // Pass paths explicitly; File's directory properties may not survive
      // structured cloning into a Web Worker in every browser.
      worker.postMessage({
        type: "inspect",
        workbook,
        photos: folder.files,
      } satisfies WorkerRequest);
    } catch {
      release();
      setState({
        ...selection,
        stage: "error",
        error: "Не удалось запустить обработку. Попробуйте обновить браузер.",
      });
    }
  }

  function selectWorkbook(workbook?: File) {
    updateSelection({ ...selectionRef.current, workbook });
  }

  function selectFolder(files?: File[]) {
    const folder: PhotoFolder | undefined = files && {
      name: files[0]?.webkitRelativePath.split("/")[0] || "Фотографии",
      size: files.reduce((sum, file) => sum + file.size, 0),
      files: files.map((file) => ({
        file,
        path:
          file.webkitRelativePath.split("/").slice(1).join("/") || file.name,
      })),
    };
    updateSelection({ ...selectionRef.current, folder });
  }

  function convert() {
    if (state.stage !== "ready" || !workerRef.current) return;
    setState((previous) => ({ ...previous, stage: "converting" }));
    workerRef.current.postMessage({ type: "convert" } satisfies WorkerRequest);
  }

  return {
    state,
    selectWorkbook,
    selectFolder,
    convert,
    reset: () => updateSelection({}),
    retry: () => updateSelection(selectionRef.current),
  };
}
