"use client";

import {
  ArrowRight,
  Check,
  CircleAlert,
  ImageIcon,
  LockKeyhole,
} from "lucide-react";

import PhotoInstructions from "@/components/photo-instructions";
import PhotoInputs from "@/components/photo-inputs";
import ConversionResults from "@/components/conversion-results";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import {
  Progress,
  ProgressLabel,
  ProgressValue,
} from "@/components/ui/progress";
import { Separator } from "@/components/ui/separator";
import { Spinner } from "@/components/ui/spinner";
import { usePhotoConverter } from "@/hooks/use-photo-converter";

export default function PhotoConverter() {
  const { state, selectWorkbook, selectFolder, convert, reset, retry } =
    usePhotoConverter();
  const progress = state.progress;
  const percentage = progress
    ? Math.min(
        99,
        Math.round(
          (progress.phase === "converting" ? 0 : 75) +
            (progress.completed / Math.max(1, progress.total)) *
              (progress.phase === "converting" ? 75 : 25),
        ),
      )
    : 0;

  return (
    <div className="mx-auto flex w-full max-w-3xl flex-1 flex-col px-5 sm:px-8">
      <header className="flex items-center justify-between gap-3 border-b py-6">
        <div className="flex items-center gap-2.5 font-semibold tracking-tight">
          <div className="flex size-8 items-center justify-center rounded-lg bg-primary text-primary-foreground">
            <ImageIcon aria-hidden="true" className="size-4" />
          </div>
          Фотообработка
        </div>
        <span className="flex items-center gap-1.5 text-xs text-muted-foreground sm:text-sm">
          <LockKeyhole aria-hidden="true" className="size-3.5" />
          На вашем устройстве
        </span>
      </header>
      <main className="flex-1 space-y-7 py-9 sm:py-12">
        <div className="space-y-3">
          <h1 className="text-3xl font-semibold tracking-tight sm:text-4xl">
            Обработка и переименование фото
          </h1>
          <p className="max-w-xl leading-relaxed text-muted-foreground">
            Преобразуйте фотографии в JPEG, сожмите и переименуйте их по списку
            из Excel.
          </p>
          <div className="flex flex-wrap items-center gap-2 pt-1">
            <Badge variant="outline" className="font-normal">
              JPG · JPEG · PNG · HEIC
            </Badge>
            <ArrowRight
              aria-hidden="true"
              className="size-3 text-muted-foreground"
            />
            <Badge variant="secondary" className="font-normal">
              JPEG · менее 500 КиБ
            </Badge>
          </div>
        </div>
        <Card className="shadow-xs">
          {state.stage !== "complete" ? (
            <CardHeader>
              <CardTitle className="text-base">
                {state.stage === "converting"
                  ? "Обработка фотографий"
                  : "Выберите таблицу и фотографии"}
              </CardTitle>
              <CardDescription>
                {state.stage === "converting"
                  ? "Не закрывайте вкладку до завершения обработки."
                  : "Добавьте список из Excel и папку с исходными фотографиями."}
              </CardDescription>
            </CardHeader>
          ) : null}
          <CardContent className="space-y-5">
            {state.stage === "complete" && state.report && state.download ? (
              <ConversionResults
                report={state.report}
                download={state.download}
                onReset={reset}
              />
            ) : (
              <>
                <PhotoInputs
                  workbook={state.workbook}
                  folder={state.folder}
                  disabled={state.stage === "converting"}
                  onWorkbook={selectWorkbook}
                  onFolder={selectFolder}
                />
                {state.error ? (
                  <Alert variant="destructive">
                    <CircleAlert />
                    <AlertTitle>Не удалось подготовить фотографии</AlertTitle>
                    <AlertDescription className="wrap-break-word">
                      {state.error}
                    </AlertDescription>
                  </Alert>
                ) : null}
                {state.stage === "ready" && state.summary ? (
                  <div className="space-y-5">
                    <div className="grid grid-cols-3 gap-3 rounded-xl bg-muted/50 p-4">
                      {[
                        [state.summary.photos, "Исходных фото"],
                        [state.summary.matched, "Совпадений"],
                        [state.summary.unmatched, "Фото вне списка"],
                      ].map(([count, label]) => (
                        <div key={label} className="space-y-1">
                          <p className="text-xl font-semibold tabular-nums">
                            {count}
                          </p>
                          <p className="text-xs text-muted-foreground">
                            {label}
                          </p>
                        </div>
                      ))}
                    </div>
                    {state.summary.issues > 0 ? (
                      <p
                        role="status"
                        className="flex items-start gap-2 text-sm text-muted-foreground"
                      >
                        <CircleAlert
                          aria-hidden="true"
                          className="mt-0.5 size-4 shrink-0"
                        />
                        Строк с ошибками: {state.summary.issues} из{" "}
                        {state.summary.rows}. Можно продолжить — подробности
                        будут в отчёте.
                      </p>
                    ) : (
                      <p className="flex items-center gap-2 text-sm text-muted-foreground">
                        <Check
                          aria-hidden="true"
                          className="size-4 text-emerald-700"
                        />
                        Таблица проверена. Можно начинать обработку.
                      </p>
                    )}
                    <Button className="h-11 w-full" onClick={convert}>
                      {state.summary.matched
                        ? `Обработать ${state.summary.matched} фото`
                        : "Создать ZIP с результатами"}
                      <ArrowRight aria-hidden="true" />
                    </Button>
                  </div>
                ) : null}
                {state.stage === "idle" ? (
                  <Button className="h-11 w-full" disabled>
                    {state.workbook
                      ? "Выберите папку с фотографиями"
                      : state.folder
                        ? "Выберите таблицу Excel"
                        : "Выберите таблицу и папку"}
                  </Button>
                ) : null}
                {state.stage === "converting" ? (
                  <div className="space-y-4">
                    <Progress value={percentage} locale="ru-RU">
                      <ProgressLabel>
                        {progress?.phase === "packing"
                          ? "Создание ZIP-архива"
                          : "Обработка фотографий"}
                      </ProgressLabel>
                      <ProgressValue />
                    </Progress>
                    <div className="flex items-center justify-between gap-3 text-xs text-muted-foreground">
                      <span className="truncate">
                        {progress?.filename || "Подготовка к обработке…"}
                      </span>
                      {progress ? (
                        <span className="shrink-0 tabular-nums">
                          {progress.completed} / {progress.total}
                        </span>
                      ) : null}
                    </div>
                    <Button
                      variant="outline"
                      className="w-full"
                      onClick={reset}
                    >
                      Отменить
                    </Button>
                  </div>
                ) : null}
                {state.stage === "inspecting" ? (
                  <p role="status" className="text-sm text-muted-foreground">
                    <Spinner className="mr-2 inline size-4" />
                    Проверяем таблицу и сопоставляем имена файлов…
                  </p>
                ) : null}
                {state.stage === "error" && state.workbook && state.folder ? (
                  <Button variant="outline" className="w-full" onClick={retry}>
                    Попробовать ещё раз
                  </Button>
                ) : null}
              </>
            )}
          </CardContent>
        </Card>
        <PhotoInstructions />
      </main>
      <footer className="space-y-4 pb-6">
        <Separator />
        <p className="text-center text-xs text-muted-foreground">
          Без отправки файлов на сервер и без регистрации. Всё на вашем
          устройстве.
        </p>
      </footer>
    </div>
  );
}
