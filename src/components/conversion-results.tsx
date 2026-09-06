import {
  Check,
  CircleAlert,
  FileArchive,
  Folder,
  Download,
  RotateCcw,
} from "lucide-react";

import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Separator } from "@/components/ui/separator";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import {
  formatBytes,
  formatDecimal,
  type ConversionReport,
  type PhotoResult,
} from "@/lib/photo-converter-types";

const statusLabels: Record<PhotoResult["status"], string> = {
  converted: "Обработано",
  manual: "Нужна проверка",
  unmatched: "Вне списка",
  missing: "Не найдено",
  skipped: "Строка пропущена",
};

export default function ConversionResults({
  report,
  download,
  onReset,
}: {
  report: ConversionReport;
  download: { url: string; name: string; size: number };
  onReset: () => void;
}) {
  const folders = [
    {
      name: "Фото после обработки",
      label: "Обработано",
      count: report.converted,
      icon: Check,
    },
    {
      name: "Непонятные фото",
      label: "Вне списка",
      count: report.unmatched,
      icon: Folder,
    },
    {
      name: "Фото для ручной обработки",
      label: "Для ручной обработки",
      count: report.manual,
      icon: CircleAlert,
    },
  ];

  return (
    <div className="space-y-6">
      <div className="flex items-start gap-3">
        <div className="flex size-10 shrink-0 items-center justify-center rounded-full bg-emerald-50 text-emerald-700">
          <Check aria-hidden="true" className="size-5" />
        </div>
        <div className="space-y-1">
          <h2 className="text-lg font-semibold">Архив готов</h2>
          <p className="text-sm text-muted-foreground">
            Обработано фото: {report.converted}. Время:{" "}
            {formatDecimal(report.elapsedMs / 1000)} с.
          </p>
        </div>
      </div>
      <div className="divide-y rounded-xl border">
        {folders.map(({ name, label, count, icon: Icon }) => (
          <div key={name} className="flex items-center gap-3 px-4 py-3">
            <Icon
              aria-hidden="true"
              className="size-4 shrink-0 text-muted-foreground"
            />
            <div className="min-w-0 flex-1">
              <p className="text-sm font-medium">{label}</p>
              <p
                className="truncate text-xs text-muted-foreground"
                title={name}
              >
                {name}
              </p>
            </div>
            <span className="text-sm tabular-nums">{count}</span>
          </div>
        ))}
      </div>
      {report.issues > 0 ? (
        <Alert>
          <CircleAlert />
          <AlertTitle>Строк с ошибками: {report.issues}</AlertTitle>
          <AlertDescription>
            Подробности — в отчёте ниже. Исходные фотографии сохранены в
            ZIP-архиве.
            {report.missing > 0
              ? ` Не найдено фотографий из списка: ${report.missing}.`
              : ""}
          </AlertDescription>
        </Alert>
      ) : null}
      <div className="space-y-3">
        <Button
          className="h-11 w-full"
          nativeButton={false}
          render={<a href={download.url} download={download.name} />}
        >
          <Download aria-hidden="true" />
          Скачать ZIP
          <span className="font-normal opacity-65">
            · {formatBytes(download.size)}
          </span>
        </Button>
        <p className="flex items-center justify-center gap-1.5 text-center text-xs text-muted-foreground">
          <FileArchive aria-hidden="true" className="size-3.5 shrink-0" />
          Внутри: исходные фото, таблица, результаты и JSON-отчёт.
        </p>
      </div>
      <Accordion>
        <AccordionItem value="results" className="border-b-0">
          <AccordionTrigger>
            Отчёт по файлам ({report.results.length})
          </AccordionTrigger>
          <AccordionContent>
            <div className="max-h-96 overflow-auto rounded-lg border">
              <Table>
                <TableHeader>
                  <TableRow>
                    <TableHead>Исходный файл</TableHead>
                    <TableHead>Результат</TableHead>
                    <TableHead>Подробности</TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {report.results.map((result, index) => (
                    <TableRow key={index}>
                      <TableCell className="max-w-52 whitespace-normal wrap-break-word">
                        {result.source || `Строка ${result.row}`}
                        {result.row ? (
                          <span className="mt-1 block text-xs text-muted-foreground">
                            Строка {result.row}
                          </span>
                        ) : null}
                      </TableCell>
                      <TableCell>
                        <Badge
                          variant={
                            result.status === "converted"
                              ? "secondary"
                              : "outline"
                          }
                        >
                          {statusLabels[result.status]}
                        </Badge>
                      </TableCell>
                      <TableCell className="max-w-60 whitespace-normal wrap-break-word text-muted-foreground">
                        {result.status === "converted" ? (
                          <>
                            {result.output}
                            <span className="mt-1 block text-xs">
                              {formatBytes(result.size ?? 0)}
                            </span>
                          </>
                        ) : (
                          result.message
                        )}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
          </AccordionContent>
        </AccordionItem>
      </Accordion>
      <Separator />
      <Button
        variant="ghost"
        className="w-full text-muted-foreground"
        onClick={onReset}
      >
        <RotateCcw aria-hidden="true" />
        Обработать другие фотографии
      </Button>
    </div>
  );
}
