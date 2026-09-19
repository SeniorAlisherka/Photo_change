import { Check, CircleAlert, Download } from "lucide-react";

import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import type { ContactConversion } from "@/lib/contact-converter-types";
import { formatBytes } from "@/lib/photo-converter-types";

export default function ContactResults({
  result,
  download,
}: {
  result: ContactConversion;
  download?: { url: string; name: string; size: number };
}) {
  const preview = result.contacts.slice(0, 5);

  return (
    <div className="min-w-0 space-y-5">
      {download ? (
        <>
          <div className="flex items-start gap-3" role="status">
            <div className="flex size-10 shrink-0 items-center justify-center rounded-full bg-emerald-50 text-emerald-700">
              <Check aria-hidden="true" className="size-5" />
            </div>
            <div className="space-y-1">
              <h2 className="text-lg font-semibold">Контакты готовы</h2>
              <p className="text-sm text-muted-foreground">
                Контактов в файле: {result.contacts.length}.
              </p>
            </div>
          </div>
          <div className="space-y-2">
            <p className="text-sm font-medium">Проверьте поля контактов</p>
            <div className="overflow-hidden rounded-lg border">
              <Table
                className="table-fixed"
                aria-label="Предпросмотр контактов"
              >
                <TableHeader>
                  <TableRow>
                    <TableHead className="w-2/5">Имя</TableHead>
                    <TableHead className="w-1/4">Фамилия</TableHead>
                    <TableHead className="whitespace-normal">
                      Телефон и почта
                    </TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {preview.map((contact) => (
                    <TableRow key={contact.row}>
                      <TableCell className="align-top whitespace-normal wrap-break-word">
                        {contact.givenName}
                      </TableCell>
                      <TableCell className="align-top whitespace-normal wrap-break-word">
                        {contact.familyName}
                      </TableCell>
                      <TableCell className="space-y-1 align-top whitespace-normal wrap-break-word">
                        {contact.phone ? <p>{contact.phone}</p> : null}
                        {contact.email ? <p>{contact.email}</p> : null}
                        {!contact.phone && !contact.email ? (
                          <span className="text-muted-foreground">
                            Не указаны
                          </span>
                        ) : null}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
            {result.contacts.length > preview.length ? (
              <p className="text-xs text-muted-foreground">
                Показаны первые {preview.length} контактов. В VCF будут все{" "}
                {result.contacts.length}.
              </p>
            ) : null}
          </div>
        </>
      ) : (
        <Alert variant="destructive">
          <CircleAlert />
          <AlertTitle>Нет контактов для экспорта</AlertTitle>
          <AlertDescription>
            Исправьте строки таблицы и выберите файл ещё раз.
          </AlertDescription>
        </Alert>
      )}
      {result.issues.length ? (
        <div className="space-y-2">
          <Alert>
            <CircleAlert />
            <AlertTitle>Строк пропущено: {result.issues.length}</AlertTitle>
            <AlertDescription>
              В этих строках не хватает данных или есть ошибки. Они не включены
              в VCF.
            </AlertDescription>
          </Alert>
          <Accordion>
            <AccordionItem value="issues" className="border-b-0">
              <AccordionTrigger>Какие строки нужно исправить?</AccordionTrigger>
              <AccordionContent>
                <ul className="max-h-64 space-y-2 overflow-y-auto text-sm text-muted-foreground">
                  {result.issues.map((issue) => (
                    <li key={issue.row} className="wrap-break-word">
                      Строка {issue.row}: {issue.message}
                    </li>
                  ))}
                </ul>
              </AccordionContent>
            </AccordionItem>
          </Accordion>
        </div>
      ) : null}
      {download ? (
        <div className="space-y-3">
          <Button
            className="h-11 w-full"
            nativeButton={false}
            render={<a href={download.url} download={download.name} />}
          >
            <Download aria-hidden="true" />
            Скачать VCF
            <span className="font-normal opacity-65">
              · {formatBytes(download.size)}
            </span>
          </Button>
          <p className="text-center text-xs text-muted-foreground">
            Скачайте файл и импортируйте его в приложение «Контакты».
          </p>
        </div>
      ) : null}
    </div>
  );
}
