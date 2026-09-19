"use client";

import { useRef } from "react";
import { ArrowRight, CircleAlert, FileSpreadsheet, X } from "lucide-react";

import ContactInstructions from "@/components/contact-instructions";
import ContactResults from "@/components/contact-results";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import {
  Attachment,
  AttachmentAction,
  AttachmentActions,
  AttachmentContent,
  AttachmentDescription,
  AttachmentMedia,
  AttachmentTitle,
} from "@/components/ui/attachment";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Spinner } from "@/components/ui/spinner";
import { useContactConverter } from "@/hooks/use-contact-converter";
import { formatBytes } from "@/lib/photo-converter-types";

export default function ContactConverter() {
  const { state, selectFile, reset } = useContactConverter();
  const inputRef = useRef<HTMLInputElement>(null);
  const busy = state.stage === "converting";

  return (
    <div className="space-y-7">
      <div className="space-y-3">
        <h1 className="text-3xl font-semibold tracking-tight sm:text-4xl">
          Контакты из Excel
        </h1>
        <p className="max-w-xl leading-relaxed text-muted-foreground">
          Создайте файл с именами, телефонами и почтой для импорта в адресную
          книгу.
        </p>
        <div className="flex items-center gap-2 pt-1">
          <Badge variant="outline" className="font-normal">
            Excel · XLSX
          </Badge>
          <ArrowRight
            aria-hidden="true"
            className="size-3 text-muted-foreground"
          />
          <Badge variant="secondary" className="font-normal">
            Контакты · VCF
          </Badge>
        </div>
      </div>
      <Card className="shadow-xs">
        <CardHeader>
          <CardTitle className="text-base">
            Выберите таблицу контактов
          </CardTitle>
          <CardDescription>
            Столбцы: ФИО, Короткое название программы, Год обучения, Номер
            телефона, Почта.
          </CardDescription>
        </CardHeader>
        <CardContent className="min-w-0 space-y-5">
          <Input
            ref={inputRef}
            type="file"
            accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            aria-label="Выбрать таблицу контактов"
            className="hidden"
            tabIndex={-1}
            onChange={(event) => {
              const file = event.currentTarget.files?.[0];
              if (file) selectFile(file);
              event.currentTarget.value = "";
            }}
          />
          {state.file ? (
            <Attachment
              className="w-full"
              state={
                busy ? "processing" : state.stage === "error" ? "error" : "done"
              }
            >
              <AttachmentMedia>
                {busy ? <Spinner /> : <FileSpreadsheet />}
              </AttachmentMedia>
              <AttachmentContent>
                <AttachmentTitle title={state.file.name}>
                  {state.file.name}
                </AttachmentTitle>
                <AttachmentDescription>
                  {formatBytes(state.file.size)}
                </AttachmentDescription>
              </AttachmentContent>
              <AttachmentActions>
                <AttachmentAction
                  aria-label="Убрать таблицу контактов"
                  onClick={reset}
                >
                  <X />
                </AttachmentAction>
              </AttachmentActions>
            </Attachment>
          ) : null}
          <Button
            variant="outline"
            className="w-full"
            onClick={() => inputRef.current?.click()}
          >
            <FileSpreadsheet aria-hidden="true" />
            {state.file ? "Заменить таблицу" : "Выбрать таблицу"}
          </Button>
          {busy ? (
            <div className="space-y-3">
              <p
                role="status"
                className="flex items-center gap-2 text-sm text-muted-foreground"
              >
                <Spinner />
                Создание контактов…
              </p>
              <Button variant="ghost" className="w-full" onClick={reset}>
                Отменить
              </Button>
            </div>
          ) : null}
          {state.error ? (
            <Alert variant="destructive">
              <CircleAlert />
              <AlertTitle>Не удалось создать контакты</AlertTitle>
              <AlertDescription className="wrap-break-word">
                {state.error}
              </AlertDescription>
            </Alert>
          ) : null}
          {state.result ? (
            <ContactResults result={state.result} download={state.download} />
          ) : null}
        </CardContent>
      </Card>
      <ContactInstructions />
    </div>
  );
}
