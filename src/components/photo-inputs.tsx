"use client";

import { useRef } from "react";
import { FileSpreadsheet, FolderOpen, X } from "lucide-react";

import {
  Attachment,
  AttachmentAction,
  AttachmentActions,
  AttachmentContent,
  AttachmentDescription,
  AttachmentMedia,
  AttachmentTitle,
} from "@/components/ui/attachment";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { formatBytes, type PhotoFolder } from "@/lib/photo-converter-types";

export default function PhotoInputs({
  workbook,
  folder,
  disabled,
  onWorkbook,
  onFolder,
}: {
  workbook?: File;
  folder?: PhotoFolder;
  disabled: boolean;
  onWorkbook: (file?: File) => void;
  onFolder: (files?: File[]) => void;
}) {
  const workbookRef = useRef<HTMLInputElement>(null);
  const folderRef = useRef<HTMLInputElement>(null);

  return (
    <div className="space-y-5">
      <div className="space-y-3">
        <h2 className="text-sm font-medium">1. Список из Excel</h2>
        <Input
          ref={workbookRef}
          type="file"
          accept=".xlsx,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          aria-label="Выбрать таблицу Excel"
          className="hidden"
          tabIndex={-1}
          disabled={disabled}
          onChange={(event) => {
            const file = event.currentTarget.files?.[0];
            if (file) onWorkbook(file);
            event.currentTarget.value = "";
          }}
        />
        {workbook ? (
          <Attachment className="w-full" state="done">
            <AttachmentMedia>
              <FileSpreadsheet />
            </AttachmentMedia>
            <AttachmentContent>
              <AttachmentTitle title={workbook.name}>
                {workbook.name}
              </AttachmentTitle>
              <AttachmentDescription>
                {formatBytes(workbook.size)}
              </AttachmentDescription>
            </AttachmentContent>
            <AttachmentActions>
              <AttachmentAction
                aria-label="Убрать таблицу"
                disabled={disabled}
                onClick={() => onWorkbook()}
              >
                <X />
              </AttachmentAction>
            </AttachmentActions>
          </Attachment>
        ) : null}
        <Button
          variant="outline"
          className="w-full"
          disabled={disabled}
          onClick={() => workbookRef.current?.click()}
        >
          <FileSpreadsheet aria-hidden="true" />
          {workbook ? "Заменить таблицу" : "Выбрать таблицу"}
        </Button>
        <p className="text-xs text-muted-foreground">
          Файл .xlsx со ссылками на фотографии и новыми именами.
        </p>
      </div>
      <div className="space-y-3">
        <h2 className="text-sm font-medium">2. Папка с фотографиями</h2>
        <Input
          ref={folderRef}
          type="file"
          {...{ webkitdirectory: "" }}
          multiple
          aria-label="Выбрать папку с фотографиями"
          className="hidden"
          tabIndex={-1}
          disabled={disabled}
          onChange={(event) => {
            if (event.currentTarget.files)
              onFolder(Array.from(event.currentTarget.files));
            event.currentTarget.value = "";
          }}
        />
        {folder ? (
          <Attachment className="w-full" state="done">
            <AttachmentMedia>
              <FolderOpen />
            </AttachmentMedia>
            <AttachmentContent>
              <AttachmentTitle title={folder.name}>
                {folder.name}
              </AttachmentTitle>
              <AttachmentDescription>
                Файлов: {folder.files.length} · {formatBytes(folder.size)}
              </AttachmentDescription>
            </AttachmentContent>
            <AttachmentActions>
              <AttachmentAction
                aria-label="Убрать папку"
                disabled={disabled}
                onClick={() => onFolder()}
              >
                <X />
              </AttachmentAction>
            </AttachmentActions>
          </Attachment>
        ) : null}
        <Button
          variant="outline"
          className="w-full"
          disabled={disabled}
          onClick={() => folderRef.current?.click()}
        >
          <FolderOpen aria-hidden="true" />
          {folder ? "Заменить папку" : "Выбрать папку"}
        </Button>
        <p className="text-xs text-muted-foreground">
          Папка с исходными фотографиями в формате JPG, JPEG, PNG или HEIC.
        </p>
      </div>
    </div>
  );
}
