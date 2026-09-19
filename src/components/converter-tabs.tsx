"use client";

import { ContactRound, ImageIcon, LockKeyhole } from "lucide-react";

import ContactConverter from "@/components/contact-converter";
import PhotoConverter from "@/components/photo-converter";
import { Separator } from "@/components/ui/separator";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";

export default function ConverterTabs() {
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
      <main className="flex-1 py-7 sm:py-9">
        <Tabs defaultValue="photos" className="gap-7">
          <TabsList aria-label="Инструменты" className="w-full">
            <TabsTrigger value="photos">
              <ImageIcon aria-hidden="true" />
              Фотообработка
            </TabsTrigger>
            <TabsTrigger value="contacts">
              <ContactRound aria-hidden="true" />
              Контакты
            </TabsTrigger>
          </TabsList>
          <TabsContent
            value="photos"
            keepMounted
            className="text-base data-hidden:hidden"
          >
            <PhotoConverter />
          </TabsContent>
          <TabsContent
            value="contacts"
            keepMounted
            className="text-base data-hidden:hidden"
          >
            <ContactConverter />
          </TabsContent>
        </Tabs>
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
