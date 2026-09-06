import Link from "next/link";

import { Button } from "@/components/ui/button";

export default function NotFoundPage() {
  return (
    <main className="flex flex-1 flex-col items-center justify-center gap-5 px-5 py-16 text-center">
      <p className="text-sm text-muted-foreground">404</p>
      <h1 className="text-3xl font-semibold tracking-tight">
        Страница не найдена
      </h1>
      <p className="text-muted-foreground">
        Проверьте адрес или вернитесь к обработке фотографий.
      </p>
      <Button nativeButton={false} render={<Link href="/" />}>
        На главную
      </Button>
    </main>
  );
}
