import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from "@/components/ui/table";
import { LINK_COLUMN, NAME_COLUMN } from "@/lib/photo-converter-types";

const exampleRows = [
  {
    source:
      "https://disk.yandex.ru/client/disk/Example?idDialog=%2Fdisk%2FExample%2Fphoto_001.jpg",
    name: "Участник_001",
  },
  {
    source:
      "https://disk.yandex.ru/client/disk/Example?idDialog=%2Fdisk%2FExample%2Fphoto_002.png",
    name: "Участник_002",
  },
  {
    source: "https://forms.yandex.ru/u/files?path=%2Fexample%2Fphoto_003.heic",
    name: "Участник_003",
  },
];

export default function PhotoInstructions() {
  return (
    <Accordion>
      <AccordionItem value="instructions">
        <AccordionTrigger>
          Как подготовить таблицу и фотографии?
        </AccordionTrigger>
        <AccordionContent className="space-y-5 text-muted-foreground">
          <p>
            Для обработки нужны таблица со ссылками и новыми именами, а также
            папка с исходными фотографиями. Ниже — пример с вымышленными
            данными.
          </p>
          <div className="space-y-3">
            <h3 className="font-medium text-foreground">
              Пример папки с фотографиями
            </h3>
            <p>
              Соберите исходные фотографии в одну папку, например «Мои
              фотографии». Фотографии во вложенных папках тоже учитываются; их
              имена должны быть уникальными.
            </p>
            <pre className="overflow-x-auto rounded-lg border bg-muted/35 p-4 font-mono text-xs leading-6 text-foreground">
              {
                "Мои фотографии/\n├── photo_001.jpg\n├── photo_002.png\n└── photo_003.heic"
              }
            </pre>
          </div>
          <div className="space-y-3">
            <h3 className="font-medium text-foreground">
              Пример таблицы Excel
            </h3>
            <p>
              Сохраните таблицу в формате .xlsx — например, «Список.xlsx».
              Название файла может быть любым. На первом листе в первой строке
              укажите заголовки точно как в примере. Каждая следующая строка
              связывает исходную фотографию с её новым именем.
            </p>
            <div className="overflow-hidden rounded-lg border text-foreground">
              <Table
                className="table-fixed"
                aria-label="Вымышленный пример таблицы Список.xlsx"
              >
                <TableHeader>
                  <TableRow className="bg-muted/35">
                    <TableHead className="w-3/5 px-3 py-3 align-top whitespace-normal">
                      {LINK_COLUMN}
                    </TableHead>
                    <TableHead className="px-3 py-3 align-top whitespace-normal">
                      {NAME_COLUMN}
                    </TableHead>
                  </TableRow>
                </TableHeader>
                <TableBody>
                  {exampleRows.map((row) => (
                    <TableRow key={row.name}>
                      <TableCell className="px-3 py-3 align-top font-mono text-xs wrap-break-word whitespace-normal">
                        {row.source}
                      </TableCell>
                      <TableCell className="px-3 py-3 align-top font-mono text-xs wrap-break-word whitespace-normal">
                        {row.name}
                      </TableCell>
                    </TableRow>
                  ))}
                </TableBody>
              </Table>
            </div>
            <p>
              В первый столбец вставьте ссылки на фотографии из вашего списка.
              По ссылке сайт определяет имя исходного файла и находит его в
              выбранной папке. Например, вторая ссылка соответствует файлу
              «photo_002.png». Все ссылки в примере вымышленные.
            </p>
            <p>
              Во втором столбце указывайте новое имя без расширения. Например,
              «photo_002.png» превратится в «Участник_002.jpg».
            </p>
          </div>
          <div className="space-y-3">
            <h3 className="font-medium text-foreground">
              Как запустить обработку
            </h3>
            <p>
              Нажмите «Выбрать таблицу» и укажите файл .xlsx. Затем нажмите
              «Выбрать папку» и укажите папку с исходными фото. После проверки
              нажмите «Обработать», дождитесь завершения и скачайте готовый ZIP.
              В нём появятся такие файлы:
            </p>
            <pre className="overflow-x-auto rounded-lg border bg-muted/35 p-4 font-mono text-xs leading-6 text-foreground">
              {
                "Фото после обработки/\n├── Участник_001.jpg\n├── Участник_002.jpg\n└── Участник_003.jpg"
              }
            </pre>
          </div>
        </AccordionContent>
      </AccordionItem>
      <AccordionItem value="output">
        <AccordionTrigger>Что будет в готовом архиве?</AccordionTrigger>
        <AccordionContent className="text-muted-foreground">
          <p>
            Исходные фотографии, таблица и три папки: готовые JPEG — в «Фото
            после обработки», фотографии вне списка — в «Непонятные фото», файлы
            с ошибками — в «Фото для ручной обработки». Копии исходных фото
            сохраняются в «Фото до обработки», а таблица — под именем
            «Список.xlsx». Файлы на вашем устройстве не изменяются.
          </p>
          <p>
            Размеры изображений в пикселях сохраняются. Качество JPEG снижается,
            пока файл не станет меньше 500 КиБ. Если этого недостаточно,
            фотография остаётся для ручной обработки. Прозрачные области
            заменяются белым фоном, метаданные не сохраняются.
          </p>
          <p>
            В JSON-отчёте будут все результаты, включая отсутствующие файлы и
            повторяющиеся имена. Обработка полностью выполняется на вашем
            устройстве.
          </p>
        </AccordionContent>
      </AccordionItem>
    </Accordion>
  );
}
