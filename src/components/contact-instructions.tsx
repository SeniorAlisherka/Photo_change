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
import {
  CONTACT_NAME_COLUMN,
  CONTACT_PROGRAM_COLUMN,
  CONTACT_YEAR_COLUMN,
  CONTACT_PHONE_COLUMN,
  CONTACT_EMAIL_COLUMN,
} from "@/lib/contact-converter-types";

const examples = [
  [CONTACT_NAME_COLUMN, "Иванов Иван Иванович"],
  [CONTACT_PROGRAM_COLUMN, "ОФД"],
  [CONTACT_YEAR_COLUMN, "2025-2026"],
  [CONTACT_PHONE_COLUMN, "+1 202 555-0101"],
  [CONTACT_EMAIL_COLUMN, "ivanov@example.com"],
];

export default function ContactInstructions() {
  return (
    <Accordion>
      <AccordionItem value="prepare">
        <AccordionTrigger>Как подготовить таблицу контактов?</AccordionTrigger>
        <AccordionContent className="space-y-4 text-muted-foreground">
          <p>
            Сохраните файл в формате .xlsx. На первом листе в первой строке
            укажите пять заголовков из примера. Каждая следующая строка —
            отдельный контакт. Порядок столбцов может быть любым.
          </p>
          <div className="overflow-hidden rounded-lg border text-foreground">
            <Table
              className="table-fixed"
              aria-label="Пример столбцов таблицы контактов"
            >
              <TableHeader>
                <TableRow>
                  <TableHead>Столбец Excel</TableHead>
                  <TableHead>Пример значения</TableHead>
                </TableRow>
              </TableHeader>
              <TableBody>
                {examples.map(([column, value]) => (
                  <TableRow key={column}>
                    <TableCell className="align-top whitespace-normal wrap-break-word">
                      {column}
                    </TableCell>
                    <TableCell className="align-top whitespace-normal wrap-break-word">
                      {value}
                    </TableCell>
                  </TableRow>
                ))}
              </TableBody>
            </Table>
          </div>
          <p>
            Все данные в примере вымышленные. Телефон указывайте с кодом страны;
            сохраняйте этот столбец в Excel как текст, чтобы не потерять «+» и
            начальные нули. В ячейке должен быть один номер или один адрес
            почты.
          </p>
          <p>
            ФИО, программа и год обучения обязательны для каждого контакта. Если
            телефона или почты нет, оставьте соответствующую ячейку пустой.
          </p>
          <div className="space-y-3 rounded-lg border bg-muted/35 p-4 text-foreground">
            <p className="font-medium">Иванов Иван Иванович ОФД 2025-2026</p>
            <dl className="space-y-2 text-sm">
              <div>
                <dt className="text-muted-foreground">Имя</dt>
                <dd>Иванов Иван Иванович</dd>
              </div>
              <div>
                <dt className="text-muted-foreground">Фамилия</dt>
                <dd>ОФД 2025-2026</dd>
              </div>
            </dl>
          </div>
          <p>
            Выберите таблицу, проверьте поля в предпросмотре и скачайте VCF.
            Затем импортируйте файл в адресную книгу на телефоне или компьютере.
            Порядок отображения имени и фамилии зависит от настроек приложения
            «Контакты».
          </p>
        </AccordionContent>
      </AccordionItem>
    </Accordion>
  );
}
