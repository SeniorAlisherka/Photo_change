export const CONTACT_NAME_COLUMN = "ФИО";
export const CONTACT_PROGRAM_COLUMN = "Короткое название программы";
export const CONTACT_YEAR_COLUMN = "Год обучения";
export const CONTACT_PHONE_COLUMN = "Номер телефона";
export const CONTACT_EMAIL_COLUMN = "Почта";
export const CONTACT_COLUMNS = [
  CONTACT_NAME_COLUMN,
  CONTACT_PROGRAM_COLUMN,
  CONTACT_YEAR_COLUMN,
  CONTACT_PHONE_COLUMN,
  CONTACT_EMAIL_COLUMN,
] as const;

export type ContactRecord = {
  row: number;
  givenName: string;
  familyName: string;
  formattedName: string;
  phone?: string;
  email?: string;
};

export type ContactIssue = {
  row: number;
  message: string;
};

export type ContactConversion = {
  contacts: ContactRecord[];
  issues: ContactIssue[];
  vcf: string;
};

export type ContactWorkerRequest = { type: "convert"; file: File };

export type ContactWorkerResponse =
  | { type: "complete"; result: ContactConversion }
  | { type: "error"; message: string };

export class ContactProcessingError extends Error {}
